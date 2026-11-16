/* Normalizador de Contratos DMS · parte 2 de 2. Script extraído de index.html sin cambios de lógica (T-014). Cargar en orden. */
        // ============================================================
        //  PROCESAMIENTO PRINCIPAL
        // ============================================================
        PROCESS_BTN.addEventListener('click', async () => {
            if (uploadedFiles.length === 0) return;
            PROCESS_BTN.disabled = true;
            PROCESS_BTN.textContent = '⏳ Procesando...';

            let validCount = 0;
            let errorCount = 0;
            processedResults = [];
            processedFileNames = [];
            ocultarCorreccion();

            log('', 'info');
            log('═══════════════════════════════════════', 'info');
            log(`🔄 Iniciando procesamiento de ${uploadedFiles.length} archivo(s)...`, 'info');

            try {
                for (const file of uploadedFiles) {
                    const buffer = await file.arrayBuffer();
                    const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });

                    const validation = validateFormulario(workbook, file.name);

                    if (validation.hasErrors) {
                        errorCount++;
                        log(`❌ ${file.name}: ${validation.errors.length} error(es) de validación`, 'error');
                        let razonSocialEncontrada = '';
                        let tieneErrorSolicitante = false;
                        for (const err of validation.errors) {
                            log(`   → ${err}`, 'error');
                            if (err.includes('CodigoSolicitanteERP') || err.includes('Código del Cliente')) {
                                tieneErrorSolicitante = true;
                                const match = err.match(/Razón Social: "([^"]+)"/);
                                if (match) {
                                    razonSocialEncontrada = match[1];
                                }
                            }
                        }

                        if (tieneErrorSolicitante && razonSocialEncontrada && razonSocialEncontrada !== 'No especificada') {
                            mostrarCorreccion(razonSocialEncontrada, file);
                            log(`💡 Podés corregir el código en la card "Corregir código del cliente"`, 'info');
                        }
                        continue;
                    }

                    const nombreData = getNombreArchivoData(workbook);
                    if (nombreData.error) {
                        errorCount++;
                        log(`❌ ${file.name}: ${nombreData.error}`, 'error');
                        continue;
                    }

                    validCount++;
                    log(`✅ ${file.name}: Validación correcta`, 'success');

                    const normalizedWB = normalizeWorkbook(workbook);

                    const sheetNames = normalizedWB.SheetNames;
                    const formularioIndex = sheetNames.indexOf('Formulario');
                    if (formularioIndex !== -1) {
                        normalizedWB.SheetNames.splice(formularioIndex, 1);
                        delete normalizedWB.Sheets['Formulario'];
                        log(`   → Hoja "Formulario" eliminada`, 'info');
                    }

                    const refIndex = normalizedWB.SheetNames.indexOf('Referencias');
                    if (refIndex !== -1) {
                        const refName = normalizedWB.SheetNames.splice(refIndex, 1)[0];
                        normalizedWB.SheetNames.push(refName);
                        log(`   → Hoja "Referencias" movida al final`, 'info');
                    }

                    limpiarDireccionesVacias(normalizedWB);

                    let outputName = `Alta de contrato DMS ${nombreData.codigoERP} ${nombreData.codigoSolicitante}`;
                    if (nombreData.clientName) {
                        outputName += ` ${nombreData.clientName}`;
                    }
                    outputName += '.xlsx';

                    processedFileNames.push(outputName);

                    const wbout = XLSX.write(normalizedWB, { bookType: 'xlsx', type: 'array' });
                    processedResults.push({
                        name: outputName,
                        data: wbout
                    });
                }

                VALID_COUNT.textContent = validCount;
                ERROR_COUNT.textContent = errorCount;

                if (processedResults.length > 0) {
                    log(`📦 Descargando ${processedResults.length} archivo(s) normalizado(s)...`, 'success');
                    for (const r of processedResults) {
                        downloadFile(r.data, r.name);
                    }
                    log(`✅ ¡Completado! Se descargaron ${processedResults.length} archivos.`, 'success');
                    generarMensaje(processedResults.length, processedFileNames);
                } else {
                    if (errorCount > 0) {
                        log(`⚠️ No se generaron archivos para descargar. Revisá los errores arriba.`, 'warning');
                    } else {
                        log(`⚠️ No se generaron archivos para descargar.`, 'warning');
                    }
                    MESSAGE_BODY.textContent = '';
                    COPY_BTN.disabled = true;
                }

            } catch (err) {
                log(`❌ ERROR: ${err.message}`, 'error');
                console.error(err);
            }

            PROCESS_BTN.disabled = false;
            PROCESS_BTN.textContent = '⚡ Procesar y Normalizar';
        });

        // ============================================================
        //  FUNCIONES AUXILIARES
        // ============================================================

        /**
         * Limpia caracteres envolventes (comillas, guiones, asteriscos, etc.)
         * del inicio y fin de un texto. Preserva caracteres del medio.
         * Ejemplo: '"CVS CARGO S.A.S."' → 'CVS CARGO S.A.S.'
         */
        function limpiarTextoEnvolvente(texto) {
            if (!texto) return '';

            let resultado = String(texto).trim();

            // Caracteres envolventes a eliminar del inicio y fin
            const envolventes = [
                '"', "'", '-', '_', '*', '–', '—', '`', '«', '»', '“', '”', '‘', '’', '(', ')'
            ];

            let cambio = true;
            while (cambio && resultado.length > 0) {
                cambio = false;

                // Verificar primer carácter
                for (const c of envolventes) {
                    if (resultado.startsWith(c)) {
                        resultado = resultado.substring(c.length).trim();
                        cambio = true;
                        break;
                    }
                }

                if (cambio) continue;

                // Verificar último carácter
                for (const c of envolventes) {
                    if (resultado.endsWith(c)) {
                        resultado = resultado.substring(0, resultado.length - c.length).trim();
                        cambio = true;
                        break;
                    }
                }
            }

            return resultado;
        }

        function limpiarDireccionesVacias(workbook) {
            const sheetName = 'Direcciones';
            if (!workbook.Sheets[sheetName]) return;

            const sheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

            if (data.length === 0) return;

            const header = data[0];
            const nuevasFilas = [header];

            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                const tieneDatos = row.some(cell => {
                    const valor = String(cell).trim();
                    return valor !== '' && valor !== '0';
                });

                if (tieneDatos) {
                    nuevasFilas.push(row);
                }
            }

            if (nuevasFilas.length === data.length) return;

            const newSheet = XLSX.utils.aoa_to_sheet(nuevasFilas);
            workbook.Sheets[sheetName] = newSheet;

            log(`   → Se eliminaron ${data.length - nuevasFilas.length} filas vacías en la hoja "Direcciones"`, 'info');
        }

        function generarMensaje(cantidad, nombres) {
            let mensaje = `Hola @Djirikian Damian Uriel, te paso ${cantidad} archivos para alta de contratos DMS\n\n`;
            mensaje += nombres.join('\n');

            MESSAGE_BODY.textContent = mensaje;
            COPY_BTN.disabled = false;
            COPY_BTN.textContent = '📋 Copiar mensaje';
            COPY_BTN.classList.remove('copiado');
        }

        function normalizarCodigoSolicitante(valor) {
            if (!valor) return null;

            const strValor = String(valor).trim();

            if (!/^\d+$/.test(strValor)) {
                return { error: 'El código debe ser numérico' };
            }

            const longitud = strValor.length;

            if (longitud === 6) {
                return { valor: strValor };
            }

            if (longitud === 10 && strValor.startsWith('0000')) {
                const normalizado = strValor.substring(4);
                return { valor: normalizado };
            }

            if (longitud === 10) {
                return { error: 'El código tiene 10 dígitos pero no comienza con 4 ceros' };
            }

            return { error: `El código tiene ${longitud} dígitos, debe tener 6 dígitos (o 10 dígitos comenzando con 4 ceros)` };
        }

        function getNombreArchivoData(workbook) {
            const sheet = workbook.Sheets['Formulario'];
            if (!sheet) {
                return { error: 'No existe la hoja "Formulario"' };
            }

            const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

            const codigoERP = String(data[2]?.[2] || '').trim();
            const codigoSolicitanteRaw = String(data[4]?.[2] || '').trim();
            const clientNameRaw = String(data[13]?.[2] || '').trim();
            const clientName = limpiarTextoEnvolvente(clientNameRaw);

            // Log si hubo limpieza
            if (clientNameRaw !== clientName) {
                log(`   → Campo "ClientName" limpiado: "${clientNameRaw}" → "${clientName}"`, 'info');
            }

            if (!codigoERP) {
                return { error: 'Campo "CodigoContratoERP" (C3) está vacío. Es obligatorio.' };
            }

            if (!codigoSolicitanteRaw) {
                const razonSocial = clientName || 'No especificada';
                return { error: `Campo "CodigoSolicitanteERP" (C5) está vacío. Es obligatorio. - Razón Social: "${razonSocial}"` };
            }

            const resultadoNormalizacion = normalizarCodigoSolicitante(codigoSolicitanteRaw);
            if (resultadoNormalizacion.error) {
                const razonSocial = clientName || 'No especificada';
                return { error: `Campo "CodigoSolicitanteERP" = "${codigoSolicitanteRaw}" - ${resultadoNormalizacion.error} - Razón Social: "${razonSocial}"` };
            }

            const codigoSolicitante = resultadoNormalizacion.valor;

            return {
                codigoERP: codigoERP,
                codigoSolicitante: codigoSolicitante,
                clientName: clientName
            };
        }

        function normalizarHasDeliveryNote(valor) {
            if (!valor) return 'SIN_REMITO';

            const valorStr = String(valor).toUpperCase().trim();

            if (['CON_REMITO', 'SIN_REMITO', 'DIGITALIZADO'].includes(valorStr)) {
                return valorStr;
            }

            if (['TRUE', '1', 'SI', 'SÍ'].includes(valorStr)) {
                return 'CON_REMITO';
            }

            if (['FALSE', '0', 'NO'].includes(valorStr)) {
                return 'SIN_REMITO';
            }

            return 'SIN_REMITO';
        }

        function validateFormulario(workbook, fileName) {
            const errors = [];
            const sheet = workbook.Sheets['Formulario'];

            if (!sheet) {
                errors.push(`No existe la hoja "Formulario"`);
                return { hasErrors: true, errors };
            }

            const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

            const filasAValidar = [];
            for (let i = 2; i <= 23; i++) filasAValidar.push(i);
            for (let i = 26; i <= 29; i++) filasAValidar.push(i);

            const refSheet = workbook.Sheets['Referencias'];
            const refData = refSheet ? XLSX.utils.sheet_to_json(refSheet, { header: 1, defval: '' }) : [];
            const referencias = extractReferences(refData);

            let huboNormalizacion = false;

            for (const rowIdx of filasAValidar) {
                if (rowIdx >= data.length) {
                    errors.push(`Fila ${rowIdx+1} no existe en el archivo`);
                    continue;
                }

                const row = data[rowIdx];
                const campo = String(row[0] || '').trim();
                let valor = String(row[2] || '').trim();

                if (!campo) continue;

                if (campo === 'HasDeliveryNote' && valor) {
                    const valorNormalizado = normalizarHasDeliveryNote(valor);
                    if (valorNormalizado !== valor) {
                        row[2] = valorNormalizado;
                        huboNormalizacion = true;
                        log(`   → Campo "HasDeliveryNote" normalizado de "${valor}" a "${valorNormalizado}"`, 'info');
                        valor = valorNormalizado;
                    }
                }

                const camposReferencia = ['TipoDeEnvio', 'Logística Solicitada', 'Continente', 'SLA', 'Plazo', 'FormularioMobileB2B'];
                if (camposReferencia.includes(campo)) {
                    const resultado = validarContraReferencias(campo, valor, referencias);

                    if (resultado.normalizado) {
                        row[2] = resultado.valor;
                        huboNormalizacion = true;
                        log(`   → Campo "${campo}" normalizado de "${valor}" a "${resultado.valor}"`, 'info');
                        valor = resultado.valor;
                    }

                    if (!resultado.valido) {
                        let valoresValidos = [];
                        if (campo === 'TipoDeEnvio') {
                            valoresValidos = [
                                'Minorista', 'Mayorista', 'Cadena de frío',
                                'Cargas completas', 'Congelado', 'Logística inversa',
                                'Paquetería', 'Retira depósito', 'Logística Asegurada'
                            ];
                        } else if (campo === 'SLA') {
                            valoresValidos = [
                                'SLA 1',
                                'SLA 2',
                                'SLA 3',
                                'SLA 4 - Logistica Asegurada V2',
                                'SLA PALLETS',
                                'SLA BIGGER'
                            ];
                        } else if (campo === 'Plazo') {
                            valoresValidos = [
                                'SLA 1 - Estándar B2B',
                                'SLA 2 - Acordado con cliente',
                                'SLA 3 - Urgente'
                            ];
                        } else {
                            const map = {
                                'Logística Solicitada': 'logistica',
                                'Continente': 'continente',
                                'FormularioMobileB2B': 'formulario'
                            };
                            const tableName = map[campo];
                            if (tableName && referencias[tableName]) {
                                valoresValidos = referencias[tableName].map(item => item.key);
                            }
                        }

                        const listaValores = valoresValidos.join(', ');
                        errors.push(`Campo "${campo}" = "${valor}" no es válido. Valores válidos: ${listaValores}`);
                    }
                }

                if (campo === 'HasDeliveryNote' && valor) {
                    const resultado = validarContraReferencias(campo, valor, referencias);

                    if (resultado.normalizado) {
                        row[2] = resultado.valor;
                        huboNormalizacion = true;
                        log(`   → Campo "HasDeliveryNote" normalizado de "${valor}" a "${resultado.valor}"`, 'info');
                        valor = resultado.valor;
                    }

                    if (!resultado.valido) {
                        errors.push(`Campo "${campo}" = "${valor}" no es válido. Valores válidos: CON_REMITO, SIN_REMITO, DIGITALIZADO`);
                    }
                }

                const razonSocial = limpiarTextoEnvolvente(String(data[13]?.[2] || '').trim()) || 'No especificada';

                const opcionales = ['CondicionDeEntrega', 'DescripcionContratoERP'];
                if (!opcionales.includes(campo) && !valor) {
                    if (campo === 'CodigoSolicitanteERP') {
                        errors.push(`Campo "${campo}" está vacío (fila ${rowIdx+1}) - Razón Social: "${razonSocial}"`);
                    } else {
                        errors.push(`Campo "${campo}" está vacío (fila ${rowIdx+1})`);
                    }
                }

                if (campo === 'CodigoContratoERP' && valor && !/^\d{9}$/.test(valor)) {
                    errors.push(`Campo "CodigoContratoERP" = "${valor}" debe tener 9 dígitos`);
                }

                if (campo === 'IDSociedadSAP' && valor && valor !== '1100') {
                    errors.push(`Campo "IDSociedadSAP" = "${valor}" debe ser "1100" para B2B`);
                }

                if (campo === 'Habilitado' && valor && valor !== '1' && valor !== '0') {
                    errors.push(`Campo "Habilitado" = "${valor}" debe ser 1 o 0`);
                }

                if (campo === 'CodigoSolicitanteERP' && valor) {
                    const razonSocial2 = limpiarTextoEnvolvente(String(data[13]?.[2] || '').trim()) || 'No especificada';
                    const resultado = normalizarCodigoSolicitante(valor);
                    if (resultado.error) {
                        errors.push(`Campo "CodigoSolicitanteERP" = "${valor}" - ${resultado.error} - Razón Social: "${razonSocial2}"`);
                    }
                }

                // ============================================
                // VALIDACIÓN CRUZADA: TipoDeEnvio ↔ SLA
                // ============================================
                const SLA_ASEGURADA = 'SLA 4 - Logistica Asegurada V2';
                const TIPO_ASEGURADA = 'Logística Asegurada';

                if (campo === 'TipoDeEnvio' && valor === TIPO_ASEGURADA) {
                    for (let j = 0; j < data.length; j++) {
                        const r = data[j];
                        if (r && String(r[0] || '').trim() === 'SLA') {
                            const slaActual = String(r[2] || '').trim();
                            if (slaActual !== SLA_ASEGURADA) {
                                r[2] = SLA_ASEGURADA;
                                huboNormalizacion = true;
                                log(`   → ⚠️ Tipo de Envío = "Logística Asegurada" → SLA auto-corregido de "${slaActual}" a "${SLA_ASEGURADA}"`, 'warning');
                            }
                            break;
                        }
                    }
                }

                if (campo === 'SLA' && valor === SLA_ASEGURADA) {
                    for (let j = 0; j < data.length; j++) {
                        const r = data[j];
                        if (r && String(r[0] || '').trim() === 'TipoDeEnvio') {
                            const tipoActual = String(r[2] || '').trim();
                            if (tipoActual !== TIPO_ASEGURADA) {
                                r[2] = TIPO_ASEGURADA;
                                huboNormalizacion = true;
                                log(`   → ⚠️ SLA = "SLA 4 - Logistica Asegurada V2" → Tipo de Envío auto-corregido de "${tipoActual}" a "${TIPO_ASEGURADA}"`, 'warning');
                            }
                            break;
                        }
                    }
                }

                // ============================================
                // VALIDACIÓN CRUZADA: Continente ↔ SLA
                // ============================================
                const CONTINENTE_PALLETS = 'Pallets';
                const CONTINENTE_BIGGER = 'Bultos Bigger';
                const SLA_PALLETS = 'SLA PALLETS';
                const SLA_BIGGER = 'SLA BIGGER';

                const continenteASla = {
                    [CONTINENTE_PALLETS]: SLA_PALLETS,
                    [CONTINENTE_BIGGER]: SLA_BIGGER
                };

                const slaAContinente = {
                    [SLA_PALLETS]: CONTINENTE_PALLETS,
                    [SLA_BIGGER]: CONTINENTE_BIGGER
                };

                if (campo === 'Continente' && continenteASla[valor]) {
                    const slaRequerido = continenteASla[valor];
                    for (let j = 0; j < data.length; j++) {
                        const r = data[j];
                        if (r && String(r[0] || '').trim() === 'SLA') {
                            const slaActual = String(r[2] || '').trim();
                            if (slaActual !== slaRequerido) {
                                r[2] = slaRequerido;
                                huboNormalizacion = true;
                                log(`   → ⚠️ Continente = "${valor}" → SLA auto-corregido de "${slaActual}" a "${slaRequerido}"`, 'warning');
                            }
                            break;
                        }
                    }
                }

                if ((campo === 'SLA' || campo === 'Plazo') && slaAContinente[valor]) {
                    const continenteRequerido = slaAContinente[valor];
                    for (let j = 0; j < data.length; j++) {
                        const r = data[j];
                        if (r && String(r[0] || '').trim() === 'Continente') {
                            const continenteActual = String(r[2] || '').trim();
                            if (continenteActual !== continenteRequerido) {
                                r[2] = continenteRequerido;
                                huboNormalizacion = true;
                                log(`   → ⚠️ SLA = "${valor}" → Continente auto-corregido de "${continenteActual}" a "${continenteRequerido}"`, 'warning');
                            }
                            break;
                        }
                    }
                }

                const booleanFields = ['AllowRescue', 'AllowWithdrawPartialValues', 'ObligatoryGauge',
                    'AllowPartialDelivery', 'IsCentralizedReturn', 'HasCollectionManagement',
                    'AllowPrintDocumentationReturn', '¿Es recambio?'
                ];
                if (booleanFields.includes(campo) && valor) {
                    const valoresValidos = ['true', 'false', 'TRUE', 'FALSE', 'True', 'False', '1', '0', 'si', 'no'];
                    if (!valoresValidos.some(v => v.toLowerCase() === valor.toLowerCase())) {
                        errors.push(`Campo "${campo}" = "${valor}" debe ser true/false o si/no`);
                    }
                }
            }

            if (huboNormalizacion) {
                const newSheet = XLSX.utils.aoa_to_sheet(data);
                workbook.Sheets['Formulario'] = newSheet;
            }

            return { hasErrors: errors.length > 0, errors };
        }

        function normalizarValorConDiccionario(campo, valor, referencias) {
            if (!valor) return { valor: valor, normalizado: false };

            const valorLimpio = String(valor).trim();

            const diccionarios = {
                'TipoDeEnvio': {
                    'minorista': 'Minorista',
                    'mayorista': 'Mayorista',
                    'cadena de frio': 'Cadena de frío',
                    'cadena de frío': 'Cadena de frío',
                    'carga completa': 'Cargas completas',
                    'cargas completa': 'Cargas completas',
                    'cargas completas': 'Cargas completas',
                    'logistica inversa': 'Logística inversa',
                    'logística inversa': 'Logística inversa',
                    'retira deposito': 'Retira depósito',
                    'retira depósito': 'Retira depósito',
                    'logistica asegurada': 'Logística Asegurada',
                    'logística asegurada': 'Logística Asegurada',
                    'paqueteria': 'Paquetería',
                    'paquetería': 'Paquetería',
                    'congelado': 'Congelado'
                },
                'SLA': {
                    'sla 1': 'SLA 1',
                    'sla1': 'SLA 1',
                    'sla 2': 'SLA 2',
                    'sla2': 'SLA 2',
                    'sla 3': 'SLA 3',
                    'sla3': 'SLA 3',
                    'sla 4 - logistica asegurada v2': 'SLA 4 - Logistica Asegurada V2',
                    'sla 4 - logística asegurada v2': 'SLA 4 - Logistica Asegurada V2',
                    'sla 4 - logistica asegurada': 'SLA 4 - Logistica Asegurada V2',
                    'sla 4 - logística asegurada': 'SLA 4 - Logistica Asegurada V2',
                    'sla4 - logistica asegurada v2': 'SLA 4 - Logistica Asegurada V2',
                    'sla 4 logistica asegurada v2': 'SLA 4 - Logistica Asegurada V2',
                    'logistica asegurada v2': 'SLA 4 - Logistica Asegurada V2',
                    'sla pallet': 'SLA PALLETS',
                    'sla pallets': 'SLA PALLETS',
                    'slapallet': 'SLA PALLETS',
                    'slapallets': 'SLA PALLETS',
                    'sla bigger': 'SLA BIGGER',
                    'sla biger': 'SLA BIGGER',
                    'slabigger': 'SLA BIGGER'
                },
                'Continente': {
                    'pallet': 'Pallets',
                    'pallets': 'Pallets',
                    'bigger': 'Bultos Bigger',
                    'bultos bigger': 'Bultos Bigger',
                    'bulto bigger': 'Bultos Bigger',
                    'bultos biger': 'Bultos Bigger',
                    'bulto biger': 'Bultos Bigger'
                },
                'HasDeliveryNote': {
                    'con remito': 'CON_REMITO',
                    'conremito': 'CON_REMITO',
                    'sin remito': 'SIN_REMITO',
                    'sinremito': 'SIN_REMITO',
                    'digitalizado': 'DIGITALIZADO'
                }
            };

            const diccionario = diccionarios[campo];
            if (diccionario) {
                const valorLower = valorLimpio.toLowerCase();
                for (const [key, value] of Object.entries(diccionario)) {
                    if (valorLower === key.toLowerCase()) {
                        const cambio = (valorLimpio !== value);
                        return { valor: value, normalizado: cambio };
                    }
                }
            }

            const map = {
                'Logística Solicitada': 'logistica',
                'Continente': 'continente',
                'SLA': 'sla',
                'Plazo': 'sla',
                'FormularioMobileB2B': 'formulario'
            };

            const tableName = map[campo];
            if (tableName && referencias[tableName]) {
                const table = referencias[tableName];
                const valorLower = valorLimpio.toLowerCase();

                for (const item of table) {
                    const itemKey = String(item.key || '').toLowerCase();
                    const itemCode = String(item.code || '').toLowerCase();
                    const itemDesc = String(item.desc || '').toLowerCase();

                    if (valorLower === itemKey ||
                        valorLower === itemCode ||
                        (itemDesc && valorLower === itemDesc)) {
                        const cambio = (valorLimpio !== item.key);
                        return { valor: item.key, normalizado: cambio };
                    }
                }
            }

            return { valor: valorLimpio, normalizado: false };
        }

        function validarContraReferencias(campo, valor, referencias) {
            const resultado = normalizarValorConDiccionario(campo, valor, referencias);
            const valorNormalizado = resultado.valor;

            if (campo === 'TipoDeEnvio') {
                const tiposValidos = [
                    'Minorista', 'Mayorista', 'Cadena de frío',
                    'Cargas completas', 'Congelado', 'Logística inversa',
                    'Paquetería', 'Retira depósito', 'Logística Asegurada'
                ];
                const esValido = tiposValidos.some(t => t === valorNormalizado);
                return {
                    valido: esValido,
                    valor: esValido ? valorNormalizado : valor,
                    normalizado: resultado.normalizado
                };
            }

            const slasValidos = [
                'SLA 1',
                'SLA 2',
                'SLA 3',
                'SLA 4 - Logistica Asegurada V2',
                'SLA PALLETS',
                'SLA BIGGER'
            ];

            if (campo === 'SLA' || campo === 'Plazo') {
                const esValido = slasValidos.includes(valorNormalizado);
                return {
                    valido: esValido,
                    valor: esValido ? valorNormalizado : valor,
                    normalizado: resultado.normalizado
                };
            }

            const map = {
                'Logística Solicitada': 'logistica',
                'Continente': 'continente',
                'FormularioMobileB2B': 'formulario'
            };

            const tableName = map[campo];
            if (!tableName) return { valido: true, valor: valor, normalizado: false };

            const table = referencias[tableName];
            if (!table || table.length === 0) return { valido: true, valor: valor, normalizado: false };

            const esValido = table.some(item => {
                return item.key === valorNormalizado ||
                    item.code === valorNormalizado ||
                    item.desc === valorNormalizado;
            });

            return {
                valido: esValido,
                valor: esValido ? valorNormalizado : valor,
                normalizado: resultado.normalizado
            };
        }

        function extractReferences(refData) {
            const refs = {};

            let currentTable = null;
            for (const row of refData) {
                const colA = String(row[0] || '').trim();
                const colB = String(row[1] || '').trim();
                const colC = String(row[2] || '').trim();
                const colE = String(row[4] || '').trim();
                const colF = String(row[5] || '').trim();
                const colG = String(row[6] || '').trim();
                const colI = String(row[8] || '').trim();
                const colJ = String(row[9] || '').trim();
                const colK = String(row[10] || '').trim();
                const colM = String(row[12] || '').trim();
                const colN = String(row[13] || '').trim();
                const colO = String(row[14] || '').trim();
                const colP = String(row[15] || '').trim();

                if (colA === 'Logistica Solicitada') {
                    currentTable = 'logistica';
                    refs.logistica = [];
                    continue;
                }
                if (colE === 'Continente') {
                    currentTable = 'continente';
                    refs.continente = [];
                    continue;
                }
                if (colI === 'SLAs') {
                    currentTable = 'sla';
                    refs.sla = [];
                    continue;
                }
                if (colM === 'Formulario Mobile B2B') {
                    currentTable = 'formulario';
                    refs.formulario = [];
                    continue;
                }

                if (currentTable === 'logistica' && colA && colB) {
                    refs.logistica.push({ key: colA, code: colB, desc: colC });
                }
                if (currentTable === 'continente' && colE && colF) {
                    refs.continente.push({ key: colE, code: colF, desc: colG });
                }
                if (currentTable === 'sla' && colI && colJ) {
                    refs.sla.push({ key: colI, code: colJ, desc: colK });
                }
                if (currentTable === 'formulario' && colM && colN) {
                    refs.formulario.push({ key: colM, code: colN, desc: colO, aplica: colP });
                }
            }

            return refs;
        }

        function normalizeWorkbook(workbook) {
            const newWB = XLSX.utils.book_new();

            let codigoSolicitante = '';
            let clientNameLimpio = '';
            const formularioSheet = workbook.Sheets['Formulario'];
            if (formularioSheet) {
                const formularioData = XLSX.utils.sheet_to_json(formularioSheet, { header: 1, defval: '' });
                for (let i = 0; i < formularioData.length; i++) {
                    const row = formularioData[i];
                    if (row && row[0]) {
                        const campo = String(row[0]).trim();
                        if (campo === 'Código del Cliente' || campo === 'CodigoSolicitanteERP') {
                            codigoSolicitante = String(row[2] || '').trim();
                            log(`   → 🔍 normalizeWorkbook: Valor leído de Formulario: "${codigoSolicitante}"`, 'info');
                            // T-014 (2026-11-16): 0000xxxxxx → xxxxxx también en la hoja Contrato (así lo recibe Damián)
                            const codigoNorm = normalizarCodigoSolicitante(codigoSolicitante);
                            if (codigoNorm && codigoNorm.valor && codigoNorm.valor !== codigoSolicitante) {
                                log(`   → CodigoSolicitanteERP normalizado de "${codigoSolicitante}" a "${codigoNorm.valor}" en hoja Contrato`, 'info');
                                codigoSolicitante = codigoNorm.valor;
                            }
                        }
                        if (campo === 'Razón Social' || campo === 'ClientName') {
                            const valorRaw = String(row[2] || '').trim();
                            clientNameLimpio = limpiarTextoEnvolvente(valorRaw);
                            if (valorRaw !== clientNameLimpio) {
                                log(`   → Campo "ClientName" limpiado en hoja Contrato: "${valorRaw}" → "${clientNameLimpio}"`, 'info');
                            }
                        }
                    }
                }
            }

            for (const sheetName of workbook.SheetNames) {
                const originalSheet = workbook.Sheets[sheetName];
                const data = XLSX.utils.sheet_to_json(originalSheet, { header: 1, defval: '', raw: false });

                if (sheetName === 'Contrato') {
                    if (data.length > 1) {
                        const rowIndex = 1;
                        if (data[rowIndex] && data[rowIndex].length > 3) {
                            if (codigoSolicitante) {
                                data[rowIndex][3] = codigoSolicitante;
                                log(`   → ✅ Actualizado CodigoSolicitanteERP en Contrato a "${codigoSolicitante}"`, 'info');
                            } else {
                                log(`   → ⚠️ No se encontró CodigoSolicitanteERP en Formulario. Valor actual: "${data[rowIndex][3] || ''}"`, 'warning');
                            }
                        }
                        // ClientName en Contrato: columna M = índice 12
                        if (data[rowIndex] && data[rowIndex].length > 12) {
                            if (clientNameLimpio) {
                                data[rowIndex][12] = clientNameLimpio;
                                log(`   → ✅ Actualizado ClientName en Contrato a "${clientNameLimpio}"`, 'info');
                            }
                        }
                        if (data[rowIndex] && data[rowIndex].length > 24) {
                            data[rowIndex][24] = 'false';
                            log(`   → AllowsStorage forzado a "false" en hoja "Contrato"`, 'info');
                        }
                    }
                }

                const newSheet = XLSX.utils.aoa_to_sheet(data);
                XLSX.utils.book_append_sheet(newWB, newSheet, sheetName);
            }

            return newWB;
        }

        function downloadFile(data, filename) {
            const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            setTimeout(() => URL.revokeObjectURL(url), 5000);
        }

        // ============================================================
        //  INICIO
        // ============================================================
        log('📌 Sistema listo. Arrastrá archivos .xlsx o hacé clic para subir.', 'info');
        log('📋 Los archivos serán validados, normalizados y descargados automáticamente.', 'info');
        log('💡 Si el código del cliente es incorrecto, aparecerá una card para corregirlo.', 'info');
