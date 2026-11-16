/* Normalizador de Contratos DMS · parte 1 de 2. Script extraído de index.html sin cambios de lógica (T-014). Cargar en orden. */
        // ============================================================
        //  CONFIGURACIÓN
        // ============================================================
        const LOG = document.getElementById('logArea');
        const FILE_LIST = document.getElementById('fileList');
        const STATS_BAR = document.getElementById('statsBar');
        const FILE_COUNT = document.getElementById('fileCount');
        const VALID_COUNT = document.getElementById('validCount');
        const ERROR_COUNT = document.getElementById('errorCount');
        const PROCESS_BTN = document.getElementById('processBtn');
        const CLEAR_BTN = document.getElementById('clearBtn');
        const MESSAGE_BODY = document.getElementById('messageBody');
        const COPY_BTN = document.getElementById('copyBtn');
        const MAIN_LAYOUT = document.getElementById('mainLayout');
        const CORRECCION_WRAPPER = document.getElementById('correccionCardWrapper');
        const RAZON_SOCIAL_DISPLAY = document.getElementById('razonSocialDisplay');
        const CODIGO_CORRECCION = document.getElementById('codigoCorreccion');
        const BTN_REPROCESAR = document.getElementById('btnReprocesarCorreccion');
        const CORRECCION_ERROR = document.getElementById('correccionError');

        let uploadedFiles = [];
        let processedResults = [];
        let processedFileNames = [];
        let razonSocialCopiada = '';
        let pendingFileForReprocess = null;

        // ============================================================
        //  FUNCIONES DE LAYOUT
        // ============================================================
        function mostrarCorreccion(razonSocial, file) {
            RAZON_SOCIAL_DISPLAY.textContent = razonSocial;
            pendingFileForReprocess = {
                file: file,
                razonSocial: razonSocial
            };
            CORRECCION_WRAPPER.style.display = 'block';
            MAIN_LAYOUT.className = 'four-column-layout';
            // Enfocar el input
            setTimeout(() => CODIGO_CORRECCION.focus(), 100);
            // Limpiar error previo
            CORRECCION_ERROR.style.display = 'none';
            CORRECCION_ERROR.textContent = '';
            // Habilitar botón
            BTN_REPROCESAR.disabled = false;
            BTN_REPROCESAR.textContent = '🔁 Reprocesar';
        }

        function ocultarCorreccion() {
            CORRECCION_WRAPPER.style.display = 'none';
            MAIN_LAYOUT.className = 'three-column-layout';
            pendingFileForReprocess = null;
            CODIGO_CORRECCION.value = '';
            CORRECCION_ERROR.style.display = 'none';
            CORRECCION_ERROR.textContent = '';
        }

        // ============================================================
        //  MODAL DE DESCARGA DE PLANTILLA
        // ============================================================
        const modalPlantilla = document.getElementById('modalPlantilla');
        const btnPlantilla = document.getElementById('btnPlantilla');
        const btnDescargarModal = document.getElementById('btnDescargarModal');
        const btnCerrarModal = document.getElementById('btnCerrarModal');
        const modalClose = document.getElementById('modalClose');

        function abrirModal() {
            modalPlantilla.classList.add('active');
            document.body.style.overflow = 'hidden';
        }

        function cerrarModal() {
            modalPlantilla.classList.remove('active');
            document.body.style.overflow = '';
        }

        btnPlantilla.addEventListener('click', abrirModal);
        btnDescargarModal.addEventListener('click', function() {
            generarPlantilla();
            cerrarModal();
        });
        btnCerrarModal.addEventListener('click', cerrarModal);
        modalClose.addEventListener('click', cerrarModal);
        modalPlantilla.addEventListener('click', function(e) {
            if (e.target === this) cerrarModal();
        });

        // ============================================================
        //  FUNCIÓN: GENERAR PLANTILLA CON EXCELJS
        // ============================================================
        async function generarPlantilla() {
            const workbook = new ExcelJS.Workbook();
            workbook.creator = 'Andreani DMS Normalizer';
            workbook.created = new Date();

            // ============================================
            // HOJA: Formulario
            // ============================================
            const ws = workbook.addWorksheet('Formulario', {
                properties: { tabColor: { argb: 'FFD71920' } },
                pageSetup: { orientation: 'portrait', fitToPage: true, margins: {
                    left: 0.7, right: 0.7, top: 0.7, bottom: 0.7, header: 0.3, footer: 0.3
                }}
            });

            // Columnas
            ws.columns = [
                { header: 'Campo', key: 'campo', width: 30 },
                { header: 'Descripción / Ayuda', key: 'descripcion', width: 45 },
                { header: 'Valor', key: 'valor', width: 35 },
                { header: 'Comentario', key: 'comentario', width: 25 }
            ];

            // Estilos predefinidos (alineados a paleta oficial Andreani DS)
            const styles = {
                header: {
                    font: { name: 'Segoe UI', bold: true, size: 11, color: { argb: 'FFFFFFFF' } },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD71920' } },
                    alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
                    border: {
                        top: { style: 'thin', color: { argb: 'FFB0B0B0' } },
                        bottom: { style: 'thin', color: { argb: 'FFB0B0B0' } },
                        left: { style: 'thin', color: { argb: 'FFB0B0B0' } },
                        right: { style: 'thin', color: { argb: 'FFB0B0B0' } }
                    }
                },
                section: {
                    font: { name: 'Segoe UI', bold: true, size: 12, color: { argb: 'FFD71920' } },
                    alignment: { horizontal: 'left', vertical: 'middle' },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFBE8E9' } }
                },
                editable: {
                    font: { name: 'Segoe UI', size: 11, color: { argb: 'FF272A2F' } },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } },
                    alignment: { horizontal: 'left', vertical: 'middle', wrapText: true },
                    border: {
                        top: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        left: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        right: { style: 'thin', color: { argb: 'FFD0D0D0' } }
                    }
                },
                locked: {
                    font: { name: 'Segoe UI', size: 11, color: { argb: 'FF818891' } },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F0F0' } },
                    alignment: { horizontal: 'center', vertical: 'middle' },
                    border: {
                        top: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        left: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        right: { style: 'thin', color: { argb: 'FFD0D0D0' } }
                    }
                },
                help: {
                    font: { name: 'Segoe UI', size: 10, color: { argb: 'FF818891' } },
                    alignment: { horizontal: 'left', vertical: 'middle', wrapText: true },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFAFAFA' } }
                },
                comment: {
                    font: { name: 'Segoe UI', size: 10, color: { argb: 'FF4E63D7' }, italic: true },
                    alignment: { horizontal: 'left', vertical: 'middle', wrapText: true },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDEFFB' } }
                }
            };

            // Datos de la plantilla
            const rows = [
                ['Campo', 'Descripción / Ayuda', 'Valor', 'Comentario'],
                ['', '🔑 DATOS DEL CONTRATO', '', ''],
                ['Código de Contrato', '9 dígitos, ej: 351005959', '', ''],
                ['Descripción del Contrato', 'Texto libre', '', ''],
                ['Código del Cliente', '6 dígitos, ej: 125383', '', ''],
                ['Sociedad SAP', 'Siempre 1100', '1100', 'No modificar'],
                ['', '📦 DATOS DEL ENVÍO', '', ''],
                ['Tipo de Envío', 'Minorista, Mayorista, Cadena de frío, Cargas completas, Congelado, Logística inversa, Paquetería, Retira depósito, Logística Asegurada', '', 'Seleccionar de la lista'],
                ['Condición de Entrega', 'Ej: Mayor de 18 + DNI', '', ''],
                ['Habilitado', '1 = Activo, 0 = Inactivo', '1', 'Dejar como 1'],
                ['Modalidad de Entrega', 'Domicilio, Sucursal, Punto de retiro', '', 'Seleccionar de la lista'],
                ['Fecha Habilitación', 'Fecha actual', '=TODAY()', 'No modificar'],
                ['Válido Desde', 'Fecha inicio contrato', '2025-01-01', 'No modificar'],
                ['Válido Hasta', 'Fecha fin contrato (si no tiene, dejar 9999-12-08)', '9999-12-08', 'No modificar'],
                ['', '🏢 DATOS DEL CLIENTE', '', ''],
                ['Razón Social', 'Nombre completo del cliente', '', ''],
                ['', '⚙️ CONFIGURACIÓN', '', ''],
                ['Nivel de Servicio (SLA)', 'SLA 1, SLA 2, SLA 3, SLA 4 - Logistica Asegurada V2, SLA PALLETS, SLA BIGGER', '', 'Seleccionar de la lista'],
                ['¿Permite Rescate?', 'true / false', '', ''],
                ['¿Permite retiro parcial de valores?', 'true / false', '', ''],
                ['¿Aforo Obligatorio?', 'true / false', '', ''],
                ['¿Permite Entrega Parcial?', 'true / false', '', ''],
                ['¿Devolución Centralizada?', 'true / false', '', ''],
                ['¿Tiene gestión de cobranza?', 'true / false', '', ''],
                ['¿Usa remito?', 'CON_REMITO, SIN_REMITO, DIGITALIZADO', '', 'Seleccionar de la lista'],
                ['¿Permite imprimir comprobante de entrega?', 'true / false', '', ''],
                ['¿Es recambio?', 'true / false', '', ''],
                ['', '', '', ''],
                ['', '📋 NIVELES DE ESFUERZO', '', ''],
                ['Logística Solicitada', 'Entrega, Logística inversa, Punto a punto, Retiro, Cambio', '', 'Seleccionar de la lista'],
                ['Continente', 'Contenedor, Bulto, Pallets, Bultos Bigger, Sobre, Bodega, etc.', '', 'Seleccionar de la lista'],
                ['Plazo', 'SLA 1 - Estándar B2B, SLA 2 - Acordado con cliente, SLA 3 - Urgente, SLA 4 - Logistica Asegurada V2, SLA PALLETS, SLA BIGGER', '', 'Seleccionar de la lista'],
                ['Formulario Mobile B2B', 'Formulario B2B 1 al 13', '', 'Seleccionar de la lista'],
                ['', '', '', ''],
                ['', '📍 DIRECCIONES (obligatorias: Devolución y Devolución de Remitos)', '', ''],
                ['Clase de Dirección', 'Calle', 'Número', 'Piso'],
                ['DEVOLUCIÓN *', '', '', ''],
                ['RESCATE', '', '', ''],
                ['REENVIO', '', '', ''],
                ['DESTRUCCION', '', '', ''],
                ['SINIESTRO', '', '', ''],
                ['DEVOLUCION_CENTRALIZADA', '', '', ''],
                ['DEVOLUCION_REMITOS *', '', '', ''],
                ['DEVOLUCION_COBROS', '', '', '']
            ];

            // Agregar filas con estilos
            rows.forEach((rowData, rowIndex) => {
                const row = ws.addRow(rowData);
                row.height = 28;

                row.eachCell((cell, colIndex) => {
                    const value = cell.value ? String(cell.value).trim() : '';

                    let style = { ...styles.editable };

                    const isHeader = rowIndex === 0;
                    const isSection = value.startsWith('🔑') || value.startsWith('📦') ||
                                     value.startsWith('🏢') || value.startsWith('⚙️') ||
                                     value.startsWith('📋') || value.startsWith('📍') ||
                                     value === '📋 NIVELES DE ESFUERZO' ||
                                     value === '📍 DIRECCIONES (obligatorias: Devolución y Devolución de Remitos)';
                    const isComment = colIndex === 3 && value;
                    const isLocked = value === '1100' || value === '1' ||
                                    value === '=TODAY()' || value === '2025-01-01' ||
                                    value === '9999-12-08' || value === 'Dejar como 1' ||
                                    value === 'No modificar' || value === 'ALERTRAN ESTANDAR MOTIVES V4' ||
                                    value === 'false' || rowData[0] === 'Sociedad SAP' ||
                                    rowData[0] === 'Habilitado' || rowData[0] === 'Fecha Habilitación' ||
                                    rowData[0] === 'Válido Desde' || rowData[0] === 'Válido Hasta';
                    const isHelp = colIndex === 1 && value && !isHeader;
                    const isEditable = colIndex === 2 && !isLocked && value !== '';

                    if (isHeader) {
                        style = { ...styles.header };
                    } else if (isSection) {
                        style = { ...styles.section };
                        ws.mergeCells(`A${rowIndex+1}:D${rowIndex+1}`);
                    } else if (isComment) {
                        style = { ...styles.comment };
                    } else if (isLocked && colIndex === 2) {
                        style = { ...styles.locked };
                    } else if (isHelp) {
                        style = { ...styles.help };
                    } else if (isEditable) {
                        style = { ...styles.editable };
                    }

                    Object.assign(cell, { style });
                });
            });

            // Bloquear celdas
            ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                    const value = cell.value ? String(cell.value).trim() : '';
                    const isLocked = value === '1100' || value === '1' ||
                                    value === '=TODAY()' || value === '2025-01-01' ||
                                    value === '9999-12-08' || value === 'Dejar como 1' ||
                                    value === 'No modificar' || value === 'ALERTRAN ESTANDAR MOTIVES V4' ||
                                    value === 'false';

                    if (isLocked && colNumber === 3) {
                        cell.protection = { locked: true };
                    } else if (colNumber === 3 && rowNumber > 1) {
                        cell.protection = { locked: false };
                    }
                });
            });

            // ============================================
            // HOJA: Contrato
            // ============================================
            const wsContrato = workbook.addWorksheet('Contrato', {
                properties: { tabColor: { argb: 'FFD71920' } }
            });

            const contratoHeaders = [
                'NombreContratoDMS', 'CodigoContratoERP', 'DescripcionContratoERP',
                'CodigoSolicitanteERP', 'IDSociedadSAP', 'TipoDeEnvio', 'CondicionDeEntrega',
                'Habilitado', 'TipoEntrega', 'FechaHabilitado', 'ValidoDesde', 'ValidoHasta',
                'ClientName', 'SLA', 'AllowRescue', 'AllowWithdrawPartialValues',
                'ObligatoryGauge', 'AllowPartialDelivery', 'IsCentralizedReturn',
                'HasCollectionManagement', 'DeliveryNote', 'AllowPrintDocumentationReturn',
                'MotiveStateMap', 'AutomaticAdmission', 'AllowsStorage', 'HasChange'
            ];

            const contratoValues = [
                '', '', '', '', '1100', '', '', '1', '', '', '', '',
                '', '', '', '', '', '', '', '', '', '',
                'ALERTRAN ESTANDAR MOTIVES V4', 'false', 'false', ''
            ];

            const headerRow = wsContrato.addRow(contratoHeaders);
            headerRow.height = 28;
            headerRow.eachCell((cell) => {
                cell.style = {
                    font: { name: 'Segoe UI', bold: true, size: 10, color: { argb: 'FFFFFFFF' } },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD71920' } },
                    alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
                    border: {
                        top: { style: 'thin', color: { argb: 'FFB0B0B0' } },
                        bottom: { style: 'thin', color: { argb: 'FFB0B0B0' } },
                        left: { style: 'thin', color: { argb: 'FFB0B0B0' } },
                        right: { style: 'thin', color: { argb: 'FFB0B0B0' } }
                    }
                };
            });

            const valueRow = wsContrato.addRow(contratoValues);
            valueRow.height = 24;
            valueRow.eachCell((cell, colIndex) => {
                const isLocked = colIndex === 5 || colIndex === 8 || colIndex === 23 || colIndex === 24;
                cell.style = isLocked ? {
                    font: { name: 'Segoe UI', size: 10, color: { argb: 'FF818891' } },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F0F0' } },
                    alignment: { horizontal: 'center', vertical: 'middle' },
                    border: {
                        top: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        left: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        right: { style: 'thin', color: { argb: 'FFD0D0D0' } }
                    }
                } : {
                    font: { name: 'Segoe UI', size: 10, color: { argb: 'FF272A2F' } },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } },
                    alignment: { horizontal: 'left', vertical: 'middle' },
                    border: {
                        top: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        left: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                        right: { style: 'thin', color: { argb: 'FFD0D0D0' } }
                    }
                };
                if (isLocked) {
                    cell.protection = { locked: true };
                }
            });

            wsContrato.columns.forEach((col, index) => {
                col.width = index === 0 ? 20 : index === 1 ? 18 : index === 2 ? 22 : index === 3 ? 18 : 15;
            });

            // ============================================
            // HOJA: Direcciones
            // ============================================
            const wsDirecciones = workbook.addWorksheet('Direcciones', {
                properties: { tabColor: { argb: 'FFD71920' } }
            });

            const direccionesHeaders = ['Street', 'Number', 'Floor', 'Department', 'City', 'PostalCode', 'State', 'Country', 'Phone', 'MailBox', 'AddressType'];
            const dirHeaderRow = wsDirecciones.addRow(direccionesHeaders);
            dirHeaderRow.height = 28;
            dirHeaderRow.eachCell((cell) => {
                cell.style = {
                    font: { name: 'Segoe UI', bold: true, size: 10, color: { argb: 'FFFFFFFF' } },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD71920' } },
                    alignment: { horizontal: 'center', vertical: 'middle' },
                    border: {
                        top: { style: 'thin', color: { argb: 'FFB0B0B0' } },
                        bottom: { style: 'thin', color: { argb: 'FFB0B0B0' } },
                        left: { style: 'thin', color: { argb: 'FFB0B0B0' } },
                        right: { style: 'thin', color: { argb: 'FFB0B0B0' } }
                    }
                };
            });

            for (let i = 1; i <= 8; i++) {
                const row = wsDirecciones.addRow(['', '', '', '', '', '', '', '', '', '', i]);
                row.height = 24;
                row.eachCell((cell, colIndex) => {
                    const isLocked = colIndex === 11;
                    cell.style = isLocked ? {
                        font: { name: 'Segoe UI', size: 10, color: { argb: 'FF818891' } },
                        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F0F0' } },
                        alignment: { horizontal: 'center', vertical: 'middle' },
                        border: {
                            top: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                            bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                            left: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                            right: { style: 'thin', color: { argb: 'FFD0D0D0' } }
                        }
                    } : {
                        font: { name: 'Segoe UI', size: 10, color: { argb: 'FF272A2F' } },
                        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } },
                        alignment: { horizontal: 'left', vertical: 'middle' },
                        border: {
                            top: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                            bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                            left: { style: 'thin', color: { argb: 'FFD0D0D0' } },
                            right: { style: 'thin', color: { argb: 'FFD0D0D0' } }
                        }
                    };
                    if (isLocked) {
                        cell.protection = { locked: true };
                    }
                });
            }

            wsDirecciones.columns.forEach((col, index) => {
                const widths = [18, 12, 10, 14, 16, 12, 14, 14, 14, 18, 14];
                col.width = widths[index] || 14;
            });

            // ============================================
            // ESCRIBIR Y DESCARGAR
            // ============================================
            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            saveAs(blob, 'Modelo de alta de contrato DMS.xlsx');
            log('📥 Plantilla descargada correctamente', 'success');
        }

        // ============================================================
        //  DRAG & DROP + INPUT
        // ============================================================
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');

        dropZone.addEventListener('click', () => fileInput.click());

        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        });
        dropZone.addEventListener('dragleave', () => {
            dropZone.classList.remove('dragover');
        });
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            const files = e.dataTransfer.files;
            handleFiles(files);
        });

        fileInput.addEventListener('change', (e) => {
            handleFiles(e.target.files);
            fileInput.value = '';
        });

        function handleFiles(fileList) {
            const valid = [];
            for (const f of fileList) {
                const ext = f.name.split('.').pop().toLowerCase();
                if (['xlsx'].includes(ext)) {
                    valid.push(f);
                } else {
                    log(`⚠️ Archivo ignorado (no es .xlsx): ${f.name}`, 'warning');
                }
            }
            if (valid.length === 0) return;
            uploadedFiles = uploadedFiles.concat(valid);
            renderFileList();
            updateStats();
            PROCESS_BTN.disabled = false;
            CLEAR_BTN.style.display = 'inline-flex';
            log(`✅ ${valid.length} archivo(s) agregado(s). Total: ${uploadedFiles.length}`, 'success');
            MESSAGE_BODY.textContent = '';
            COPY_BTN.disabled = true;
            COPY_BTN.classList.remove('copiado');
            COPY_BTN.textContent = '📋 Copiar mensaje';
            razonSocialCopiada = '';
            ocultarCorreccion();
        }

        function renderFileList() {
            FILE_LIST.innerHTML = uploadedFiles.map((f, i) =>
                `<li>📄 ${f.name} <span class="file-badge">${(f.size/1024).toFixed(1)} KB</span></li>`
            ).join('');
        }

        function updateStats() {
            STATS_BAR.style.display = 'flex';
            FILE_COUNT.textContent = uploadedFiles.length;
        }

        // ============================================================
        //  LIMPIAR
        // ============================================================
        CLEAR_BTN.addEventListener('click', () => {
            uploadedFiles = [];
            processedResults = [];
            processedFileNames = [];
            renderFileList();
            STATS_BAR.style.display = 'none';
            PROCESS_BTN.disabled = true;
            CLEAR_BTN.style.display = 'none';
            MESSAGE_BODY.textContent = '';
            COPY_BTN.disabled = true;
            COPY_BTN.classList.remove('copiado');
            COPY_BTN.textContent = '📋 Copiar mensaje';
            LOG.textContent = '📋 Todo limpiado. Esperando archivos...';
            razonSocialCopiada = '';
            ocultarCorreccion();
        });

        // ============================================================
        //  LOG
        // ============================================================
        function log(msg, type = 'info') {
            const types = {
                success: 'log-success',
                warning: 'log-warning',
                error: 'log-error',
                info: 'log-info'
            };
            const cls = types[type] || 'log-info';
            LOG.innerHTML += `<span class="${cls}">${msg}</span>\n`;
            LOG.scrollTop = LOG.scrollHeight;
        }

        // ============================================================
        //  COPIAR RAZÓN SOCIAL
        // ============================================================
        function copiarRazonSocial(texto) {
            const match = texto.match(/Razón Social: "([^"]+)"/);
            if (match) {
                const razonSocial = match[1];
                if (razonSocial && razonSocial !== 'No especificada') {
                    navigator.clipboard.writeText(razonSocial).then(() => {
                        const msg = `✅ Razón Social "${razonSocial}" copiada al portapapeles!`;
                        if (razonSocialCopiada !== razonSocial) {
                            razonSocialCopiada = razonSocial;
                            log(msg, 'success');
                        }
                    }).catch(() => {
                        const textarea = document.createElement('textarea');
                        textarea.value = razonSocial;
                        document.body.appendChild(textarea);
                        textarea.select();
                        document.execCommand('copy');
                        document.body.removeChild(textarea);
                        const msg = `✅ Razón Social "${razonSocial}" copiada al portapapeles!`;
                        if (razonSocialCopiada !== razonSocial) {
                            razonSocialCopiada = razonSocial;
                            log(msg, 'success');
                        }
                    });
                }
            }
        }

        LOG.addEventListener('click', function(e) {
            const target = e.target;
            const errorElement = target.closest('.log-error');
            if (errorElement) {
                const texto = errorElement.textContent;
                if (texto.includes('Razón Social:')) {
                    copiarRazonSocial(texto);
                }
            }
        });

        // ============================================================
        //  FUNCIÓN: REPROCESAR ARCHIVO CON NUEVO CÓDIGO
        // ============================================================
        async function reprocesarConCodigo(archivo, nuevoCodigo, razonSocial) {
            const validacion = normalizarCodigoSolicitante(nuevoCodigo);
            if (validacion.error) {
                CORRECCION_ERROR.textContent = `❌ ${validacion.error}`;
                CORRECCION_ERROR.style.display = 'block';
                return;
            }

            const codigoNormalizado = validacion.valor;
            CORRECCION_ERROR.style.display = 'none';
            BTN_REPROCESAR.disabled = true;
            BTN_REPROCESAR.textContent = '⏳ Procesando...';

            log(`🔄 Reprocesando con código "${codigoNormalizado}" para "${razonSocial}"...`, 'info');

            try {
                const buffer = await archivo.arrayBuffer();
                const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });

                const sheet = workbook.Sheets['Formulario'];
                if (!sheet) {
                    log(`❌ No se encuentra la hoja "Formulario" en el archivo`, 'error');
                    BTN_REPROCESAR.disabled = false;
                    BTN_REPROCESAR.textContent = '🔁 Reprocesar';
                    return;
                }

                const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

                let filaEncontrada = false;
                let filaIndex = -1;

                for (let i = 0; i < data.length; i++) {
                    const row = data[i];
                    if (row && row[0]) {
                        const campo = String(row[0]).trim();
                        if (campo === 'Código del Cliente' || campo === 'CodigoSolicitanteERP') {
                            filaIndex = i;
                            filaEncontrada = true;
                            break;
                        }
                    }
                }

                if (!filaEncontrada || filaIndex === -1) {
                    log(`❌ No se encontró la fila "Código del Cliente" en el archivo`, 'error');
                    BTN_REPROCESAR.disabled = false;
                    BTN_REPROCESAR.textContent = '🔁 Reprocesar';
                    return;
                }

                log(`   → Actualizando código en fila ${filaIndex + 1} de "${String(data[filaIndex][2] || '').trim()}" a "${codigoNormalizado}"`, 'info');
                data[filaIndex][2] = codigoNormalizado;

                const newSheet = XLSX.utils.aoa_to_sheet(data);
                workbook.Sheets['Formulario'] = newSheet;

                const verificarData = XLSX.utils.sheet_to_json(workbook.Sheets['Formulario'], { header: 1, defval: '' });
                const valorVerificado = String(verificarData[filaIndex]?.[2] || '').trim();
                if (valorVerificado !== codigoNormalizado) {
                    log(`⚠️ Advertencia: El código no se actualizó correctamente. Esperado: "${codigoNormalizado}", Actual: "${valorVerificado}"`, 'warning');
                } else {
                    log(`   → ✅ Código actualizado correctamente a "${codigoNormalizado}"`, 'success');
                }

                const validation = validateFormulario(workbook, archivo.name);
                if (validation.hasErrors) {
                    log(`❌ El archivo sigue teniendo errores después de la corrección:`, 'error');
                    for (const err of validation.errors) {
                        log(`   → ${err}`, 'error');
                    }
                    BTN_REPROCESAR.disabled = false;
                    BTN_REPROCESAR.textContent = '🔁 Reprocesar';
                    return;
                }

                const nombreData = getNombreArchivoData(workbook);
                if (nombreData.error) {
                    log(`❌ ${nombreData.error}`, 'error');
                    BTN_REPROCESAR.disabled = false;
                    BTN_REPROCESAR.textContent = '🔁 Reprocesar';
                    return;
                }

                const normalizedWB = normalizeWorkbook(workbook);

                const sheetNames = normalizedWB.SheetNames;
                const formularioIndex = sheetNames.indexOf('Formulario');
                if (formularioIndex !== -1) {
                    normalizedWB.SheetNames.splice(formularioIndex, 1);
                    delete normalizedWB.Sheets['Formulario'];
                }

                const refIndex = normalizedWB.SheetNames.indexOf('Referencias');
                if (refIndex !== -1) {
                    const refName = normalizedWB.SheetNames.splice(refIndex, 1)[0];
                    normalizedWB.SheetNames.push(refName);
                }

                limpiarDireccionesVacias(normalizedWB);

                let outputName = `Alta de contrato DMS ${nombreData.codigoERP} ${nombreData.codigoSolicitante}`;
                if (nombreData.clientName) {
                    outputName += ` ${nombreData.clientName}`;
                }
                outputName += '.xlsx';

                const wbout = XLSX.write(normalizedWB, { bookType: 'xlsx', type: 'array' });
                downloadFile(wbout, outputName);

                log(`✅ ¡Archivo reprocesado correctamente! Descargando "${outputName}"`, 'success');

                processedFileNames.push(outputName);
                generarMensaje(processedFileNames.length, processedFileNames);

                ocultarCorreccion();

            } catch (err) {
                log(`❌ Error al reprocesar: ${err.message}`, 'error');
                console.error(err);
            }

            BTN_REPROCESAR.disabled = false;
            BTN_REPROCESAR.textContent = '🔁 Reprocesar';
        }

        // ============================================================
        //  EVENTOS DEL CAMPO DE CORRECCIÓN
        // ============================================================
        BTN_REPROCESAR.addEventListener('click', function() {
            const codigo = CODIGO_CORRECCION.value.trim();
            if (!codigo) {
                CORRECCION_ERROR.textContent = '⚠️ Ingresá un código válido';
                CORRECCION_ERROR.style.display = 'block';
                return;
            }
            if (pendingFileForReprocess) {
                reprocesarConCodigo(
                    pendingFileForReprocess.file,
                    codigo,
                    pendingFileForReprocess.razonSocial
                );
            }
        });

        CODIGO_CORRECCION.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                BTN_REPROCESAR.click();
            }
        });

        // ============================================================
        //  COPIAR MENSAJE
        // ============================================================
        COPY_BTN.addEventListener('click', async () => {
            const texto = MESSAGE_BODY.textContent;
            if (!texto) return;

            try {
                await navigator.clipboard.writeText(texto);
                COPY_BTN.textContent = '✅ ¡Copiado!';
                COPY_BTN.classList.add('copiado');
                setTimeout(() => {
                    COPY_BTN.textContent = '📋 Copiar mensaje';
                    COPY_BTN.classList.remove('copiado');
                }, 3000);
            } catch (err) {
                const textarea = document.createElement('textarea');
                textarea.value = texto;
                document.body.appendChild(textarea);
                textarea.select();
                document.execCommand('copy');
                document.body.removeChild(textarea);
                COPY_BTN.textContent = '✅ ¡Copiado!';
                COPY_BTN.classList.add('copiado');
                setTimeout(() => {
                    COPY_BTN.textContent = '📋 Copiar mensaje';
                    COPY_BTN.classList.remove('copiado');
                }, 3000);
            }
        });
