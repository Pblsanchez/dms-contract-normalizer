"""Normalizador de contratos DMS - puerto a Python (T-020).

Copia de las reglas de dms-contract-normalizer (js/normalizador-1.js + js/normalizador-2.js,
iguales a masivo/masivo-reglas.js), para que la automatizacion aplique exactamente lo mismo
que el normalizador diario. Si cambia una regla alla, hay que copiarla aca tambien.

Uso: python3 normalizador_dms.py <entrada.xlsx | datos.json> <carpeta_salida>
(datos.json = texto que devuelve el Office Script LeerAltaContratoDMS.ts del flujo de T-020)
Salida: JSON con ok, errores, log, archivo, codigoERP, solicitante (6 digitos), razonSocial.
Sin datos reales en este archivo (D-003).
"""
import datetime
import json
import os
import re
import sys

import openpyxl

ENVOLVENTES = ['"', "'", '-', '_', '*', '\u2013', '\u2014', '`', '\u00ab', '\u00bb',
               '\u201c', '\u201d', '\u2018', '\u2019', '(', ')']
TIPOS_VALIDOS = ['Minorista', 'Mayorista', 'Cadena de fr\u00edo', 'Cargas completas', 'Congelado',
                 'Log\u00edstica inversa', 'Paqueter\u00eda', 'Retira dep\u00f3sito', 'Log\u00edstica Asegurada']
SLAS_VALIDOS = ['SLA 1', 'SLA 2', 'SLA 3', 'SLA 4 - Logistica Asegurada V2', 'SLA PALLETS', 'SLA BIGGER']
DICCIONARIOS = {
    'TipoDeEnvio': {
        'minorista': 'Minorista', 'mayorista': 'Mayorista',
        'cadena de frio': 'Cadena de fr\u00edo', 'cadena de fr\u00edo': 'Cadena de fr\u00edo',
        'carga completa': 'Cargas completas', 'cargas completa': 'Cargas completas',
        'cargas completas': 'Cargas completas',
        'logistica inversa': 'Log\u00edstica inversa', 'log\u00edstica inversa': 'Log\u00edstica inversa',
        'retira deposito': 'Retira dep\u00f3sito', 'retira dep\u00f3sito': 'Retira dep\u00f3sito',
        'logistica asegurada': 'Log\u00edstica Asegurada', 'log\u00edstica asegurada': 'Log\u00edstica Asegurada',
        'paqueteria': 'Paqueter\u00eda', 'paqueter\u00eda': 'Paqueter\u00eda', 'congelado': 'Congelado'},
    'SLA': {
        'sla 1': 'SLA 1', 'sla1': 'SLA 1', 'sla 2': 'SLA 2', 'sla2': 'SLA 2',
        'sla 3': 'SLA 3', 'sla3': 'SLA 3',
        'sla 4 - logistica asegurada v2': 'SLA 4 - Logistica Asegurada V2',
        'sla 4 - log\u00edstica asegurada v2': 'SLA 4 - Logistica Asegurada V2',
        'sla 4 - logistica asegurada': 'SLA 4 - Logistica Asegurada V2',
        'sla 4 - log\u00edstica asegurada': 'SLA 4 - Logistica Asegurada V2',
        'sla4 - logistica asegurada v2': 'SLA 4 - Logistica Asegurada V2',
        'sla 4 logistica asegurada v2': 'SLA 4 - Logistica Asegurada V2',
        'logistica asegurada v2': 'SLA 4 - Logistica Asegurada V2',
        'sla pallet': 'SLA PALLETS', 'sla pallets': 'SLA PALLETS', 'slapallet': 'SLA PALLETS',
        'slapallets': 'SLA PALLETS', 'sla bigger': 'SLA BIGGER', 'sla biger': 'SLA BIGGER',
        'slabigger': 'SLA BIGGER'},
    'Continente': {
        'pallet': 'Pallets', 'pallets': 'Pallets', 'bigger': 'Bultos Bigger',
        'bultos bigger': 'Bultos Bigger', 'bulto bigger': 'Bultos Bigger',
        'bultos biger': 'Bultos Bigger', 'bulto biger': 'Bultos Bigger'},
    'HasDeliveryNote': {
        'con remito': 'CON_REMITO', 'conremito': 'CON_REMITO', 'sin remito': 'SIN_REMITO',
        'sinremito': 'SIN_REMITO', 'digitalizado': 'DIGITALIZADO'},
}
MAPA_REF = {'Log\u00edstica Solicitada': 'logistica', 'Continente': 'continente', 'SLA': 'sla',
            'Plazo': 'sla', 'FormularioMobileB2B': 'formulario'}
CAMPOS_REFERENCIA = ['TipoDeEnvio', 'Log\u00edstica Solicitada', 'Continente', 'SLA', 'Plazo', 'FormularioMobileB2B']
BOOLEANOS = ['AllowRescue', 'AllowWithdrawPartialValues', 'ObligatoryGauge', 'AllowPartialDelivery',
             'IsCentralizedReturn', 'HasCollectionManagement', 'AllowPrintDocumentationReturn', '\u00bfEs recambio?']


def js_str(v):
    """Equivalente a String(v || '') de SheetJS con raw:true (valor crudo)."""
    if v is None or v == '' or v is False or (not isinstance(v, bool) and isinstance(v, (int, float)) and v == 0):
        return ''  # en JS, 0 / '' / null / false son "falsy" -> ''
    if v is True:
        return 'true'
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, datetime.datetime):
        return v.isoformat()
    return str(v)


def fmt_val(cell):
    """Equivalente a sheet_to_json(raw:false): texto como se ve en Excel (General / m/d/yy)."""
    v = cell.value
    if v is None:
        return ''
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    if isinstance(v, (datetime.datetime, datetime.date)):
        return f'{v.month}/{v.day}/{v.year % 100:02d}'
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    return str(v)


def hoja_a_filas(ws, crudo=True):
    filas = []
    for row in ws.iter_rows():
        filas.append([(c.value if crudo else fmt_val(c)) for c in row])
    while filas and all((x is None or x == '') for x in filas[-1]):
        filas.pop()
    return filas


def celda(data, r, c):
    if r < len(data) and c < len(data[r]):
        return data[r][c]
    return None


def limpiar_texto_envolvente(texto):
    if not texto:
        return ''
    r = str(texto).strip()
    cambio = True
    while cambio and r:
        cambio = False
        for c in ENVOLVENTES:
            if r.startswith(c):
                r = r[len(c):].strip(); cambio = True; break
        if cambio:
            continue
        for c in ENVOLVENTES:
            if r.endswith(c):
                r = r[:-len(c)].strip(); cambio = True; break
    return r


def normalizar_codigo_solicitante(valor):
    if not valor:
        return None
    s = str(valor).strip()
    if not re.fullmatch(r'\d+', s):
        return {'error': 'El c\u00f3digo debe ser num\u00e9rico'}
    n = len(s)
    if n == 6:
        return {'valor': s}
    if n == 10 and s.startswith('0000'):
        return {'valor': s[4:]}
    if n == 10:
        return {'error': 'El c\u00f3digo tiene 10 d\u00edgitos pero no comienza con 4 ceros'}
    return {'error': f'El c\u00f3digo tiene {n} d\u00edgitos, debe tener 6 d\u00edgitos (o 10 d\u00edgitos comenzando con 4 ceros)'}


def normalizar_has_delivery_note(valor):
    if not valor:
        return 'SIN_REMITO'
    v = str(valor).upper().strip()
    if v in ('CON_REMITO', 'SIN_REMITO', 'DIGITALIZADO'):
        return v
    if v in ('TRUE', '1', 'SI', 'S\u00cd'):
        return 'CON_REMITO'
    return 'SIN_REMITO'


def extraer_referencias(ref):
    refs, actual = {}, None
    for row in ref:
        g = lambda i: str(row[i] if i < len(row) and row[i] not in (None,) else '').strip()
        A, B, C, E, F, G, I, J, K, M, N, O, P = (g(i) for i in (0, 1, 2, 4, 5, 6, 8, 9, 10, 12, 13, 14, 15))
        if A == 'Logistica Solicitada':
            actual = 'logistica'; refs[actual] = []; continue
        if E == 'Continente':
            actual = 'continente'; refs[actual] = []; continue
        if I == 'SLAs':
            actual = 'sla'; refs[actual] = []; continue
        if M == 'Formulario Mobile B2B':
            actual = 'formulario'; refs[actual] = []; continue
        if actual == 'logistica' and A and B:
            refs['logistica'].append({'key': A, 'code': B, 'desc': C})
        if actual == 'continente' and E and F:
            refs['continente'].append({'key': E, 'code': F, 'desc': G})
        if actual == 'sla' and I and J:
            refs['sla'].append({'key': I, 'code': J, 'desc': K})
        if actual == 'formulario' and M and N:
            refs['formulario'].append({'key': M, 'code': N, 'desc': O, 'aplica': P})
    return refs


def normalizar_con_diccionario(campo, valor, refs):
    if not valor:
        return valor, False
    limpio = str(valor).strip()
    dic = DICCIONARIOS.get(campo)
    if dic:
        low = limpio.lower()
        for k, v in dic.items():
            if low == k.lower():
                return v, limpio != v
    t = MAPA_REF.get(campo)
    if t and refs.get(t):
        low = limpio.lower()
        for it in refs[t]:
            if low in (it['key'].lower(), it['code'].lower()) or (it['desc'] and low == it['desc'].lower()):
                return it['key'], limpio != it['key']
    return limpio, False


def validar_contra_referencias(campo, valor, refs):
    vn, norm = normalizar_con_diccionario(campo, valor, refs)
    if campo == 'TipoDeEnvio':
        ok = vn in TIPOS_VALIDOS
        return ok, (vn if ok else valor), norm
    if campo in ('SLA', 'Plazo'):
        ok = vn in SLAS_VALIDOS
        return ok, (vn if ok else valor), norm
    t = {'Log\u00edstica Solicitada': 'logistica', 'Continente': 'continente', 'FormularioMobileB2B': 'formulario'}.get(campo)
    if not t or not refs.get(t):
        return True, valor, False
    ok = any(vn in (it['key'], it['code'], it['desc']) for it in refs[t])
    return ok, (vn if ok else valor), norm


def validar_formulario(form, refs, log):
    errores = []
    filas = list(range(2, 24)) + list(range(26, 30))
    rs = limpiar_texto_envolvente(js_str(celda(form, 13, 2)).strip()) or 'No especificada'

    def buscar(nombre):
        for r in form:
            if r and js_str(r[0]).strip() == nombre:
                return r
        return None

    for i in filas:
        if i >= len(form):
            errores.append(f'Fila {i+1} no existe en el archivo'); continue
        row = form[i]
        while len(row) < 3:
            row.append(None)
        campo = js_str(row[0]).strip()
        valor = js_str(row[2]).strip()
        if not campo:
            continue
        if campo == 'HasDeliveryNote' and valor:
            vn = normalizar_has_delivery_note(valor)
            if vn != valor:
                row[2] = vn; log.append(f'HasDeliveryNote: "{valor}" -> "{vn}"'); valor = vn
        if campo in CAMPOS_REFERENCIA:
            ok, v2, norm = validar_contra_referencias(campo, valor, refs)
            if norm:
                row[2] = v2; log.append(f'{campo}: "{valor}" -> "{v2}"'); valor = v2
            if not ok:
                if campo == 'TipoDeEnvio':
                    vv = TIPOS_VALIDOS
                elif campo == 'SLA':
                    vv = SLAS_VALIDOS
                elif campo == 'Plazo':
                    vv = ['SLA 1 - Est\u00e1ndar B2B', 'SLA 2 - Acordado con cliente', 'SLA 3 - Urgente']
                else:
                    t = {'Log\u00edstica Solicitada': 'logistica', 'Continente': 'continente', 'FormularioMobileB2B': 'formulario'}.get(campo)
                    vv = [it['key'] for it in refs.get(t, [])] if t else []
                errores.append(f'Campo "{campo}" = "{valor}" no es v\u00e1lido. Valores v\u00e1lidos: {", ".join(vv)}')
        if campo == 'HasDeliveryNote' and valor:
            ok, v2, norm = validar_contra_referencias(campo, valor, refs)
            if not ok:
                errores.append(f'Campo "{campo}" = "{valor}" no es v\u00e1lido. Valores v\u00e1lidos: CON_REMITO, SIN_REMITO, DIGITALIZADO')
        if campo not in ('CondicionDeEntrega', 'DescripcionContratoERP') and not valor:
            if campo == 'CodigoSolicitanteERP':
                errores.append(f'Campo "{campo}" est\u00e1 vac\u00edo (fila {i+1}) - Raz\u00f3n Social: "{rs}"')
            else:
                errores.append(f'Campo "{campo}" est\u00e1 vac\u00edo (fila {i+1})')
        if campo == 'CodigoContratoERP' and valor and not re.fullmatch(r'\d{9}', valor):
            errores.append(f'Campo "CodigoContratoERP" = "{valor}" debe tener 9 d\u00edgitos')
        if campo == 'IDSociedadSAP' and valor and valor != '1100':
            errores.append(f'Campo "IDSociedadSAP" = "{valor}" debe ser "1100" para B2B')
        if campo == 'Habilitado' and valor and valor not in ('1', '0'):
            errores.append(f'Campo "Habilitado" = "{valor}" debe ser 1 o 0')
        if campo == 'CodigoSolicitanteERP' and valor:
            r = normalizar_codigo_solicitante(valor)
            if r and r.get('error'):
                errores.append(f'Campo "CodigoSolicitanteERP" = "{valor}" - {r["error"]} - Raz\u00f3n Social: "{rs}"')
        # Cruzadas TipoDeEnvio <-> SLA
        if campo == 'TipoDeEnvio' and valor == 'Log\u00edstica Asegurada':
            r = buscar('SLA')
            if r is not None and js_str(r[2]).strip() != 'SLA 4 - Logistica Asegurada V2':
                log.append(f'SLA auto-corregido a "SLA 4 - Logistica Asegurada V2"'); r[2] = 'SLA 4 - Logistica Asegurada V2'
        if campo == 'SLA' and valor == 'SLA 4 - Logistica Asegurada V2':
            r = buscar('TipoDeEnvio')
            if r is not None and js_str(r[2]).strip() != 'Log\u00edstica Asegurada':
                log.append('TipoDeEnvio auto-corregido a "Log\u00edstica Asegurada"'); r[2] = 'Log\u00edstica Asegurada'
        # Cruzadas Continente <-> SLA
        c2s = {'Pallets': 'SLA PALLETS', 'Bultos Bigger': 'SLA BIGGER'}
        s2c = {v: k for k, v in c2s.items()}
        if campo == 'Continente' and valor in c2s:
            r = buscar('SLA')
            if r is not None and js_str(r[2]).strip() != c2s[valor]:
                log.append(f'SLA auto-corregido a "{c2s[valor]}"'); r[2] = c2s[valor]
        if campo in ('SLA', 'Plazo') and valor in s2c:
            r = buscar('Continente')
            if r is not None and js_str(r[2]).strip() != s2c[valor]:
                log.append(f'Continente auto-corregido a "{s2c[valor]}"'); r[2] = s2c[valor]
        if campo in BOOLEANOS and valor:
            if valor.lower() not in ('true', 'false', '1', '0', 'si', 'no'):
                errores.append(f'Campo "{campo}" = "{valor}" debe ser true/false o si/no')
    return errores


def libro_desde_json(texto):
    """Arma un libro en memoria con la salida del Office Script LeerAltaContratoDMS
    (todas las hojas con sus valores; las celdas con formato de fecha vienen como serial de Excel)."""
    datos = json.loads(texto)
    if isinstance(datos, str):  # el flujo puede devolverlo doblemente serializado
        datos = json.loads(datos)
    base = datetime.datetime(1899, 12, 30)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for h in datos['hojas']:
        ws = wb.create_sheet(h['nombre'])
        fechas = {(r, c) for r, c in h.get('fechas', [])}
        for r, fila in enumerate(h['valores']):
            for c, v in enumerate(fila):
                if v == '' or v is None:
                    continue
                if (r, c) in fechas and isinstance(v, (int, float)):
                    v = base + datetime.timedelta(days=float(v))
                ws.cell(row=r + 1, column=c + 1, value=v)
    return wb


def procesar(entrada, carpeta_salida):
    """entrada: ruta a un .xlsx, un libro openpyxl o el texto JSON del Office Script."""
    log = []
    if isinstance(entrada, openpyxl.Workbook):
        wb_f = entrada
    elif isinstance(entrada, str) and entrada.lstrip().startswith(('{', '"')):
        wb_f = libro_desde_json(entrada)
    else:
        wb_f = openpyxl.load_workbook(entrada, data_only=True)
    if 'Formulario' not in wb_f.sheetnames:
        return {'ok': False, 'errores': ['No existe la hoja "Formulario"'], 'log': log}
    form = hoja_a_filas(wb_f['Formulario'])
    ref = hoja_a_filas(wb_f['Referencias']) if 'Referencias' in wb_f.sheetnames else []
    refs = extraer_referencias(ref)
    errores = validar_formulario(form, refs, log)
    if errores:
        return {'ok': False, 'errores': errores, 'log': log}
    # getNombreArchivoData
    codigo_erp = js_str(celda(form, 2, 2)).strip()
    sol_raw = js_str(celda(form, 4, 2)).strip()
    cn_raw = js_str(celda(form, 13, 2)).strip()
    client = limpiar_texto_envolvente(cn_raw)
    if not codigo_erp:
        return {'ok': False, 'errores': ['Campo "CodigoContratoERP" (C3) est\u00e1 vac\u00edo. Es obligatorio.'], 'log': log}
    if not sol_raw:
        return {'ok': False, 'errores': [f'Campo "CodigoSolicitanteERP" (C5) est\u00e1 vac\u00edo. Es obligatorio. - Raz\u00f3n Social: "{client or "No especificada"}"'], 'log': log}
    r = normalizar_codigo_solicitante(sol_raw)
    if r.get('error'):
        return {'ok': False, 'errores': [f'Campo "CodigoSolicitanteERP" = "{sol_raw}" - {r["error"]} - Raz\u00f3n Social: "{client or "No especificada"}"'], 'log': log}
    sol = r['valor']
    if sol != sol_raw:
        log.append(f'CodigoSolicitanteERP: "{sol_raw}" -> "{sol}"')
    if cn_raw != client:
        log.append(f'ClientName limpiado: "{cn_raw}" -> "{client}"')
    # normalizeWorkbook: todas las hojas como texto (raw:false), sin Formulario, Referencias al final
    salida = openpyxl.Workbook()
    salida.remove(salida.active)
    orden = [s for s in wb_f.sheetnames if s not in ('Formulario', 'Referencias')]
    if 'Referencias' in wb_f.sheetnames:
        orden.append('Referencias')
    for nombre in orden:
        data = hoja_a_filas(wb_f[nombre], crudo=False)
        if nombre == 'Contrato' and len(data) > 1:
            fila = data[1]
            if len(fila) > 3:
                fila[3] = sol
            if len(fila) > 12 and client:
                fila[12] = client
            if len(fila) > 24:
                fila[24] = 'false'
        if nombre == 'Direcciones' and data:
            antes = len(data)
            data = [data[0]] + [f for f in data[1:] if any(str(x).strip() not in ('', '0') for x in f)]
            if len(data) != antes:
                log.append(f'Se eliminaron {antes - len(data)} filas vac\u00edas en "Direcciones"')
        ws = salida.create_sheet(nombre)
        for f in data:
            ws.append([(x if x != '' else None) for x in f])
    nombre_salida = f'Alta de contrato DMS {codigo_erp} {sol}' + (f' {client}' if client else '') + '.xlsx'
    os.makedirs(carpeta_salida, exist_ok=True)
    ruta = os.path.join(carpeta_salida, nombre_salida)
    salida.save(ruta)
    return {'ok': True, 'errores': [], 'log': log, 'archivo': ruta, 'codigoERP': codigo_erp,
            'solicitante': sol, 'razonSocial': client}


if __name__ == '__main__':
    arg = sys.argv[1]
    if arg.lower().endswith('.json'):
        with open(arg, encoding='utf-8') as fh:
            arg = fh.read()
    print(json.dumps(procesar(arg, sys.argv[2]), ensure_ascii=False, indent=1))
