"""T-029 E3 (2026-11-25): Continente "Pallet" en el normalizador.

En la hoja Referencias del modelo el continente es "Pallet" (codigo DMS 4), pero el normalizador
pasaba "pallet"/"pallets" a "Pallets", que no existe en Referencias: un contrato de pallets daba
error de validacion, y la regla Continente <-> SLA tampoco funcionaba con "Pallet".
Ademas, en el modelo V9/V10 el nombre "Continente" esta en la columna B: la correccion
SLA -> Continente (D-107.1) no encontraba la fila. Se busca con campo_de_fila.

Este script lo corre el workflow .github/workflows/parches.yml: (1) prueba el normalizador actual
con datos sinteticos, (2) aplica los reemplazos (idempotente), (3) vuelve a probar y (4) deja el
resultado en .github/parches/resultado-e3.md. Sin datos reales (D-003).
"""
import datetime
import importlib.util
import os
import sys
import tempfile

import openpyxl

RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
SALIDA = os.path.join(RAIZ, '.github', 'parches', 'resultado-e3.md')

NOTA_PY = ('Regla E3 de T-029 (2026-11-25): el Continente es "Pallet" (asi figura en Referencias, codigo 4);\n'
           '"pallet"/"pallets" se normalizan a "Pallet" y la regla Continente <-> SLA usa Pallet <-> SLA PALLETS.\n'
           'La correccion SLA -> Continente encuentra la fila aunque el nombre este en la columna B (modelo V9/V10).\n')

BUSCAR_VIEJO = ('    def buscar(nombre):\n'
                '        for r in form:\n'
                '            if r and js_str(r[0]).strip() == nombre:\n'
                '                return r\n'
                '        return None\n')
BUSCAR_NUEVO = ('    def buscar(nombre):\n'
                '        for k, r in enumerate(form):  # T-029 E3: tambien con el nombre en la columna B\n'
                '            if r and campo_de_fila(r, k) == nombre:\n'
                '                return r\n'
                '        return None\n')

REEMPLAZOS = {
    'automatizacion/normalizador_dms.py': [
        ("'pallet': 'Pallets', 'pallets': 'Pallets',", "'pallet': 'Pallet', 'pallets': 'Pallet',"),
        ("c2s = {'Pallets': 'SLA PALLETS'", "c2s = {'Pallet': 'SLA PALLETS'"),
        (BUSCAR_VIEJO, BUSCAR_NUEVO),
        ('Reglas D-110 (2026-11-20, Pablo)', NOTA_PY + 'Reglas D-110 (2026-11-20, Pablo)'),
    ],
    'js/normalizador-2.js': [
        ("'pallet': 'Pallets',", "'pallet': 'Pallet',"),
        ("'pallets': 'Pallets',", "'pallets': 'Pallet',"),
        ("const CONTINENTE_PALLETS = 'Pallets';",
         "const CONTINENTE_PALLETS = 'Pallet'; // T-029 E3: asi figura en Referencias (codigo 4)"),
    ],
    'masivo/masivo-reglas.js': [
        ("'pallet': 'Pallets',", "'pallet': 'Pallet',"),
        ("'pallets': 'Pallets',", "'pallets': 'Pallet',"),
        ("const CONTINENTE_PALLETS = 'Pallets';",
         "const CONTINENTE_PALLETS = 'Pallet'; // T-029 E3: asi figura en Referencias (codigo 4)"),
    ],
    'js/normalizador-1.js': [
        ("'Contenedor, Bulto, Pallets, Bultos Bigger, Sobre, Bodega, etc.'",
         "'Contenedor, Bulto/Paquete, Pallet, Bultos Bigger, Sobre, Bodega, etc.'"),
    ],
    'index.html': [
        ('js/normalizador-1.js?v=2.1"', 'js/normalizador-1.js?v=2.2"'),
        ('js/normalizador-2.js?v=2.2"', 'js/normalizador-2.js?v=2.3"'),
    ],
    'masivo/index.html': [
        ('masivo-reglas.js?v=1.1"', 'masivo-reglas.js?v=1.2"'),
    ],
}

BOOLS = [('AllowRescue', 'true'), ('AllowWithdrawPartialValues', 'false'), ('ObligatoryGauge', 'false'),
         ('AllowPartialDelivery', 'true'), ('IsCentralizedReturn', 'true'), ('HasCollectionManagement', 'false'),
         ('HasDeliveryNote', 'CON_REMITO'), ('AllowPrintDocumentationReturn', 'true')]
CONTRATO_HDR = ['NombreContratoDMS', 'CodigoContratoERP', 'DescripcionContratoERP', 'CodigoSolicitanteERP',
                'IDSociedadSAP', 'TipoDeEnvio', 'CondicionDeEntrega', 'Habilitado', 'TipoEntrega',
                'FechaHabilitado', 'ValidoDesde', 'ValidoHasta', 'ClientName', 'SLA', 'AllowRescue',
                'AllowWithdrawPartialValues', 'ObligatoryGauge', 'AllowPartialDelivery', 'IsCentralizedReturn',
                'HasCollectionManagement', 'DeliveryNote', 'AllowPrintDocumentationReturn', 'MotiveStateMap',
                'AutomaticAdmission', 'AllowsStorage', 'HasChange']


def libro_sintetico(continente, sla, ruta):
    """Excel con la estructura del modelo V9/V10 y datos inventados."""
    wb = openpyxl.Workbook()
    f = wb.active
    f.title = 'Formulario'
    filas = [('Campos a Completar', 'Campos a Completar', 'Completar'), ('NombreContratoDMS', '', ''),
             ('CodigoContratoERP', '', '351000001'), ('DescripcionContratoERP', '', 'PRUEBA SINTETICA'),
             ('CodigoSolicitanteERP', '', '100001'), ('IDSociedadSAP', '', '1100'),
             ('TipoDeEnvio', '', 'Paqueter\u00eda'), ('CondicionDeEntrega', '', ''), ('Habilitado', '', '1'),
             ('TipoEntrega', '', 'Domicilio'), ('FechaHabilitado', '', datetime.datetime(2026, 11, 25)),
             ('ValidoDesde', '', datetime.datetime(2025, 1, 1)), ('ValidoHasta', '', datetime.datetime(9999, 12, 8)),
             ('ClientName', '', 'CLIENTE PRUEBA SA'), ('SLA', '', sla)]
    filas += [(a, '', b) for a, b in BOOLS]
    filas += [('', '\u00bfEs recambio?', 'false'), ('', '', ''), ('', 'Niveles de esfuerzo', 'Seleccionar Opci\u00f3n'),
              ('', 'Log\u00edstica Solicitada', 'Entrega'), ('', 'Continente', continente),
              ('', 'Plazo', 'SLA 1 - Estandar B2B'), ('', 'FormularioMobileB2B', 'Formulario B2B 1')]
    for i, (a, b, c) in enumerate(filas, start=1):
        f.cell(i, 1, a or None)
        f.cell(i, 2, b or None)
        f.cell(i, 3, c if c != '' else None)
    c = wb.create_sheet('Contrato')
    c.append(CONTRATO_HDR)
    c.append(['x', '351000001', 'PRUEBA SINTETICA', '100001', '1100', 'Paqueter\u00eda', '', '1', 'Domicilio',
              '', '', '', 'CLIENTE PRUEBA SA', sla, 'true', 'false', 'false', 'true', 'true', 'false',
              'CON_REMITO', 'true', 'ALERTRAN ESTANDAR MOTIVES V4', 'false', 'false', 'false'])
    d = wb.create_sheet('Direcciones')
    d.append(['Street', 'Number', 'Floor', 'Department', 'City', 'PostalCode', 'State', 'Country', 'Phone',
              'MailBox', 'AddressType'])
    d.append(['Calle Falsa', '123', '0', '0', 'CABA', '1000', 'BUENOS AIRES', 'ARGENTINA', '0', '0', '1'])
    r = wb.create_sheet('Referencias')
    r['E1'] = 'Continente'
    r['F1'] = 'C\u00f3digo DMS'
    for k, (n, cod) in enumerate([('Bulto/Paquete', '3'), ('Pallet', '4'), ('Bultos Bigger', '10'),
                                  ('Pallets irregulares', '12')], start=2):
        r.cell(k, 5, n)
        r.cell(k, 6, cod)
    wb.save(ruta)


CASOS = [
    ('Continente "Pallet" con SLA 1', 'Pallet', 'SLA 1', 'Pallet', 'SLA PALLETS'),
    ('Continente "pallets" (mal escrito)', 'pallets', 'SLA 1', 'Pallet', 'SLA PALLETS'),
    ('SLA PALLETS con Continente "Bulto/Paquete"', 'Bulto/Paquete', 'SLA PALLETS', 'Pallet', 'SLA PALLETS'),
    ('Continente "Bultos Bigger" (no cambia)', 'Bultos Bigger', 'SLA 1', 'Bultos Bigger', 'SLA BIGGER'),
    ('Continente "Pallets irregulares" (no cambia)', 'Pallets irregulares', 'SLA 1', 'Pallets irregulares', 'SLA 1'),
]


def probar():
    spec = importlib.util.spec_from_file_location('norm', os.path.join(RAIZ, 'automatizacion', 'normalizador_dms.py'))
    norm = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(norm)
    out, todos = [], True
    with tempfile.TemporaryDirectory() as tmp:
        for i, (nombre, cont, sla, cont_ok, sla_ok) in enumerate(CASOS):
            ruta = os.path.join(tmp, f'caso{i}.xlsx')
            libro_sintetico(cont, sla, ruta)
            res = norm.procesar(ruta, os.path.join(tmp, f'sal{i}'))
            sla_final = cont_final = None
            if res['ok']:
                ws = openpyxl.load_workbook(res['archivo'])['Contrato']
                hdr = [x.value for x in ws[1]]
                sla_final = ws.cell(2, hdr.index('SLA') + 1).value
                wbf = openpyxl.load_workbook(ruta, data_only=True)
                form = norm.hoja_a_filas(wbf['Formulario'])
                refs = norm.extraer_referencias(norm.hoja_a_filas(wbf['Referencias']))
                norm.validar_formulario(form, refs, [])
                cont_final = norm.js_str(form[27][2]).strip()
            ok = res['ok'] and sla_final == sla_ok and cont_final == cont_ok
            todos = todos and ok
            errs = '; '.join(res['errores'])[:160].replace('|', '/')
            out.append(f"| {nombre} | {'OK' if ok else 'FALLA'} | {res['ok']} | {cont_final} | {sla_final} | {errs} |")
    return todos, out


def aplicar():
    cambios = []
    for archivo, pares in REEMPLAZOS.items():
        ruta = os.path.join(RAIZ, archivo)
        with open(ruta, encoding='utf-8', newline='') as fh:
            txt = fh.read()
        orig = txt
        for viejo, nuevo in pares:
            corto = viejo.strip().splitlines()[0][:60]
            n = txt.count(viejo)
            if n == 1:
                txt = txt.replace(viejo, nuevo)
                cambios.append(f'- `{archivo}`: aplicado `{corto}`')
            elif n == 0 and nuevo in txt:
                cambios.append(f'- `{archivo}`: ya estaba aplicado `{corto}`')
            else:
                raise SystemExit(f'{archivo}: "{corto}" aparece {n} veces; no se toca nada')
        if txt != orig:
            with open(ruta, 'w', encoding='utf-8', newline='') as fh:
                fh.write(txt)
    return cambios


if __name__ == '__main__':
    hdr = ['| Caso | Resultado | ok | Continente final | SLA en Contrato | Errores |', '|---|---|---|---|---|---|']
    antes_ok, antes = probar()
    cambios = aplicar()
    despues_ok, despues = probar()
    md = ['# T-029 E3 - Continente "Pallet" en el normalizador', '',
          f'Corrida: {datetime.datetime.utcnow().isoformat()}Z', '', '## Antes del cambio', ''] + hdr + antes + \
         ['', '## Cambios', ''] + cambios + ['', '## Despues del cambio', ''] + hdr + despues + \
         ['', f'**Resultado final: {"OK" if despues_ok else "FALLA"}**', '']
    with open(SALIDA, 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(md))
    print('\n'.join(md))
    sys.exit(0 if despues_ok else 1)
