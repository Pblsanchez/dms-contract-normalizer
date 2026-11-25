# T-029 E3 - Continente "Pallet" en el normalizador

Corrida: 2026-11-25T18:54:09.312481Z

## Antes del cambio

| Caso | Resultado | ok | Continente final | SLA en Contrato | Errores |
|---|---|---|---|---|---|
| Continente "Pallet" con SLA 1 | FALLA | False | None | None | Campo "Continente" = "Pallet" no es válido. Valores válidos: Bulto/Paquete, Pallet, Bultos Bigger, Pallets irregulares |
| Continente "pallets" (mal escrito) | FALLA | False | None | None | Campo "Continente" = "pallets" no es válido. Valores válidos: Bulto/Paquete, Pallet, Bultos Bigger, Pallets irregulares |
| SLA PALLETS con Continente "Bulto/Paquete" | FALLA | True | Bulto/Paquete | SLA PALLETS |  |
| Continente "Bultos Bigger" (no cambia) | OK | True | Bultos Bigger | SLA BIGGER |  |
| Continente "Pallets irregulares" (no cambia) | OK | True | Pallets irregulares | SLA 1 |  |

## Cambios

- `automatizacion/normalizador_dms.py`: aplicado `'pallet': 'Pallets', 'pallets': 'Pallets',`
- `automatizacion/normalizador_dms.py`: aplicado `c2s = {'Pallets': 'SLA PALLETS'`
- `automatizacion/normalizador_dms.py`: aplicado `def buscar(nombre):`
- `automatizacion/normalizador_dms.py`: aplicado `Reglas D-110 (2026-11-20, Pablo)`
- `js/normalizador-2.js`: aplicado `'pallet': 'Pallets',`
- `js/normalizador-2.js`: aplicado `'pallets': 'Pallets',`
- `js/normalizador-2.js`: aplicado `const CONTINENTE_PALLETS = 'Pallets';`
- `masivo/masivo-reglas.js`: aplicado `'pallet': 'Pallets',`
- `masivo/masivo-reglas.js`: aplicado `'pallets': 'Pallets',`
- `masivo/masivo-reglas.js`: aplicado `const CONTINENTE_PALLETS = 'Pallets';`
- `js/normalizador-1.js`: aplicado `'Contenedor, Bulto, Pallets, Bultos Bigger, Sobre, Bodega, e`
- `index.html`: aplicado `js/normalizador-1.js?v=2.1"`
- `index.html`: aplicado `js/normalizador-2.js?v=2.2"`
- `masivo/index.html`: aplicado `masivo-reglas.js?v=1.1"`

## Despues del cambio

| Caso | Resultado | ok | Continente final | SLA en Contrato | Errores |
|---|---|---|---|---|---|
| Continente "Pallet" con SLA 1 | OK | True | Pallet | SLA PALLETS |  |
| Continente "pallets" (mal escrito) | OK | True | Pallet | SLA PALLETS |  |
| SLA PALLETS con Continente "Bulto/Paquete" | OK | True | Pallet | SLA PALLETS |  |
| Continente "Bultos Bigger" (no cambia) | OK | True | Bultos Bigger | SLA BIGGER |  |
| Continente "Pallets irregulares" (no cambia) | OK | True | Pallets irregulares | SLA 1 |  |

**Resultado final: OK**
