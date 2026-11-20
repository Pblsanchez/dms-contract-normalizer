/**
 * Office Script "LeerAltaContratoDMS" (T-020, Agente Andreani).
 * Lee TODAS las hojas del Excel de alta de contrato DMS y devuelve su contenido como texto JSON.
 * No modifica el archivo. Lo usa el flujo "T-020 Leer alta de contrato" para pasarle los datos al agente.
 */
function main(workbook: ExcelScript.Workbook): string {
  const hojas: { nombre: string; valores: (string | number | boolean)[][]; fechas: number[][] }[] = [];
  for (const ws of workbook.getWorksheets()) {
    const usado = ws.getUsedRange(true);
    if (!usado) {
      hojas.push({ nombre: ws.getName(), valores: [], fechas: [] });
      continue;
    }
    const filas = usado.getRowIndex() + usado.getRowCount();
    const columnas = usado.getColumnIndex() + usado.getColumnCount();
    const rango = ws.getRangeByIndexes(0, 0, filas, columnas);
    const valores = rango.getValues() as (string | number | boolean)[][];
    const formatos = rango.getNumberFormats();
    const fechas: number[][] = [];
    for (let r = 0; r < valores.length; r++) {
      for (let c = 0; c < valores[r].length; c++) {
        const f = String(formatos[r][c]).toLowerCase();
        if (typeof valores[r][c] === "number" && f !== "general" && /[dy]/.test(f)) {
          fechas.push([r, c]);
        }
      }
    }
    hojas.push({ nombre: ws.getName(), valores: valores, fechas: fechas });
  }
  return JSON.stringify({ version: 1, hojas: hojas });
}
