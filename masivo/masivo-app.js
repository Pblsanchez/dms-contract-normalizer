/* ==MASIVO-START== Carga masiva por carpeta (T-014, migración Cruz del Sur).
   Lee los .xlsx de una carpeta de a uno, aplica las mismas reglas que el normalizador (REGLAS, copia textual),
   escribe cada archivo normalizado directo en la carpeta de salida (File System Access API, Chrome/Edge)
   y deja un Excel de resumen. Reanudable: saltea lo que ya está en la carpeta de salida. Detecta contratos duplicados.
   Los archivos no salen de la PC: todo ocurre en el navegador. */
var MASIVO = (function () {
  "use strict";
  var RX_SALIDA = /^Alta de contrato DMS (\d+) /;
  var DESTINATARIO = "Hola @Djirikian Damian Uriel";

  /* ---------- Lógica (sin DOM, testeable en Node con XLSX simulado) ---------- */
  function hoja(wb, nombre) { var s = wb.Sheets[nombre]; return s ? XLSX.utils.sheet_to_json(s, { header: 1, defval: "" }) : []; }
  function datosFormulario(wb) {
    var d = hoja(wb, "Formulario"), v = function (r) { return String((d[r] && d[r][2]) || "").trim(); };
    return { contrato: v(2), descripcion: v(3), solicitante: v(4), razon: REGLAS.limpiarTextoEnvolvente(v(13)) };
  }
  /* Mismo flujo que el botón «Procesar y Normalizar» del normalizador, para UN archivo */
  function procesar(wb, nombreOrigen) {
    var lineas = [];
    REGLAS.setLog(function (m, t) { lineas.push({ m: String(m).replace(/^\s*→\s*/, "").trim(), t: t || "info" }); });
    var base = datosFormulario(wb);
    var r = { origen: nombreOrigen, estado: "", errores: [], correcciones: [], salida: "", datos: base, wbout: null };
    try {
      var val = REGLAS.validateFormulario(wb, nombreOrigen);
      if (val.hasErrors) { r.estado = "Con error"; r.errores = val.errors.slice(); }
      else {
        var nd = REGLAS.getNombreArchivoData(wb);
        if (nd.error) { r.estado = "Con error"; r.errores = [nd.error]; }
        else {
          var nwb = REGLAS.normalizeWorkbook(wb);
          var i = nwb.SheetNames.indexOf("Formulario");
          if (i !== -1) { nwb.SheetNames.splice(i, 1); delete nwb.Sheets["Formulario"]; }
          var j = nwb.SheetNames.indexOf("Referencias");
          if (j !== -1) nwb.SheetNames.push(nwb.SheetNames.splice(j, 1)[0]);
          REGLAS.limpiarDireccionesVacias(nwb);
          var nombre = "Alta de contrato DMS " + nd.codigoERP + " " + nd.codigoSolicitante + (nd.clientName ? " " + nd.clientName : "") + ".xlsx";
          r.estado = "OK"; r.salida = nombre; r.contrato = nd.codigoERP; r.wb = nwb;
        }
      }
    } catch (e) { r.estado = "Con error"; r.errores = ["No se pudo leer el archivo: " + (e && e.message ? e.message : e)]; }
    r.correcciones = lineas.filter(function (l) { return /normalizado|auto-corregido|limpiado|eliminaron|forzado/i.test(l.m); })
      .map(function (l) { return l.m.replace(/^[⚠️✅\s]+/u, ""); });
    REGLAS.setLog(null);
    return r;
  }
  /* Estado contra lo ya escrito: «Ya procesado» si el mismo archivo de salida existe; «Duplicado» si el contrato ERP ya salió con otro nombre */
  function clasificar(r, existentes, porContrato) {
    if (r.estado !== "OK") return r;
    if (existentes[r.salida]) { r.estado = "Ya procesado"; return r; }
    var prev = porContrato[r.contrato];
    if (prev) { r.estado = "Duplicado"; r.salida = ""; r.errores = ["Contrato " + r.contrato + " ya generado desde «" + prev + "»"]; return r; }
    return r;
  }
  function indexarSalida(nombres) {
    var existentes = {}, porContrato = {};
    nombres.forEach(function (n) { existentes[n] = 1; var m = RX_SALIDA.exec(n); if (m) porContrato[m[1]] = n; });
    return { existentes: existentes, porContrato: porContrato };
  }
  var COLS = ["N°", "Archivo origen", "Estado", "Contrato ERP", "Código cliente", "Razón social", "Descripción contrato", "Archivo generado", "Errores", "Correcciones automáticas"];
  function filasResumen(res) {
    return res.map(function (r, i) {
      return [i + 1, r.origen, r.estado, r.datos.contrato, r.datos.solicitante, r.datos.razon, r.datos.descripcion, r.salida, r.errores.join(" | "), r.correcciones.join(" | ")];
    });
  }
  function contar(res) {
    var c = { total: res.length, OK: 0, "Ya procesado": 0, "Con error": 0, Duplicado: 0 };
    res.forEach(function (r) { c[r.estado] = (c[r.estado] || 0) + 1; });
    return c;
  }
  /* Para Damián: solo el aviso de archivos nuevos en la carpeta compartida (el resumen es para Pablo / referente CDS) */
  function mensaje(n, carpeta) {
    return DESTINATARIO + ", hay " + n + (n === 1 ? " archivo nuevo" : " archivos nuevos") + " para alta de contratos DMS (migración Cruz del Sur)" +
      (carpeta ? " en la carpeta «" + carpeta + "»" : "") + ".";
  }
  /* OneDrive «solo en la nube»: al leer, Windows baja el archivo y puede fallar la 1.ª lectura (InvalidStateError / NotReadableError). Reintenta. */
  async function leerConReintento(handle) {
    var ultimo;
    for (var k = 0; k < 4; k++) {
      try { var f = await handle.getFile(); return { name: f.name, buf: await f.arrayBuffer() }; }
      catch (e) { ultimo = e; await new Promise(function (ok) { setTimeout(ok, 1500 * (k + 1)); }); }
    }
    throw ultimo;
  }
  function sello(d) { var p = function (x) { return ("0" + x).slice(-2); }; return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()) + " " + p(d.getHours()) + "-" + p(d.getMinutes()); }

  /* ---------- Vista ---------- */
  var $ = function (id) { return document.getElementById(id); };
  var dirIn = null, dirOut = null, archivos = [], resultados = [], detener = false, corriendo = false, ultimoResumen = null;
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  async function listarXlsx(dir) {
    var out = [];
    for await (var entry of dir.values()) {
      if (entry.kind === "file" && /\.xlsx$/i.test(entry.name) && entry.name.indexOf("~$") !== 0) out.push(entry);
    }
    out.sort(function (a, b) { return a.name.localeCompare(b.name, "es"); });
    return out;
  }
  async function nombresSalida() {
    var n = [];
    for await (var entry of dirOut.values()) if (entry.kind === "file") n.push(entry.name);
    return n;
  }
  async function escribir(nombre, datos) {
    var fh = await dirOut.getFileHandle(nombre, { create: true });
    var w = await fh.createWritable(); await w.write(datos); await w.close();
  }
  function pintarCarpetas() {
    $("in-nombre").textContent = dirIn ? dirIn.name + " · " + archivos.length + " archivo" + (archivos.length === 1 ? "" : "s") + " .xlsx" : "Sin elegir";
    $("out-nombre").textContent = dirOut ? dirOut.name : "Sin elegir";
    $("btn-procesar").disabled = !(dirIn && dirOut && archivos.length) || corriendo;
    if (dirIn && dirOut && dirIn.name === dirOut.name) aviso("La carpeta de salida tiene el mismo nombre que la de origen. Conviene usar una carpeta distinta para no mezclar originales y normalizados.", "warning");
  }
  function aviso(t, tipo) { var a = $("aviso"); a.textContent = t; a.className = "ms-aviso ms-aviso--" + (tipo || "info"); a.hidden = !t; }
  var BADGE = { OK: "ok", "Ya procesado": "neutro", "Con error": "error", Duplicado: "warning" };
  function pintarFila(r, i) {
    var tr = document.createElement("tr");
    tr.innerHTML = "<td class=\"n\">" + (i + 1) + "</td><td>" + esc(r.origen) + "</td><td><span class=\"ms-badge ms-badge--" + BADGE[r.estado] + "\">" + esc(r.estado) + "</span></td>" +
      "<td>" + esc(r.datos.contrato) + "</td><td>" + esc(r.salida || "—") + "</td><td>" + esc(r.errores.join(" · ")) + "</td>";
    $("tabla-cuerpo").appendChild(tr);
  }
  function pintarContadores() {
    var c = contar(resultados);
    $("c-total").textContent = c.total + " / " + archivos.length; $("c-ok").textContent = c.OK; $("c-ya").textContent = c["Ya procesado"];
    $("c-err").textContent = c["Con error"]; $("c-dup").textContent = c.Duplicado;
    var p = archivos.length ? Math.round(c.total / archivos.length * 100) : 0;
    $("barra").style.width = p + "%"; $("barra-wrap").setAttribute("aria-valuenow", String(p));
  }
  async function elegir(modo) {
    if (!window.showDirectoryPicker) { aviso("Este navegador no permite elegir carpetas. Usá Google Chrome o Microsoft Edge en la computadora.", "error"); return; }
    try {
      var d = await window.showDirectoryPicker({ mode: modo === "in" ? "read" : "readwrite", id: "cds-" + modo });
      if (modo === "in") { dirIn = d; archivos = await listarXlsx(d); } else dirOut = d;
      aviso("", ""); pintarCarpetas();
    } catch (e) { if (e && e.name !== "AbortError") aviso("No se pudo abrir la carpeta: " + e.message, "error"); }
  }
  async function correr() {
    corriendo = true; detener = false; resultados = []; $("tabla-cuerpo").innerHTML = ""; $("resultados").hidden = false;
    $("btn-procesar").disabled = true; $("btn-detener").hidden = false; $("btn-resumen").disabled = true; $("btn-copiar").disabled = true;
    aviso("Procesando… no cierres esta pestaña. Podés detener cuando quieras y volver a correr: retoma donde quedó.", "info");
    var idx = indexarSalida(await nombresSalida()), nuevos = 0;
    for (var i = 0; i < archivos.length && !detener; i++) {
      $("estado-actual").textContent = "Procesando " + (i + 1) + " de " + archivos.length + ": " + archivos[i].name;
      var r = null;
      try {
        var file = await leerConReintento(archivos[i]);
        var wb = XLSX.read(file.buf, { type: "array", cellDates: false });
        r = clasificar(procesar(wb, file.name), idx.existentes, idx.porContrato);
        if (r.estado === "OK") {
          await escribir(r.salida, XLSX.write(r.wb, { bookType: "xlsx", type: "array" }));
          idx.existentes[r.salida] = 1; idx.porContrato[r.contrato] = r.salida; nuevos++;
        }
      } catch (e) {
        r = r || { origen: archivos[i].name, datos: {}, salida: "", correcciones: [], errores: [] };
        r.estado = "Con error"; r.errores = ["No se pudo procesar o escribir: " + (e && e.message ? e.message : e)]; r.salida = "";
      }
      delete r.wb; resultados.push(r); pintarFila(r, i); pintarContadores();
      await new Promise(function (ok) { setTimeout(ok, 0); });
    }
    var c = contar(resultados);
    $("estado-actual").textContent = (detener ? "Detenido. " : "Terminado. ") + c.OK + " generados · " + c["Ya procesado"] + " ya procesados · " + c["Con error"] + " con error · " + c.Duplicado + " duplicados.";
    try { await guardarResumen(); } catch (e) { aviso("Los archivos se generaron, pero no se pudo guardar el resumen en la carpeta: " + e.message + ". Usá «Descargar resumen».", "warning"); }
    $("mensaje").textContent = nuevos ? mensaje(nuevos, dirOut.name) : "";
    $("btn-copiar").disabled = !nuevos; $("btn-resumen").disabled = false; $("btn-detener").hidden = true;
    corriendo = false; pintarCarpetas();
    if (!detener) aviso(c["Con error"] ? "Corregí los archivos con error en la carpeta de origen y volvé a procesar: solo se generan los pendientes." : "Listo. Revisá el resumen en la carpeta de salida.", c["Con error"] ? "warning" : "success");
  }
  function libroResumen() {
    var ws = XLSX.utils.aoa_to_sheet([COLS].concat(filasResumen(resultados)));
    ws["!cols"] = [6, 40, 14, 14, 14, 30, 34, 60, 70, 70].map(function (w) { return { wch: w }; });
    ws["!autofilter"] = { ref: "A1:J" + (resultados.length + 1) };
    var wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Resumen");
    return wb;
  }
  async function guardarResumen() {
    ultimoResumen = "Resumen carga masiva " + sello(new Date()) + ".xlsx";
    await escribir(ultimoResumen, XLSX.write(libroResumen(), { bookType: "xlsx", type: "array" }));
  }
  function descargarResumen() { XLSX.writeFile(libroResumen(), ultimoResumen || ("Resumen carga masiva " + sello(new Date()) + ".xlsx")); }
  async function copiar() {
    var t = $("mensaje").textContent;
    try { await navigator.clipboard.writeText(t); $("btn-copiar").textContent = "Copiado"; setTimeout(function () { $("btn-copiar").textContent = "Copiar mensaje"; }, 2500); }
    catch (e) { aviso("No se pudo copiar automáticamente. Seleccioná el texto y copialo.", "warning"); }
  }
  function iniciar() {
    if (typeof document === "undefined" || !$("btn-in")) return;
    if (!window.showDirectoryPicker) aviso("Este navegador no permite elegir carpetas. Usá Google Chrome o Microsoft Edge en la computadora.", "error");
    if (typeof XLSX === "undefined") aviso("No se pudo cargar la librería de Excel (necesita internet). Recargá la página.", "error");
    $("btn-in").addEventListener("click", function () { elegir("in"); });
    $("btn-out").addEventListener("click", function () { elegir("out"); });
    $("btn-procesar").addEventListener("click", correr);
    $("btn-detener").addEventListener("click", function () { detener = true; $("estado-actual").textContent = "Deteniendo al terminar el archivo actual…"; });
    $("btn-resumen").addEventListener("click", descargarResumen);
    $("btn-copiar").addEventListener("click", copiar);
    window.addEventListener("beforeunload", function (e) { if (corriendo) { e.preventDefault(); e.returnValue = ""; } });
    pintarCarpetas();
  }
  iniciar();
  return { leerConReintento: leerConReintento, procesar: procesar, clasificar: clasificar, indexarSalida: indexarSalida, filasResumen: filasResumen, contar: contar, mensaje: mensaje, COLS: COLS };
})();
if (typeof module !== "undefined" && module.exports) module.exports = MASIVO;
/* ==MASIVO-END== */
