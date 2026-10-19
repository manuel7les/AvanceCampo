/**
 * Avance de Campo — código para la Hoja de Google.
 * Pegar en Extensiones → Apps Script, guardar y publicar como "Aplicación web"
 * (Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario).
 *
 * Pestañas que crea solo:
 *  - Reportes: una fila por cada reporte de los auxiliares.
 *  - Labores:  lista de labores y su unidad (puedes editarla).
 *  - Lotes:    lista de lotes (puedes editarla).
 */

const COLS = ['ID', 'Fecha', 'Auxiliar', 'Labor', 'Lote', 'Cantidad', 'Unidad', 'Nota',
              'Registrado', 'Última edición', 'Dispositivo'];

function doGet(e) {
  let p = {};
  try { p = JSON.parse((e && e.parameter && e.parameter.q) || '{}'); } catch (err) {}
  return salida(atender(p));
}

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) {}
  return salida(atender(p));
}

function salida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function atender(p) {
  try {
    preparar();
    switch (p.accion) {
      case 'config':   return Object.assign({ ok: true }, leerConfig());
      case 'mis':      return { ok: true, reportes: misReportes(p.dispositivo) };
      case 'guardar':  return guardar(p);
      case 'eliminar': return eliminar(p);
      default:         return { ok: true, app: 'AvanceCampo' };
    }
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

/** Crea las pestañas si faltan. También puedes ejecutarla a mano una vez. */
function preparar() {
  const ss = SpreadsheetApp.getActive();
  let rep = ss.getSheetByName('Reportes');
  if (!rep) {
    rep = ss.insertSheet('Reportes', 0);
    rep.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold')
      .setBackground('#2E6B3F').setFontColor('#FFFFFF');
    rep.setFrozenRows(1);
    rep.getRange('B:B').setNumberFormat('dd/mm/yyyy');
    rep.getRange('I:J').setNumberFormat('dd/mm/yyyy hh:mm');
    rep.hideColumns(1); rep.hideColumns(11);
    const hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
    if (hoja1 && hoja1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(hoja1);
  }
  if (!ss.getSheetByName('Labores')) {
    const sh = ss.insertSheet('Labores');
    sh.getRange(1, 1, 7, 2).setValues([
      ['Labor', 'Unidad'],
      ['Poda', 'plantas'], ['Riego', 'ha'], ['Deshierbo', 'ha'],
      ['Fumigación', 'ha'], ['Fertilización', 'ha'], ['Cosecha', 'jabas']
    ]);
    sh.getRange(1, 1, 1, 2).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  if (!ss.getSheetByName('Lotes')) {
    const sh = ss.insertSheet('Lotes');
    sh.getRange(1, 1, 5, 1).setValues([['Lote'], ['Lote 1'], ['Lote 2'], ['Lote 3'], ['Lote 4']]);
    sh.getRange(1, 1, 1, 1).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
}

function leerConfig() {
  const ss = SpreadsheetApp.getActive();
  const lab = ss.getSheetByName('Labores').getDataRange().getValues().slice(1)
    .filter(r => String(r[0]).trim())
    .map(r => ({ nombre: String(r[0]).trim(), unidad: String(r[1] || '').trim() || '%' }));
  const lotes = ss.getSheetByName('Lotes').getDataRange().getValues().slice(1)
    .map(r => String(r[0]).trim()).filter(Boolean);
  return { labores: lab, lotes: lotes };
}

function texto(v) { return v == null ? '' : String(v).slice(0, 300); }

function fechaTexto(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
}

function aFecha(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  if (!m) throw new Error('Fecha no válida');
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function misReportes(dispositivo) {
  if (!dispositivo) return [];
  const sh = SpreadsheetApp.getActive().getSheetByName('Reportes');
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const desde = new Date(Date.now() - 60 * 24 * 3600 * 1000); // últimos 60 días
  return sh.getRange(2, 1, n, COLS.length).getValues()
    .filter(r => String(r[10]) === String(dispositivo) && (!(r[1] instanceof Date) || r[1] >= desde))
    .map(r => ({
      id: String(r[0]), fecha: fechaTexto(r[1]), auxiliar: String(r[2]),
      labor: String(r[3]), lote: String(r[4]), cantidad: Number(r[5]) || 0,
      unidad: String(r[6]), nota: String(r[7] || ''),
      creado: r[8] instanceof Date ? r[8].getTime() : 0,
      editado: r[9] instanceof Date ? r[9].getTime() : 0
    }));
}

function filaDe(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function guardar(p) {
  const r = p.reporte || {};
  if (!r.id || !p.dispositivo) throw new Error('Faltan datos');
  const cantidad = Number(r.cantidad);
  if (!(cantidad > 0)) throw new Error('Cantidad no válida');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName('Reportes');
    const fila = filaDe(sh, r.id);
    const ahora = new Date();
    const valores = [String(r.id), aFecha(r.fecha), texto(p.auxiliar), texto(r.labor), texto(r.lote),
                     cantidad, texto(r.unidad), texto(r.nota), ahora, '', String(p.dispositivo)];
    if (fila > 0) {
      const actual = sh.getRange(fila, 1, 1, COLS.length).getValues()[0];
      if (String(actual[10]) !== String(p.dispositivo)) throw new Error('Este reporte es de otro auxiliar');
      valores[8] = actual[8];
      valores[9] = ahora;
      sh.getRange(fila, 1, 1, COLS.length).setValues([valores]);
    } else {
      sh.appendRow(valores);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function eliminar(p) {
  if (!p.id || !p.dispositivo) throw new Error('Faltan datos');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName('Reportes');
    const fila = filaDe(sh, p.id);
    if (fila < 0) return { ok: true };
    if (String(sh.getRange(fila, 11).getValue()) !== String(p.dispositivo)) throw new Error('Este reporte es de otro auxiliar');
    sh.deleteRow(fila);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}
