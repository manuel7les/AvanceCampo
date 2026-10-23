/**
 * Avance de Campo — código para la Hoja de Google.
 * Pegar en Extensiones → Apps Script, guardar y publicar como "Aplicación web"
 * (Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario).
 *
 * Pestañas que crea solo:
 *  - Reportes: una fila por cada reporte de los auxiliares.
 *  - Labores:  Labor | Unidad | Por válvulas (Sí/No).
 *  - Lotes:    Lote | Válvulas (ej.: 35-41  o  35,36,37,40).
 */

const COLS = ['ID', 'Fecha', 'Auxiliar', 'Labor', 'Lote', 'Cantidad', 'Unidad',
              'Hora inicio', 'Hora fin', 'Jornales inicio', 'Ingresos/retiros de jornales',
              'Jornales final', 'Horas-hombre',
              'Observaciones (válvulas)', 'Nota', 'Registrado', 'Última edición',
              'Dispositivo', 'Datos válvulas', 'Datos jornales'];
const OCULTAS = ['ID', 'Dispositivo', 'Datos válvulas', 'Datos jornales'];
const COLS_TEXTO = ['Hora inicio', 'Hora fin', 'Datos jornales', 'Datos válvulas'];

const LOTES = ['L1A', 'L2A', 'L3A', 'L4A', 'L5A', 'L6A', 'L7A', 'L8A', 'L9A', 'L10A', 'L11A',
               'L12A', 'L12B', 'L13A', 'L13B', 'L14A', 'L15A', 'L16A'];

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

/** Crea o completa las pestañas. También puedes ejecutarla a mano una vez. */
function preparar() {
  const ss = SpreadsheetApp.getActive();
  let rep = ss.getSheetByName('Reportes');
  if (!rep) {
    rep = ss.insertSheet('Reportes', 0);
    rep.getRange(1, 1, 1, COLS.length).setValues([COLS]);
    rep.setFrozenRows(1);
    const hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
    if (hoja1 && hoja1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(hoja1);
  }
  // agrega columnas que falten (si la hoja se creó con una versión anterior)
  const cab = cabecera(rep);
  const faltan = COLS.filter(c => !(c in cab));
  if (faltan.length) {
    const desde = rep.getLastColumn() + 1;
    rep.getRange(1, desde, 1, faltan.length).setValues([faltan]);
  }
  if (faltan.length || rep.getRange(1, 1).getFontWeight() !== 'bold') {
    const c = cabecera(rep);
    rep.getRange(1, 1, 1, rep.getLastColumn()).setFontWeight('bold')
      .setBackground('#2E6B3F').setFontColor('#FFFFFF');
    rep.getRange(2, c['Fecha'] + 1, rep.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy');
    rep.getRange(2, c['Registrado'] + 1, rep.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    rep.getRange(2, c['Última edición'] + 1, rep.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    rep.setColumnWidth(c['Observaciones (válvulas)'] + 1, 320);
    rep.setColumnWidth(c['Ingresos/retiros de jornales'] + 1, 260);
    COLS_TEXTO.forEach(n => rep.getRange(2, c[n] + 1, rep.getMaxRows() - 1, 1).setNumberFormat('@'));
    OCULTAS.forEach(n => rep.hideColumns(c[n] + 1));
  }

  let lab = ss.getSheetByName('Labores');
  if (!lab) {
    lab = ss.insertSheet('Labores');
    lab.getRange(1, 1, 5, 3).setValues([
      ['Labor', 'Unidad', 'Por válvulas'],
      ['Deshierbo manual', 'plantas', 'Sí'], ['Deshierbo a lampa', 'plantas', 'Sí'],
      ['Encamado', 'm lineales', 'No'], ['Cosecha', 'jabas', 'No']
    ]);
    lab.getRange(1, 1, 1, 3).setFontWeight('bold');
    lab.setFrozenRows(1);
  } else if (!String(lab.getRange(1, 3).getValue()).trim()) {
    lab.getRange(1, 3).setValue('Por válvulas').setFontWeight('bold');
  }

  let lot = ss.getSheetByName('Lotes');
  if (!lot) {
    lot = ss.insertSheet('Lotes');
    lot.getRange('B:B').setNumberFormat('@'); // texto, para que "35-41" no se vuelva fecha
    const filas = [['Lote', 'Válvulas']].concat(LOTES.map(l => [l, l === 'L8A' ? '35-41' : '']));
    lot.getRange(1, 1, filas.length, 2).setValues(filas);
    lot.getRange(1, 1, 1, 2).setFontWeight('bold');
    lot.setFrozenRows(1);
    lot.setColumnWidth(2, 220);
  } else if (!String(lot.getRange(1, 2).getValue()).trim()) {
    lot.getRange('B:B').setNumberFormat('@');
    lot.getRange(1, 2).setValue('Válvulas').setFontWeight('bold');
  }
}

function cabecera(sh) {
  const n = Math.max(sh.getLastColumn(), 1);
  const out = {};
  sh.getRange(1, 1, 1, n).getValues()[0].forEach((v, i) => { if (v !== '') out[String(v)] = i; });
  return out;
}

/** "35-41" → 35..41 · "35,36,40" → lista · "35-38, 40" → mezcla */
function leerValvulas(v) {
  const out = [];
  String(v == null ? '' : v).split(/[,;\s]+/).forEach(parte => {
    const p = parte.trim();
    if (!p) return;
    const m = /^(\d+)\s*[-–]\s*(\d+)$/.exec(p);
    if (m) {
      const a = Number(m[1]), b = Number(m[2]);
      if (b >= a && b - a < 300) for (let i = a; i <= b; i++) out.push(String(i));
    } else {
      out.push(p);
    }
  });
  return out.filter((x, i) => out.indexOf(x) === i);
}

function leerConfig() {
  const ss = SpreadsheetApp.getActive();
  const labores = ss.getSheetByName('Labores').getDataRange().getValues().slice(1)
    .filter(r => String(r[0]).trim())
    .map(r => ({
      nombre: String(r[0]).trim(),
      unidad: String(r[1] || '').trim() || '%',
      valvulas: /^s[ií]/i.test(String(r[2] || '').trim())
    }));
  const lotes = ss.getSheetByName('Lotes').getDataRange().getDisplayValues().slice(1)
    .filter(r => String(r[0]).trim())
    .map(r => ({ nombre: String(r[0]).trim(), valvulas: leerValvulas(r[1]) }));
  return { labores: labores, lotes: lotes };
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

/** "35=T;36=M;37=S" → "Terminadas: 35–36 · A medias: 37 · Sin tocar: 38–41" */
function rango(list) {
  const out = [];
  let i = 0;
  while (i < list.length) {
    let j = i;
    while (j + 1 < list.length && /^\d+$/.test(list[j]) && /^\d+$/.test(list[j + 1]) &&
           Number(list[j + 1]) === Number(list[j]) + 1) j++;
    out.push(j - i >= 2 ? list[i] + '–' + list[j] : list.slice(i, j + 1).join(', '));
    i = j + 1;
  }
  return out.join(', ');
}

function observacion(datos) {
  const pares = String(datos || '').split(';').filter(Boolean).map(p => p.split('='));
  if (!pares.length) return '';
  const de = k => pares.filter(p => p[1] === k).map(p => p[0]);
  const partes = [];
  const t = de('T'), m = de('M'), s = de('S');
  if (t.length) partes.push('Terminadas: ' + rango(t));
  if (m.length) partes.push('A medias: ' + rango(m));
  if (s.length) partes.push('Sin tocar: ' + rango(s));
  return partes.join(' · ');
}

/* ---------- Horario y jornales ---------- */
function horaTexto(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  const m = /^(\d{1,2}):(\d{2})/.exec(String(v || ''));
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : '';
}
function aMin(h) { const m = /^(\d{1,2}):(\d{2})$/.exec(h || ''); return m ? Number(m[1]) * 60 + Number(m[2]) : null; }
function leerMovs(str) {
  return String(str || '').split(';').map(x => /^([IR])(\d+)@(\d{1,2}:\d{2})$/.exec(x.trim())).filter(Boolean)
    .map(m => ({ t: m[1], n: Number(m[2]), h: horaTexto(m[3]) }))
    .sort((a, b) => aMin(a.h) - aMin(b.h));
}
/** Calcula jornales al final, horas-hombre y el texto de movimientos. */
function calcularJornales(hi, hf, jor, movs) {
  const a = aMin(hi), z = aMin(hf);
  let n = Number(jor) || 0, t = a, hh = 0;
  movs.forEach(m => {
    const mm = aMin(m.h);
    if (t != null && mm >= t) { hh += n * (mm - t); t = mm; }
    n = Math.max(0, n + (m.t === 'I' ? m.n : -m.n));
  });
  const horas = (a != null && z != null && z > a && t != null) ? Math.round((hh + n * (z - t)) / 60 * 100) / 100 : '';
  const texto = movs.map(m => (m.t === 'I' ? 'Ingresan ' : 'Salen ') + m.n + ' a las ' + m.h).join(' · ');
  return { final: n, horas: horas, texto: texto };
}

function misReportes(dispositivo) {
  if (!dispositivo) return [];
  const sh = SpreadsheetApp.getActive().getSheetByName('Reportes');
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const c = cabecera(sh);
  const desde = new Date(Date.now() - 60 * 24 * 3600 * 1000); // últimos 60 días
  return sh.getRange(2, 1, n, sh.getLastColumn()).getValues()
    .filter(r => String(r[c['Dispositivo']]) === String(dispositivo) &&
                 (!(r[c['Fecha']] instanceof Date) || r[c['Fecha']] >= desde))
    .map(r => ({
      id: String(r[c['ID']]), fecha: fechaTexto(r[c['Fecha']]), auxiliar: String(r[c['Auxiliar']]),
      labor: String(r[c['Labor']]), lote: String(r[c['Lote']]), cantidad: Number(r[c['Cantidad']]) || 0,
      unidad: String(r[c['Unidad']]), nota: String(r[c['Nota']] || ''),
      valvulas: String(r[c['Datos válvulas']] || ''),
      hi: horaTexto(r[c['Hora inicio']]), hf: horaTexto(r[c['Hora fin']]),
      jor: r[c['Jornales inicio']] === '' ? '' : Number(r[c['Jornales inicio']]),
      movs: String(r[c['Datos jornales']] || ''),
      creado: r[c['Registrado']] instanceof Date ? r[c['Registrado']].getTime() : 0,
      editado: r[c['Última edición']] instanceof Date ? r[c['Última edición']].getTime() : 0
    }));
}

function filaDe(sh, col, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, col + 1, n, 1).getValues();
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
    const c = cabecera(sh);
    const ancho = sh.getLastColumn();
    const fila = filaDe(sh, c['ID'], r.id);
    const ahora = new Date();
    let fil = new Array(ancho).fill('');
    if (fila > 0) {
      fil = sh.getRange(fila, 1, 1, ancho).getValues()[0];
      if (String(fil[c['Dispositivo']]) !== String(p.dispositivo)) throw new Error('Este reporte es de otro auxiliar');
      fil[c['Última edición']] = ahora;
    } else {
      fil[c['Registrado']] = ahora;
    }
    const datos = texto(r.valvulas).slice(0, 2000);
    fil[c['ID']] = String(r.id);
    fil[c['Fecha']] = aFecha(r.fecha);
    fil[c['Auxiliar']] = texto(p.auxiliar);
    fil[c['Labor']] = texto(r.labor);
    fil[c['Lote']] = texto(r.lote);
    fil[c['Cantidad']] = cantidad;
    fil[c['Unidad']] = texto(r.unidad);
    fil[c['Observaciones (válvulas)']] = observacion(datos);
    fil[c['Nota']] = texto(r.nota);
    fil[c['Dispositivo']] = String(p.dispositivo);
    fil[c['Datos válvulas']] = datos;
    const hi = horaTexto(r.hi), hf = horaTexto(r.hf);
    const movs = leerMovs(r.movs);
    const tieneJor = String(r.jor == null ? '' : r.jor).trim() !== '';
    const jc = calcularJornales(hi, hf, r.jor, movs);
    fil[c['Hora inicio']] = hi;
    fil[c['Hora fin']] = hf;
    fil[c['Jornales inicio']] = tieneJor ? Number(r.jor) : '';
    fil[c['Ingresos/retiros de jornales']] = jc.texto;
    fil[c['Jornales final']] = tieneJor ? jc.final : '';
    fil[c['Horas-hombre']] = tieneJor ? jc.horas : '';
    fil[c['Datos jornales']] = movs.map(m => m.t + m.n + '@' + m.h).join(';');
    if (fila > 0) sh.getRange(fila, 1, 1, ancho).setValues([fil]);
    else sh.appendRow(fil);
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
    const c = cabecera(sh);
    const fila = filaDe(sh, c['ID'], p.id);
    if (fila < 0) return { ok: true };
    if (String(sh.getRange(fila, c['Dispositivo'] + 1).getValue()) !== String(p.dispositivo)) {
      throw new Error('Este reporte es de otro auxiliar');
    }
    sh.deleteRow(fila);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}
