/* entreno-core.js — lógica compartida de la planilla ENTRENO 25 (la usan panel.html y mobile.html).
   Solo depende de los worksheets de ExcelJS; no toca el DOM ni el almacenamiento.

   Modelo de la hoja del año:
   - Cada semana es un "bloque": una fila de encabezado (A = "DÍA") seguida de los grupos
     Día 1, Día 2, Día 3, Día 4 y MOVILIDAD.
   - Si en una misma semana se repite un día (ej.: dos veces Día 1 porque se salteó pierna),
     la repetición va como un grupo "extra" al final del mismo bloque (mismo DÍA/GRUPO, copia de
     las filas de la plantilla), así todo consta en la misma semana. */

const COLS = { dia:1, fecha:2, grupo:3, orden:4, ejercicio:5, peso:6, series:7, reps:8, mrep:9, mkg:10, rm:11, ss:12, fallo:13, notas:14, polea1:15, polea2:16, ayuno:17 };
const COLLETTERS = {1:'A',2:'B',3:'C',4:'D',5:'E',6:'F',7:'G',8:'H',9:'I',10:'J',11:'K',12:'L',13:'M',14:'N',15:'O',16:'P',17:'Q'};
const MOVILIDAD = 'MOVILIDAD';
const SESIONES_OBJETIVO = 4;   // objetivo: 4 entrenamientos por semana
const DIAS_MAX_SEMANA = 6;     // una repetición dentro de los 6 días del primer entreno cuenta como "misma semana"

// ---------- lectura de la estructura ----------
function nombreEjercicio(ws, r) { return ws.getCell(r, COLS.ejercicio).text.trim(); }

function parseStructure(ws) {
  const rowCount = ws.rowCount;
  const headerRows = [];
  for (let r = 1; r <= rowCount; r++) {
    if (ws.getCell(r, COLS.dia).text === 'DÍA') headerRows.push(r);
  }
  const weekBlocks = [];
  for (let i = 0; i < headerRows.length; i++) {
    const start = headerRows[i] + 1;
    const end = (i + 1 < headerRows.length) ? headerRows[i+1] - 1 : rowCount;
    if (start > end) { weekBlocks.push({ headerRow: headerRows[i], start, end, groups: [] }); continue; }
    const groups = [];
    let gStart = start;
    let gGrupo = ws.getCell(start, COLS.grupo).text;
    for (let r = start + 1; r <= end + 1; r++) {
      const grupo = (r <= end) ? ws.getCell(r, COLS.grupo).text : null;
      // un grupo también termina cuando vuelve a aparecer su primer ejercicio: es una
      // repetición del mismo día pegada a la anterior (sesión extra)
      const primero = (r <= end && grupo === gGrupo && grupo !== MOVILIDAD) ? nombreEjercicio(ws, gStart) : '';
      const repiteInicio = primero !== '' && nombreEjercicio(ws, r) === primero;
      if (grupo !== gGrupo || repiteInicio) {
        groups.push({ grupo: gGrupo, dia: ws.getCell(gStart, COLS.dia).text, start: gStart, end: r - 1 });
        gStart = r;
        gGrupo = grupo;
      }
    }
    weekBlocks.push({ headerRow: headerRows[i], start, end, groups });
  }
  return { headerRows, weekBlocks, rowCount };
}
function isGroupBlank(ws, group) {
  for (let r = group.start; r <= group.end; r++) {
    if (ws.getCell(r, COLS.peso).text || ws.getCell(r, COLS.reps).text) return false;
  }
  return true;
}
function _fmtDDMMYYYY(d) {
  return String(d.getUTCDate()).padStart(2,'0') + '/' + String(d.getUTCMonth()+1).padStart(2,'0') + '/' + d.getUTCFullYear();
}
function cellDateDDMMYYYY(cell) {
  const v = cell.value;
  if (!v) return '';
  if (v instanceof Date) return _fmtDDMMYYYY(v);
  // número suelto en una celda de fecha = serie de Excel (pasa con lo recién escrito en memoria)
  if (typeof v === 'number' && v > 20000 && v < 80000) return _fmtDDMMYYYY(new Date(Math.round((v - 25569) * 86400000)));
  return cell.text;
}
function fechaKey(dd_mm_yyyy) {
  if (!dd_mm_yyyy) return null;
  const parts = dd_mm_yyyy.split('/');
  if (parts.length !== 3) return null;
  const [d,m,y] = parts.map(Number);
  if (!d || !m || !y) return null;
  return y * 10000 + m * 100 + d;
}
function _keyToUTC(k) { return Date.UTC(Math.floor(k / 10000), Math.floor(k / 100) % 100 - 1, k % 100); }
function diasEntreKeys(a, b) { return Math.round((_keyToUTC(a) - _keyToUTC(b)) / 86400000); }

function leerEjercicios(ws, g) {
  const out = [];
  for (let r = g.start; r <= g.end; r++) {
    const peso = ws.getCell(r, COLS.peso).text;
    const reps = ws.getCell(r, COLS.reps).text;
    if (!peso && !reps) continue;
    out.push({
      nombre: nombreEjercicio(ws, r), peso, reps,
      orden: ws.getCell(r, COLS.orden).text,
      ss: ws.getCell(r, COLS.ss).text,
      fallo: ws.getCell(r, COLS.fallo).text,
      notas: ws.getCell(r, COLS.notas).text
    });
  }
  return out;
}

// sesiones (entrenamientos) reales de un bloque/semana, en orden cronológico. Un mismo día
// puede aparecer más de una vez (repetición = sesión extra). Los fragmentos de un mismo día
// partidos por una celda GRUPO vacía de más (visto en 2025) se unen a la sesión anterior.
function sesionesDeBloque(wb, ws) {
  const todas = [];
  wb.groups.forEach(g => {
    if (g.grupo === MOVILIDAD) return;
    const dia = Number(g.dia);
    if (![1, 2, 3, 4].includes(dia)) return;
    const fechaRaw = cellDateDDMMYYYY(ws.getCell(g.start, COLS.fecha));
    const ejercicios = leerEjercicios(ws, g);
    let previa = null;
    for (let i = todas.length - 1; i >= 0; i--) if (todas[i].dia === dia) { previa = todas[i]; break; }
    if (previa && (fechaRaw === '' || fechaRaw === previa.fechaRaw)) {
      previa.ejercicios.push(...ejercicios);
      if (!previa.grupo && g.grupo) previa.grupo = g.grupo;
      return;
    }
    todas.push({ dia, grupo: g.grupo, fechaRaw, fecha: fechaRaw, ejercicios });
  });
  const validas = todas.filter(s => s.fechaRaw && fechaKey(s.fechaRaw) !== null && s.ejercicios.length);
  validas.sort((a, b) => (fechaKey(a.fecha) - fechaKey(b.fecha)) || (a.dia - b.dia));
  const vistos = {};
  validas.forEach(s => { s.extra = !!vistos[s.dia]; vistos[s.dia] = true; });
  return validas;
}

function extractAll(structure, ws) {
  const byDia = { 1: [], 2: [], 3: [], 4: [] };
  const exerciseHistory = {};
  const allDates = new Set();

  structure.weekBlocks.forEach(wb => {
    wb.groups.forEach(g => {
      if (g.grupo === MOVILIDAD) return;
      const dia = Number(g.dia);
      if (!dia || !byDia[dia]) return;
      const fecha = cellDateDDMMYYYY(ws.getCell(g.start, COLS.fecha));
      if (!fecha || fechaKey(fecha) === null) return;
      const ejercicios = leerEjercicios(ws, g);
      ejercicios.forEach(e => {
        if (!exerciseHistory[e.nombre]) exerciseHistory[e.nombre] = [];
        exerciseHistory[e.nombre].push({ fecha, peso: e.peso, reps: e.reps, notas: e.notas });
      });
      if (ejercicios.length) {
        byDia[dia].push({ fecha, grupo: g.grupo, ejercicios });
        allDates.add(fecha);
      }
    });
  });

  Object.keys(byDia).forEach(d => byDia[d].sort((a, b) => fechaKey(b.fecha) - fechaKey(a.fecha)));
  Object.keys(exerciseHistory).forEach(n => exerciseHistory[n].sort((a, b) => fechaKey(a.fecha) - fechaKey(b.fecha)));
  return { byDia, exerciseHistory, allDates };
}

// una entrada por semana: sesiones (todas, con repeticiones) y dias[d] = sesiones de ese día
function extractWeeks(structure, ws) {
  const weeks = [];
  structure.weekBlocks.forEach(wb => {
    const sesiones = sesionesDeBloque(wb, ws);
    if (!sesiones.length) return;
    const dias = { 1: [], 2: [], 3: [], 4: [] };
    sesiones.forEach(s => dias[s.dia].push(s));
    const fechas = sesiones.map(s => s.fecha).sort((a, b) => fechaKey(a) - fechaKey(b));
    const label = fechas[0] === fechas[fechas.length - 1] ? fechas[0] : (fechas[0] + ' – ' + fechas[fechas.length - 1]);
    weeks.push({ label, sortKey: fechaKey(fechas[fechas.length - 1]), dias, sesiones });
  });
  weeks.sort((a, b) => b.sortKey - a.sortKey);
  return weeks;
}

// a qué día pertenece cada ejercicio, según la plantilla de la semana más reciente
function getCurrentTemplateByDia(structure, ws) {
  const map = {};
  let idx = structure.weekBlocks.length - 1;
  while (idx >= 0 && structure.weekBlocks[idx].groups.length === 0) idx--;
  if (idx < 0) return map;
  structure.weekBlocks[idx].groups.forEach(g => {
    if (g.grupo === MOVILIDAD) return;
    const dia = Number(g.dia);
    if (!dia) return;
    for (let r = g.start; r <= g.end; r++) {
      const nombre = nombreEjercicio(ws, r);
      if (nombre) map[nombre] = dia;
    }
  });
  return map;
}

// fecha -> sesiones de ese día (normalmente una), para el popover del calendario
function buildSessionByDate(weeks) {
  const map = {};
  weeks.forEach(w => {
    w.sesiones.forEach(s => {
      const key = fechaKey(s.fecha);
      if (key === null) return;
      if (!map[key]) map[key] = [];
      map[key].push({ dia: s.dia, grupo: s.grupo, fecha: s.fecha, ejercicios: s.ejercicios, extra: !!s.extra });
    });
  });
  return map;
}

// resume cada semana en un objeto plano (sin referencias al worksheet), para poder concatenar
// hojas/años y calcular la racha sobre el historial completo
function extractWeekFlags(structure, ws) {
  return structure.weekBlocks.map(wb => {
    const dayGroups = wb.groups.filter(g => g.grupo !== MOVILIDAD);
    const fechas = dayGroups.map(g => cellDateDDMMYYYY(ws.getCell(g.start, COLS.fecha))).filter(f => f && fechaKey(f) !== null);
    fechas.sort((a, b) => fechaKey(a) - fechaKey(b));
    const rango = fechas.length ? (fechas[0] === fechas[fechas.length - 1] ? fechas[0] : (fechas[0] + ' – ' + fechas[fechas.length - 1])) : '(sin fechas)';
    const sesiones = sesionesDeBloque(wb, ws).length;
    return {
      hasDayGroups: dayGroups.length > 0,
      sesiones,
      // semana completa = llegó al objetivo de sesiones (4), sean o no los 4 días distintos
      allFilled: sesiones >= SESIONES_OBJETIVO,
      anyFilled: dayGroups.some(g => !isGroupBlank(ws, g)),
      rango
    };
  });
}
function computeRachaFromFlags(weekFlagsChrono) {
  const blocks = weekFlagsChrono.filter(b => b.hasDayGroups);
  let streak = 0;
  const detalle = [];
  for (let i = blocks.length - 1; i >= 0; i--) {
    const b = blocks[i];
    if (b.allFilled) { streak++; detalle.push({ rango: b.rango, estado: 'completa' }); continue; }
    if (i === blocks.length - 1 && b.anyFilled) { detalle.push({ rango: b.rango, estado: 'en curso' }); continue; }
    detalle.push({ rango: b.rango, estado: b.anyFilled ? 'incompleta' : 'vacía' });
    break;
  }
  return { streak, detalle };
}

// "Retorno 75%" (o una nota suelta "75%") -> "75"; sirve para mostrar el tag de modo retorno
function parseRetornoTag(notas) {
  const t = String(notas || '');
  let m = /retorno\s+(\d+)\s*%/i.exec(t);
  if (m) return m[1];
  m = /^\s*(\d{2})\s*%\s*$/.exec(t);
  return (m && Number(m[1]) < 100) ? m[1] : null;
}

// ---------- escritura ----------
function parseNum(v) {
  if (v === null || v === undefined || v === '') return null;
  return parseFloat(String(v).replace(',', '.'));
}
function toSheetNumberString(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim().replace('.', ',');
}
function fmtDelta(n) {
  if (n === null || Number.isNaN(n)) return '';
  const r = Math.round(n * 100) / 100;
  return toSheetNumberString(r);
}
function dateToDDMMYYYY(iso) {
  const [y,m,d] = iso.split('-');
  return d + '/' + m + '/' + y;
}
function ddmmyyyyToDate(str) {
  const [d,m,y] = str.split('/').map(Number);
  return new Date(Date.UTC(y, m - 1, d));
}
function diaGroups(wb, dia) {
  return wb.groups.filter(g => g.grupo !== MOVILIDAD && g.dia === String(dia));
}
function findDayGroup(wb, dia) { return diaGroups(wb, dia)[0] || null; }
function findMovGroup(wb) { return wb.groups.find(g => g.grupo === MOVILIDAD) || null; }

// último valor cargado de un ejercicio ANTES de la fila indicada (la sesión más reciente
// anterior del mismo día, incluidas las repeticiones extra)
function findPrevValue(structure, ws, beforeRow, exerciseName, isMov, dia) {
  const name = exerciseName.trim();
  for (let b = structure.weekBlocks.length - 1; b >= 0; b--) {
    const wb = structure.weekBlocks[b];
    for (let gi = wb.groups.length - 1; gi >= 0; gi--) {
      const g = wb.groups[gi];
      if (g.start >= beforeRow) continue;
      const coincide = isMov ? g.grupo === MOVILIDAD : (g.grupo !== MOVILIDAD && g.dia === String(dia));
      if (!coincide) continue;
      for (let r = g.start; r <= g.end; r++) {
        if (nombreEjercicio(ws, r) === name) {
          const peso = ws.getCell(r, COLS.peso).text;
          const reps = ws.getCell(r, COLS.reps).text;
          if (peso || reps) return { peso, reps };
        }
      }
    }
  }
  return null;
}

// ¿un día que ya está cargado en la semana actual corresponde a una repetición o a semana nueva?
function sugerirModoSemana(wb, ws, fechaEntry) {
  const ses = sesionesDeBloque(wb, ws);
  if (!ses.length) return { modo: 'extra' };
  const k0 = Math.min(...ses.map(s => fechaKey(s.fecha)));
  const dias = diasEntreKeys(fechaKey(fechaEntry), k0);
  if (dias < 0) return { modo: 'error', motivo: 'La fecha (' + fechaEntry + ') es anterior al inicio de la semana actual de la planilla. Revisala a mano.' };
  if (ses.length < SESIONES_OBJETIVO && dias <= DIAS_MAX_SEMANA) return { modo: 'extra' };
  return { modo: 'nueva' };
}

// Planifica las escrituras de UNA entrada (JSON del celu) sobre la semana actual.
// opts.modo: 'auto' (default) | 'extra' | 'nueva' — qué hacer si el día ya está cargado en la semana.
// Devuelve { writes, errors, avisos, needsNewWeek, extra }.
function computeWrites(structure, ws, entry, opts) {
  opts = opts || {};
  const writes = [], errors = [], avisos = [];
  let blockIndex = structure.weekBlocks.length - 1;
  while (blockIndex >= 0 && structure.weekBlocks[blockIndex].groups.length === 0) blockIndex--;
  const wb = structure.weekBlocks[blockIndex];
  if (!wb) {
    errors.push('No hay una semana preparada todavía.');
    return { writes, errors, avisos };
  }
  const fechaEntry = dateToDDMMYYYY(entry.fecha);
  const cands = diaGroups(wb, entry.dia);
  if (!cands.length) { errors.push('No encontré el bloque del día ' + entry.dia + ' en la semana actual.'); return { writes, errors, avisos }; }

  const info = cands.map(g => ({ g, blank: isGroupBlank(ws, g), fecha: cellDateDDMMYYYY(ws.getCell(g.start, COLS.fecha)) }));
  // 1) misma fecha ya cargada -> se reescribe esa sesión; 2) primer grupo vacío; 3) grupo con datos pero sin fecha
  const hit = info.find(x => !x.blank && x.fecha === fechaEntry) || info.find(x => x.blank) || info.find(x => !x.fecha);
  let group = hit ? hit.g : null;
  let insertOp = null;

  if (!group) {
    let modo = opts.modo || 'auto';
    const sug = sugerirModoSemana(wb, ws, fechaEntry);
    if (modo === 'auto') modo = sug.modo;
    if (modo === 'error') { errors.push(sug.motivo); return { writes, errors, avisos }; }
    if (modo === 'nueva') {
      errors.push('El día ' + entry.dia + ' ya está cargado en la semana actual (' + info.map(x => x.fecha).filter(Boolean).join(', ') + '): esta sesión va en una semana nueva.');
      return { writes, errors, avisos, needsNewWeek: true };
    }
    // modo 'extra': se agrega una copia de las filas del día al final de la misma semana
    const tpl = cands[0];
    const n = tpl.end - tpl.start + 1;
    const at = wb.end + 1;
    insertOp = { op: 'insertGroup', at, n, templateStart: tpl.start, label: 'Repetición de Día ' + entry.dia + ' en la misma semana (' + n + ' filas nuevas)' };
    group = { start: at, end: at + n - 1, virtualDesde: tpl.start };
  }

  const n = group.end - group.start + 1;
  const nameAt = (row) => nombreEjercicio(ws, group.virtualDesde ? group.virtualDesde + (row - group.start) : row);
  if (entry.ejercicios.some(e => e.posicionIdeal > n)) {
    errors.push('La cantidad de ejercicios no coincide con la plantilla de la planilla (' + n + ' filas). Revisá el orden.');
  }

  entry.ejercicios.forEach(ex => {
    const row = group.start + (ex.posicionIdeal - 1);
    const rowExerciseName = nameAt(row);
    if (rowExerciseName !== ex.nombre.trim()) {
      errors.push('Fila ' + row + ': se esperaba "' + rowExerciseName + '" pero el registro dice "' + ex.nombre + '". Revisá la plantilla del celular.');
      return;
    }
    const ordenStr = String(ex.ordenReal) + (ex.ordenReal !== ex.posicionIdeal ? '*' : '');
    const esAccesoria = ex.maquina === 'accesoria';
    writes.push({ row, col: 'B', value: fechaEntry, isDate: true, label: ex.nombre });
    writes.push({ row, col: 'D', value: ordenStr, label: ex.nombre });
    if (ex.ss) writes.push({ row, col: 'L', value: 'S', label: ex.nombre });
    if (ex.fallo) writes.push({ row, col: 'M', value: 'S', label: ex.nombre });
    const notaRetorno = ex.retornoAplicado ? 'Retorno ' + ex.retornoAplicado + '%' : '';
    const notaFinal = [notaRetorno, ex.notas].filter(Boolean).join(' · ');
    if (notaFinal) writes.push({ row, col: 'N', value: notaFinal, label: ex.nombre });
    writes.push({ row, col: 'Q', value: entry.ayuno ? 'Sí' : 'No', label: ex.nombre });

    if (esAccesoria) {
      writes.push({ row, col: 'O', value: toSheetNumberString(ex.peso), label: ex.nombre + ' (accesoria)' });
      writes.push({ row, col: 'P', value: toSheetNumberString(ex.reps), label: ex.nombre + ' (accesoria)' });
    } else {
      writes.push({ row, col: 'F', value: toSheetNumberString(ex.peso), label: ex.nombre });
      writes.push({ row, col: 'H', value: toSheetNumberString(ex.reps), label: ex.nombre });
      const prev = findPrevValue(structure, ws, group.start, ex.nombre, false, entry.dia);
      if (prev) {
        const dPeso = parseNum(ex.peso) - parseNum(prev.peso || '0');
        const dReps = parseNum(ex.reps) - parseNum(prev.reps || '0');
        writes.push({ row, col: 'I', value: fmtDelta(dReps), label: ex.nombre });
        writes.push({ row, col: 'J', value: fmtDelta(dPeso), label: ex.nombre });
      }
    }
  });

  const movGroup = findMovGroup(wb);
  if (movGroup && entry.movilidad && entry.movilidad.length) {
    entry.movilidad.forEach(mx => {
      let found = null;
      for (let r = movGroup.start; r <= movGroup.end; r++) {
        if (nombreEjercicio(ws, r) === mx.nombre.trim()) { found = r; break; }
      }
      if (!found) { errors.push('No encontré el ejercicio de movilidad "' + mx.nombre + '" en la plantilla.'); return; }
      // la movilidad tiene una sola fila por ejercicio en la semana: si ya tiene datos de otra
      // fecha (una repetición del día), no se pisa
      const yaTiene = ws.getCell(found, COLS.peso).text || ws.getCell(found, COLS.reps).text;
      const fechaYa = cellDateDDMMYYYY(ws.getCell(found, COLS.fecha));
      if (yaTiene && fechaYa && fechaYa !== fechaEntry) {
        avisos.push('Movilidad "' + mx.nombre + '" ya tenía datos del ' + fechaYa + ': no se pisó.');
        return;
      }
      writes.push({ row: found, col: 'B', value: fechaEntry, isDate: true, label: mx.nombre });
      writes.push({ row: found, col: 'F', value: toSheetNumberString(mx.peso), label: mx.nombre });
      writes.push({ row: found, col: 'H', value: toSheetNumberString(mx.reps), label: mx.nombre });
      if (mx.fallo) writes.push({ row: found, col: 'M', value: 'S', label: mx.nombre });
      if (mx.notas) writes.push({ row: found, col: 'N', value: mx.notas, label: mx.nombre });

      const prev = findPrevValue(structure, ws, movGroup.start, mx.nombre, true, entry.dia);
      if (prev) {
        const dPeso = parseNum(mx.peso) - parseNum(prev.peso || '0');
        const dReps = parseNum(mx.reps) - parseNum(prev.reps || '0');
        writes.push({ row: found, col: 'I', value: fmtDelta(dReps), label: mx.nombre });
        writes.push({ row: found, col: 'J', value: fmtDelta(dPeso), label: mx.nombre });
      }
    });
  }

  if (insertOp) writes.unshift(insertOp);
  return { writes, errors, avisos, extra: !!insertOp };
}

// copia las filas de una plantilla a filas nuevas (mismos textos y formatos, sin los datos cargados)
const COLUMNAS_A_LIMPIAR = ['B','D','F','H','I','J','K','L','M','N','O','P','Q'];
function copiarFilasPlantilla(ws, srcStart, insertAt, n) {
  ws.spliceRows(insertAt, 0, ...Array.from({ length: n }, () => []));
  for (let i = 0; i < n; i++) {
    const src = ws.getRow(srcStart + i);
    const dst = ws.getRow(insertAt + i);
    for (let c = 1; c <= 17; c++) {
      const sc = src.getCell(c);
      const dc = dst.getCell(c);
      dc.value = sc.value;
      if (sc.style) dc.style = JSON.parse(JSON.stringify(sc.style));
      if (sc.numFmt) dc.numFmt = sc.numFmt;
    }
    dst.height = src.height;
  }
  COLUMNAS_A_LIMPIAR.forEach(colLetter => {
    const colNum = Number(Object.keys(COLLETTERS).find(k => COLLETTERS[k] === colLetter));
    for (let i = 0; i < n; i++) ws.getRow(insertAt + i).getCell(colNum).value = null;
  });
}

function applyWrites(ws, writes) {
  writes.filter(w => w.op === 'insertGroup').forEach(w => copiarFilasPlantilla(ws, w.templateStart, w.at, w.n));
  writes.filter(w => !w.op).forEach(w => {
    const colNum = Number(Object.keys(COLLETTERS).find(k => COLLETTERS[k] === w.col));
    const cell = ws.getCell(w.row, colNum);
    if (w.isDate) {
      cell.value = ddmmyyyyToDate(w.value);   // fecha real (UTC), Excel la guarda como serie
      cell.numFmt = cell.numFmt || 'd/m/yyyy';
    } else {
      cell.value = w.value;
    }
  });
}

function duplicateHeaderRow(ws, srcHeaderRow, atRow) {
  ws.spliceRows(atRow, 0, []);
  const srcHeader = ws.getRow(srcHeaderRow);
  const dstHeader = ws.getRow(atRow);
  for (let c = 1; c <= 17; c++) {
    const sc = srcHeader.getCell(c);
    const dc = dstHeader.getCell(c);
    dc.value = sc.value;
    if (sc.style) dc.style = JSON.parse(JSON.stringify(sc.style));
  }
  dstHeader.height = srcHeader.height;
}

// filas de la semana "estándar" de un bloque: hasta antes de la primera repetición (sesión extra)
function rangoPlantilla(block) {
  const vistos = new Set();
  let end = block.end;
  for (const g of block.groups) {
    if (g.grupo === MOVILIDAD) continue;
    const clave = g.dia + '|' + String(g.grupo).trim();
    if (vistos.has(clave)) { end = g.start - 1; break; }
    vistos.add(clave);
  }
  return { start: block.start, end };
}

function prepareNewWeek(ws, structure) {
  const lastBlock = structure.weekBlocks[structure.weekBlocks.length - 1];
  let templateBlock, insertAfterRow;

  if (lastBlock.groups.length === 0) {
    // ya existe un encabezado vacío al final, esperando la próxima semana
    templateBlock = structure.weekBlocks[structure.weekBlocks.length - 2];
    if (!templateBlock) throw new Error('No encontré una semana completa anterior para copiar.');
    insertAfterRow = lastBlock.headerRow;
  } else {
    // la última semana ya tiene datos: se crea un encabezado nuevo después, usando esta misma de plantilla
    templateBlock = lastBlock;
    const newHeaderRow = lastBlock.end + 1;
    duplicateHeaderRow(ws, lastBlock.headerRow, newHeaderRow);
    insertAfterRow = newHeaderRow;
  }

  const rango = rangoPlantilla(templateBlock);
  const n = rango.end - rango.start + 1;
  const insertAt = insertAfterRow + 1;
  copiarFilasPlantilla(ws, rango.start, insertAt, n);

  // encabezado vacío al final, listo para la próxima "semana nueva"
  const trailingHeaderRow = insertAt + n;
  duplicateHeaderRow(ws, insertAfterRow, trailingHeaderRow);
  return { insertAt, insertEnd: insertAt + n - 1, newHeaderRow: trailingHeaderRow };
}

// Aplica UNA entrada a la hoja en memoria: decide sola si va en la semana actual, como repetición
// dentro de la misma semana, o si hay que abrir una semana nueva (y la abre).
// Devuelve { ok, errors, avisos, pasos, writes, extra }.
function aplicarEntrada(ws, entry, opts) {
  opts = opts || {};
  const pasos = [];
  let structure = parseStructure(ws);
  let res = computeWrites(structure, ws, entry, opts);
  if (res.needsNewWeek) {
    const r = prepareNewWeek(ws, structure);
    pasos.push('Semana nueva abierta (filas ' + r.insertAt + '–' + r.insertEnd + ')');
    structure = parseStructure(ws);
    res = computeWrites(structure, ws, entry, Object.assign({}, opts, { modo: 'auto' }));
  }
  if (res.errors.length) return { ok: false, errors: res.errors, avisos: res.avisos || [], pasos, writes: res.writes };
  applyWrites(ws, res.writes);
  if (res.extra) pasos.push('Día ' + entry.dia + ' repetido: se sumó a la misma semana');
  return { ok: true, errors: [], avisos: res.avisos || [], pasos, writes: res.writes, extra: !!res.extra };
}

// referencias del celu: último peso/reps de cada ejercicio de la semana actual hacia atrás
function buildSnapshot(structure, ws) {
  let blockIndex = structure.weekBlocks.length - 1;
  while (blockIndex >= 0 && structure.weekBlocks[blockIndex].groups.length === 0) blockIndex--;
  if (blockIndex < 0) return {};
  const templateBlock = structure.weekBlocks[blockIndex];
  const rango = rangoPlantilla(templateBlock);

  const names = [];
  for (let r = rango.start; r <= rango.end; r++) {
    const n = nombreEjercicio(ws, r);
    if (n && !names.includes(n)) names.push(n);
  }

  const snapshot = {};
  names.forEach(nombre => {
    for (let b = blockIndex; b >= 0; b--) {
      const wb = structure.weekBlocks[b];
      for (let gi = wb.groups.length - 1; gi >= 0; gi--) {
        const g = wb.groups[gi];
        for (let r = g.start; r <= g.end; r++) {
          if (nombreEjercicio(ws, r) !== nombre) continue;
          const peso = ws.getCell(r, COLS.peso).text;
          const reps = ws.getCell(r, COLS.reps).text;
          const notas = ws.getCell(r, COLS.notas).text;
          if (peso || reps) { snapshot[nombre] = { peso, reps, notas }; return; }
        }
      }
    }
  });
  return snapshot;
}
