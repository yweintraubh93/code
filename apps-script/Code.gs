/**
 * @OnlyCurrentDoc
 *
 * Cafetería — cuentas desde fotos de las hojas escritas a mano.
 *
 * Este script vive dentro de la hoja de Google de las cuentas (Extensiones > Apps Script).
 * - Publica una página (web app) para tomar/subir la foto desde el celular.
 * - Manda la foto a Claude, que lee la letra y la empareja con la lista de alumnos.
 * - Tú revisas y confirmas; al guardar se actualizan DEBE / A FAVOR, la pestaña BARZILAY
 *   y se anota todo en la pestaña "Movimientos".
 */

const CONFIG = {
  HOJA_ALUMNOS: 'Alumnos',
  HOJA_BARZILAY: 'BARZILAY',
  HOJA_MOVIMIENTOS: 'Movimientos',

  MODELO: 'claude-opus-5-5',
  ESFUERZO: 'medium',
  API_URL: 'https://api.anthropic.com/v1/messages',
  API_VERSION: '2026-10-21',

  // Familias que siempre se cobran a su cuenta familiar (renglón con solo el apellido),
  // aunque el niño ponga su nombre.
  FAMILIAS_CUENTA_UNICA: ['DELGADILLO'],
};

const PROP_API_KEY = 'ANTHROPIC_API_KEY';

const COLUMNAS_MOVIMIENTOS = [
  'Registrado', 'Lote', 'Fecha hoja', 'Cuenta', 'ID cuenta', 'Tipo', 'Monto',
  'DEBE antes', 'A FAVOR antes', 'DEBE después', 'A FAVOR después',
  'Texto original', 'Archivo', 'Huella', 'Estado', 'Clave',
];

/* ------------------------------------------------------------------ */
/* Menú en la hoja y página web                                        */
/* ------------------------------------------------------------------ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('☕ Cafetería')
    .addItem('Abrir app para fotos', 'mostrarLinkApp')
    .addItem('Configurar clave de API', 'configurarClaveApi')
    .addSeparator()
    .addItem('Deshacer el último lote guardado', 'deshacerUltimoLoteDesdeMenu')
    .addToUi();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Cuentas Cafetería')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function configurarClaveApi() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(
    'Clave de API de Anthropic',
    'Pega aquí tu clave (empieza con "sk-ant-"). Se guarda solo dentro de esta hoja.',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const clave = r.getResponseText().trim();
  if (!/^sk-ant-/.test(clave)) {
    ui.alert('Esa clave no parece correcta. Debe empezar con "sk-ant-".');
    return;
  }
  PropertiesService.getScriptProperties().setProperty(PROP_API_KEY, clave);
  ui.alert('Listo, la clave quedó guardada.');
}

function mostrarLinkApp() {
  const url = ScriptApp.getService().getUrl();
  const ui = SpreadsheetApp.getUi();
  if (!url) {
    ui.alert('Todavía no está publicada la app. En Apps Script: Implementar > Nueva implementación > App web.');
    return;
  }
  const html = HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif">Abre este link en tu celular y guárdalo en tu pantalla de inicio:</p>' +
    '<p style="font-family:sans-serif;word-break:break-all"><a href="' + url + '" target="_blank">' + url + '</a></p>')
    .setWidth(420).setHeight(160);
  ui.showModalDialog(html, 'App de la cafetería');
}

function deshacerUltimoLoteDesdeMenu() {
  const ui = SpreadsheetApp.getUi();
  const lote = ultimoLote_();
  if (!lote) { ui.alert('No hay lotes guardados para deshacer.'); return; }
  const ok = ui.alert('Deshacer', '¿Deshacer el lote ' + lote + '? Se revierten todos sus movimientos.', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  const r = deshacerLote(lote);
  ui.alert('Se revirtieron ' + r.revertidos + ' movimientos.');
}

/* ------------------------------------------------------------------ */
/* Funciones que llama la página                                       */
/* ------------------------------------------------------------------ */

/** Lista de cuentas para la página (buscador y saldos). */
function obtenerCuentas() {
  return listarCuentas_().map(c => ({
    id: c.id, etiqueta: c.etiqueta, clave: c.clave, tipoCuenta: c.tipoCuenta,
    debe: c.debe || 0, afavor: c.afavor || 0,
  }));
}

/**
 * Lee una hoja (una foto o una página de PDF) con Claude.
 * @param {{data:string, mediaType:string, nombreArchivo:string, huella:string, hoy:string}} p
 */
function leerHoja(p) {
  const clave = PropertiesService.getScriptProperties().getProperty(PROP_API_KEY);
  if (!clave) throw new Error('Falta la clave de API. En la hoja: menú ☕ Cafetería > Configurar clave de API.');

  const cuentas = listarCuentas_();
  const resultado = llamarClaude_(clave, cuentas, p);

  const porId = {};
  cuentas.forEach(c => { porId[c.id] = c; });
  const movimientos = (resultado.movimientos || []).map(m => {
    const cuentaId = m.cuenta_id && porId[m.cuenta_id] ? m.cuenta_id : null;
    const candidatos = (m.candidatos || []).filter(id => porId[id]);
    return {
      textoOriginal: m.texto_original || '',
      nombreLeido: m.nombre_leido || '',
      tipo: m.tipo === 'deposito' ? 'deposito' : 'cargo',
      monto: typeof m.monto === 'number' && m.monto > 0 ? m.monto : null,
      tachado: !!m.tachado,
      cuentaId: cuentaId,
      candidatos: candidatos,
      confianza: m.confianza || 'baja',
      nota: m.nota || '',
    };
  });

  return {
    fechaHoja: resultado.fecha_hoja || '',
    movimientos: movimientos,
    yaProcesada: !!p.huella && huellaYaUsada_(p.huella),
  };
}

/**
 * Guarda los movimientos confirmados.
 * @param {{archivo:string, huellas:string[], movimientos:Array<{cuentaId:string, clave:string,
 *   tipo:string, monto:number, fecha:string, textoOriginal:string}>}} lote
 */
function guardarMovimientos(lote) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shMov = hojaMovimientos_(ss);
    const loteId = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    const ahora = new Date();
    const huella = (lote.huellas || []).join(' ');
    const resumen = { cargos: 0, depositos: 0, cantidad: 0, lote: loteId };
    const movs = lote.movimientos || [];

    // 1) Validar todo antes de tocar la hoja.
    const ids = {};
    listarCuentas_().forEach(c => { ids[c.id] = c; });
    movs.forEach(m => {
      const quien = m.textoOriginal || m.cuentaId;
      if (!(Number(m.monto) > 0)) throw new Error('Monto inválido en: ' + quien);
      if (m.tipo !== 'cargo' && m.tipo !== 'deposito') throw new Error('Tipo inválido en: ' + quien);
      const esB = String(m.cuentaId).indexOf('B:') === 0;
      if (esB && !ids[m.cuentaId]) throw new Error('No encuentro esa cuenta en la pestaña BARZILAY: ' + quien);
      if (esB && m.tipo === 'deposito') throw new Error('Los depósitos de los Barzilay se anotan a mano en su pestaña: ' + quien);
      if (!esB && !/^A:\d+$/.test(String(m.cuentaId))) throw new Error('Falta elegir la cuenta en: ' + quien);
    });

    // 2) Aplicar y anotar cada movimiento en el momento (si algo falla, lo ya hecho queda registrado).
    movs.forEach(m => {
      const monto = Number(m.monto);
      const r = aplicarMovimiento_(ss, m.cuentaId, m.clave, m.tipo, monto, m.fecha);
      shMov.appendRow([
        ahora, loteId, m.fecha || '', r.etiqueta, r.id, m.tipo, monto,
        r.antes.debe, r.antes.afavor, r.despues.debe, r.despues.afavor,
        m.textoOriginal || '', lote.archivo || '', huella, 'OK', r.clave,
      ]);
      resumen.cantidad++;
      if (m.tipo === 'cargo') resumen.cargos += monto; else resumen.depositos += monto;
    });
    return resumen;
  } finally {
    lock.releaseLock();
  }
}

/** Revierte todos los movimientos de un lote. */
function deshacerLote(loteId) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = hojaMovimientos_(ss);
    const datos = sh.getDataRange().getValues();
    const c = indiceColumnas_(datos[0], COLUMNAS_MOVIMIENTOS);
    let revertidos = 0;
    for (let i = datos.length - 1; i >= 1; i--) {
      const fila = datos[i];
      if (String(fila[c['Lote']]) !== String(loteId) || fila[c['Estado']] !== 'OK') continue;
      revertirMovimiento_(ss, fila, c);
      sh.getRange(i + 1, c['Estado'] + 1).setValue('DESHECHO ' +
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
      revertidos++;
    }
    return { revertidos: revertidos };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ */
/* Cuentas                                                             */
/* ------------------------------------------------------------------ */

function listarCuentas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return leerAlumnos_(ss).concat(leerBarzilay_(ss));
}

function leerAlumnos_(ss) {
  const sh = ss.getSheetByName(CONFIG.HOJA_ALUMNOS);
  if (!sh) throw new Error('No encuentro la pestaña "' + CONFIG.HOJA_ALUMNOS + '".');
  const datos = sh.getDataRange().getValues();
  const col = columnasAlumnos_(datos[0]);
  // Apellidos que tienen alumnos: un renglón con solo ese apellido es la cuenta familiar.
  const apellidosConAlumnos = {};
  for (let i = 1; i < datos.length; i++) {
    if (limpiar_(datos[i][col.nombre])) apellidosConAlumnos[normalizar_(datos[i][col.paterno])] = true;
  }
  const cuentas = [];
  for (let i = 1; i < datos.length; i++) {
    const f = datos[i];
    const paterno = limpiar_(f[col.paterno]);
    const materno = limpiar_(f[col.materno]);
    const nombre = limpiar_(f[col.nombre]);
    if (!paterno && !materno && !nombre) continue;
    const grado = col.grado >= 0 ? limpiar_(f[col.grado]) : '';
    const esFamiliar = !nombre && !materno && !!apellidosConAlumnos[normalizar_(paterno)];
    const partes = [paterno, materno, nombre].filter(Boolean).join(' ');
    let etiqueta = partes;
    if (esFamiliar) etiqueta = paterno + ' (cuenta familiar)';
    else if (grado) etiqueta += ' — ' + grado;
    cuentas.push({
      id: 'A:' + (i + 1),
      fila: i + 1,
      tipoCuenta: esFamiliar ? 'familia' : (grado ? 'alumno' : 'adulto'),
      paterno: paterno, materno: materno, nombre: nombre, grado: grado,
      mama: col.mama >= 0 ? limpiar_(f[col.mama]) : '',
      debe: aNumero_(f[col.debe]),
      afavor: aNumero_(f[col.afavor]),
      etiqueta: etiqueta,
      clave: claveAlumno_(paterno, materno, nombre),
    });
  }
  return cuentas;
}

function columnasAlumnos_(encabezado) {
  const h = encabezado.map(x => normalizar_(x));
  const buscar = (opciones, obligatoria) => {
    for (const o of opciones) {
      const i = h.indexOf(normalizar_(o));
      if (i >= 0) return i;
    }
    for (const o of opciones) {
      const i = h.findIndex(x => x.indexOf(normalizar_(o)) === 0);
      if (i >= 0) return i;
    }
    if (obligatoria) throw new Error('No encuentro la columna "' + opciones[0] + '" en la pestaña Alumnos.');
    return -1;
  };
  return {
    paterno: buscar(['Apellido paterno'], true),
    materno: buscar(['Apellido materno'], true),
    nombre: buscar(['Nombre(s)', 'Nombre'], true),
    mama: buscar(['Mamá', 'Mama'], false),
    grado: buscar(['Grado'], false),
    debe: buscar(['DEBE'], true),
    afavor: buscar(['A FAVOR'], true),
  };
}

function leerBarzilay_(ss) {
  const sh = ss.getSheetByName(CONFIG.HOJA_BARZILAY);
  if (!sh) return [];
  const encabezado = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const cuentas = [];
  encabezado.forEach((v, i) => {
    const nombre = limpiar_(v);
    if (i === 0 || !nombre || /\d/.test(nombre) || nombre.indexOf(' ') >= 0) return;
    cuentas.push({
      id: 'B:' + nombre.toUpperCase(),
      tipoCuenta: 'barzilay',
      columna: i + 1,
      nombre: nombre,
      etiqueta: 'Barzilay ' + capitalizar_(nombre) + ' (pestaña BARZILAY)',
      clave: 'B|' + nombre.toUpperCase(),
      debe: 0, afavor: 0,
    });
  });
  return cuentas;
}

function claveAlumno_(paterno, materno, nombre) {
  return [paterno, materno, nombre].map(normalizar_).join('|');
}

/* ------------------------------------------------------------------ */
/* Aplicar movimientos                                                 */
/* ------------------------------------------------------------------ */

/**
 * Regla de saldos (no se pierde nada de lo que ya tienes):
 *  - Cargo: primero se descuenta de A FAVOR; lo que falte se suma a DEBE.
 *  - Depósito: primero se descuenta de DEBE; lo que sobre se suma a A FAVOR.
 */
function calcularSaldo(debe, afavor, tipo, monto) {
  debe = redondear_(debe || 0);
  afavor = redondear_(afavor || 0);
  if (tipo === 'cargo') {
    const usa = Math.min(afavor, monto);
    afavor = redondear_(afavor - usa);
    debe = redondear_(debe + (monto - usa));
  } else {
    const usa = Math.min(debe, monto);
    debe = redondear_(debe - usa);
    afavor = redondear_(afavor + (monto - usa));
  }
  return { debe: debe, afavor: afavor };
}

function aplicarMovimiento_(ss, cuentaId, clave, tipo, monto, fecha) {
  if (String(cuentaId).indexOf('B:') === 0) return aplicarBarzilay_(ss, cuentaId, tipo, monto, fecha);

  const sh = ss.getSheetByName(CONFIG.HOJA_ALUMNOS);
  const datos = sh.getDataRange().getValues();
  const col = columnasAlumnos_(datos[0]);
  const fila = filaDeCuenta_(datos, col, cuentaId, clave);
  const f = datos[fila - 1];
  const antes = { debe: aNumero_(f[col.debe]), afavor: aNumero_(f[col.afavor]) };
  const despues = calcularSaldo(antes.debe, antes.afavor, tipo, monto);
  escribirMonto_(sh.getRange(fila, col.debe + 1), despues.debe);
  escribirMonto_(sh.getRange(fila, col.afavor + 1), despues.afavor);

  const etiqueta = [limpiar_(f[col.paterno]), limpiar_(f[col.materno]), limpiar_(f[col.nombre])].filter(Boolean).join(' ');
  const claveFila = claveAlumno_(limpiar_(f[col.paterno]), limpiar_(f[col.materno]), limpiar_(f[col.nombre]));
  return { id: 'A:' + fila, clave: claveFila, etiqueta: etiqueta, antes: antes, despues: despues };
}

/** Número de renglón de la cuenta; si los renglones se movieron, la vuelve a buscar por nombre. */
function filaDeCuenta_(datos, col, cuentaId, clave) {
  const claveEn = i => {
    const f = datos[i - 1];
    return f ? claveAlumno_(limpiar_(f[col.paterno]), limpiar_(f[col.materno]), limpiar_(f[col.nombre])) : '';
  };
  const fila = parseInt(String(cuentaId).replace('A:', ''), 10);
  if (!clave || claveEn(fila) === clave) return fila;
  for (let i = 2; i <= datos.length; i++) if (claveEn(i) === clave) return i;
  throw new Error('No encontré la cuenta "' + clave.replace(/\|/g, ' ').trim() + '" en Alumnos.');
}

/** BARZILAY: cada renglón es un día del mes; cada columna un hijo. Se suma el consumo en su celda. */
function aplicarBarzilay_(ss, cuentaId, tipo, monto, fecha) {
  const sh = ss.getSheetByName(CONFIG.HOJA_BARZILAY);
  if (!sh) throw new Error('No encuentro la pestaña ' + CONFIG.HOJA_BARZILAY);
  const nombre = cuentaId.replace('B:', '');
  const cuenta = leerBarzilay_(ss).filter(c => c.id === cuentaId)[0];
  if (!cuenta) throw new Error('No encuentro a ' + nombre + ' en la pestaña BARZILAY.');
  if (tipo !== 'cargo') {
    throw new Error('Los depósitos de los Barzilay se anotan a mano en su pestaña (no sé dónde van).');
  }
  const dia = diaDeFecha_(fecha);
  const fila = filaDelDia_(sh, dia);
  const celda = sh.getRange(fila, cuenta.columna);
  const antes = aNumero_(celda.getValue());
  const despues = redondear_(antes + monto);
  escribirMonto_(celda, despues);
  return {
    id: cuentaId,
    clave: 'B|' + nombre,
    etiqueta: 'BARZILAY ' + nombre + ' día ' + dia,
    antes: { debe: antes, afavor: '' },
    despues: { debe: despues, afavor: '' },
  };
}

function filaDelDia_(sh, dia) {
  const ultima = Math.max(sh.getLastRow(), 1);
  const colA = sh.getRange(1, 1, ultima, 1).getValues();
  for (let i = colA.length - 1; i >= 1; i--) {
    if (Number(colA[i][0]) === dia && String(colA[i][0]).trim() !== '') return i + 1;
  }
  // Nuevo día: se agrega debajo del último renglón con número de día.
  let ultimaConDia = 1;
  for (let i = 1; i < colA.length; i++) if (String(colA[i][0]).trim() !== '' && !isNaN(Number(colA[i][0]))) ultimaConDia = i + 1;
  const nueva = ultimaConDia + 1;
  if (nueva <= sh.getLastRow() && String(sh.getRange(nueva, 1).getValue()).trim() !== '') sh.insertRowAfter(ultimaConDia);
  sh.getRange(nueva, 1).setValue(dia);
  return nueva;
}

function revertirMovimiento_(ss, fila, c) {
  const id = String(fila[c['ID cuenta']]);
  const monto = Number(fila[c['Monto']]);
  const tipo = fila[c['Tipo']];
  if (id.indexOf('B:') === 0) {
    const sh = ss.getSheetByName(CONFIG.HOJA_BARZILAY);
    const cuenta = leerBarzilay_(ss).filter(x => x.id === id)[0];
    const dia = diaDeFecha_(fila[c['Fecha hoja']]);
    const colA = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
    for (let i = colA.length - 1; i >= 1; i--) {
      if (Number(colA[i][0]) === dia) {
        const celda = sh.getRange(i + 1, cuenta.columna);
        escribirMonto_(celda, Math.max(0, redondear_(aNumero_(celda.getValue()) - monto)));
        return;
      }
    }
    return;
  }
  const sh = ss.getSheetByName(CONFIG.HOJA_ALUMNOS);
  const datos = sh.getDataRange().getValues();
  const col = columnasAlumnos_(datos[0]);
  const n = filaDeCuenta_(datos, col, id, String(fila[c['Clave']] || ''));
  const f = datos[n - 1];
  const actual = { debe: aNumero_(f[col.debe]), afavor: aNumero_(f[col.afavor]) };
  let nuevo;
  // Si nadie tocó la cuenta después, se regresa exactamente a como estaba.
  if (actual.debe === aNumero_(fila[c['DEBE después']]) && actual.afavor === aNumero_(fila[c['A FAVOR después']])) {
    nuevo = { debe: aNumero_(fila[c['DEBE antes']]), afavor: aNumero_(fila[c['A FAVOR antes']]) };
  } else {
    nuevo = calcularSaldo(actual.debe, actual.afavor, tipo === 'cargo' ? 'deposito' : 'cargo', monto);
  }
  escribirMonto_(sh.getRange(n, col.debe + 1), nuevo.debe);
  escribirMonto_(sh.getRange(n, col.afavor + 1), nuevo.afavor);
}

/* ------------------------------------------------------------------ */
/* Pestaña Movimientos                                                 */
/* ------------------------------------------------------------------ */

function hojaMovimientos_(ss) {
  let sh = ss.getSheetByName(CONFIG.HOJA_MOVIMIENTOS);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.HOJA_MOVIMIENTOS);
    sh.getRange(1, 1, 1, COLUMNAS_MOVIMIENTOS.length).setValues([COLUMNAS_MOVIMIENTOS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function huellaYaUsada_(huella) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_MOVIMIENTOS);
  if (!sh || sh.getLastRow() < 2) return false;
  const datos = sh.getDataRange().getValues();
  const c = indiceColumnas_(datos[0], COLUMNAS_MOVIMIENTOS);
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][c['Estado']] === 'OK' && String(datos[i][c['Huella']]).split(' ').indexOf(huella) >= 0) return true;
  }
  return false;
}

function ultimoLote_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.HOJA_MOVIMIENTOS);
  if (!sh || sh.getLastRow() < 2) return null;
  const datos = sh.getDataRange().getValues();
  const c = indiceColumnas_(datos[0], COLUMNAS_MOVIMIENTOS);
  for (let i = datos.length - 1; i >= 1; i--) if (datos[i][c['Estado']] === 'OK') return String(datos[i][c['Lote']]);
  return null;
}

function indiceColumnas_(encabezado, nombres) {
  const idx = {};
  nombres.forEach(n => {
    const i = encabezado.indexOf(n);
    if (i < 0) throw new Error('Falta la columna "' + n + '" en Movimientos.');
    idx[n] = i;
  });
  return idx;
}

/* ------------------------------------------------------------------ */
/* Claude                                                              */
/* ------------------------------------------------------------------ */

const ESQUEMA_LECTURA = {
  type: 'object',
  properties: {
    fecha_hoja: { type: 'string', description: 'Fecha escrita en la hoja en formato AAAA-MM-DD, o "" si no hay.' },
    movimientos: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          texto_original: { type: 'string' },
          nombre_leido: { type: 'string' },
          tipo: { type: 'string', enum: ['cargo', 'deposito'] },
          monto: { anyOf: [{ type: 'number' }, { type: 'null' }] },
          tachado: { type: 'boolean' },
          cuenta_id: { anyOf: [{ type: 'string' }, { type: 'null' }] },
          candidatos: { type: 'array', items: { type: 'string' } },
          confianza: { type: 'string', enum: ['alta', 'media', 'baja'] },
          nota: { type: 'string' },
        },
        required: ['texto_original', 'nombre_leido', 'tipo', 'monto', 'tachado', 'cuenta_id', 'candidatos', 'confianza', 'nota'],
        additionalProperties: false,
      },
    },
  },
  required: ['fecha_hoja', 'movimientos'],
  additionalProperties: false,
};

function instrucciones_(cuentas) {
  const lista = cuentas.map(c => {
    const extra = c.tipoCuenta === 'alumno' && c.mama ? ' (mamá ' + c.mama + ')' : '';
    return c.id + ' | ' + c.etiqueta + extra;
  }).join('\n');

  return [
    'Eres el asistente de una cafetería escolar en México. Los alumnos anotan a mano, en una hoja,',
    'lo que consumen "fiado" (en pesos). Tu trabajo es leer la foto de la hoja y convertir cada anotación',
    'en movimientos para sus cuentas.',
    '',
    'CÓMO SE ESCRIBE EN LAS HOJAS',
    '- Arriba suele venir la fecha, a veces con mes en número romano (24/XI/26 = 2026-11-24).',
    '- Cada renglón es: nombre (a veces solo apellido, apodo o con faltas) y uno o más montos.',
    '- Varios montos en el mismo renglón separados por "-", "+", "," o espacios ("25-25", "25+25+20")',
    '  son VARIOS CONSUMOS: un movimiento "cargo" por cada monto. El guion separa, no resta.',
    '- Un número ANTES del nombre ("60 Sophie 25-25") es dinero que DEJÓ: primero un "deposito" por ese',
    '  número y luego un "cargo" por cada monto que sigue.',
    '- "depositó", "deposito", "dejé/deje", "pagó/pagado", "abono", o un monto con "+" adelante ("+800")',
    '  son "deposito".',
    '- "$", "pesos" o "=" solo indican el monto. "5 pesos" = 5.',
    '- Si un renglón está tachado, repórtalo igual con tachado=true.',
    '- Si un número está corregido o encimado, usa la lectura más probable, baja la confianza y explícalo en "nota".',
    '- Si no hay monto (p.ej. "1 quesadilla"), monto=null y explica en "nota".',
    '- Ignora logotipos, marcas del cuaderno y garabatos sin nombre ni monto.',
    '',
    'CÓMO ELEGIR LA CUENTA (cuenta_id)',
    '- Usa SOLO los IDs de la lista <cuentas>. Nunca inventes un ID.',
    '- Compara apellidos y nombres con tolerancia a faltas de ortografía y letra difícil.',
    '- Las cuentas "(cuenta familiar)" son de toda la familia. Si solo viene el apellido y existe cuenta',
    '  familiar con ese apellido, usa la cuenta familiar.',
    '- Estas familias SIEMPRE van a su cuenta familiar, aunque escriban el nombre del niño: ' +
      CONFIG.FAMILIAS_CUENTA_UNICA.join(', ') + '.',
    '- Los hermanos Barzilay (Dor, Shay, Noy, Or) tienen cuentas "B:..." propias; úsalas siempre para ellos.',
    '  "Dov" casi siempre es Dor.',
    '- Si dudas entre varias cuentas, pon cuenta_id=null y en "candidatos" los IDs posibles (máximo 5),',
    '  del más probable al menos probable. Si estás razonablemente seguro, pon cuenta_id y candidatos=[].',
    '- confianza: "alta" si nombre y monto se leen claros y la cuenta es inequívoca; "media" si hubo que',
    '  interpretar algo; "baja" si es una suposición.',
    '',
    'SALIDA',
    '- Un elemento en "movimientos" por cada cargo o depósito, en el orden de la hoja.',
    '- texto_original: el renglón tal como lo lees (igual para todos los movimientos de ese renglón).',
    '- nota: breve, en español, solo si hay algo que revisar; si no, "".',
    '',
    '<cuentas>',
    lista,
    '</cuentas>',
  ].join('\n');
}

function llamarClaude_(clave, cuentas, p) {
  const esPdf = p.mediaType === 'application/pdf';
  const archivo = esPdf
    ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: p.data } }
    : { type: 'image', source: { type: 'base64', media_type: p.mediaType, data: p.data } };

  const cuerpo = {
    model: CONFIG.MODELO,
    max_tokens: 16000,
    output_config: {
      effort: CONFIG.ESFUERZO,
      format: { type: 'json_schema', schema: ESQUEMA_LECTURA },
    },
    fallbacks: 'default',
    system: [{ type: 'text', text: instrucciones_(cuentas), cache_control: { type: 'ephemeral' } }],
    messages: [{
      role: 'user',
      content: [
        archivo,
        { type: 'text', text: 'Hoy es ' + (p.hoy || '') + '. Lee esta hoja de la cafetería y devuelve los movimientos.' },
      ],
    }],
  };

  const opciones = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': clave,
      'anthropic-version': CONFIG.API_VERSION,
      'anthropic-beta': 'server-side-fallback-2026-07-01',
    },
    payload: JSON.stringify(cuerpo),
    muteHttpExceptions: true,
  };

  let resp;
  for (let intento = 1; intento <= 3; intento++) {
    resp = UrlFetchApp.fetch(CONFIG.API_URL, opciones);
    const code = resp.getResponseCode();
    if (code === 200) break;
    const reintentable = code === 429 || code >= 500;
    if (!reintentable || intento === 3) throw new Error(mensajeErrorApi_(code, resp.getContentText()));
    Utilities.sleep(2000 * intento);
  }

  const json = JSON.parse(resp.getContentText());
  if (json.stop_reason === 'refusal') throw new Error('La IA no quiso procesar esta imagen. Intenta con otra foto.');
  if (json.stop_reason === 'max_tokens') throw new Error('La hoja es demasiado larga. Intenta tomar la foto en dos partes.');
  const texto = (json.content || []).filter(b => b.type === 'text').map(b => b.text).join('');
  try {
    return JSON.parse(texto);
  } catch (e) {
    throw new Error('No pude entender la respuesta de la IA. Intenta de nuevo.');
  }
}

function mensajeErrorApi_(code, cuerpo) {
  let detalle = '';
  try { detalle = JSON.parse(cuerpo).error.message; } catch (e) { detalle = String(cuerpo).slice(0, 200); }
  if (code === 401) return 'La clave de API no es válida. Vuelve a configurarla desde el menú ☕ Cafetería.';
  if (code === 400 && /credit|balance/i.test(detalle)) return 'Se acabó el saldo de la cuenta de Anthropic. Recarga crédito en console.anthropic.com.';
  if (code === 413) return 'La foto es demasiado grande.';
  if (code === 429) return 'Demasiadas solicitudes seguidas. Espera un minuto e intenta de nuevo.';
  return 'Error de la IA (' + code + '): ' + detalle;
}

/* ------------------------------------------------------------------ */
/* Utilidades                                                          */
/* ------------------------------------------------------------------ */

function limpiar_(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); }

function normalizar_(v) {
  return limpiar_(v).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s*-\s*/g, '-');
}

function capitalizar_(s) { s = String(s).toLowerCase(); return s.charAt(0).toUpperCase() + s.slice(1); }

function redondear_(n) { return Math.round(Number(n) * 100) / 100; }

/** Convierte "$1,650.00", "$1.165,00", 52 o "" en número. */
function aNumero_(v) {
  if (typeof v === 'number') return redondear_(v);
  let s = limpiar_(v).replace(/[$\s]/g, '');
  if (!s) return 0;
  const coma = s.lastIndexOf(','), punto = s.lastIndexOf('.');
  if (coma > punto) s = s.replace(/\./g, '').replace(',', '.');
  else s = s.replace(/,/g, '');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : redondear_(n);
}

/** Escribe un monto respetando el estilo de la hoja: 0 se deja vacío. */
function escribirMonto_(rango, valor) {
  rango.setValue(valor ? valor : '');
}

function diaDeFecha_(fecha) {
  if (fecha instanceof Date) return fecha.getDate();
  const m = String(fecha || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return parseInt(m[3], 10);
  return new Date().getDate();
}
