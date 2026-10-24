/**
 * Planilla de consultas y contrataciones de Fabián Gi.
 *
 * Hojas que crea:
 *  - Consultas:      una fila por consulta que llega del formulario (acá marcás el estado)
 *  - Contrataciones: lista automática de shows contratados, ordenada por fecha del evento
 *  - Panel:          estadísticas y gráficos (consultas vs contrataciones)
 *  - Previsión:      cuántas consultas necesitás para cerrar los shows que querés
 *
 * Instrucciones al final del archivo.
 */

const HOJA = 'Consultas';
const TZ = 'America/Argentina/Buenos_Aires';
const ESTADOS = ['Nueva', 'Cotizada', 'Contratada', 'Perdida'];
const ENCABEZADOS = [
  'Fecha consulta', 'Mes', 'Estado', 'Tipo de evento', 'Nombre', 'WhatsApp', 'Email',
  'Fecha del evento', 'Horario', 'Localidad', 'Personas', 'Tipo de magia',
  'Cómo nos conoció', 'Detalle', 'Precio cotizado', 'Notas', 'Presupuesto enviado',
  'Fecha contratación'
];
// Columnas: A Fecha consulta, B Mes, C Estado, D Tipo, E Nombre, F WhatsApp, G Email,
// H Fecha evento, I Horario, J Localidad, K Personas, L Tipo magia, M Origen,
// N Detalle, O Precio, P Notas, Q Presupuesto enviado, R Fecha contratación

/* ======================= RECEPCIÓN DEL FORMULARIO ======================= */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.web) return respuesta_(true);                       // trampa anti-spam
    if (!d.nombre || !d.whatsapp || !d.tipo) return respuesta_(false);

    lock.waitLock(15000);
    const ahora = new Date();
    const sh = hoja_();
    sh.appendRow([
      ahora,
      Utilities.formatDate(ahora, TZ, 'yyyy-MM'),
      'Nueva',
      limpiar_(d.tipo), limpiar_(d.nombre), limpiar_(d.whatsapp), limpiar_(d.email),
      limpiar_(d.fechaEvento), limpiar_(d.horario), limpiar_(d.localidad),
      Number(d.personas) || '', limpiar_(d.tipoMagia), limpiar_(d.origen),
      limpiar_(d.detalle, 1500), '', '', limpiar_(d.presupuestoNombre), ''
    ]);
    // WhatsApp como enlace: un clic y se abre el chat con el cliente
    const digitos = String(d.whatsapp).replace(/\D/g, '');
    sh.getRange(sh.getLastRow(), 6)
      .setFormula('=HYPERLINK("https://wa.me/' + wa_(d.whatsapp) + '","' + digitos + '")');
    lock.releaseLock();
    avisar_(d);
    return respuesta_(true);
  } catch (err) {
    try { lock.releaseLock(); } catch (x) {}
    return respuesta_(false);
  }
}

function respuesta_(ok) {
  return ContentService.createTextOutput(JSON.stringify({ ok: ok }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Evita que un texto escrito por un visitante se interprete como fórmula
function limpiar_(v, max) {
  let s = String(v === undefined || v === null ? '' : v).trim().slice(0, max || 300);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

// Número argentino de celular para wa.me: 549 + código de área + número
function wa_(n) {
  let s = String(n || '').replace(/\D/g, '').replace(/^0+/, '');
  if (s.indexOf('54') === 0) s = s.slice(2);
  if (s.indexOf('9') === 0 && s.length > 10) s = s.slice(1);
  return '549' + s;
}

function avisar_(d) {
  try {
    const para = Session.getEffectiveUser().getEmail();
    const cuerpo = [
      'Nueva consulta de ' + d.nombre,
      'Evento: ' + d.tipo,
      'Fecha: ' + (d.fechaEvento || 'sin fecha') + ' ' + (d.horario || ''),
      'Localidad: ' + d.localidad,
      'Presupuesto que vio: ' + (d.presupuestoNombre || 'ninguno'),
      'Detalle: ' + (d.detalle || '-'),
      'WhatsApp: https://wa.me/' + wa_(d.whatsapp)
    ].join('\n');
    const manual = String(d.presupuestoNombre || '').indexOf('A medida') === 0;
    const asunto = (manual ? '⚠ ENVIAR PRESUPUESTO A MEDIDA - ' : 'Nueva consulta: ') + d.tipo + ' - ' + d.nombre;
    MailApp.sendEmail(para, asunto, (manual ? 'Evento en otra localidad: hay que calcular el traslado y mandar el presupuesto a mano.\n\n' : '') + cuerpo);
  } catch (err) { /* el aviso falla en silencio: la consulta ya quedó guardada */ }
}

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(HOJA) || ss.insertSheet(HOJA);
}

/* ======================= MENÚ (los "macros") ======================= */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🎩 Magia')
    .addItem('✅ Marcar como Contratada', 'marcarContratada')
    .addItem('📨 Marcar como Cotizada', 'marcarCotizada')
    .addItem('❌ Marcar como Perdida', 'marcarPerdida')
    .addSeparator()
    .addItem('🔄 Actualizar Panel y Previsión', 'actualizarPaneles')
    .addToUi();
}
function marcarContratada() { cambiarEstado_('Contratada', true); }
function marcarCotizada()   { cambiarEstado_('Cotizada', false); }
function marcarPerdida()    { cambiarEstado_('Perdida', false); }

function cambiarEstado_(estado, pedirPrecio) {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== HOJA) { ui.alert('Primero seleccioná una fila de la hoja "Consultas".'); return; }
  const rango = sh.getActiveRange();
  const ini = rango.getRow(), n = rango.getNumRows();
  let cambiadas = 0;
  for (let r = ini; r < ini + n; r++) {
    if (r < 2 || !sh.getRange(r, 5).getValue()) continue;
    if (pedirPrecio) {
      const resp = ui.prompt('Contratación de ' + sh.getRange(r, 5).getValue(),
        'Precio cerrado (solo números). Dejalo vacío para no cambiar el valor actual.', ui.ButtonSet.OK_CANCEL);
      if (resp.getSelectedButton() !== ui.Button.OK) break;
      const num = parseFloat(String(resp.getResponseText()).replace(/\./g, '').replace(',', '.'));
      if (!isNaN(num)) sh.getRange(r, 15).setValue(num);
    }
    sh.getRange(r, 3).setValue(estado);
    if (estado === 'Contratada' && !sh.getRange(r, 18).getValue()) sh.getRange(r, 18).setValue(new Date());
    cambiadas++;
  }
  if (!cambiadas) ui.alert('No había ninguna consulta seleccionada.');
}

// También funciona si cambiás el estado con el desplegable
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== HOJA || e.range.getColumn() !== 3 || e.range.getRow() < 2) return;
    const fila = e.range.getRow();
    if (e.value === 'Contratada' && !sh.getRange(fila, 18).getValue()) sh.getRange(fila, 18).setValue(new Date());
  } catch (err) {}
}

/* ======================= PREPARACIÓN (una sola vez) ======================= */
function preparar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = hoja_();

  sh.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS])
    .setFontWeight('bold').setBackground('#2b0b19').setFontColor('#ffffff');
  sh.setFrozenRows(1);

  sh.getRange('C2:C2000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).setAllowInvalid(false).build());
  sh.getRange('A2:A2000').setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange('R2:R2000').setNumberFormat('dd/mm/yyyy');
  sh.getRange('O2:O2000').setNumberFormat('$#,##0');

  const reglas = [['Nueva', '#fff2cc'], ['Cotizada', '#cfe2f3'], ['Contratada', '#b7e1cd'], ['Perdida', '#f4cccc']]
    .map(r => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(r[0])
      .setBackground(r[1]).setRanges([sh.getRange('C2:C2000')]).build());
  sh.setConditionalFormatRules(reglas);
  sh.setColumnWidths(1, 13, 130);
  sh.setColumnWidth(14, 420);
  sh.setColumnWidths(15, 4, 150);

  actualizarPaneles();
  ss.setActiveSheet(ss.getSheetByName('Panel'));
}

function actualizarPaneles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  armarContrataciones_(ss);
  armarPanel_(ss);
  armarPrevision_(ss);
}

/* ======================= HOJA "Contrataciones" ======================= */
function armarContrataciones_(ss) {
  const c = ss.getSheetByName('Contrataciones') || ss.insertSheet('Contrataciones');
  c.clear();
  c.getRange('A1').setValue('Shows contratados').setFontSize(18).setFontWeight('bold');
  c.getRange('A2').setValue('Se completa sola cuando marcás una consulta como "Contratada". Ordenada por fecha del evento.');
  c.getRange('A4').setFormula(
    `=IFERROR(QUERY(Consultas!A1:R,"select E, D, H, I, J, K, L, O, F where C = 'Contratada' order by H ` +
    `label E 'Nombre', D 'Tipo de evento', H 'Fecha del evento', I 'Horario', J 'Localidad', K 'Personas', ` +
    `L 'Tipo de magia', O 'Precio', F 'WhatsApp'",1),"Todavía no hay shows contratados.")`);
  c.getRange('C5:C500').setNumberFormat('dd/mm/yyyy');
  c.getRange('H5:H500').setNumberFormat('$#,##0');
  c.setColumnWidths(1, 9, 140);
  c.setColumnWidth(1, 200);
  c.setColumnWidth(2, 240);
  c.setFrozenRows(4);
}

/* ======================= HOJA "Panel" ======================= */
function letra_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// Tabla: nombre | consultas | contratadas | tasa de contratación
function bloque_(p, titulo, colOrigen, colIni, filaTitulo) {
  const a = letra_(colIni), b = letra_(colIni + 1), c = letra_(colIni + 2);
  const f = filaTitulo + 2;
  p.getRange(filaTitulo, colIni).setValue(titulo).setFontWeight('bold');
  p.getRange(filaTitulo + 1, colIni, 1, 4).setValues([[titulo.split(' ')[0], 'Consultas', 'Contratadas', 'Tasa']])
    .setFontWeight('bold').setBackground('#eee4d8');
  const src = 'Consultas!' + colOrigen + '2:' + colOrigen;
  p.getRange(f, colIni).setFormula(`=IFERROR(SORT(UNIQUE(FILTER(${src},${src}<>""))),"")`);
  p.getRange(f, colIni + 1).setFormula(`=ARRAYFORMULA(IF(${a}${f}:${a}="","",COUNTIF(${src},${a}${f}:${a})))`);
  p.getRange(f, colIni + 2).setFormula(
    `=ARRAYFORMULA(IF(${a}${f}:${a}="","",COUNTIFS(${src},${a}${f}:${a},Consultas!C2:C,"Contratada")))`);
  p.getRange(f, colIni + 3).setFormula(
    `=ARRAYFORMULA(IF(${a}${f}:${a}="","",IF(${b}${f}:${b}=0,0,${c}${f}:${c}/${b}${f}:${b})))`);
  p.getRange(f, colIni + 3, 300, 1).setNumberFormat('0%');
}

function armarPanel_(ss) {
  const p = ss.getSheetByName('Panel') || ss.insertSheet('Panel');
  p.clear();
  p.getCharts().forEach(c => p.removeChart(c));

  p.getRange('A1').setValue('Panel de consultas y contrataciones').setFontSize(18).setFontWeight('bold');

  const kpis = [
    ['Consultas totales', '=COUNTA(Consultas!E2:E)'],
    ['Nuevas (sin cotizar)', '=COUNTIF(Consultas!C2:C,"Nueva")'],
    ['Cotizadas', '=COUNTIF(Consultas!C2:C,"Cotizada")'],
    ['Contratadas', '=COUNTIF(Consultas!C2:C,"Contratada")'],
    ['Perdidas', '=COUNTIF(Consultas!C2:C,"Perdida")'],
    ['Tasa de contratación (sobre todas)', '=IFERROR(B6/B3,0)'],
    ['Tasa sobre consultas ya cerradas', '=IFERROR(B6/(B6+B7),0)'],
    ['Facturación (contratadas)', '=SUMIF(Consultas!C2:C,"Contratada",Consultas!O2:O)'],
    ['Precio promedio contratado', '=IFERROR(B10/B6,0)'],
    ['Días promedio hasta contratar',
      '=IFERROR(AVERAGE(FILTER(ARRAYFORMULA(Consultas!R2:R-Consultas!A2:A),Consultas!C2:C="Contratada",Consultas!R2:R<>"")),0)']
  ];
  p.getRange('A3:B12').setFormulas(kpis);
  p.getRange('A3:A12').setFontWeight('bold');
  p.getRange('B8:B9').setNumberFormat('0.0%');
  p.getRange('B10:B11').setNumberFormat('$#,##0');
  p.getRange('B12').setNumberFormat('0.0');

  // Tablas (abajo, para que puedan crecer sin chocar con los gráficos)
  bloque_(p, 'Mes', 'B', 1, 63);
  bloque_(p, 'Tipo de evento', 'D', 6, 63);
  bloque_(p, 'Origen (cómo te conocieron)', 'M', 11, 63);
  bloque_(p, 'Localidad', 'J', 16, 63);
  p.getRange('A63').setValue('Mes (año-mes de la consulta)');

  const grafico = (rangos, tipo, titulo, fila, col, leyenda, encabezados, extra) => {
    const b = p.newChart().setChartType(tipo);
    rangos.forEach(r => b.addRange(p.getRange(r)));
    b.setNumHeaders(encabezados).setPosition(fila, col, 0, 0)
      .setOption('title', titulo).setOption('width', 480).setOption('height', 300)
      .setOption('legend', { position: leyenda });
    if (extra) Object.keys(extra).forEach(k => b.setOption(k, extra[k]));
    p.insertChart(b.build());
  };
  const COL = Charts.ChartType.COLUMN;
  grafico(['A64:C76'], COL, 'Consultas vs contratadas por mes', 14, 1, 'bottom', 1);
  grafico(['A4:B7'], Charts.ChartType.PIE, 'Estado de las consultas', 14, 5, 'right', 0);
  grafico(['F64:H76'], COL, 'Consultas vs contratadas por tipo de evento', 30, 1, 'bottom', 1);
  grafico(['F64:F76', 'I64:I76'], COL, 'Tasa de contratación por tipo de evento', 30, 5, 'none', 1, { vAxis: { format: '0%' } });
  grafico(['K64:M76'], COL, 'Consultas vs contratadas por origen', 46, 1, 'bottom', 1);
  grafico(['P64:R76'], Charts.ChartType.BAR, 'Consultas vs contratadas por localidad', 46, 5, 'bottom', 1);

  p.setColumnWidths(1, 19, 120);
  p.setColumnWidth(1, 250);
}

/* ======================= HOJA "Previsión" ======================= */
function armarPrevision_(ss) {
  const v = ss.getSheetByName('Previsión') || ss.insertSheet('Previsión');
  v.clear();
  v.getRange('A1').setValue('Previsión: ¿cuántas consultas necesito?').setFontSize(18).setFontWeight('bold');
  v.getRange('A2').setValue('Completá las celdas amarillas. Todo lo demás se calcula con tu historial real.');

  const filas = [
    ['Tu tasa de contratación actual', '=Panel!B8'],
    ['Contrataciones que querés cerrar por mes', 4],
    ['Tasa a usar en la previsión (podés escribir otro % para simular)', '=B3'],
    ['Consultas que necesitás por mes', '=IFERROR(ROUNDUP(B4/B5,0),"Faltan datos")'],
    ['Consultas recibidas este mes', '=COUNTIF(Consultas!B2:B,YEAR(TODAY())&"-"&TEXT(MONTH(TODAY()),"00"))'],
    ['Te faltan este mes', '=IFERROR(MAX(0,B6-B7),"")'],
    ['Contrataciones esperadas este mes con las consultas que ya tenés', '=IFERROR(ROUND(B7*B5,1),"")'],
    ['Valor promedio de cada consulta', '=IFERROR(Panel!B11*B5,0)']
  ];
  v.getRange('A3:B10').setValues(filas.map(f => [f[0], '']));
  filas.forEach((f, i) => {
    const celda = v.getRange(4 - 1 + i, 2);
    if (typeof f[1] === 'string') celda.setFormula(f[1]); else celda.setValue(f[1]);
  });
  v.getRange('A3:A10').setFontWeight('bold');
  v.getRange('B3').setNumberFormat('0.0%');
  v.getRange('B5').setNumberFormat('0.0%');
  v.getRange('B10').setNumberFormat('$#,##0');
  v.getRange('B4:B5').setBackground('#fff2a8');
  v.getRange('A12').setValue('Ojo: con pocas consultas la tasa es poco confiable, y las consultas más recientes todavía pueden cerrar. ' +
    'Cuanto más historial tengas, más precisa es la previsión.').setFontStyle('italic').setWrap(true);
  v.getRange('A12:B12').merge();

  v.getRange('A15').setValue('Consultas necesarias según el objetivo').setFontWeight('bold');
  v.getRange('A16:B16').setValues([['Contrataciones por mes', 'Consultas necesarias']]).setFontWeight('bold').setBackground('#eee4d8');
  for (let i = 1; i <= 10; i++) {
    v.getRange(16 + i, 1).setValue(i);
    v.getRange(16 + i, 2).setFormula('=IFERROR(ROUNDUP(A' + (16 + i) + '/$B$5,0),"-")');
  }
  v.getRange('D15').setValue('Contrataciones esperadas según las consultas').setFontWeight('bold');
  v.getRange('D16:E16').setValues([['Consultas por mes', 'Contrataciones esperadas']]).setFontWeight('bold').setBackground('#eee4d8');
  [5, 10, 15, 20, 30, 40].forEach((n, i) => {
    v.getRange(17 + i, 4).setValue(n);
    v.getRange(17 + i, 5).setFormula('=IFERROR(ROUND(D' + (17 + i) + '*$B$5,1),"-")');
  });
  v.setColumnWidth(1, 460); v.setColumnWidth(2, 150); v.setColumnWidth(3, 40);
  v.setColumnWidth(4, 160); v.setColumnWidth(5, 200);
}

/*
 * INSTRUCCIONES
 * 1. Creá una planilla nueva en Google Sheets (por ejemplo "Consultas Fabián Gi").
 *    Si ya tenés una con este script, pegá este código encima del anterior.
 * 2. Menú Extensiones > Apps Script. Borrá lo que haya y pegá todo este archivo. Guardá.
 * 3. Elegí la función "preparar" arriba y tocá Ejecutar. Aceptá los permisos.
 *    Se crean las hojas Consultas, Contrataciones, Panel y Previsión.
 * 4. Implementar > Nueva implementación > tipo "Aplicación web".
 *    Ejecutar como: Yo. Quién tiene acceso: Cualquier persona. Implementar.
 *    (Si ya la tenías implementada: Implementar > Administrar implementaciones > editar > Nueva versión.)
 * 5. Copiá la URL que termina en /exec y pegala en CONFIG.ENDPOINT de cotizar.html.
 * 6. Cerrá y volvé a abrir la planilla: aparece el menú "🎩 Magia".
 *
 * USO DIARIO
 * - Cada consulta entra como fila "Nueva" y te llega un mail. El WhatsApp es un enlace al chat.
 * - Seleccioná la fila y usá el menú 🎩 Magia > "Marcar como Contratada" (te pide el precio cerrado),
 *   "Cotizada" o "Perdida". También podés usar el desplegable de la columna Estado.
 * - Los shows contratados aparecen solos en la hoja "Contrataciones".
 * - Panel y Previsión se recalculan solos.
 */
