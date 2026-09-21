var REG = 'Postulaciones';
var DASH = 'Dashboard';
var FILAS = 400;
var PRIMARIO = '#4F46E5';
var PRIMARIO_CLARO = '#EEF2FF';
var TEXTO_CLARO = '#FFFFFF';
var ESTADOS = ['Enviado', 'Visto', 'Entrevista', 'Prueba técnica', 'Oferta', 'Contratada', 'Rechazada', 'Sin respuesta', 'Descartada'];
var ACTIVOS = ['Enviado', 'Visto', 'Entrevista', 'Prueba técnica'];
var CANALES = ['LinkedIn', 'Computrabajo', 'Portal de empresa', 'Email directo', 'Referencia', 'Otro'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🚀 Búsqueda')
    .addItem('➕ Nueva postulación', 'abrirPostulacion')
    .addSeparator()
    .addItem('⚙️ Configurar planilla', 'setupBuscador')
    .addToUi();
  try { avisarEventos2Dias(); } catch (e) {}
}

function onEdit(e) {
  var sh = e.range.getSheet();
  var r = e.range.getRow();
  var c = e.range.getColumn();
  if (sh.getName() !== REG || r < 2 || r > FILAS) return;
  if ((c === 2 || c === 3) && !sh.getRange(r, 1).getValue()) {
    sh.getRange(r, 1).setValue(siguienteId());
  }
}

function setupBuscador() {
  var ss = SpreadsheetApp.getActive();
  prepararHojas(ss);
  configurarPostulaciones(ss);
  buildDashboard(ss);
  ss.getSheetByName(DASH).activate();
  ss.getUi().alert('✅ ¡Listo!', 'Planilla configurada.\n\nUsá 🚀 Búsqueda ▸ ➕ Nueva postulación.', ss.getUi().ButtonSet.OK);
}

function prepararHojas(ss) {
  var base = ss || SpreadsheetApp.getActiveSpreadsheet();
  var reg = base.getSheetByName(REG);
  if (!reg) {
    var s1 = base.getSheetByName('Sheet1');
    if (s1) s1.setName(REG); else base.insertSheet(REG);
  }
  if (!base.getSheetByName(DASH)) base.insertSheet(DASH);
  base.getSheets().forEach(function (s) {
    if (s.getName() !== REG && s.getName() !== DASH) base.deleteSheet(s);
  });
  base.setActiveSheet(base.getSheetByName(REG));
  base.moveActiveSheet(1);
}

function configurarPostulaciones(ss) {
  var sh = ss.getSheetByName(REG);
  sh.clear();
  sh.clearFormats();
  sh.clearConditionalFormatRules();
  var enc = ['ID', 'Empresa', 'Puesto', 'Fecha de envío', 'Canal', 'Estado', 'Salario pretendido', 'Salario ofrecido', 'Próxima entrevista', 'Notas'];
  sh.getRange(1, 1, 1, enc.length).setValues([enc]);
  sh.getRange(1, 1, 1, enc.length).setBackground(PRIMARIO).setFontColor(TEXTO_CLARO).setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 32);
  [40, 150, 150, 100, 110, 100, 110, 110, 110, 160].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange('A2:A' + FILAS).setHorizontalAlignment('center');
  sh.getRange('D2:D' + FILAS).setNumberFormat('dd/mm/yyyy');
  sh.getRange('G2:H' + FILAS).setNumberFormat('"$"#,##0').setHorizontalAlignment('right');
  sh.getRange('I2:I' + FILAS).setNumberFormat('dd/mm/yyyy');
  sh.getRange('J2:J' + FILAS).setWrap(true);
  sh.getRange('E2:E' + FILAS).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(CANALES, true).build());
  sh.getRange('F2:F' + FILAS).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).build());
  var rules = [];
  [['Enviado', '#DBEAFE', '#1E3A8A'], ['Visto', '#FEF3C7', '#92400E'], ['Entrevista', '#EDE9FE', '#6B21A8'], ['Prueba técnica', '#FCE7F3', '#9D174D'], ['Oferta', '#D1FAE5', '#065F46'], ['Contratada', '#A7F3D0', '#047857'], ['Rechazada', '#FEE2E2', '#991B1B'], ['Sin respuesta', '#F3F4F6', '#374151'], ['Descartada', '#E5E7EB', '#6B7280']].forEach(function (c) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(c[0]).setBackground(c[1]).setFontColor(c[2]).setRanges([sh.getRange('F2:F' + FILAS)]).build());
  });
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=MOD(ROW(),2)=0').setBackground('#F8FAFC').setRanges([sh.getRange('A2:J' + FILAS)]).build());
  sh.setConditionalFormatRules(rules);
  sh.getRange(2, 1, FILAS - 1, enc.length).createFilter();
}

function buildDashboard(ss) {
  var sh = ss.getSheetByName(DASH);
  sh.clear();
  sh.clearFormats();
  [140, 55, 140, 15, 110, 45, 50, 130].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange('A1:H1').merge().setValue('🧭 Dashboard · Búsqueda de Dai').setBackground(PRIMARIO).setFontColor(TEXTO_CLARO).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 34);
  var cards = [
    ['TOTAL CVS', '=COUNTA(' + REG + '!B2:B' + FILAS + ')', 1],
    ['ESTE MES', '=COUNTIFS(' + REG + '!D2:D' + FILAS + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),' + REG + '!D2:D' + FILAS + ',"<"&EOMONTH(TODAY(),0)+1)', 3],
    ['EN PROCESO', '=SUMPRODUCT((' + ACTIVOS.map(function (n) { return '(' + REG + '!F2:F' + FILAS + '="' + n + '")'; }).join('+') + '))', 5],
    ['PRÓX. ENTREV', '=COUNTIFS(' + REG + '!I2:I' + FILAS + ',">="&TODAY(),' + REG + '!I2:I' + FILAS + ',"<"&TODAY()+8)', 7]
  ];
  cards.forEach(function (c) {
    sh.getRange(3, c[2], 1, 2).merge().setValue(c[0]).setBackground(PRIMARIO).setFontColor(TEXTO_CLARO).setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(4, c[2], 1, 2).merge().setFormula(c[1]).setBackground(PRIMARIO_CLARO).setFontColor('#1E1B4B').setFontSize(18).setFontWeight('bold').setNumberFormat('#,##0').setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  sh.setRowHeight(3, 20);
  sh.setRowHeight(4, 34);
  sh.getRange('A6:C6').merge().setValue('📊 POR ESTADO').setBackground(PRIMARIO_CLARO).setFontColor('#4338CA').setFontWeight('bold').setFontSize(11);
  ['Estado', 'Cant.', 'Barra'].forEach(function (h, i) { sh.getRange(7, 1 + i).setValue(h).setBackground(PRIMARIO).setFontColor(TEXTO_CLARO).setFontWeight('bold').setFontSize(9).setHorizontalAlignment(i === 1 ? 'center' : 'left'); });
  sh.setRowHeight(7, 20);
  for (var s = 0; s < ESTADOS.length; s++) {
    sh.getRange(8 + s, 1).setValue(ESTADOS[s]).setFontSize(10);
    sh.getRange(8 + s, 2).setFormula('=COUNTIF(' + REG + '!F2:F' + FILAS + ',"' + ESTADOS[s] + '")').setHorizontalAlignment('center').setFontSize(10);
    sh.getRange(8 + s, 3).setFormula('=IF(B' + (8 + s) + '=0,"",REPT("█",ROUND(B' + (8 + s) + '/IF(B' + (8 + ESTADOS.length) + '=0,1,B' + (8 + ESTADOS.length) + ')*20)))').setFontColor(PRIMARIO).setFontSize(9);
    sh.setRowHeight(8 + s, 20);
  }
  sh.getRange(8 + ESTADOS.length, 1).setValue('Total').setFontWeight('bold').setFontSize(10);
  sh.getRange(8 + ESTADOS.length, 2).setFormula('=SUM(B8:B' + (7 + ESTADOS.length) + ')').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(10);
  sh.getRange('E6:H6').merge().setValue('🏆 POR CANAL').setBackground(PRIMARIO_CLARO).setFontColor('#4338CA').setFontWeight('bold').setFontSize(11);
  ['Canal', 'CVs', '%', 'Barra'].forEach(function (h, i) { sh.getRange(7, 5 + i).setValue(h).setBackground(PRIMARIO).setFontColor(TEXTO_CLARO).setFontWeight('bold').setFontSize(9).setHorizontalAlignment(i <= 1 ? 'left' : 'center'); });
  for (var ch = 0; ch < CANALES.length; ch++) {
    sh.getRange(8 + ch, 5).setValue(CANALES[ch]).setFontSize(10);
    sh.getRange(8 + ch, 6).setFormula('=COUNTIF(' + REG + '!E2:E' + FILAS + ',"' + CANALES[ch] + '")').setHorizontalAlignment('center').setFontSize(10);
    sh.getRange(8 + ch, 7).setFormula('=IF(F' + (8 + ch) + '=0,"",TEXT(F' + (8 + ch) + '/F' + (8 + CANALES.length) + ',"0%"))').setHorizontalAlignment('center').setFontSize(9);
    sh.getRange(8 + ch, 8).setFormula('=IF(F' + (8 + ch) + '=0,"",REPT("█",ROUND(F' + (8 + ch) + '/IF(F' + (8 + CANALES.length) + '=0,1,F' + (8 + CANALES.length) + ')*20)))').setFontColor(PRIMARIO).setFontSize(9);
    sh.setRowHeight(8 + ch, 20);
  }
  sh.getRange(8 + CANALES.length, 5).setValue('Total').setFontWeight('bold').setFontSize(10);
  sh.getRange(8 + CANALES.length, 6).setFormula('=SUM(F8:F' + (7 + CANALES.length) + ')').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(10);
  sh.getRange('A19:H19').merge().setValue('📅 PRÓXIMOS EVENTOS · 7 días').setBackground(PRIMARIO_CLARO).setFontColor('#4338CA').setFontWeight('bold').setFontSize(11);
  for (var ev = 0; ev < 6; ev++) {
    sh.getRange(20 + ev, 1, 1, 8).merge().setFontSize(10).setFontColor('#374151').setVerticalAlignment('middle');
    sh.setRowHeight(20 + ev, 22);
  }
  refreshUpcoming(sh);
}

function refreshUpcoming(sh) {
  if (!sh) sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASH);
  if (!sh) return;
  var lista = leerProximos(0, 7);
  for (var i = 0; i < 6; i++) sh.getRange(20 + i, 1, 1, 8).clearContent().setBackground('#FFFFFF').setFontColor('#374151').setFontWeight('normal');
  if (!lista.length) { sh.getRange(20, 1, 1, 8).merge().setValue('Sin eventos próximos 🎈'); return; }
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  lista.slice(0, 6).forEach(function (e, j) {
    var diff = Math.round((e.fecha - hoy) / 86400000);
    var cuando = diff === 0 ? 'HOY' : (diff === 1 ? 'MAÑANA' : Utilities.formatDate(e.fecha, tz, 'EEE d'));
    sh.getRange(20 + j, 1, 1, 8).merge().setValue('▶ ' + cuando + ' · ' + e.empresa);
    if (diff <= 1) sh.getRange(20 + j, 1).setFontColor('#B45309').setFontWeight('bold').setBackground('#FEF3C7');
    else sh.getRange(20 + j, 1).setFontColor('#4338CA');
  });
}

function leerProximos(desde, hastaDias) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REG);
  var out = [], fin = sh.getLastRow();
  if (fin < 2) return out;
  var datos = sh.getRange(2, 1, fin - 1, 10).getValues();
  var hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  datos.forEach(function (row) {
    var f = row[8];
    if (!(f instanceof Date)) return;
    f.setHours(0, 0, 0, 0);
    var diff = (f - hoy) / 86400000;
    if (diff >= desde && diff <= hastaDias) out.push({ fecha: new Date(f), empresa: row[1] || '—' });
  });
  return out.sort(function (a, b) { return a.fecha - b.fecha; });
}

function avisarEventos2Dias() {
  var lista = leerProximos(0, 2);
  if (!lista.length) return;
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  var lineas = lista.map(function (e) {
    var diff = Math.round((e.fecha - hoy) / 86400000);
    var cuando = diff === 0 ? 'HOY' : (diff === 1 ? 'MAÑANA' : Utilities.formatDate(e.fecha, tz, 'EEEE d'));
    return '• ' + cuando + ' · ' + e.empresa;
  });
  SpreadsheetApp.getUi().alert('📅 ¡Tenés entrevistas próximas!', lineas.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function siguienteId() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REG);
  var vals = sh.getRange('A2:A' + FILAS).getValues();
  var max = 0;
  vals.forEach(function (r) { var v = Number(r[0]); if (v > max) max = v; });
  return max + 1;
}

function abrirPostulacion() {
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile('postulacion').setTitle('➕ Nueva postulación'));
}

function parseFecha(s) {
  if (!s) return '';
  var p = s.split('-');
  return new Date(p[0], p[1] - 1, p[2]);
}

function appendPostulacion(obj) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REG);
  var id = siguienteId();
  var fila = 2;
  var col = sh.getRange(2, 2, FILAS - 1, 1).getValues();
  while (fila <= FILAS && col[fila - 2][0]) fila++;
  if (fila > FILAS) throw new Error('No hay más filas disponibles.');
  sh.getRange(fila, 1, 1, 10).setValues([[
    id, obj.empresa || '', obj.puesto || '',
    parseFecha(obj.fecha) || new Date(), obj.canal || 'Otro', obj.estado || 'Enviado',
    obj.salarioPre || '', obj.salarioOf || '', parseFecha(obj.proxEntrevista) || '', obj.notas || ''
  ]]);
  return { id: id };
}
