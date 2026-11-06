// Google Sheet: Extensions > Apps Script > الصق هذا الكود > Deploy > New deployment
// Type: Web app | Execute as: Me | Who has access: Anyone  → انسخ الرابط (ينتهي بـ /exec)
function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Colis') || ss.insertSheet('Colis');
  sh.clear();
  sh.getRange('A:A').setNumberFormat('@');
  sh.getRange('E:E').setNumberFormat('@');
  var head = ['Tracking','Destinataire','Wilaya','Commune','Téléphone','Montant','Statut','Premier échec','Échecs (date + raison)','Remarque'];
  var rows = data.colis.map(function (c) {
    return [c.tracking, c.nom, c.wilaya, c.commune, c.tel, c.montant, c.statut, c.firstFail || '',
      (c.echecs || []).map(function (x) { return x.date + ' ' + x.raison; }).join(' | '), c.remarque || ''];
  });
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#ea6a0a').setFontColor('#ffffff');
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  return ContentService.createTextOutput('ok');
}
