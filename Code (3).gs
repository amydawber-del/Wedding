// ============================================================
// Wedding Venue Tracker — Apps Script Backend v3
// ============================================================
// SETUP: Deploy → Manage Deployments → edit → New version → Deploy
// URL stays the same — no need to update HTML files.
// ============================================================
//
// 📊 VENUE COMPARISON columns:
//   A  Venue Name       B  Date Viewed      C  Hire Cost (£)
//   D  Cost per Head    E  Guest Count      F  Catering Total (formula)
//   G  Drinks Est       H  Extras           I  Est. Total (formula)
//   J  Cost per Guest   K  All Inclusive?   L  Hidden Cost Notes
//   M  Score /10        N  Shortlist        O  My Status (★/~/✗)
//   P  Partner Score    Q  Follow-up Qs
//
// 📝 VENUE NOTES columns:
//   A  Venue Name       B  Date Viewed      C  Location
//   D  Gut Feeling
//   E  Pricing — Confirmed    F  Pricing — Notes
//   G  Food — Confirmed       H  Food — Notes
//   I  Logistics — Confirmed  J  Logistics — Notes
//   K  Suppliers — Confirmed  L  Suppliers — Notes
//   M  Stays — Confirmed      N  Stays — Notes
//   O  Hidden Costs — Conf    P  Hidden Costs — Notes
//   Q  Vibe — Confirmed       R  Vibe — Notes
//   S  Overall Notes
// ============================================================

var SHEET_COMPARISON = '📊 Venue Comparison';
var SHEET_NOTES      = '📝 Venue Notes';

// ── GET — returns all venue data as JSON for the dashboard ───────────────
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;

  if (action === 'getData') {
    return getVenueData();
  }
  if (action === 'updateStatus') {
    return updateVenueStatus(e.parameter.venue, e.parameter.field, e.parameter.value);
  }
  return jsonResponse({ status: 'ok', message: 'Venue Tracker v3 live.' });
}

function getVenueData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws1 = ss.getSheetByName(SHEET_COMPARISON);
  var ws2 = ss.getSheetByName(SHEET_NOTES);

  if (!ws1) return jsonResponse({ status: 'error', message: 'Comparison sheet not found.' });

  var compData = ws1.getDataRange().getValues();
  var notesData = ws2 ? ws2.getDataRange().getValues() : [];

  // Build notes lookup by venue name (col A = index 0)
  var notesMap = {};
  for (var n = 2; n < notesData.length; n++) {
    var row = notesData[n];
    if (row[0]) notesMap[row[0].toString().trim()] = {
      location:         row[2]  || '',
      gutFeeling:       row[3]  || '',
      basicsConfirmed:  row[4]  || '',
      basicsNotes:      row[5]  || '',
      foodConfirmed:    row[6]  || '',
      foodNotes:        row[7]  || '',
      logisticsConf:    row[8]  || '',
      logisticsNotes:   row[9]  || '',
      suppliersConf:    row[10] || '',
      suppliersNotes:   row[11] || '',
      staysConf:        row[12] || '',
      staysNotes:       row[13] || '',
      hiddenConf:       row[14] || '',
      hiddenNotes:      row[15] || '',
      vibeConf:         row[16] || '',
      vibeNotes:        row[17] || '',
      overallNotes:     row[18] || '',
    };
  }

  var venues = [];
  // Data rows start at row 4 (index 3) based on our sheet structure
  for (var i = 3; i < compData.length; i++) {
    var r = compData[i];
    if (!r[0] || r[0].toString().trim() === '') continue;
    var name = r[0].toString().trim();
    var notes = notesMap[name] || {};

    // Risk scoring based on hidden costs checklist
    var hiddenConf = (notes.hiddenConf || '').toLowerCase();
    var riskItems = ['tables','linen','glassware','staff','cake cutting','sound','lighting','cleaning','security'];
    var unconfirmed = riskItems.filter(function(item){ return hiddenConf.indexOf(item) === -1; }).length;
    var riskLevel = unconfirmed >= 6 ? 'High' : unconfirmed >= 3 ? 'Medium' : 'Low';
    var hasSupplierRestrictions = (notes.suppliersConf || '').toLowerCase().indexOf('mandatory') !== -1;
    var hasNoRainPlan = (notes.logisticsConf || '').toLowerCase().indexOf('wet weather') === -1;
    var flags = [];
    if (unconfirmed >= 5) flags.push('Hidden cost risks unconfirmed');
    if (hasSupplierRestrictions) flags.push('Mandatory supplier list');
    if (hasNoRainPlan) flags.push('No wet weather plan confirmed');

    var hireCost   = toNum(r[2]);
    var cph        = toNum(r[3]);
    var guests     = toNum(r[4]);
    var drinks     = toNum(r[6]);
    var extras     = toNum(r[7]);
    var totalCost  = hireCost + (cph * guests) + drinks + extras;
    var cpg        = guests > 0 ? totalCost / guests : 0;

    venues.push({
      name:           name,
      dateViewed:     r[1] ? Utilities.formatDate(new Date(r[1]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      hireCost:       hireCost,
      costPerHead:    cph,
      guestCount:     guests,
      drinksEst:      drinks,
      extras:         extras,
      totalCost:      Math.round(totalCost),
      costPerGuest:   Math.round(cpg),
      allInclusive:   r[10] || '',
      hiddenCostNotes:r[11] || '',
      score:          toNum(r[12]),
      shortlist:      r[13] || '',
      myStatus:       r[14] || '',
      partnerScore:   toNum(r[15]),
      followUpQs:     r[16] || '',
      riskLevel:      riskLevel,
      riskFlags:      flags,
      notes:          notes,
    });
  }

  return jsonResponse({ status: 'success', venues: venues });
}

// ── updateStatus — PATCH a single field on the Comparison sheet ──────────
function updateVenueStatus(venueName, field, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws1 = ss.getSheetByName(SHEET_COMPARISON);
  if (!ws1) return jsonResponse({ status: 'error', message: 'Sheet not found.' });

  var fieldColMap = { myStatus: 15, partnerScore: 16, shortlist: 14, followUpQs: 17 };
  var col = fieldColMap[field];
  if (!col) return jsonResponse({ status: 'error', message: 'Unknown field.' });

  var data = ws1.getDataRange().getValues();
  for (var i = 3; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === venueName.toString().trim()) {
      ws1.getRange(i + 1, col).setValue(value);
      return jsonResponse({ status: 'success' });
    }
  }
  return jsonResponse({ status: 'error', message: 'Venue not found.' });
}

// ── POST — receives checklist submissions ────────────────────────────────
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    writeToSheets(payload);
    return jsonResponse({ status: 'success', message: 'Venue saved.' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function writeToSheets(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currencyFmt = '£#,##0;(£#,##0);"-"';

  var ws1 = ss.getSheetByName(SHEET_COMPARISON);
  if (!ws1) throw new Error('Sheet "' + SHEET_COMPARISON + '" not found.');
  ensureComparisonHeaders(ws1);

  var r1 = getNextDataRow(ws1, 4);
  ws1.getRange(r1, 1).setValue(p.venueName || '');
  ws1.getRange(r1, 2).setValue(p.dateViewed ? new Date(p.dateViewed) : '').setNumberFormat('dd/mm/yyyy');
  ws1.getRange(r1, 3).setValue(toNum(p.hireCost));
  ws1.getRange(r1, 4).setValue(toNum(p.costPerHead));
  ws1.getRange(r1, 5).setValue(toNum(p.guestCount));
  ws1.getRange(r1, 6).setFormula('=D' + r1 + '*E' + r1);
  ws1.getRange(r1, 7).setValue(toNum(p.drinksEst));
  ws1.getRange(r1, 8).setValue(toNum(p.extras));
  ws1.getRange(r1, 9).setFormula('=C' + r1 + '+F' + r1 + '+G' + r1 + '+H' + r1);
  ws1.getRange(r1, 10).setFormula('=IF(E' + r1 + '=0,"",I' + r1 + '/E' + r1 + ')');
  ws1.getRange(r1, 11).setValue('');
  ws1.getRange(r1, 12).setValue(p.sneakyChecked || '');
  ws1.getRange(r1, 13).setValue(p.score || '');
  ws1.getRange(r1, 14).setValue('');
  ws1.getRange(r1, 15).setValue('');   // myStatus
  ws1.getRange(r1, 16).setValue('');   // partnerScore
  ws1.getRange(r1, 17).setValue('');   // followUpQs

  [3,4,6,7,8,9,10].forEach(function(col){ ws1.getRange(r1, col).setNumberFormat(currencyFmt); });
  ws1.getRange(r1, 12).setWrap(true);

  var ws2 = ss.getSheetByName(SHEET_NOTES);
  if (!ws2) throw new Error('Sheet "' + SHEET_NOTES + '" not found.');
  ensureNotesHeaders(ws2);

  var r2 = getNextDataRow(ws2, 3);
  ws2.getRange(r2, 1).setValue(p.venueName || '');
  ws2.getRange(r2, 2).setValue(p.dateViewed ? new Date(p.dateViewed) : '').setNumberFormat('dd/mm/yyyy');
  ws2.getRange(r2, 3).setValue(p.location || '');
  ws2.getRange(r2, 4).setValue(p.gutFeeling || '');

  var sectionData = [
    [p.basicsChecked,    p.basicsNotes],
    [p.foodChecked,      p.foodNotes],
    [p.logisticsChecked, p.logisticsNotes],
    [p.suppliersChecked, p.suppliersNotes],
    [p.staysChecked,     p.staysNotes],
    [p.sneakyChecked,    p.sneakyNotes],
    [p.vibeChecked,      p.vibeNotes],
  ];
  sectionData.forEach(function(sec, i) {
    var col = 5 + (i * 2);
    ws2.getRange(r2, col).setValue(sec[0] || '');
    ws2.getRange(r2, col + 1).setValue(sec[1] || '');
  });
  ws2.getRange(r2, 19).setValue(p.overallNotes || '');
  ws2.getRange(r2, 1, 1, 19).setWrap(true);
}

function ensureComparisonHeaders(ws) {
  if (ws.getRange(3, 1).getValue() !== '') return;
  var h = ['Venue Name','Date Viewed','Hire Cost (£)','Cost per Head (£)','Guest Count',
    'Catering Total (£)','Drinks Estimate (£)','Extras (£)','Estimated Total (£)',
    'Cost per Guest (£)','All Inclusive?','Hidden Cost Notes','Score /10','Shortlist',
    'My Status','Partner Score','Follow-up Questions'];
  ws.getRange(3, 1, 1, h.length).setValues([h]);
}

function ensureNotesHeaders(ws) {
  if (ws.getRange(2, 1).getValue() !== '') return;
  var h = ['Venue Name','Date Viewed','Location','Gut Feeling',
    'Pricing — Confirmed','Pricing — Notes','Food — Confirmed','Food — Notes',
    'Logistics — Confirmed','Logistics — Notes','Suppliers — Confirmed','Suppliers — Notes',
    'Stays — Confirmed','Stays — Notes','Hidden Costs — Confirmed','Hidden Costs — Notes',
    'Vibe — Confirmed','Vibe — Notes','Overall Notes'];
  ws.getRange(2, 1, 1, h.length).setValues([h]);
  var hdr = ws.getRange(2, 1, 1, h.length);
  hdr.setFontWeight('bold').setBackground('#9b6b7a').setFontColor('#ffffff').setWrap(true);
  for (var i = 0; i < 7; i++) {
    ws.getRange(3, 5 + i*2, 100, 2).setBackground(i % 2 === 0 ? '#f5dde2' : '#fdf6f8');
  }
}

function getNextDataRow(sheet, startRow) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(startRow, 1, Math.max(lastRow - startRow + 2, 1), 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0] || data[i][0].toString().trim() === '') return startRow + i;
  }
  return lastRow + 1;
}

function toNum(val) { var n = parseFloat(val); return isNaN(n) ? 0 : n; }

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
