/***********************
 * Project Delivery Log — Versión optimizada (completa)
 * Autor: Amílcar + Asistente Senior
 ***********************/

const NAMES = {
  rooms: 'rooms',
  roomsMaster: 'rooms_master',
  deliveries: 'deliveries_log',
  report: 'report_matrix',
  users: 'users',
  sourceWide: 'source_wide',
  status: 'status_overview',
};

const BED_BASE_CODES = new Set(['G-113A','G-113','G-114A','G-114']); // bases Queen ADA, Queen, King ADA, King

// Orden de columnas visibles en report_matrix
const ITEM_COLUMNS = [
  'G-113 Queen Bed Base',
  'BedNumbers (Queen)',
  'G-114 King Bed Base',
  'BedNumbers (King)',
  'G-105B Queen Headboard',
  'G-105A King Headboard',
  'G-103 RAF Desk',
  'G-103 LAF Desk',
  'G-103 Desk NO REF',
  'G-112 RAF Luggage',
  'G-112 LAF Luggage',
  'G-111 Side Chair',
  'G-101 Side Table',
  'G-102 Desk Chair',
  'G-104 Nightstand',
  'G-106 Mirror',
  'G-108 Pta Cover',
  'G-100 Lounge Chair',
  'MB-190 Round Wall Light',
  'MB-280 Oval Wall Light',
  'MB-180 Reading Light',
];

/* =========================
   CACHÉ (servidor) — rendimiento
========================= */
const CACHE_TTL = 600; // 10 min
function cacheGet_(key) {
  try {
    const raw = CacheService.getScriptCache().get(key);
    if (!raw) return null;
    if (raw === '__bust__') return null;
    return JSON.parse(raw);
  } catch (_) { return null; }
}
function cachePut_(key, obj, ttlSec) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(obj), Math.max(5, ttlSec || CACHE_TTL));
  } catch (err) { Logger.log('cachePut_ error: ' + err); }
}
function cacheBust_(key) {
  try {
    CacheService.getScriptCache().put(key, '__bust__', 5);
  } catch (err) { Logger.log('cacheBust_ error: ' + err); }
}

/* =========================
   MENÚ ÚNICO
========================= */
function onOpen() { buildMenu(); }

function buildMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Delivery Log')
    .addItem('Validar fila actual (debug)','menuValidateActiveRow')
    .addSeparator()
    .addItem('Generar Reporte (fecha)','menuBuildReportForDate')
    .addItem('Limpiar report_matrix','menuClearReport')
    .addSeparator()
    .addItem('Exportar PDF (fecha)','menuExportPdfForDate')
    .addSeparator()
    .addItem('Generar Estado General','menuBuildStatusOverview')
    .addSeparator()
    .addItem('Importar datos (source_wide → rooms, rooms_master)','menuImportFromSourceWide')
    .addToUi();
}

/* =========================
   VALIDACIÓN EN TIEMPO REAL + DROPDOWNS DEPENDIENTES
========================= */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== NAMES.deliveries) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row <= 1) return;

    // Validar fila entera
    validateDeliveryRow(row, true);

    // Dropdowns dependientes
    if (col === 2) { // Floor
      applyRoomValidationForRow(row);
      sh.getRange(row, 3).clearContent().clearDataValidations(); // Room
      sh.getRange(row, 4).clearContent().clearDataValidations(); // ItemCode
    } else if (col === 3) { // Room
      sh.getRange(row, 4).clearContent().clearDataValidations();
      applyItemValidationForRow(row);
    }
  } catch (err) {
    Logger.log(err);
  }
}

/* =========================
   MENÚS — Acciones directas
========================= */
function menuValidateActiveRow() {
  const sh = getSheet(NAMES.deliveries);
  const r = sh.getActiveRange();
  if (!r) return;
  const row = r.getRow();
  if (row <= 1) return uiAlert('Selecciona una fila de datos (A2 en adelante).');
  const ok = validateDeliveryRow(row, false);
  uiAlert(ok ? 'Fila válida ✅' : 'Fila con errores ❌ (revisa celdas resaltadas).');
}

function menuBuildReportForDate() {
  const dateStr = uiPrompt('Fecha del reporte (YYYY-MM-DD):');
  if (!dateStr) return;
  if (!parseIsoDate(dateStr)) return uiAlert('Fecha inválida. Usa YYYY-MM-DD.');
  buildReportMatrixForDate(dateStr);
  uiAlert('report_matrix generado para ' + dateStr + ' ✅');
}

function menuClearReport() {
  const sh = getSheet(NAMES.report);
  sh.clearContents();
  setReportHeaders(sh);
  uiAlert('report_matrix limpio ✅');
}

function menuExportPdfForDate() {
  const dateStr = uiPrompt('Fecha del reporte (YYYY-MM-DD):');
  if (!dateStr) return;
  if (!parseIsoDate(dateStr)) return uiAlert('Fecha inválida. Usa YYYY-MM-DD.');
  buildReportMatrixForDate(dateStr);
  const url = exportReportMatrixToPdf(dateStr);
  uiAlert('PDF generado ✅\n\n' + url);
}

function menuBuildStatusOverview() {
  buildStatusOverview();
  uiAlert('Estado general generado en "'+NAMES.status+'". ✅');
}

function menuImportFromSourceWide() {
  importFromSourceWide();
  uiAlert('Importación completada ✅\nSe actualizaron "rooms" y "rooms_master".');
}

/* =========================
   VALIDACIONES DE FILAS (deliveries_log)
========================= */
function validateDeliveryRow(row, silent) {
  const sh = getSheet(NAMES.deliveries);
  // A..H: Date, Floor, Room, ItemCode, QuantityDelivered, BedNumbers, VerifiedBy, Notes
  const vals = sh.getRange(row, 1, 1, 8).getValues()[0];
  const [date, floor, room, item, qty, bedNumbers, verifiedBy] = vals;

  // Reset fondo
  sh.getRange(row,1,1,8).setBackground(null).setNote('');

  let ok = true;

  // 1) Fecha
  if (!(date instanceof Date)) { markError(sh, row, 1, 'Fecha inválida'); ok = false; }

  // 2) Floor existe
  const floors = getColumnValues(getSheet(NAMES.rooms), 1);
  if (!floors.has(String(floor))) { markError(sh, row, 2, 'Floor no existe'); ok = false; }

  // 3) Room pertenece al Floor
  const pair = String(floor)+'|'+String(room);
  const roomPairs = getPairsFloorRoom();
  if (!roomPairs.has(pair)) { markError(sh, row, 3, 'Room no corresponde al Floor'); ok = false; }

  // 4) Item válido para cuarto
  const validItems = getValidItemsForRoom(String(floor), String(room));
  if (!validItems.has(String(item))) { markError(sh, row, 4, 'ItemCode no válido para el cuarto'); ok = false; }

  // 5) Quantity >= 0
  if (isNaN(qty) || Number(qty) < 0) { markError(sh, row, 5, 'Cantidad inválida'); ok = false; }

  // 6) BedNumbers requerido para bases + conteo
  if (BED_BASE_CODES.has(String(item))) {
    const needed = Number(qty);
    const list = (String(bedNumbers || '')).trim();
    if (!list) { markError(sh, row, 6, 'BedNumbers requerido'); ok = false; }
    else {
      const count = list.split(',').map(s => s.trim()).filter(Boolean).length;
      if (count !== needed) markWarn(sh, row, 6, 'Cantidad de BedNumbers ≠ QuantityDelivered');
    }
  }

  // 7) VerifiedBy si viene debe existir en users!A
  if (verifiedBy) {
    const users = getColumnValues(getSheet(NAMES.users), 1);
    if (!users.has(String(verifiedBy))) markWarn(sh, row, 7, 'VerifiedBy no está en users');
  }

  if (!silent && !ok) SpreadsheetApp.getUi().alert('La fila '+row+' contiene errores.');
  return ok;
}

/* =========================
   DROPDOWNS DEPENDIENTES
========================= */
function applyRoomValidationForRow(row) {
  const shDel = getSheet(NAMES.deliveries);
  const floor = String(shDel.getRange(row, 2).getValue() || '').trim();
  const cell = shDel.getRange(row, 3); // Room
  if (!floor) { cell.clearDataValidations(); return; }

  const shRooms = getSheet(NAMES.rooms);
  const data = shRooms.getDataRange().getValues();
  const head = data.shift();
  const iF = head.indexOf('Floor');
  const iR = head.indexOf('Room');
  const roomsForFloor = [];
  data.forEach(r => {
    if (String(r[iF]).trim() === floor) {
      const rm = String(r[iR]).trim();
      if (rm) roomsForFloor.push(rm);
    }
  });

  if (!roomsForFloor.length) { cell.clearDataValidations(); return; }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(roomsForFloor, true)
    .setAllowInvalid(false)
    .setHelpText('Selecciona un Room válido para el Floor '+floor)
    .build();
  cell.setDataValidation(rule);
}

function applyItemValidationForRow(row) {
  const shDel = getSheet(NAMES.deliveries);
  const floor = String(shDel.getRange(row, 2).getValue() || '').trim();
  const room  = String(shDel.getRange(row, 3).getValue() || '').trim();
  const cell  = shDel.getRange(row, 4); // ItemCode
  if (!floor || !room) { cell.clearDataValidations(); return; }

  const shM = getSheet(NAMES.roomsMaster);
  const vals = shM.getDataRange().getValues();
  if (vals.length < 2) { cell.clearDataValidations(); return; }
  const h = vals[0];
  const iF = h.indexOf('Floor'), iR = h.indexOf('Room'), iC = h.indexOf('ItemCode'), iQ = h.indexOf('QuantityRequired');

  const items = [];
  for (let i=1;i<vals.length;i++){
    const r = vals[i];
    if (String(r[iF]).trim()===floor && String(r[iR]).trim()===room) {
      const code = String(r[iC]).trim();
      const qty  = Number(r[iQ] || 0);
      if (code && qty>0) items.push(code);
    }
  }
  if (!items.length) { cell.clearDataValidations(); return; }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(items, true)
    .setAllowInvalid(false)
    .setHelpText('Selecciona un ItemCode válido para '+floor+'-'+room)
    .build();
  cell.setDataValidation(rule);
}

/* =========================
   REPORT MATRIX (por fecha)
========================= */
function buildReportMatrixForDate(dateStr) {
  const shReport = getOrCreateSheet(NAMES.report);
  shReport.clearContents();
  setReportHeaders(shReport);

  const shD = getSheet(NAMES.deliveries);
  const data = shD.getDataRange().getValues();
  const head = data.shift(); // Date, Floor, Room, Item, Qty, Bed, ...
  const IDX = {Date:0, Floor:1, Room:2, Item:3, Qty:4, Bed:5};

  const matrix = new Map(); // key FR -> {floor,room,items:Map, bedQ:[], bedK:[]}
  const itemNameByCode = getItemNameByCode();
  const isQueen = new Set(['G-113','G-113A']);
  const isKing  = new Set(['G-114','G-114A']);

  data.forEach(row => {
    const d = row[IDX.Date];
    if (!isSameISODate(d, dateStr)) return;
    const floor = String(row[IDX.Floor]).trim();
    const room  = String(row[IDX.Room]).trim();
    const itemCode = String(row[IDX.Item]).trim();
    const qty = Number(row[IDX.Qty] || 0);
    const beds = String(row[IDX.Bed] || '').trim();

    const key = floor+'|'+room;
    if (!matrix.has(key)) matrix.set(key, {floor, room, items:new Map(), bedQ:[], bedK:[]});
    const rec = matrix.get(key);

    const colName = codeToReportColumn(itemCode, itemNameByCode);
    if (colName) {
      const prev = rec.items.get(colName) || 0;
      rec.items.set(colName, prev + qty);
    }
    if (beds) {
      if (isQueen.has(itemCode)) rec.bedQ.push(beds);
      if (isKing.has(itemCode))  rec.bedK.push(beds);
    }
  });

  const rowsOut = [];
  matrix.forEach(rec => {
    const row = new Array(3 + ITEM_COLUMNS.length + 1).fill('');
    row[0]=dateStr; row[1]=rec.floor; row[2]=rec.room;
    let roomTotal = 0;
    ITEM_COLUMNS.forEach((col, idx) => {
      const abs = 3 + idx;
      if (col === 'BedNumbers (Queen)') row[abs] = rec.bedQ.join(' | ');
      else if (col === 'BedNumbers (King)') row[abs] = rec.bedK.join(' | ');
      else {
        const q = rec.items.get(col) || 0;
        row[abs] = q > 0 ? q : '';
        if (q > 0) roomTotal += q;
      }
    });
    row[row.length - 1] = roomTotal;
    rowsOut.push(row);
  });

  rowsOut.sort((a,b)=> (Number(a[1])||0)-(Number(b[1])||0) || (Number(a[2])||0)-(Number(b[2])||0));
  if (rowsOut.length) shReport.getRange(2,1,rowsOut.length,rowsOut[0].length).setValues(rowsOut);

  addTotalsRow(shReport);
}

function setReportHeaders(sh) {
  const headers = ['Date','Floor','Room', ...ITEM_COLUMNS, 'Room Total'];
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function addTotalsRow(sh) {
  const lastRow = Math.max(sh.getLastRow(), 1);
  const lastCol = sh.getLastColumn();
  const totalsRow = lastRow + 1;
  if (lastRow === 1) { sh.getRange(totalsRow,1).setValue('Total'); return; }

  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  const rowVals = new Array(lastCol).fill('');
  rowVals[0] = 'Total';
  for (let c=4; c<=lastCol-1; c++) {
    const colHeader = headers[c-1];
    if (/^BedNumbers/.test(colHeader)) continue; // texto
    rowVals[c-1] = '=SUM(' + sh.getRange(2,c,lastRow-1,1).getA1Notation() + ')';
  }
  sh.getRange(totalsRow,1,1,lastCol).setValues([rowVals]).setFontWeight('bold');
}

/* =========================
   EXPORTAR PDF
========================= */
function exportReportMatrixToPdf(dateStr) {
  const sh = getSheet(NAMES.report);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error('report_matrix está vacío. Genera primero el reporte para la fecha.');

  const headers = values[0];
  const rows = values.slice(1).filter(r => isSameISODate(r[0], dateStr));
  if (!rows.length) throw new Error('No hay filas en report_matrix para la fecha: ' + dateStr);

  const floorMap = new Map();
  rows.forEach(r => {
    const floor = String(r[1]).trim();
    if (!floorMap.has(floor)) floorMap.set(floor, []);
    floorMap.get(floor).push(r);
  });

  const doc = DocumentApp.create('Delivery Log — ' + dateStr);
  const body = doc.getBody();
  body.clear();

  body.appendParagraph('Delivery Log — Project Summary').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Date: ' + dateStr).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('This document summarizes the items delivered by floor and room for the selected date. The Project Manager certifies that all listed items were received in good condition, as detailed below.')
    .setSpacingAfter(10);

  floorMap.forEach((rowsForFloor, floor) => {
    body.appendParagraph('Floor ' + floor).setHeading(DocumentApp.ParagraphHeading.HEADING3);

    const tableData = [];
    tableData.push(headers.map(h=>String(h)));
    rowsForFloor.forEach(r => {
      const clean = r.map((v, i) => i<=2 ? v : (typeof v==='number' ? (v>0?v:'') : (v||'')));
      tableData.push(clean.map(x=>String(x)));
    });

    const table = body.appendTable(tableData);
    table.setBorderWidth(0.5);
    table.getRow(0).editAsText().setBold(true);
    body.appendParagraph('');
  });

  body.appendParagraph('Delivery Confirmation & Sign-Off').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('I hereby confirm that all items listed in this delivery log have been delivered in full and in good condition.');

  const signTable = body.appendTable([
    ['Delivered by (Warehouse/PM)', '', 'Received by (Hotel/PM)', ''],
    ['Name:', '______________________________', 'Name:', '______________________________'],
    ['Signature:', '______________________________', 'Signature:', '______________________________'],
    ['Date:', '______________________________', 'Date:', '______________________________'],
  ]);
  signTable.setBorderWidth(0.5);

  body.appendParagraph('\nNOTES:\n\n____________________________________________________________________________\n\n____________________________________________________________________________\n');

  doc.saveAndClose();

  const blob = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
  blob.setName('Delivery Log - ' + dateStr + '.pdf');
  const pdfFile = DriveApp.createFile(blob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // Opcional: borrar el Doc original
  // DriveApp.getFileById(doc.getId()).setTrashed(true);
  return pdfFile.getUrl();
}

/* =========================
   ESTADO GENERAL (status_overview)
========================= */
function buildStatusOverview() {
  const shStatus = getOrCreateSheet(NAMES.status);
  const roomsMeta = readRoomsMeta();   // Map FR -> {roomType, ada}
  const master    = readRoomsMaster();  // Map FR -> Map<ItemCode, qtyReq>
  const delivered = readDeliveries();   // Map FR -> {byItem:Map, lastDate:Date|null}

  const out = [];
  master.forEach((reqMap, keyFR) => {
    const [floor,room] = keyFR.split('|');
    const meta = roomsMeta.get(keyFR) || {};
    const reqTotal = sumMap(reqMap);

    const d = delivered.get(keyFR);
    let deliveredTotalCapped = 0;
    let lastDate = '';
    if (d) {
      reqMap.forEach((reqQty, code) => {
        const got = Number(d.byItem.get(code) || 0);
        deliveredTotalCapped += Math.min(reqQty, got);
      });
      lastDate = d.lastDate ? formatISO(d.lastDate) : '';
    }

    const missing = [];
    reqMap.forEach((reqQty, code) => {
      const got = Number(d?.byItem.get(code) || 0);
      if (reqQty > got) missing.push(code + ' x' + (reqQty - got));
    });

    let status = 'Not Delivered';
    if (deliveredTotalCapped > 0 && missing.length > 0) status = 'Partial';
    if (missing.length === 0 && reqTotal > 0) status = 'Complete';

    const pct = reqTotal > 0 ? deliveredTotalCapped / reqTotal : 0;
    out.push([Number(floor), room, meta.roomType || '', meta.ada || '', reqTotal, deliveredTotalCapped, pct, status, lastDate, missing.join('; '), '']);
  });

  out.sort((a,b)=> a[0]-b[0] || (Number(a[1])||0)-(Number(b[1])||0));

  const headers = ['Floor','Room','RoomType','ADA','Required','Delivered','%','Status','LastDate','Missing','Notes'];
  shStatus.clearContents();
  shStatus.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  if (out.length) {
    shStatus.getRange(2,1,out.length,headers.length).setValues(out);
    const pctCol = headers.indexOf('%') + 1;
    shStatus.getRange(2, pctCol, out.length, 1).setNumberFormat('0.0%');
  } else {
    shStatus.getRange(2,1).setValue('No data found. Check rooms_master and deliveries_log.');
  }
  shStatus.setFrozenRows(1);
}

/* =========================
   IMPORTADOR source_wide → rooms / rooms_master
========================= */
function importFromSourceWide(){
  const ss = SpreadsheetApp.getActive();
  const shSrc = ss.getSheetByName(NAMES.sourceWide);
  if (!shSrc) throw new Error('No existe la hoja "'+NAMES.sourceWide+'".');

  const vals = shSrc.getDataRange().getValues();
  if (vals.length < 2) throw new Error('La hoja "'+NAMES.sourceWide+'" no tiene datos.');

  // Cabeceras fijas de metadatos
  const HEAD_FLOOR='Floor', HEAD_ROOMNUM='Room Number', HEAD_TYPE='Room Type', HEAD_ADA='ADA Compliance', HEAD_DESK='Desk Orientation', HEAD_LUGG='Luggage Rack Orientation';
  const head = vals[0].map(s => String(s).trim());
  const idx = {
    floor: head.indexOf(HEAD_FLOOR),
    roomn: head.indexOf(HEAD_ROOMNUM),
    rtype: head.indexOf(HEAD_TYPE),
    ada:   head.indexOf(HEAD_ADA),
    desk:  head.indexOf(HEAD_DESK),
    lugg:  head.indexOf(HEAD_LUGG),
  };
  Object.keys(idx).forEach(k=>{
    if (idx[k] < 0) throw new Error('Falta la columna requerida: ' + (
      k==='floor'?HEAD_FLOOR:
      k==='roomn'?HEAD_ROOMNUM:
      k==='rtype'?HEAD_TYPE:
      k==='ada'?HEAD_ADA:
      k==='desk'?HEAD_DESK:
      k==='lugg'?HEAD_LUGG:k
    ));
  });

  // Todas las demás columnas se consideran ítems
  const itemCols = [];
  for (let c=0; c<head.length; c++){
    if (Object.values(idx).includes(c)) continue;
    const name = String(head[c]).replace(/\s+/g,' ').trim();
    if (name) itemCols.push({c, name});
  }
  if (!itemCols.length) throw new Error('No se detectaron columnas de ítems en "'+NAMES.sourceWide+'".');

  // ROOMS (únicos por Floor|Room)
  const roomsMap = new Map();
  for (let r=1; r<vals.length; r++){
    const row = vals[r];
    const floor = String(row[idx.floor]).trim();
    const room  = String(row[idx.roomn]).trim();
    if (!floor || !room) continue;
    const key = floor+'|'+room;
    if (!roomsMap.has(key)) {
      roomsMap.set(key, {
        floor,
        room,
        roomType: String(row[idx.rtype]||'').trim(),
        ada:      String(row[idx.ada]  ||'').trim(),
        desk:     String(row[idx.desk] ||'').trim(),
        lugg:     String(row[idx.lugg] ||'').trim(),
      });
    }
  }

  // Volcar ROOMS
  const shRooms = getOrCreateSheet(NAMES.rooms);
  shRooms.getDataRange().clearDataValidations();
  shRooms.clearContents();
  const roomsHeaders = ['Floor','Room','RoomType','ADA','DeskOrientation','LuggageOrientation'];
  shRooms.getRange(1,1,1,roomsHeaders.length).setValues([roomsHeaders]).setFontWeight('bold');

  const roomsOut = Array.from(roomsMap.values())
    .sort((a,b)=> Number(a.floor)-Number(b.floor) || (Number(a.room)||0)-(Number(b.room)||0))
    .map(o => [o.floor, o.room, o.roomType, o.ada, o.desk, o.lugg]);
  if (roomsOut.length) shRooms.getRange(2,1,roomsOut.length,roomsHeaders.length).setValues(roomsOut);
  shRooms.setFrozenRows(1);

  // Volcar ROOMS_MASTER (formato largo)
  const shMaster = getOrCreateSheet(NAMES.roomsMaster);
  shMaster.getDataRange().clearDataValidations();
  shMaster.clearContents();
  const masterHeaders = ['Floor','Room','RoomType','ADA','DeskOrientation','LuggageOrientation','ItemCode','QuantityRequired'];
  shMaster.getRange(1,1,1,masterHeaders.length).setValues([masterHeaders]).setFontWeight('bold');

  const masterOut = [];
  for (let r=1; r<vals.length; r++){
    const row = vals[r];
    const floor = String(row[idx.floor]).trim();
    const room  = String(row[idx.roomn]).trim();
    if (!floor || !room) continue;

    const roomType = String(row[idx.rtype]||'').trim();
    const ada      = String(row[idx.ada]  ||'').trim();
    const desk     = String(row[idx.desk] ||'').trim();
    const lugg     = String(row[idx.lugg] ||'').trim();

    itemCols.forEach(ic=>{
      const qty = Number(row[ic.c] || 0);
      if (qty > 0) masterOut.push([floor, room, roomType, ada, desk, lugg, ic.name, qty]);
    });
  }
  masterOut.sort((a,b)=> Number(a[0])-Number(b[0]) || (Number(a[1])||0)-(Number(b[1])||0) || String(a[6]).localeCompare(String(b[6])));
  if (masterOut.length) shMaster.getRange(2,1,masterOut.length,masterHeaders.length).setValues(masterOut);
  shMaster.setFrozenRows(1);

  // Re-aplicar VALIDACIONES (solo desde fila 2)
  const roomsDataRows  = Math.max(0, shRooms.getLastRow()  - 1);
  const masterDataRows = Math.max(0, shMaster.getLastRow() - 1);

  if (roomsDataRows > 0) {
    // ADA (Yes/No) col D
    const adaRange = shRooms.getRange(2, 4, roomsDataRows, 1);
    const ruleADA = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes','No'], true)
      .setAllowInvalid(false)
      .setHelpText('Use Yes or No')
      .build();
    adaRange.setDataValidation(ruleADA);

    // DeskOrientation (RAF/LAF/NO REF) col E
    const deskRange = shRooms.getRange(2, 5, roomsDataRows, 1);
    const ruleDesk = SpreadsheetApp.newDataValidation()
      .requireValueInList(['RAF','LAF','NO REF'], true)
      .setAllowInvalid(false)
      .setHelpText('Use RAF, LAF or NO REF')
      .build();
    deskRange.setDataValidation(ruleDesk);

    // LuggageOrientation (RAF/LAF) col F
    const luggRange = shRooms.getRange(2, 6, roomsDataRows, 1);
    const ruleLugg = SpreadsheetApp.newDataValidation()
      .requireValueInList(['RAF','LAF'], true)
      .setAllowInvalid(false)
      .setHelpText('Use RAF or LAF')
      .build();
    luggRange.setDataValidation(ruleLugg);
  }

  if (masterDataRows > 0) {
    // QuantityRequired numérico ≥ 0 (col H)
    const qtyRange = shMaster.getRange(2, 8, masterDataRows, 1);
    const ruleQty = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThanOrEqualTo(0)
      .setAllowInvalid(false)
      .setHelpText('Enter a number ≥ 0')
      .build();
    qtyRange.setDataValidation(ruleQty);
    qtyRange.setNumberFormat('0'); // sin decimales
  }
}

/* =========================
   HELPERS GENERALES
========================= */
function getSheet(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('No existe sheet: '+name);
  return sh;
}
function getOrCreateSheet(name){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}
function getSheetSafe(name) { return getSheet(name); }

function uiAlert(msg){ SpreadsheetApp.getUi().alert(msg); }
function uiPrompt(msg){
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(msg);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  const v = r.getResponseText().trim();
  return v || null;
}

function markError(sh, row, col, note) { sh.getRange(row, col).setBackground('#ffd6d6').setNote(note || 'Error'); }
function markWarn(sh, row, col, note)  { sh.getRange(row, col).setBackground('#fff4cc').setNote(note || 'Aviso'); }

function getColumnValues(sh, colIdx) {
  const last = sh.getLastRow();
  if (last < 2) return new Set();
  const vals = sh.getRange(2, colIdx, last-1, 1).getValues().flat().map(v=>String(v).trim()).filter(Boolean);
  return new Set(vals);
}
function getPairsFloorRoom() {
  const sh = getSheet(NAMES.rooms);
  const last = sh.getLastRow();
  if (last < 2) return new Set();
  const rng = sh.getRange(2,1,last-1,2).getValues();
  const set = new Set();
  rng.forEach(r => {
    const f = String(r[0]).trim();
    const rm = String(r[1]).trim();
    if (f && rm) set.add(f+'|'+rm);
  });
  return set;
}
function getValidItemsForRoom(floor, room) {
  const sh = getSheet(NAMES.roomsMaster);
  const data = sh.getDataRange().getValues();
  const head = data.shift();
  const iF = head.indexOf('Floor'), iR=head.indexOf('Room'), iC=head.indexOf('ItemCode'), iQ=head.indexOf('QuantityRequired');
  const set = new Set();
  data.forEach(r=>{
    if (String(r[iF])===floor && String(r[iR])===room) {
      const code = String(r[iC]).trim();
      const qty  = Number(r[iQ] || 0);
      if (code && qty>0) set.add(code);
    }
  });
  return set;
}
function getItemNameByCode() {
  const map = new Map();
  map.set('G-113','G-113 Queen Bed Base');
  map.set('G-113A','G-113 Queen Bed Base');
  map.set('G-114','G-114 King Bed Base');
  map.set('G-114A','G-114 King Bed Base');

  map.set('G-105B','G-105B Queen Headboard');
  map.set('G-105A','G-105A King Headboard');

  map.set('G-103 RAF Desk','G-103 RAF Desk');
  map.set('G-103 LAF Desk','G-103 LAF Desk');
  map.set('G-103  Desk NO REF','G-103 Desk NO REF');
  map.set('G-103 Desk NO REF','G-103 Desk NO REF');

  map.set('G-112 RAF Luggage Rack','G-112 RAF Luggage');
  map.set('G-112 LAF Luggage Rack','G-112 LAF Luggage');

  map.set('G-111','G-111 Side Chair');
  map.set('G-101','G-101 Side Table');
  map.set('G-102','G-102 Desk Chair');
  map.set('G-104','G-104 Nightstand');
  map.set('G-106','G-106 Mirror');
  map.set('G-108','G-108 Pta Cover');
  map.set('G-100','G-100 Lounge Chair');

  map.set('MB-190','MB-190 Round Wall Light');
  map.set('MB-280','MB-280 Oval Wall Light');
  map.set('MB-180','MB-180 Reading Light');
  return map;
}
function codeToReportColumn(itemCode, map) {
  const key = String(itemCode || '').trim();
  return map.get(key) || null;
}
function isSameISODate(d, isoStr) {
  if (!(d instanceof Date)) return false;
  const y=d.getFullYear(), m=('0'+(d.getMonth()+1)).slice(-2), day=('0'+d.getDate()).slice(-2);
  return `${y}-${m}-${day}` === isoStr;
}
function parseIsoDate(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s).trim());
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2])-1, Number(m[3]));
  return isNaN(d.getTime()) ? null : d;
}
function sumMap(m){ let s=0; m.forEach(v=>s+=Number(v||0)); return s; }
function formatISO(d){ if(!(d instanceof Date))return''; const y=d.getFullYear(),m=('0'+(d.getMonth()+1)).slice(-2),dy=('0'+d.getDate()).slice(-2); return `${y}-${m}-${dy}`; }

/* =========================
   LECTURAS PARA STATUS
========================= */
function readRoomsMeta() {
  const sh = getSheetSafe(NAMES.rooms);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return new Map();
  const head = vals.shift();
  const idx = {floor: head.indexOf('Floor'), room: head.indexOf('Room'), type: head.indexOf('RoomType'), ada: head.indexOf('ADA')};
  const map = new Map();
  vals.forEach(r=>{
    const f=String(r[idx.floor]||'').trim();
    const rm=String(r[idx.room]||'').trim();
    if(!f||!rm)return;
    const key=f+'|'+rm;
    map.set(key,{roomType:String(r[idx.type]||'').trim(), ada:String(r[idx.ada]||'').trim()});
  });
  return map;
}
function readRoomsMaster() {
  const sh = getSheetSafe(NAMES.roomsMaster);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return new Map();
  const head = vals.shift();
  const iF = head.indexOf('Floor'), iR = head.indexOf('Room'), iC=head.indexOf('ItemCode'), iQ=head.indexOf('QuantityRequired');
  const map = new Map();
  vals.forEach(r=>{
    const f=String(r[iF]||'').trim(), rm=String(r[iR]||'').trim(), code=String(r[iC]||'').trim(), q=Number(r[iQ]||0);
    if(!f||!rm||!code||q<=0)return;
    const key=f+'|'+rm;
    if(!map.has(key))map.set(key,new Map());
    const m=map.get(key);
    m.set(code,(m.get(code)||0)+q);
  });
  return map;
}
function readDeliveries() {
  const sh = getSheetSafe(NAMES.deliveries);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return new Map();
  const head = vals.shift();
  const iDate=head.indexOf('Date'), iF=head.indexOf('Floor'), iR=head.indexOf('Room'), iC=head.indexOf('ItemCode'), iQ=head.indexOf('QuantityDelivered');
  const map = new Map();
  vals.forEach(r=>{
    const d=r[iDate], f=String(r[iF]||'').trim(), rm=String(r[iR]||'').trim(), code=String(r[iC]||'').trim(), q=Number(r[iQ]||0);
    if(!f||!rm||!code||q<=0)return;
    const key=f+'|'+rm;
    if(!map.has(key))map.set(key,{byItem:new Map(), lastDate:null});
    const obj=map.get(key);
    obj.byItem.set(code,(obj.byItem.get(code)||0)+q);
    if(d instanceof Date){ if(!obj.lastDate || d.getTime()>obj.lastDate.getTime()) obj.lastDate=d; }
  });
  return map;
}

/***** === API JSON READ-ONLY === *****/

// doGet con acciones: kanban_status, room_detail, users_list
function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    const action = (params.action || 'kanban_status').trim();
    const force = String(params.force || 'false').toLowerCase() === 'true'; // parámetro opcional

    if (action === 'kanban_status') {
      const payload = apiGetKanbanStatus_(force);
      return json_(payload);
    }

    if (action === 'room_detail') {
      const floor = String(params.floor || '').trim();
      const room  = String(params.room  || '').trim();
      if (!floor || !room) return json_({error:'Missing floor/room'}, 400);
      const payload = apiGetRoomDetail_(floor, room, force);
      return json_(payload);
    }

    if (action === 'users_list') {
      return json_({ users: apiUsersList_() });
    }

    return json_({error:'Unknown action'}, 404);
  } catch (err) {
    return json_({error:String(err)}, 500);
  }
}

// Construye la matriz de estado con caché
function apiGetKanbanStatus_(force = false) {
  const cacheKey = 'kanban_status_v1';
  if (!force) {
    const cached = cacheGet_(cacheKey);
    if (cached) {
      cached.cached = true;
      return cached;
    }
  }

  // Generar/actualizar status_overview
  try { buildStatusOverview(); } catch (_) {}

  const sh = SpreadsheetApp.getActive().getSheetByName(NAMES.status || 'status_overview');
  if (!sh || sh.getLastRow() < 2) return { items: [], floors: [], generatedAt: new Date().toISOString(), cached:false };

  const vals = sh.getDataRange().getValues();
  const head = vals[0];
  const rows = vals.slice(1);

  const idx = {
    Floor: head.indexOf('Floor'),
    Room: head.indexOf('Room'),
    RoomType: head.indexOf('RoomType'),
    ADA: head.indexOf('ADA'),
    Required: head.indexOf('Required'),
    Delivered: head.indexOf('Delivered'),
    Pct: head.indexOf('%'),
    Status: head.indexOf('Status'),
    LastDate: head.indexOf('LastDate'),
    Missing: head.indexOf('Missing'),
  };

  const items = rows
    .filter(r => String(r[idx.Floor]||'').trim())
    .map(r => ({
      floor: Number(r[idx.Floor]) || 0,
      room: String(r[idx.Room]||'').trim(),
      roomType: String(r[idx.RoomType]||'').trim(),
      ada: String(r[idx.ADA]||'').trim(),
      required: Number(r[idx.Required]||0),
      delivered: Number(r[idx.Delivered]||0),
      pct: Number(r[idx.Pct]||0),
      status: String(r[idx.Status]||'').trim(), // Not Delivered | Partial | Complete
      lastDeliveryDate: String(r[idx.LastDate]||'').trim(),
      missing: String(r[idx.Missing]||'').trim()
    }));

  items.sort((a,b) => (a.floor-b.floor) || ((Number(a.room)||0)-(Number(b.room)||0)));
  const floors = Array.from(new Set(items.map(i=>i.floor))).sort((a,b)=>a-b);

  const payload = { items, floors, generatedAt: new Date().toISOString(), cached:false };
  cachePut_(cacheKey, payload);
  return payload;
}

// Devuelve el detalle por cuarto con caché
function apiGetRoomDetail_(floor, room, force = false) {
  const cacheKey = `room_detail_${floor}_${room}`;
  if (!force) {
    const cached = cacheGet_(cacheKey);
    if (cached) {
      cached.cached = true;
      return cached;
    }
  }

  const master = readRoomsMaster(); // Map floor|room -> Map<ItemCode, qtyReq>
  const key = floor + '|' + room;
  const reqMap = master.get(key) || new Map();

  const delivered = readDeliveries(); // Map key -> { byItem:Map<code,qty>, lastDate }
  const delObj = delivered.get(key) || { byItem:new Map(), lastDate:null };
  const gotMap = delObj.byItem;

  const rows = [];
  const bedCodesQueen = new Set(['G-113','G-113A']);
  const bedCodesKing  = new Set(['G-114','G-114A']);
  const bedNums = readBedNumbersForRoom_(floor, room); // {queen:[], king:[]}

  reqMap.forEach((qtyReq, code) => {
    const qtyGot = Number(gotMap.get(code) || 0);
    rows.push({
      itemCode: code,
      qtyRequired: Number(qtyReq || 0),
      qtyDelivered: qtyGot,
      isBedQueen: bedCodesQueen.has(code),
      isBedKing:  bedCodesKing.has(code)
    });
  });

  rows.sort((a,b)=> String(a.itemCode).localeCompare(String(b.itemCode)));

  const payload = {
    floor,
    room,
    rows,
    bedNumbers: {
      queen: bedNums.queen.join(' | '),
      king:  bedNums.king.join(' | ')
    },
    cached:false
  };

  cachePut_(cacheKey, payload);
  return payload;
}

// Extrae BedNumbers por cuarto desde deliveries_log
function readBedNumbersForRoom_(floor, room) {
  const sh = SpreadsheetApp.getActive().getSheetByName(NAMES.deliveries);
  if (!sh || sh.getLastRow() < 2) return { queen:[], king:[] };

  const vals = sh.getDataRange().getValues();
  const head = vals.shift();
  const iF = head.indexOf('Floor');
  const iR = head.indexOf('Room');
  const iC = head.indexOf('ItemCode');
  const iB = head.indexOf('BedNumbers');

  const queen = [];
  const king = [];
  const queenCodes = new Set(['G-113','G-113A']);
  const kingCodes  = new Set(['G-114','G-114A']);

  vals.forEach(r => {
    if (String(r[iF])===String(floor) && String(r[iR])===String(room)) {
      const code = String(r[iC]||'').trim();
      const beds = String(r[iB]||'').trim();
      if (!beds) return;
      if (queenCodes.has(code)) queen.push(beds);
      if (kingCodes.has(code))  king.push(beds);
    }
  });
  return {queen, king};
}

// Helper: responde JSON
function json_(obj, status) {
  return ContentService
    .createTextOutput(JSON.stringify(obj || {}))
    .setMimeType(ContentService.MimeType.JSON);
}

/***** === API WRITE (doPost) para registrar entregas === *****/

function doPost(e) {
  try {
    const out = ContentService.createTextOutput().setMimeType(ContentService.MimeType.JSON);
    const params = (e && e.parameter) ? e.parameter : {};

    // API KEY
    const apiKey = String(params.apiKey || '');
    if (!validateApiKey_(apiKey)) {
      out.setContent(JSON.stringify({ error: 'Unauthorized' }));
      return out;
    }

    const action = String(params.action || '');
    if (action === 'submit_delivery') {
      const payloadStr = String(params.payload || '{}');
      const payload = JSON.parse(payloadStr);
      const result = apiSubmitDelivery_(payload);
      out.setContent(JSON.stringify({ ok: true, result }));
      return out;
    }

    out.setContent(JSON.stringify({ error: 'Unknown action' }));
    return out;
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function validateApiKey_(incoming) {
  const props = PropertiesService.getScriptProperties();
  const expected = props.getProperty('API_KEY') || '';
  return expected && incoming && String(incoming) === String(expected);
}

/**
 * payload esperado:
 * {
 *   dateISO: "2025-10-22",     // opcional, si no viene usa hoy
 *   floor: "8",
 *   room: "801",
 *   verifiedBy: "Milka Hernández", // opcional
 *   notes: "Entrega parcial",
 *   items: [
 *     { itemCode: "G-113", qtyDelivered: 2, bedNumbers: "Q-1782, Q-1783" },
 *     { itemCode: "G-104", qtyDelivered: 1 }
 *   ]
 * }
 */
function apiSubmitDelivery_(payload) {
  if (!payload) throw new Error('Missing payload');
  const floor = String(payload.floor || '').trim();
  const room  = String(payload.room  || '').trim();
  if (!floor || !room) throw new Error('floor/room required');

  // Validaciones básicas contra catálogo
  const pair = floor + '|' + room;
  const roomPairs = getPairsFloorRoom();
  if (!roomPairs.has(pair)) throw new Error('Room does not exist in rooms');

  const sh = getSheet(NAMES.deliveries);
  const dateObj = payload.dateISO ? parseIsoDate(payload.dateISO) : new Date();
  if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) throw new Error('Invalid date');

  const verifiedBy = String(payload.verifiedBy || '').trim();
  const notes = String(payload.notes || '').trim();

  // Construir filas a insertar (una por ítem)
  const rows = [];
  const validItems = getValidItemsForRoom(floor, room); // Set de códigos válidos
  (payload.items || []).forEach(it => {
    const code = String(it.itemCode || '').trim();
    const qty  = Number(it.qtyDelivered || 0);
    const beds = String(it.bedNumbers || '').trim();

    if (!code || !validItems.has(code)) {
      throw new Error('Invalid ItemCode for room: ' + code);
    }
    if (isNaN(qty) || qty < 0) throw new Error('Invalid qtyDelivered for ' + code);

    // Si es base de cama, BedNumbers obligatorio
    if (BED_BASE_CODES.has(code)) {
      if (!beds) throw new Error('BedNumbers required for ' + code);
      const count = beds.split(',').map(s=>s.trim()).filter(Boolean).length;
      if (count !== qty) throw new Error('BedNumbers count must match qty for ' + code);
    }

    rows.push([
      dateObj,            // Date (A)
      floor,              // Floor (B)
      room,               // Room (C)
      code,               // ItemCode (D)
      qty,                // QuantityDelivered (E)
      beds || '',         // BedNumbers (F)
      verifiedBy || '',   // VerifiedBy (G)
      notes || ''         // Notes (H)
    ]);
  });

  if (!rows.length) return { inserted: 0 };

  // Append en bloque
  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rows.length, 8).setValues(rows);

  // Validación visual (opcional)
  for (let i = 0; i < rows.length; i++) {
    const r = startRow + i;
    validateDeliveryRow(r, true);
  }

  // 🔥 Invalida caché para que frontend vea datos nuevos
  cacheBust_('kanban_status_v1');
  cacheBust_(`room_detail_${floor}_${room}`);

  return { inserted: rows.length };
}

// Lee users!A como lista simple (sin vacíos)
function apiUsersList_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(NAMES.users);
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1)
                .getValues()
                .flat()
                .map(v => String(v).trim())
                .filter(Boolean);
  return vals;
}
