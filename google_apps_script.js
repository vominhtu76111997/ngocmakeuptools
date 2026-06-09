/**
 * ============================================
 * RUBIES BEAUTY — Google Apps Script Backend
 * ============================================
 * 
 * SETUP:
 * 1. Tạo Google Sheet mới
 * 2. Vào Extensions > Apps Script
 * 3. Dán code này vào
 * 4. Thêm sheet "Items" với header: id | name | category | brand | shade | price | qty | purchaseDate | expiryDate | status | notes | photo | createdAt | updatedAt
 * 5. Deploy > New deployment > Web app > Anyone can access
 * 6. Copy URL, dán vào Settings của app
 * 
 * IMPROVEMENTS OVER FINANCE APP:
 * - Delta sync: chỉ trả items thay đổi sau timestamp
 * - Batch upsert: xử lý nhiều items cùng lúc
 * - Batch delete: xóa nhiều items cùng lúc
 * - Smart row lookup: cache row index bằng ID
 * - CORS headers tự động
 */

const SHEET_NAME = 'Items';
const HEADERS = ['id','name','category','brand','shade','price','qty','purchaseDate','expiryDate','status','notes','photo','createdAt','updatedAt'];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getAllItems(sinceTimestamp) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { items: [], total: 0 };
  
  const headers = data[0];
  const items = [];
  const since = sinceTimestamp ? Number(sinceTimestamp) : 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // skip empty rows
    
    const item = {};
    headers.forEach((h, j) => {
      item[h] = row[j];
    });
    
    // Convert numeric fields
    item.price = Number(item.price) || 0;
    item.qty = Number(item.qty) || 1;
    item.createdAt = Number(item.createdAt) || 0;
    item.updatedAt = Number(item.updatedAt) || 0;
    
    // Delta sync: only return items updated after 'since'
    if (since === 0 || item.updatedAt > since) {
      items.push(item);
    }
  }
  
  return { items, total: data.length - 1 };
}

function upsertItems(items) {
  if (!items || !items.length) return { updated: 0 };
  
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Build ID -> row index map
  const idRowMap = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) idRowMap[data[i][0]] = i + 1; // 1-indexed
  }
  
  let updated = 0;
  const newRows = [];
  
  items.forEach(item => {
    const rowData = headers.map(h => item[h] !== undefined ? item[h] : '');
    const existingRow = idRowMap[item.id];
    
    if (existingRow) {
      // Update existing row
      sheet.getRange(existingRow, 1, 1, headers.length).setValues([rowData]);
      updated++;
    } else {
      // Collect new rows for batch append
      newRows.push(rowData);
      updated++;
    }
  });
  
  // Batch append new rows
  if (newRows.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newRows.length, headers.length).setValues(newRows);
  }
  
  return { updated };
}

function deleteItems(ids) {
  if (!ids || !ids.length) return { deleted: 0 };
  
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  
  // Find rows to delete (reverse order to avoid index shifting)
  const rowsToDelete = [];
  for (let i = 1; i < data.length; i++) {
    if (ids.includes(String(data[i][0]))) {
      rowsToDelete.push(i + 1);
    }
  }
  
  // Delete from bottom up
  rowsToDelete.sort((a, b) => b - a);
  rowsToDelete.forEach(row => sheet.deleteRow(row));
  
  return { deleted: rowsToDelete.length };
}

// === HTTP HANDLERS ===

function doGet(e) {
  const action = e.parameter.action || 'getAll';
  const since = e.parameter.since || '0';
  
  let result;
  switch (action) {
    case 'getAll':
      result = getAllItems(since);
      break;
    default:
      result = { error: 'Unknown action' };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Invalid JSON' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  let result;
  switch (payload.action) {
    case 'sync':
      // Combined: upsert items + delete items
      const upsertResult = upsertItems(payload.items || []);
      const deleteResult = deleteItems(payload.deletedIds || []);
      result = {
        success: true,
        updated: upsertResult.updated,
        deleted: deleteResult.deleted,
        timestamp: Date.now()
      };
      break;
      
    case 'save':
      result = upsertItems([payload.item]);
      break;
      
    case 'saveBatch':
      result = upsertItems(payload.items);
      break;
      
    case 'delete':
      result = deleteItems([payload.id]);
      break;
      
    case 'deleteAll':
      const sheet = getSheet();
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
      result = { success: true, deleted: lastRow - 1 };
      break;
      
    default:
      result = { error: 'Unknown action: ' + payload.action };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
