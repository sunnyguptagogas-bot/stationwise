// ═══════════════════════════════════════════════════════════
// StationWise Google Apps Script — Full Version with Photos
// ═══════════════════════════════════════════════════════════

const SHEET_NAME = 'All Reports';
const SESSIONS_SHEET = 'Device Sessions';
const PHOTO_FOLDER_NAME = 'StationWise DU Meter Photos';

function getOrCreateSheet(name, headers, headerColor) {
  const doc = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = doc.getSheetByName(name);
  if (!sheet) {
    sheet = doc.insertSheet(name);
    if (headers) {
      sheet.appendRow(headers);
      sheet.getRange(1,1,1,headers.length)
        .setBackground(headerColor||'#1b5e20')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function getPhotoFolder(){
  const folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if(folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

function doPost(e) {
  try {
    const lock = LockService.getPublicLock();
    lock.waitLock(30000);
    const data = JSON.parse(e.postData.contents);

    // ── PHOTO UPLOAD ──
    if (data.type === 'photo') {
      const photoUrl = handlePhotoUpload(data);
      lock.releaseLock();
      return jsonResponse({status:'ok', photoUrl});
    }

    // ── DEVICE REGISTRATION ──
    if (data.type === 'registerDevice') {
      handleDeviceRegistration(data);
      lock.releaseLock();
      return jsonResponse({status:'ok'});
    }

    // ── FORCE LOGOUT / BLOCK / UNBLOCK ──
    if (['forceLogout','blockDevice','unblockDevice'].includes(data.type)) {
      const status = data.type==='blockDevice'?'Blocked':data.type==='forceLogout'?'Force_Logout':'Active';
      updateDeviceStatus(data.deviceId||data.erp, status);
      lock.releaseLock();
      return jsonResponse({status:'ok'});
    }

    // ── REPORT SUBMISSION ──
    const sheet = getOrCreateSheet(SHEET_NAME, [
      'Timestamp','Sync Time','ERP Code','Station Name','State','SS Name',
      'Shift','Date','Month','Year',
      'Shift 1 Sales','Shift 2 Sales','Shift 3 Sales',
      'Open A1','Open B1','Open A2','Open B2',
      'Close A1','Close B1','Close A2','Close B2',
      'Total Sale','Rtn Testing','Free Sale','Net Sale (L)',
      'RSP DU1-N1','RSP DU1-N2','RSP DU2-N1','RSP DU2-N2',
      'Prev Cash','Total Sales Amt','Card Sale','Paytm',
      'Free Sales Amt','Cash Sale Amt','Bank Deposit','Cash In Hand',
      'Dip Reading','Stock Ltrs','DG Open','DG Close','DG Hours',
      'Comments','Photo URL','Report ID','Status','Edit History'
    ]);

    const reportId = data.reportId || data.id || '';
    
    // ── SERVER-SIDE DUPLICATE PREVENTION ──
    if (sheet.getLastRow() > 1) {
      const lastRow = sheet.getLastRow()-1;
      const allData = sheet.getRange(2,1,lastRow,sheet.getLastColumn()).getValues();
      
      for (let i=0; i<allData.length; i++) {
        const row = allData[i];
        const rowId = row[42]||'';  // Report ID column
        const rowErp = row[2]||'';  // ERP Code
        const rowDate = row[7]||''; // Date
        const rowMonth = row[8]||''; // Month
        const rowYear = row[9]||''; // Year
        const rowShift = row[6]||''; // Shift
        
        // Match by Report ID (exact match)
        if (reportId && rowId === reportId) {
          sheet.getRange(i+2,2).setValue(new Date().toLocaleString());
          if(data.photoUrl) sheet.getRange(i+2,42).setValue(data.photoUrl);
          // Update audit trail
          const existingHistory = row[44]||'';
          const newHistory = existingHistory ? existingHistory+' | Edited:'+new Date().toLocaleString()+' by '+data.erp : 'Edited:'+new Date().toLocaleString()+' by '+data.erp;
          sheet.getRange(i+2,45).setValue(newHistory);
          lock.releaseLock();
          return jsonResponse({status:'updated', message:'Report updated (same ID)'});
        }
        
        // Match by ERP+Date+Month+Year+Shift (same report different device)
        if (rowErp === (data.erp||'') && 
            rowDate.toString() === (data.date||'').toString() &&
            rowMonth === (data.month||'') &&
            rowYear.toString() === (data.year||'').toString() &&
            rowShift === (data.shift||'')) {
          // Update the existing row instead of creating duplicate
          sheet.getRange(i+2,2).setValue(new Date().toLocaleString());
          if(data.photoUrl) sheet.getRange(i+2,42).setValue(data.photoUrl);
          // Update key fields that may have changed
          sheet.getRange(i+2,22).setValue(data.totalSale||row[21]);
          sheet.getRange(i+2,25).setValue(data.netSale||row[24]);
          sheet.getRange(i+2,36).setValue(data.cashInHand||row[35]);
          lock.releaseLock();
          return jsonResponse({status:'duplicate_prevented', message:'Duplicate report prevented - existing row updated'});
        }
      }
    }

    sheet.appendRow([
      data.submittedAt||new Date().toISOString(),
      new Date().toLocaleString(),
      data.erp||'', data.stationName||'', data.state||'', data.ssName||'',
      data.shift||'', data.date||'', data.month||'', data.year||'',
      data.s1sale||'0.00', data.s2sale||'0.00', data.s3sale||'0.00',
      data.openA1||'', data.openB1||'', data.openA2||'NA', data.openB2||'NA',
      data.closeA1||'', data.closeB1||'', data.closeA2||'NA', data.closeB2||'NA',
      data.totalSale||'0.00', data.rtnTesting||'0.00', data.freeSale||'0.00', data.netSale||'0.00',
      data.rsp1||'', data.rsp2||'NA', data.rsp3||'NA', data.rsp4||'NA',
      data.prevCash||'0.00', data.totalAmt||'0.00', data.cardSale||'0.00',
      data.paytmSale||'0.00', data.freeAmt||'0.00', data.cashSale||'0.00',
      data.bankDeposit||'0.00', data.cashInHand||'0.00',
      data.dipReading||'0', data.stockLtrs||'0.00',
      data.dgOpen||'0.00', data.dgClose||'0.00', data.dgHours||'0.00',
      data.comments||'', data.photoUrl||'', reportId, 'Synced', data.editHistory||''
    ]);

    const newRow = sheet.getLastRow();
    if (newRow%2===0) sheet.getRange(newRow,1,1,sheet.getLastColumn()).setBackground('#f1f8e9');

    lock.releaseLock();
    return jsonResponse({status:'success', row:newRow});

  } catch(err) {
    return jsonResponse({status:'error', message:err.toString()});
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action||'';

    if (action==='checkDevice') {
      return jsonResponse(getDeviceStatus(e.parameter.deviceId||''));
    }
    if (action==='getSessions') {
      return jsonResponse({status:'ok', sessions:getAllSessions()});
    }

    // Get reports
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet||sheet.getLastRow()<=1) return jsonResponse({status:'ok',count:0,rows:[]});
    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
    const rows = data.map(row=>{
      const obj={};
      headers.forEach((h,i)=>{obj[h]=row[i];});
      return obj;
    });
    return jsonResponse({status:'ok',count:rows.length,rows});
  } catch(err) {
    return jsonResponse({status:'error',message:err.toString()});
  }
}

// ── PHOTO UPLOAD TO DRIVE ──
function handlePhotoUpload(data) {
  try {
    const folder = getPhotoFolder();
    const blob = Utilities.newBlob(
      Utilities.base64Decode(data.base64Data),
      data.mimeType||'image/jpeg',
      data.fileName||('DU_'+Date.now()+'.jpg')
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    // Return direct view URL
    const fileId = file.getId();
    return 'https://drive.google.com/file/d/'+fileId+'/view';
  } catch(e) {
    return '';
  }
}

// ── DEVICE FUNCTIONS ──
function handleDeviceRegistration(data) {
  const sheet = getOrCreateSheet(SESSIONS_SHEET, [
    'Device ID','ERP Code','SS Name','Station Name','State',
    'First Seen','Last Seen','User Agent','Status'
  ], '#1a3a5c');
  const deviceId = data.deviceId||'';
  if (!deviceId) return;
  if (sheet.getLastRow()>1) {
    const ids = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
    for (let i=0;i<ids.length;i++) {
      if (ids[i][0]===deviceId) {
        sheet.getRange(i+2,7).setValue(new Date().toLocaleString());
        return;
      }
    }
  }
  sheet.appendRow([deviceId,data.erp||'',data.ssName||'',data.stationName||'',data.state||'',
    new Date().toLocaleString(),new Date().toLocaleString(),(data.userAgent||'').substring(0,100),'Active']);
  sheet.getRange(sheet.getLastRow(),1,1,sheet.getLastColumn()).setBackground('#fff9c4');
}

function getDeviceStatus(deviceId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS_SHEET);
  if (!sheet||sheet.getLastRow()<=1) return {blocked:false,forceLogout:false};
  const data = sheet.getRange(2,1,sheet.getLastRow()-1,9).getValues();
  for (const row of data) {
    if (row[0]===deviceId) {
      const status = (row[8]||'').toString().toLowerCase();
      return {blocked:status==='blocked',forceLogout:status==='force_logout'};
    }
  }
  return {blocked:false,forceLogout:false};
}

function updateDeviceStatus(id, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS_SHEET);
  if (!sheet||sheet.getLastRow()<=1) return;
  const ids = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
  for (let i=0;i<ids.length;i++) {
    if (ids[i][0]===id||ids[i][1]===id) {
      sheet.getRange(i+2,9).setValue(status);
      const color=status==='Blocked'?'#fde8e8':status==='Force_Logout'?'#fdf4e3':'#f1f8e9';
      sheet.getRange(i+2,1,1,sheet.getLastColumn()).setBackground(color);
    }
  }
}

function getAllSessions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS_SHEET);
  if (!sheet||sheet.getLastRow()<=1) return {};
  const data = sheet.getRange(2,1,sheet.getLastRow()-1,9).getValues();
  const sessions={};
  data.forEach(row=>{
    sessions[row[0]]={deviceId:row[0],erp:row[1],ssName:row[2],stationName:row[3],
      state:row[4],firstSeen:row[5],lastSeen:row[6],
      blocked:(row[8]||'').toLowerCase()==='blocked',
      forceLogout:(row[8]||'').toLowerCase()==='force_logout'};
  });
  return sessions;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
