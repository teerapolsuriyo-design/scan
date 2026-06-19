// ============================================================
//  GOOGLE APPS SCRIPT — REST API Backend
//  ระบบบันทึกและตรวจสอบการเข้าแนะแนวโรงเรียน
//  วิธีใช้: Deploy > New deployment > Web App
//           Execute as: Me | Who has access: Anyone
// ============================================================

function doGet(e) {
  const action = e.parameter.action;
  let result;

  if (action === 'getVisits') {
    result = getVisits();
  } else if (action === 'getConfig') {
    result = getConfig();
  } else {
    result = { error: 'Unknown action: ' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Invalid JSON body' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = data.action;
  let result;

  if (action === 'addVisit') {
    result = addVisit(data);
  } else if (action === 'updateVisit') {
    result = updateVisit(data);
  } else if (action === 'deleteVisit') {
    result = deleteVisit(data.id);
  } else if (action === 'saveConfig') {
    result = saveConfig(data.lat, data.lng, data.radius);
  } else {
    result = { error: 'Unknown action: ' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- ส่วนจัดการข้อมูลการแนะแนว (Guidance Visits) ---
function getOrCreateVisitsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('GuidanceVisits');
  if (!sheet) {
    sheet = ss.insertSheet('GuidanceVisits');
    sheet.appendRow([
      'ID', 
      'Date', 
      'SchoolName', 
      'Province', 
      'District', 
      'Subdistrict', 
      'Status', 
      'LectorName', 
      'Notes', 
      'Latitude', 
      'Longitude', 
      'Timestamp'
    ]);
    // จัดรูปแบบแถวหัวตารางให้เป็นตัวหนา
    sheet.getRange(1, 1, 1, 12).setFontWeight('bold');
  }
  return sheet;
}

function getVisits() {
  try {
    const sheet = getOrCreateVisitsSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const visits = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const visit = {};
      let hasData = false;
      for (let j = 0; j < headers.length; j++) {
        visit[headers[j]] = row[j];
        if (row[j] !== '') hasData = true;
      }
      if (hasData) {
        // แปลงรูปแบบวันที่ (Date) ให้เป็น String YYYY-MM-DD เพื่อใช้งานใน input type="date"
        if (visit.Date instanceof Date) {
          visit.Date = Utilities.formatDate(visit.Date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        if (visit.Timestamp instanceof Date) {
          visit.Timestamp = Utilities.formatDate(visit.Timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        }
        visits.push(visit);
      }
    }
    return visits;
  } catch (err) {
    return { error: err.message };
  }
}

function addVisit(data) {
  try {
    const sheet = getOrCreateVisitsSheet();
    const id = 'V' + new Date().getTime();
    const now = new Date();
    
    sheet.appendRow([
      id,
      data.date || '',
      data.schoolName || '',
      data.province || '',
      data.district || '',
      data.subdistrict || '',
      data.status || 'รอดำเนินการ',
      data.lectorName || '',
      data.notes || '',
      data.lat || '',
      data.lng || '',
      now
    ]);
    
    return { success: true, message: 'บันทึกข้อมูลแนะแนวสำเร็จ', id: id };
  } catch (err) {
    return { error: err.message };
  }
}

function updateVisit(data) {
  try {
    const sheet = getOrCreateVisitsSheet();
    const sheetData = sheet.getDataRange().getValues();
    const id = data.id;
    if (!id) return { error: 'ไม่พบ ID ที่ต้องการแก้ไข' };

    let rowIndex = -1;
    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][0] === id) {
        rowIndex = i + 1; // แปลงเป็น 1-indexed และข้ามหัวตาราง
        break;
      }
    }

    if (rowIndex === -1) {
      return { error: 'ไม่พบข้อมูลที่ตรงกับ ID: ' + id };
    }

    const now = new Date();
    
    if (data.date !== undefined) sheet.getRange(rowIndex, 2).setValue(data.date);
    if (data.schoolName !== undefined) sheet.getRange(rowIndex, 3).setValue(data.schoolName);
    if (data.province !== undefined) sheet.getRange(rowIndex, 4).setValue(data.province);
    if (data.district !== undefined) sheet.getRange(rowIndex, 5).setValue(data.district);
    if (data.subdistrict !== undefined) sheet.getRange(rowIndex, 6).setValue(data.subdistrict);
    if (data.status !== undefined) sheet.getRange(rowIndex, 7).setValue(data.status);
    if (data.lectorName !== undefined) sheet.getRange(rowIndex, 8).setValue(data.lectorName);
    if (data.notes !== undefined) sheet.getRange(rowIndex, 9).setValue(data.notes);
    if (data.lat !== undefined) sheet.getRange(rowIndex, 10).setValue(data.lat);
    if (data.lng !== undefined) sheet.getRange(rowIndex, 11).setValue(data.lng);
    sheet.getRange(rowIndex, 12).setValue(now); // อัปเดตเวลาอัปเดตล่าสุด

    return { success: true, message: 'แก้ไขข้อมูลสำเร็จ' };
  } catch (err) {
    return { error: err.message };
  }
}

function deleteVisit(id) {
  try {
    const sheet = getOrCreateVisitsSheet();
    const sheetData = sheet.getDataRange().getValues();
    if (!id) return { error: 'ไม่พบ ID ที่ต้องการลบ' };

    let rowIndex = -1;
    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][0] === id) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { error: 'ไม่พบข้อมูลที่ตรงกับ ID: ' + id };
    }

    sheet.deleteRow(rowIndex);
    return { success: true, message: 'ลบข้อมูลสำเร็จ' };
  } catch (err) {
    return { error: err.message };
  }
}

// --- ส่วนจัดการ Config ดั้งเดิม (เผื่อใช้งาน) ---
function saveConfig(lat, lng, radius) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');

  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.getRange('A1:B1').setValues([['Parameter', 'Value']]);
    sheet.getRange('A2').setValue('Target Latitude');
    sheet.getRange('A3').setValue('Target Longitude');
    sheet.getRange('A4').setValue('Allowed Radius (KM)');
    sheet.setColumnWidth(1, 150);
  }

  sheet.getRange('B2').setValue(lat);
  sheet.getRange('B3').setValue(lng);
  sheet.getRange('B4').setValue(radius);

  return { success: true, message: 'บันทึกการตั้งค่าลง Google Sheets เรียบร้อย' };
}

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');

  let config = { lat: 0, lng: 0, radius: 0.5 };

  if (sheet) {
    const latVal = sheet.getRange('B2').getValue();
    const lngVal = sheet.getRange('B3').getValue();
    const radiusVal = sheet.getRange('B4').getValue();

    if (latVal !== '') config.lat = parseFloat(latVal);
    if (lngVal !== '') config.lng = parseFloat(lngVal);
    if (radiusVal !== '') config.radius = parseFloat(radiusVal);
  }

  return config;
}
