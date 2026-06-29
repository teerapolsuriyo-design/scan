// ============================================================
//  GOOGLE APPS SCRIPT — REST API Backend
//  ระบบบันทึกและตรวจสอบการเข้าแนะแนวโรงเรียน
//  วิธีใช้: Deploy > New deployment > Web App
//           Execute as: Me | Who has access: Anyone
// ============================================================

const SPREADSHEET_ID = '1l4BJct9oesnkxiqe7rHxtdJ8Q1ltItb1AAU7qdUGQ_Y';

function getSpreadsheet() {
  try {
    const activeSs = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSs) {
      return activeSs;
    }
  } catch (e) {
    // Fail silently, fall back to openById
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function doGet(e) {
  const action = (e && e.parameter) ? e.parameter.action : null;

  if (!action) {
    // Serve index.html as Web App UI
    const template = HtmlService.createTemplateFromFile('index');
    template.webAppUrl = ScriptApp.getService().getUrl();
    return template.evaluate()
      .setTitle('ระบบบันทึกและตรวจสอบการเข้าแนะแนวโรงเรียน')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  let result;

  if (action === 'getVisits') {
    result = getVisits();
  } else if (action === 'getConfig') {
    result = getConfig();
  } else if (action === 'getSchools') {
    result = getSchools();
  } else {
    result = { error: 'Unknown action: ' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Invalid JSON body' }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeaders({
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type'
      });
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
  } else if (action === 'addSchool') {
    result = addSchool(data);
  } else if (action === 'login') {
    result = loginUser(data.username, data.password);
  } else if (action === 'getUsers') {
    result = getUsers();
  } else if (action === 'addUser') {
    result = addUser(data);
  } else if (action === 'updateUser') {
    result = updateUser(data);
  } else if (action === 'deleteUser') {
    result = deleteUser(data.username);
  } else {
    result = { error: 'Unknown action: ' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
}

// --- ส่วนจัดการข้อมูลการแนะแนว (Guidance Visits) ---
function getOrCreateVisitsSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('GuidanceVisits');
  
  const defaultHeaders = [
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
    'YearBE',
    'Phone',
    'Teacher',
    'Students',
    'Timestamp'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('GuidanceVisits');
    sheet.appendRow(defaultHeaders);
    sheet.getRange(1, 1, 1, defaultHeaders.length).setFontWeight('bold');
  } else {
    // ระบบอัปเกรดคอลัมน์ให้อัตโนมัติในกรณีชีตเดิมไม่มีคอลัมน์ใหม่
    const existingHeaders = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
    defaultHeaders.forEach(header => {
      if (existingHeaders.indexOf(header) === -1) {
        // เพิ่มคอลัมน์ก่อนหน้าคอลัมน์ Timestamp
        const timestampIdx = existingHeaders.indexOf('Timestamp');
        if (timestampIdx !== -1) {
          sheet.insertColumnBefore(timestampIdx + 1);
          sheet.getRange(1, timestampIdx + 1).setValue(header).setFontWeight('bold');
          existingHeaders.splice(timestampIdx, 0, header); // อัปเดตโครงสร้างชั่วคราว
        } else {
          sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header).setFontWeight('bold');
          existingHeaders.push(header);
        }
      }
    });
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
        // แปลงรูปแบบวันที่ (Date) ให้เป็น String YYYY-MM-DD
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
    
    // โหลด headers เพื่อทำ mapping
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowValues = new Array(headers.length).fill('');
    const colMap = {};
    headers.forEach((h, idx) => {
      colMap[h] = idx;
    });

    if (colMap['ID'] !== undefined) rowValues[colMap['ID']] = id;
    if (colMap['Date'] !== undefined) rowValues[colMap['Date']] = data.date || '';
    if (colMap['SchoolName'] !== undefined) rowValues[colMap['SchoolName']] = data.schoolName || '';
    if (colMap['Province'] !== undefined) rowValues[colMap['Province']] = data.province || '';
    if (colMap['District'] !== undefined) rowValues[colMap['District']] = data.district || '';
    if (colMap['Subdistrict'] !== undefined) rowValues[colMap['Subdistrict']] = data.subdistrict || '';
    if (colMap['Status'] !== undefined) rowValues[colMap['Status']] = data.status || 'รอดำเนินการ';
    if (colMap['LectorName'] !== undefined) rowValues[colMap['LectorName']] = data.lectorName || '';
    if (colMap['Notes'] !== undefined) rowValues[colMap['Notes']] = data.notes || '';
    if (colMap['Latitude'] !== undefined) rowValues[colMap['Latitude']] = data.lat || '';
    if (colMap['Longitude'] !== undefined) rowValues[colMap['Longitude']] = data.lng || '';
    if (colMap['YearBE'] !== undefined) rowValues[colMap['YearBE']] = data.yearBE || '';
    if (colMap['Phone'] !== undefined) rowValues[colMap['Phone']] = data.phone || '';
    if (colMap['Teacher'] !== undefined) rowValues[colMap['Teacher']] = data.teacher || '';
    if (colMap['Students'] !== undefined) rowValues[colMap['Students']] = data.students || '';
    if (colMap['Timestamp'] !== undefined) rowValues[colMap['Timestamp']] = now;

    sheet.appendRow(rowValues);
    
    return { success: true, message: 'บันทึกข้อมูลแนะแนวสำเร็จ', id: id };
  } catch (err) {
    return { error: err.message };
  }
}

function updateVisit(data) {
  try {
    const sheet = getOrCreateVisitsSheet();
    const sheetData = sheet.getDataRange().getValues();
    const headers = sheetData[0];
    
    // สร้างแผนผังชื่อคอลัมน์ไปเป็นลำดับคอลัมน์ (1-indexed)
    const colMap = {};
    headers.forEach((h, idx) => {
      colMap[h] = idx + 1;
    });

    const id = data.id;
    if (!id) return { error: 'ไม่พบ ID ที่ต้องการแก้ไข' };

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

    const now = new Date();
    
    const fields = {
      'Date': data.date,
      'SchoolName': data.schoolName,
      'Province': data.province,
      'District': data.district,
      'Subdistrict': data.subdistrict,
      'Status': data.status,
      'LectorName': data.lectorName,
      'Notes': data.notes,
      'Latitude': data.lat,
      'Longitude': data.lng,
      'YearBE': data.yearBE,
      'Phone': data.phone,
      'Teacher': data.teacher,
      'Students': data.students
    };

    for (const [key, val] of Object.entries(fields)) {
      if (val !== undefined && colMap[key] !== undefined) {
        sheet.getRange(rowIndex, colMap[key]).setValue(val);
      }
    }
    
    if (colMap['Timestamp'] !== undefined) {
      sheet.getRange(rowIndex, colMap['Timestamp']).setValue(now);
    }

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

// --- ส่วนจัดการข้อมูลโรงเรียนใหม่ (Schools Database Sheet) ---
function getOrCreateSchoolsSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('Schools');
  
  const defaultHeaders = [
    'SchoolName', 
    'Province', 
    'District', 
    'Subdistrict', 
    'Phone', 
    'Teacher', 
    'Students', 
    'Notes'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('Schools');
    sheet.appendRow(defaultHeaders);
    sheet.getRange(1, 1, 1, defaultHeaders.length).setFontWeight('bold');
  }
  return sheet;
}

function getSchools() {
  try {
    const sheet = getOrCreateSchoolsSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const schools = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const school = {};
      let hasData = false;
      for (let j = 0; j < headers.length; j++) {
        school[headers[j]] = row[j];
        if (row[j] !== '') hasData = true;
      }
      if (hasData) {
        schools.push(school);
      }
    }
    return schools;
  } catch (err) {
    return { error: err.message };
  }
}

function addSchool(data) {
  try {
    const sheet = getOrCreateSchoolsSheet();
    const sheetData = sheet.getDataRange().getValues();
    
    const name = (data.schoolName || '').trim();
    const province = (data.province || '').trim();
    const district = (data.district || '').trim();
    const subdistrict = (data.subdistrict || '').trim();
    
    if (!name || !province || !district || !subdistrict) {
      return { error: 'ข้อมูลโรงเรียนไม่ครบถ้วน' };
    }
    
    // ตรวจสอบข้อมูลโรงเรียนซ้ำในชีต (ดูจากชื่อ จังหวัด และอำเภอ)
    let exists = false;
    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][0].toString().trim().toLowerCase() === name.toLowerCase() &&
          sheetData[i][1].toString().trim().toLowerCase() === province.toLowerCase() &&
          sheetData[i][2].toString().trim().toLowerCase() === district.toLowerCase()) {
        exists = true;
        break;
      }
    }
    
    if (exists) {
      return { success: true, message: 'โรงเรียนนี้มีอยู่แล้วในระบบ Google Sheets' };
    }
    
    sheet.appendRow([
      name,
      province,
      district,
      subdistrict,
      data.phone || '',
      data.teacher || '',
      data.students || '',
      data.notes || ''
    ]);
    
    return { success: true, message: 'บันทึกข้อมูลโรงเรียนลงฐานข้อมูล Google Sheets สำเร็จ' };
  } catch (err) {
    return { error: err.message };
  }
}

// --- ส่วนจัดการ Config ดั้งเดิม ---
function saveConfig(lat, lng, radius) {
  const ss = getSpreadsheet();
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
  const ss = getSpreadsheet();
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

// --- ส่วนจัดการบัญชีผู้ใช้งาน (User Management Sheet) ---
function getOrCreateUsersSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('Users');
  
  const defaultHeaders = ['Username', 'Password', 'Name', 'Role'];
  
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.appendRow(defaultHeaders);
    sheet.getRange(1, 1, 1, defaultHeaders.length).setFontWeight('bold');
    
    // บัญชีผู้ใช้เริ่มต้น
    sheet.appendRow(['admin', 'admin1234', 'ผู้ดูแลระบบ', 'Admin']);
  } else {
    // ตรวจสอบว่ามีผู้ใช้ admin หรือไม่ และอัปเดตรหัสผ่านให้เป็น admin1234
    const data = sheet.getDataRange().getValues();
    let adminIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toLowerCase() === 'admin') {
        adminIndex = i;
        break;
      }
    }
    if (adminIndex === -1) {
      sheet.appendRow(['admin', 'admin1234', 'ผู้ดูแลระบบ', 'Admin']);
    } else {
      if (data[adminIndex][1] && data[adminIndex][1].toString().trim() !== 'admin1234') {
        sheet.getRange(adminIndex + 1, 2).setValue('admin1234');
      }
    }
  }
  return sheet;
}

function loginUser(username, password) {
  try {
    const sheet = getOrCreateUsersSheet();
    const data = sheet.getDataRange().getValues();
    
    const uName = (username || '').trim().toLowerCase();
    const uPass = (password || '').trim();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const sheetUser = row[0].toString().trim().toLowerCase();
      const sheetPass = row[1].toString().trim();
      
      if (sheetUser === uName && sheetPass === uPass) {
        return {
          success: true,
          user: {
            username: row[0].toString().trim(),
            name: row[2].toString().trim(),
            role: row[3].toString().trim()
          }
        };
      }
    }
    return { error: 'ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง' };
  } catch (err) {
    return { error: err.message };
  }
}

function getUsers() {
  try {
    const sheet = getOrCreateUsersSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const users = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] !== '') {
        users.push({
          username: row[0].toString().trim(),
          password: row[1].toString().trim(), // คืนรหัสผ่านให้ผู้ใช้ Admin แก้ไขได้
          name: row[2].toString().trim(),
          role: row[3].toString().trim()
        });
      }
    }
    return users;
  } catch (err) {
    return { error: err.message };
  }
}

function addUser(data) {
  try {
    const sheet = getOrCreateUsersSheet();
    const sheetData = sheet.getDataRange().getValues();
    
    const username = (data.username || '').trim();
    const password = (data.password || '').trim();
    const name = (data.name || '').trim();
    const role = (data.role || 'User').trim();
    
    if (!username || !password || !name) {
      return { error: 'กรอกข้อมูลผู้ใช้งานไม่ครบถ้วน' };
    }
    
    // ตรวจสอบชื่อผู้ใช้งานซ้ำ
    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][0].toString().trim().toLowerCase() === username.toLowerCase()) {
        return { error: 'ชื่อผู้ใช้งานนี้มีอยู่ในระบบแล้ว' };
      }
    }
    
    sheet.appendRow([username, password, name, role]);
    return { success: true, message: 'เพิ่มบัญชีผู้ใช้งานสำเร็จ' };
  } catch (err) {
    return { error: err.message };
  }
}

function updateUser(data) {
  try {
    const sheet = getOrCreateUsersSheet();
    const sheetData = sheet.getDataRange().getValues();
    
    const username = (data.username || '').trim();
    const password = (data.password || '').trim();
    const name = (data.name || '').trim();
    const role = (data.role || 'User').trim();
    
    if (!username || !password || !name) {
      return { error: 'ข้อมูลสำหรับแก้ไขไม่ครบถ้วน' };
    }
    
    let rowIndex = -1;
    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][0].toString().trim().toLowerCase() === username.toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { error: 'ไม่พบบัญชีผู้ใช้งานที่ต้องการแก้ไข' };
    }
    
    sheet.getRange(rowIndex, 2).setValue(password);
    sheet.getRange(rowIndex, 3).setValue(name);
    sheet.getRange(rowIndex, 4).setValue(role);
    
    return { success: true, message: 'แก้ไขบัญชีผู้ใช้งานสำเร็จ' };
  } catch (err) {
    return { error: err.message };
  }
}

function deleteUser(username) {
  try {
    const sheet = getOrCreateUsersSheet();
    const sheetData = sheet.getDataRange().getValues();
    
    const uName = (username || '').trim().toLowerCase();
    
    // ห้ามลบ admin เริ่มต้น
    if (uName === 'admin') {
      return { error: 'ไม่สามารถลบบัญชีผู้ดูแลระบบหลัก (admin) ได้' };
    }
    
    let rowIndex = -1;
    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][0].toString().trim().toLowerCase() === uName) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { error: 'ไม่พบบัญชีผู้ใช้งานที่ต้องการลบ' };
    }
    
    sheet.deleteRow(rowIndex);
    return { success: true, message: 'ลบบัญชีผู้ใช้งานสำเร็จ' };
  } catch (err) {
    return { error: err.message };
  }
}

