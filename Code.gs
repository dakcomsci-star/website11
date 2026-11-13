/**
 * ==============================================================================
 * ระบบจัดการหอพักและค่าน้ำค่าไฟ (Dormitory Management System)
 * Google Apps Script Backend (Code.gs)
 * Google Sheets Database + Google Drive Storage + RESTful API
 * ==============================================================================
 */

// Global Sheet Names
const SHEETS = {
  SETTINGS: 'Settings',
  ROOM_TYPES: 'RoomTypes',
  ROOMS: 'Rooms',
  TENANTS: 'Tenants',
  CONTRACTS: 'Contracts',
  METER_READINGS: 'MeterReadings',
  BILLS: 'Bills'
};

/**
 * Helper: แปลงค่าเดือน/ปี ทุกรูปแบบให้เป็น 'YYYY-MM' เสมอ (เช่น '2026-10')
 * ป้องกันปัญหา Google Sheets แปลง '2026-10' เป็น Date Object หรือมีเวลาติดมา
 */
function normalizeMonthYear(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    return `${y}-${m}`;
  }
  const str = String(val).trim();
  if (str === 'all') return 'all';
  
  // Match YYYY-MM or YYYY/MM (e.g. 2026-10, 2026-10-01, 2026-10-01 00:00:00)
  const match = str.match(/(\d{4})[-\/](\d{1,2})/);
  if (match) {
    let y = parseInt(match[1], 10);
    if (y > 2400) y -= 543; // แปลงปี พ.ศ. เป็น ค.ศ.
    const m = String(parseInt(match[2], 10)).padStart(2, '0');
    return `${y}-${m}`;
  }
  return str;
}

/**
 * Web App Entry Point: GET Request
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest(e.parameter.action, e.parameter);
  }

  // Render Web App Page
  const template = HtmlService.createTemplateFromFile('Index');
  try {
    template.initialData = getDashboardData();
  } catch (err) {
    template.initialData = {
      success: false,
      error: err.message,
      settings: { dorm_name: 'ระบบจัดการหอพักและค่าน้ำค่าไฟ' },
      currentMonth: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM'),
      stats: { totalRooms: 0, occupiedRooms: 0, vacantRooms: 0, maintenanceRooms: 0, occupancyRate: 0, monthIncome: 0, monthUnpaid: 0, unpaidCount: 0, pendingMeterCount: 0 }
    };
  }

  return template.evaluate()
    .setTitle('ระบบจัดการหอพักและค่าน้ำค่าไฟ (Google Sheets)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Web App Entry Point: POST Request
 */
function doPost(e) {
  try {
    let payload = {};
    if (e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        payload = e.parameter || {};
      }
    } else if (e.parameter) {
      payload = e.parameter;
    }

    const action = payload.action;
    return handleApiRequest(action, payload);
  } catch (error) {
    return createJsonResponse({ success: false, error: error.message });
  }
}

/**
 * API Router / Dispatcher
 */
function handleApiRequest(action, payload) {
  try {
    let result = null;

    switch (action) {
      case 'initDatabase':
        result = initDatabase();
        break;
      case 'getDashboardData':
        result = getDashboardData(payload.month);
        break;
      case 'getSettings':
        result = getSettings();
        break;
      case 'saveSettings':
        result = saveSettings(payload.settings);
        break;
      case 'getRoomTypes':
        result = getRoomTypes();
        break;
      case 'saveRoomType':
        result = saveRoomType(payload.roomType);
        break;
      case 'deleteRoomType':
        result = deleteRoomType(payload.id);
        break;
      case 'getRoomsList':
        result = getRoomsList();
        break;
      case 'saveRoom':
        result = saveRoom(payload.room);
        break;
      case 'deleteRoom':
        result = deleteRoom(payload.id);
        break;
      case 'importRooms':
        result = importRooms(payload.rooms);
        break;
      case 'getTenants':
        result = getTenants();
        break;
      case 'getContracts':
        result = getContracts();
        break;
      case 'createContract':
        result = createContract(payload.contract);
        break;
      case 'updateContract':
        result = updateContract(payload.contract);
        break;
      case 'terminateContract':
        result = terminateContract(payload.id, payload.actual_end_date);
        break;
      case 'deleteContract':
        result = deleteContract(payload.id);
        break;
      case 'getMeterReadings':
        result = getMeterReadings(payload.month);
        break;
      case 'saveMeterReadings':
        result = saveMeterReadings(payload.month, payload.readings);
        break;
      case 'getBills':
        result = getBills(payload.month, payload.status);
        break;
      case 'getBillDetail':
        result = getBillDetail(payload.bill_id);
        break;
      case 'generateBills':
        result = generateBills(payload.month, payload.due_date);
        break;
      case 'recordPayment':
        result = recordPayment(payload.bill_id, payload.payment_method, payload.slip_base64, payload.slip_name, payload.paid_at);
        break;
      case 'deleteBill':
        result = deleteBill(payload.bill_id);
        break;
      case 'getSpreadsheetInfo':
        const ss = getSpreadsheet();
        result = { id: ss.getId(), name: ss.getName(), url: ss.getUrl() };
        break;
      default:
        throw new Error('ไม่พบคำสั่ง API ที่ระบุ: ' + action);
    }

    return createJsonResponse({ success: true, data: result });
  } catch (err) {
    return createJsonResponse({ success: false, error: err.message });
  }
}

/**
 * Return JSON Response with MIME Type
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 1-Click Database Setup: สร้างชีตและโครงสร้างตารางพร้อมข้อมูลเริ่มต้น
 */
function initDatabase() {
  const ss = getSpreadsheet();
  setupDatabaseStructure(ss);
  return { 
    success: true, 
    spreadsheetUrl: ss.getUrl(),
    message: 'สร้างโครงสร้างตารางฐานข้อมูล Google Sheets ครบถ้วน 7 ตารางเรียบร้อยแล้ว!' 
  };
}

/**
 * สร้างโครงสร้างตาราง 7 แผ่นงานและข้อมูลตัวอย่างเริ่มต้น
 */
function setupDatabaseStructure(ss) {
  function setupSheet(name, headers, initialData = []) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#1e293b')
                 .setFontColor('#ffffff')
                 .setFontWeight('bold');
      sheet.setFrozenRows(1);
      
      if (initialData.length > 0) {
        initialData.forEach(row => sheet.appendRow(row));
      }
    }
    return sheet;
  }

  // 1. Settings Sheet
  setupSheet(SHEETS.SETTINGS, ['key', 'value', 'description'], [
    ['dorm_name', 'หอพักสุขสบาย วิลเลจ', 'ชื่อหอพัก'],
    ['address', '123/45 ซอยสุขใจ 9 ถ.พัฒนาการ แขวงสวนหลวง เขตสวนหลวง กทม. 10250', 'ที่อยู่หอพัก'],
    ['phone', '081-234-5678', 'เบอร์โทรศัพท์ติดต่อ'],
    ['water_rate_default', '18', 'ค่าน้ำต่อหน่วย (บาท)'],
    ['elec_rate_default', '8', 'ค่าไฟต่อหน่วย (บาท)'],
    ['common_fee_default', '0', 'ค่าบริการส่วนกลาง (บาท/เดือน)'],
    ['promptpay_id', '0812345678', 'หมายเลขพร้อมเพย์ (เบอร์มือถือ 10 หลัก หรือ เลขบัตร ปชช. 13 หลัก)'],
    ['promptpay_name', 'นายสมหมาย ใจดี', 'ชื่อบัญชีพร้อมเพย์'],
    ['qr_type', 'generate', 'ประเภท QR Code'],
    ['admin_pin', 'admin123', 'รหัสผ่านสำหรับเข้าสู่ระบบ']
  ]);

  // 2. RoomTypes Sheet
  setupSheet(SHEETS.ROOM_TYPES, ['id', 'name', 'base_price', 'description', 'created_at'], [
    ['1', 'ห้องพัดลม Standard', 2500, 'พัดลมติดผนัง, ระเบียง, ห้องน้ำในตัว, ซิงค์ล้างจาน', new Date()],
    ['2', 'ห้องแอร์ Standard', 3500, 'เครื่องปรับอากาศ Inverter, เตียง 5 ฟุต, ตู้เสื้อผ้า, โต๊ะเครื่องแป้ง', new Date()],
    ['3', 'ห้องแอร์ Deluxe / VIP', 4500, 'แอร์, เฟอร์นิเจอร์บิวท์อิน, เครื่องทำน้ำอุ่น, สมาร์ตทีวี, ตู้เย็น', new Date()]
  ]);

  // 3. Rooms Sheet
  setupSheet(SHEETS.ROOMS, ['id', 'room_number', 'floor', 'room_type_id', 'custom_price', 'status', 'notes', 'created_at'], [
    ['1', '101', 1, '1', 2500, 'occupied', 'ห้องริม ใกล้ประตูทางเข้า', new Date()],
    ['2', '102', 1, '1', 2500, 'vacant', 'ห้องสะอาด พร้อมเข้าอยู่ทันที', new Date()],
    ['3', '201', 2, '2', 3500, 'occupied', '', new Date()],
    ['4', '202', 2, '2', 3500, 'vacant', 'วิวระเบียงโล่ง', new Date()],
    ['5', '301', 3, '3', 4500, 'occupied', '', new Date()],
    ['6', '302', 3, '3', 4500, 'maintenance', 'แอร์ไม่เย็น อยู่ระหว่างซ่อมบำรุง', new Date()]
  ]);

  // 4. Tenants Sheet
  setupSheet(SHEETS.TENANTS, ['id', 'fullname', 'phone', 'id_card', 'notes', 'created_at'], [
    ['1', 'นายสมชาย มุ่งมั่น', '089-111-2233', '1-1002-34567-89-0', '', new Date()],
    ['2', 'น.ส.ดวงใจ รักสงบ', '086-444-5566', '3-2001-98765-43-2', '', new Date()],
    ['3', 'นายณัฐพล ชัยชนะ', '081-777-8899', '1-5099-12345-67-8', '', new Date()]
  ]);

  // 5. Contracts Sheet
  const now = new Date();
  const year = now.getFullYear();
  setupSheet(SHEETS.CONTRACTS, ['id', 'contract_code', 'room_id', 'tenant_id', 'start_date', 'end_date', 'actual_end_date', 'monthly_rent', 'deposit_amount', 'water_rate', 'elec_rate', 'status', 'created_at'], [
    ['1', 'CNT-' + year + '-0001', '1', '1', year + '-01-01', (year + 1) + '-01-01', '', 2500, 5000, 18, 8, 'active', new Date()],
    ['2', 'CNT-' + year + '-0002', '3', '2', year + '-02-01', (year + 1) + '-02-01', '', 3500, 7000, 18, 8, 'active', new Date()],
    ['3', 'CNT-' + year + '-0003', '5', '3', year + '-03-01', (year + 1) + '-03-01', '', 4500, 9000, 18, 8, 'active', new Date()]
  ]);

  // 6. MeterReadings Sheet
  setupSheet(SHEETS.METER_READINGS, ['id', 'room_id', 'month_year', 'water_prev', 'water_curr', 'water_units', 'elec_prev', 'elec_curr', 'elec_units', 'reading_date', 'created_at']);

  // 7. Bills Sheet
  setupSheet(SHEETS.BILLS, ['id', 'bill_code', 'contract_id', 'month_year', 'rent_amount', 'water_units', 'water_rate', 'water_amount', 'elec_units', 'elec_rate', 'elec_amount', 'common_fee', 'other_amount', 'other_desc', 'discount_amount', 'total_amount', 'due_date', 'status', 'payment_method', 'paid_at', 'slip_url', 'created_at']);

  // Remove default Sheet1
  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch (e) {}
  }

  // Ensure Drive folder for slips
  getOrCreateUploadsFolder();
}

/**
 * โฟลเดอร์ Google Drive สำหรับเก็บสลิปการชำระเงิน
 */
function getOrCreateUploadsFolder() {
  const folderName = 'Dormitory_Slips_Uploads';
  const folders = DriveApp.getFoldersByName(folderName);
  let folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  return folder;
}

/**
 * ดึง Spreadsheet ที่ผูกไว้ หรือสร้างใหม่อัตโนมัติ (Zero-Config)
 */
function getSpreadsheet() {
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}

  if (!ss) {
    const propId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (propId) {
      try {
        ss = SpreadsheetApp.openById(propId);
      } catch (e) {}
    }
  }

  if (!ss) {
    try {
      ss = SpreadsheetApp.create('ระบบจัดการหอพักและค่าน้ำค่าไฟ (Google Sheets Database)');
      PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
      setupDatabaseStructure(ss);
    } catch (createErr) {
      throw new Error('ไม่พบ Google Sheet และสร้างอัตโนมัติไม่สำเร็จ: ' + createErr.message);
    }
  }
  return ss;
}

/**
 * Helper: ดึงข้อมูลจากชีตเป็น Array of Objects พร้อม normalize ฟิลด์วันที่และเดือนปี
 */
function getSheetData(sheetName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);
  return rows.map((row, idx) => {
    let obj = { _rowNumber: idx + 2 };
    headers.forEach((h, i) => {
      let val = row[i];
      if (h === 'month_year') {
        val = normalizeMonthYear(val);
      } else if (val instanceof Date) {
        if (val.getHours() === 0 && val.getMinutes() === 0 && val.getSeconds() === 0) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        }
      } else if (typeof val === 'string') {
        val = val.trim();
      }
      obj[h] = val;
    });
    return obj;
  });
}

/**
 * Helper: ดึงข้อมูลการตั้งค่าระบบ
 */
function getSettings() {
  const rows = getSheetData(SHEETS.SETTINGS);
  let settings = {};
  rows.forEach(r => {
    settings[r.key] = r.value;
  });
  return settings;
}

/**
 * Helper: บันทึกการตั้งค่าระบบ
 */
function saveSettings(settingsObj) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.SETTINGS);
  if (!sheet) throw new Error('ไม่พบชีต Settings');

  const data = sheet.getDataRange().getValues();
  const keys = Object.keys(settingsObj);

  keys.forEach(k => {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() == k) {
        sheet.getRange(i + 1, 2).setValue(settingsObj[k]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([k, settingsObj[k], '']);
    }
  });

  return { success: true, message: 'บันทึกการตั้งค่าระบบเรียบร้อยแล้ว' };
}

/**
 * ข้อมูลสำหรับ Dashboard
 */
function getDashboardData(month) {
  const currentMonth = normalizeMonthYear(month) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const settings = getSettings();
  const rooms = getRoomsList();
  const bills = getBills(currentMonth);
  const meterReadings = getMeterReadings(currentMonth);

  let totalRooms = rooms.length;
  let occupiedRooms = rooms.filter(r => r.status === 'occupied').length;
  let vacantRooms = rooms.filter(r => r.status === 'vacant').length;
  let maintenanceRooms = rooms.filter(r => r.status === 'maintenance').length;
  let occupancyRate = totalRooms > 0 ? Math.round((occupiedRooms / totalRooms) * 100) : 0;

  let monthIncome = 0;
  let monthUnpaid = 0;
  let unpaidCount = 0;

  bills.forEach(b => {
    const total = Number(b.total_amount) || 0;
    if (b.status === 'paid') {
      monthIncome += total;
    } else {
      monthUnpaid += total;
      unpaidCount++;
    }
  });

  // Check pending meters (occupied rooms without meter recorded this month)
  const recordedRoomIds = new Set(meterReadings.filter(m => m.is_recorded).map(m => String(m.room_id).trim()));
  const occupiedList = rooms.filter(r => r.status === 'occupied');
  const pendingMeterCount = occupiedList.filter(r => !recordedRoomIds.has(String(r.id).trim())).length;

  return {
    success: true,
    currentMonth: currentMonth,
    settings: settings,
    stats: {
      totalRooms,
      occupiedRooms,
      vacantRooms,
      maintenanceRooms,
      occupancyRate,
      monthIncome,
      monthUnpaid,
      unpaidCount,
      pendingMeterCount
    },
    unpaidBills: bills.filter(b => b.status !== 'paid').slice(0, 10),
    rooms: rooms
  };
}

/**
 * ดึงรายการประเภทห้องพัก
 */
function getRoomTypes() {
  return getSheetData(SHEETS.ROOM_TYPES);
}

/**
 * บันทึก/แก้ไขประเภทห้องพัก
 */
function saveRoomType(rt) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ROOM_TYPES);
  if (!sheet) throw new Error('ไม่พบชีต RoomTypes');

  if (rt.id) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(rt.id).trim()) {
        sheet.getRange(i + 1, 2).setValue(rt.name);
        sheet.getRange(i + 1, 3).setValue(Number(rt.base_price));
        sheet.getRange(i + 1, 4).setValue(rt.description || '');
        return { success: true, message: 'แก้ไขประเภทห้องพักเรียบร้อยแล้ว' };
      }
    }
  }

  const newId = String(Date.now());
  sheet.appendRow([newId, rt.name, Number(rt.base_price), rt.description || '', new Date()]);
  return { success: true, message: 'เพิ่มประเภทห้องพักเรียบร้อยแล้ว', id: newId };
}

/**
 * ลบประเภทห้องพัก
 */
function deleteRoomType(id) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ROOM_TYPES);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบประเภทห้องพักเรียบร้อยแล้ว' };
    }
  }
  throw new Error('ไม่พบประเภทห้องพักที่ต้องการลบ');
}

/**
 * ดึงรายการห้องพักทั้งหมด พร้อมข้อมูลประเภทห้องและผู้เช่าปัจจุบัน
 */
function getRoomsList() {
  const rooms = getSheetData(SHEETS.ROOMS);
  const roomTypes = getRoomTypes();
  const typeMap = {};
  roomTypes.forEach(t => { typeMap[String(t.id).trim()] = t; });

  const contracts = getSheetData(SHEETS.CONTRACTS).filter(c => c.status === 'active');
  const tenants = getSheetData(SHEETS.TENANTS);
  const tenantMap = {};
  tenants.forEach(t => { tenantMap[String(t.id).trim()] = t; });

  const roomContractMap = {};
  contracts.forEach(c => {
    roomContractMap[String(c.room_id).trim()] = {
      contract: c,
      tenant: tenantMap[String(c.tenant_id).trim()] || null
    };
  });

  return rooms.map(r => {
    const rType = typeMap[String(r.room_type_id).trim()] || { name: 'ไม่ระบุ', base_price: 0 };
    const contractInfo = roomContractMap[String(r.id).trim()];
    return {
      ...r,
      type_name: rType.name,
      base_price: rType.base_price,
      effective_price: r.custom_price ? Number(r.custom_price) : Number(rType.base_price),
      tenant_name: contractInfo && contractInfo.tenant ? contractInfo.tenant.fullname : null,
      tenant_phone: contractInfo && contractInfo.tenant ? contractInfo.tenant.phone : null,
      contract_id: contractInfo && contractInfo.contract ? contractInfo.contract.id : null
    };
  });
}

/**
 * บันทึก/แก้ไขห้องพัก
 */
function saveRoom(room) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ROOMS);
  if (!sheet) throw new Error('ไม่พบชีต Rooms');

  if (room.id) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(room.id).trim()) {
        sheet.getRange(i + 1, 2).setValue(String(room.room_number));
        sheet.getRange(i + 1, 3).setValue(Number(room.floor));
        sheet.getRange(i + 1, 4).setValue(String(room.room_type_id));
        sheet.getRange(i + 1, 5).setValue(room.custom_price ? Number(room.custom_price) : '');
        sheet.getRange(i + 1, 6).setValue(room.status || 'vacant');
        sheet.getRange(i + 1, 7).setValue(room.notes || '');
        return { success: true, message: 'แก้ไขข้อมูลห้องพักเรียบร้อยแล้ว' };
      }
    }
  }

  const newId = String(Date.now());
  sheet.appendRow([
    newId,
    String(room.room_number),
    Number(room.floor || 1),
    String(room.room_type_id),
    room.custom_price ? Number(room.custom_price) : '',
    room.status || 'vacant',
    room.notes || '',
    new Date()
  ]);

  return { success: true, message: 'เพิ่มห้องพักใหม่เรียบร้อยแล้ว', id: newId };
}

/**
 * ลบห้องพัก
 */
function deleteRoom(id) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ROOMS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบห้องพักเรียบร้อยแล้ว' };
    }
  }
  throw new Error('ไม่พบห้องพักที่ต้องการลบ');
}

/**
 * นำเข้าข้อมูลห้องพักแบบชุด (CSV / Batch)
 */
function importRooms(roomsArray) {
  if (!Array.isArray(roomsArray) || roomsArray.length === 0) {
    throw new Error('ไม่มีข้อมูลห้องพักสำหรับนำเข้า');
  }

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ROOMS);
  let imported = 0;

  roomsArray.forEach(r => {
    if (r.room_number) {
      const newId = String(Date.now() + Math.floor(Math.random() * 1000));
      sheet.appendRow([
        newId,
        String(r.room_number),
        Number(r.floor || 1),
        String(r.room_type_id || '1'),
        r.custom_price ? Number(r.custom_price) : '',
        r.status || 'vacant',
        r.notes || '',
        new Date()
      ]);
      imported++;
    }
  });

  return { success: true, message: `นำเข้าข้อมูลห้องพักเรียบร้อยแล้วจำนวน ${imported} ห้อง` };
}

/**
 * ดึงรายชื่อผู้เช่าทั้งหมด
 */
function getTenants() {
  return getSheetData(SHEETS.TENANTS);
}

/**
 * ดึงรายการสัญญาเช่าทั้งหมด พร้อมข้อมูลห้องและผู้เช่า
 */
function getContracts() {
  const contracts = getSheetData(SHEETS.CONTRACTS);
  const rooms = getSheetData(SHEETS.ROOMS);
  const tenants = getSheetData(SHEETS.TENANTS);

  const roomMap = {};
  rooms.forEach(r => { roomMap[String(r.id).trim()] = r; });

  const tenantMap = {};
  tenants.forEach(t => { tenantMap[String(t.id).trim()] = t; });

  return contracts.map(c => {
    const r = roomMap[String(c.room_id).trim()] || { room_number: 'N/A' };
    const t = tenantMap[String(c.tenant_id).trim()] || { fullname: 'N/A', phone: '-', id_card: '-' };
    return {
      ...c,
      room_number: r.room_number,
      tenant_name: t.fullname,
      tenant_phone: t.phone,
      tenant_id_card: t.id_card
    };
  });
}

/**
 * ทำสัญญาเช่าใหม่ (ย้ายเข้า)
 */
function createContract(contractData) {
  const ss = getSpreadsheet();
  const tenantSheet = ss.getSheetByName(SHEETS.TENANTS);
  const contractSheet = ss.getSheetByName(SHEETS.CONTRACTS);
  const roomSheet = ss.getSheetByName(SHEETS.ROOMS);

  // 1. Save or Find Tenant
  let tenantId = contractData.tenant_id;
  if (!tenantId) {
    tenantId = String(Date.now());
    tenantSheet.appendRow([
      tenantId,
      contractData.fullname,
      contractData.phone,
      contractData.id_card || '',
      contractData.tenant_notes || '',
      new Date()
    ]);
  }

  // 2. Generate Contract Code
  const year = new Date().getFullYear();
  const cData = contractSheet.getDataRange().getValues();
  const contractCode = 'CNT-' + year + '-' + String(cData.length).padStart(4, '0');
  const contractId = String(Date.now() + 1);

  // 3. Save Contract
  contractSheet.appendRow([
    contractId,
    contractCode,
    String(contractData.room_id),
    String(tenantId),
    contractData.start_date,
    contractData.end_date || '',
    '',
    Number(contractData.monthly_rent),
    Number(contractData.deposit_amount || 0),
    Number(contractData.water_rate || 18),
    Number(contractData.elec_rate || 8),
    'active',
    new Date()
  ]);

  // 4. Update Room Status to 'occupied'
  const rData = roomSheet.getDataRange().getValues();
  for (let i = 1; i < rData.length; i++) {
    if (String(rData[i][0]).trim() === String(contractData.room_id).trim()) {
      roomSheet.getRange(i + 1, 6).setValue('occupied');
      break;
    }
  }

  return { success: true, message: 'บันทึกสัญญาเช่าใหม่และเปลี่ยนสถานะห้องเป็นมีผู้เช่าแล้ว', contract_id: contractId };
}

/**
 * แก้ไขสัญญาเช่า
 */
function updateContract(c) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CONTRACTS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(c.id).trim()) {
      sheet.getRange(i + 1, 8).setValue(Number(c.monthly_rent));
      sheet.getRange(i + 1, 9).setValue(Number(c.deposit_amount));
      sheet.getRange(i + 1, 10).setValue(Number(c.water_rate));
      sheet.getRange(i + 1, 11).setValue(Number(c.elec_rate));
      return { success: true, message: 'แก้ไขสัญญาเช่าเรียบร้อยแล้ว' };
    }
  }
  throw new Error('ไม่พบสัญญาเช่าที่ต้องการแก้ไข');
}

/**
 * Check-out คืนห้อง (สิ้นสุดสัญญา)
 */
function terminateContract(contractId, actualEndDate) {
  const ss = getSpreadsheet();
  const contractSheet = ss.getSheetByName(SHEETS.CONTRACTS);
  const roomSheet = ss.getSheetByName(SHEETS.ROOMS);

  const cData = contractSheet.getDataRange().getValues();
  let roomId = null;

  for (let i = 1; i < cData.length; i++) {
    if (String(cData[i][0]).trim() === String(contractId).trim()) {
      roomId = cData[i][2];
      contractSheet.getRange(i + 1, 7).setValue(actualEndDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
      contractSheet.getRange(i + 1, 12).setValue('terminated');
      break;
    }
  }

  if (roomId) {
    const rData = roomSheet.getDataRange().getValues();
    for (let j = 1; j < rData.length; j++) {
      if (String(rData[j][0]).trim() === String(roomId).trim()) {
        roomSheet.getRange(j + 1, 6).setValue('vacant');
        break;
      }
    }
  }

  return { success: true, message: 'สิ้นสุดสัญญาเช่า (Check-out) และเปลี่ยนสถานะห้องเป็นห้องว่างแล้ว' };
}

/**
 * ลบสัญญาเช่า
 */
function deleteContract(id) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CONTRACTS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบสัญญาเช่าเรียบร้อยแล้ว' };
    }
  }
  throw new Error('ไม่พบสัญญาเช่าที่ต้องการลบ');
}

/**
 * ดึงข้อมูลมิเตอร์น้ำ-ไฟประจำเดือน (พร้อม normalize month_year)
 */
function getMeterReadings(monthYear) {
  const targetMonth = normalizeMonthYear(monthYear) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const allReadings = getSheetData(SHEETS.METER_READINGS);
  const rooms = getRoomsList();

  const parts = targetMonth.split('-');
  let y = parseInt(parts[0], 10);
  let m = parseInt(parts[1], 10) - 1;
  if (m === 0) { m = 12; y--; }
  const prevMonthStr = `${y}-${String(m).padStart(2, '0')}`;

  const currentMap = {};
  const prevMap = {};

  allReadings.forEach(r => {
    const rMonth = normalizeMonthYear(r.month_year);
    const rRoomId = String(r.room_id).trim();
    if (rMonth === targetMonth) {
      currentMap[rRoomId] = r;
    }
    if (rMonth === prevMonthStr) {
      prevMap[rRoomId] = r;
    }
  });

  return rooms.map(r => {
    const rId = String(r.id).trim();
    const curr = currentMap[rId] || {};
    const prev = prevMap[rId] || {};

    const waterPrev = curr.water_prev !== undefined && curr.water_prev !== '' ? Number(curr.water_prev) : (prev.water_curr !== undefined ? Number(prev.water_curr) : 0);
    const elecPrev = curr.elec_prev !== undefined && curr.elec_prev !== '' ? Number(curr.elec_prev) : (prev.elec_curr !== undefined ? Number(prev.elec_curr) : 0);

    return {
      room_id: r.id,
      room_number: r.room_number,
      floor: r.floor,
      status: r.status,
      tenant_name: r.tenant_name,
      month_year: targetMonth,
      water_prev: waterPrev,
      water_curr: curr.water_curr !== undefined ? curr.water_curr : '',
      water_units: curr.water_units !== undefined ? curr.water_units : 0,
      elec_prev: elecPrev,
      elec_curr: curr.elec_curr !== undefined ? curr.elec_curr : '',
      elec_units: curr.elec_units !== undefined ? curr.elec_units : 0,
      reading_date: curr.reading_date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      is_recorded: curr.water_curr !== undefined && curr.water_curr !== ''
    };
  });
}

/**
 * บันทึกการจดมิเตอร์น้ำ-ไฟ
 */
function saveMeterReadings(monthYear, readingsArray) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.METER_READINGS);
  const data = sheet.getDataRange().getValues();
  const targetMonth = normalizeMonthYear(monthYear);

  readingsArray.forEach(item => {
    const wPrev = Number(item.water_prev) || 0;
    const wCurr = Number(item.water_curr) || 0;
    const wUnits = Math.max(0, wCurr - wPrev);

    const ePrev = Number(item.elec_prev) || 0;
    const eCurr = Number(item.elec_curr) || 0;
    const eUnits = Math.max(0, eCurr - ePrev);

    let foundRow = -1;
    for (let i = 1; i < data.length; i++) {
      const rowRoomId = String(data[i][1]).trim();
      const rowMonth = normalizeMonthYear(data[i][2]);
      if (rowRoomId === String(item.room_id).trim() && rowMonth === targetMonth) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow > 0) {
      sheet.getRange(foundRow, 4).setValue(wPrev);
      sheet.getRange(foundRow, 5).setValue(wCurr);
      sheet.getRange(foundRow, 6).setValue(wUnits);
      sheet.getRange(foundRow, 7).setValue(ePrev);
      sheet.getRange(foundRow, 8).setValue(eCurr);
      sheet.getRange(foundRow, 9).setValue(eUnits);
      sheet.getRange(foundRow, 10).setValue(item.reading_date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    } else {
      const newId = String(Date.now() + Math.floor(Math.random() * 1000));
      sheet.appendRow([
        newId,
        String(item.room_id),
        "'" + targetMonth,
        wPrev,
        wCurr,
        wUnits,
        ePrev,
        eCurr,
        eUnits,
        item.reading_date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        new Date()
      ]);
    }
  });

  return { success: true, message: 'บันทึกเลขอ่านมิเตอร์น้ำ-ไฟเรียบร้อยแล้ว' };
}

/**
 * ดึงรายการบิลตามเดือนและสถานะ (แก้ปัญหาไม่แสดงบิล 100%)
 */
function getBills(monthYear, statusFilter) {
  const bills = getSheetData(SHEETS.BILLS);
  const contracts = getSheetData(SHEETS.CONTRACTS);
  const rooms = getSheetData(SHEETS.ROOMS);
  const tenants = getSheetData(SHEETS.TENANTS);

  const cMap = {};
  contracts.forEach(c => {
    if (c.id !== undefined && c.id !== null) {
      cMap[String(c.id).trim()] = c;
    }
  });

  const rMap = {};
  rooms.forEach(r => {
    if (r.id !== undefined && r.id !== null) {
      rMap[String(r.id).trim()] = r;
    }
  });

  const tMap = {};
  tenants.forEach(t => {
    if (t.id !== undefined && t.id !== null) {
      tMap[String(t.id).trim()] = t;
    }
  });

  const targetMonth = normalizeMonthYear(monthYear);
  let filtered = bills;

  // กรองตามเดือน (ถ้ามีระบุ และไม่ใช่ 'all')
  if (targetMonth && targetMonth !== 'all') {
    filtered = filtered.filter(b => normalizeMonthYear(b.month_year) === targetMonth);
  }

  // กรองตามสถานะการชำระเงิน
  if (statusFilter && statusFilter !== 'all') {
    filtered = filtered.filter(b => String(b.status).trim() === statusFilter);
  }

  return filtered.map(b => {
    const cId = String(b.contract_id || '').trim();
    const c = cMap[cId] || {};
    const rId = String(c.room_id || '').trim();
    const r = rMap[rId] || {};
    const tId = String(c.tenant_id || '').trim();
    const t = tMap[tId] || {};

    return {
      ...b,
      month_year: normalizeMonthYear(b.month_year),
      room_number: r.room_number || (c.room_id ? 'ห้อง ' + c.room_id : 'N/A'),
      tenant_name: t.fullname || 'N/A',
      tenant_phone: t.phone || '-'
    };
  });
}

/**
 * ดึงรายละเอียดบิลสำหรับพิมพ์ใบแจ้งหนี้ / ใบเสร็จ
 */
function getBillDetail(billId) {
  const bills = getSheetData(SHEETS.BILLS);
  const b = bills.find(x => String(x.id).trim() === String(billId).trim());
  if (!b) throw new Error('ไม่พบบิลรหัส: ' + billId);

  const contracts = getSheetData(SHEETS.CONTRACTS);
  const rooms = getSheetData(SHEETS.ROOMS);
  const tenants = getSheetData(SHEETS.TENANTS);
  const settings = getSettings();

  const cMap = {};
  contracts.forEach(c => { cMap[String(c.id).trim()] = c; });

  const rMap = {};
  rooms.forEach(r => { rMap[String(r.id).trim()] = r; });

  const tMap = {};
  tenants.forEach(t => { tMap[String(t.id).trim()] = t; });

  const c = cMap[String(b.contract_id || '').trim()] || {};
  const r = rMap[String(c.room_id || '').trim()] || {};
  const t = tMap[String(c.tenant_id || '').trim()] || {};

  // Generate PromptPay EMVCo Payload
  let promptpayPayload = '';
  if (settings.promptpay_id) {
    promptpayPayload = generatePromptPayPayload(settings.promptpay_id, b.total_amount);
  }

  return {
    ...b,
    month_year: normalizeMonthYear(b.month_year),
    room_number: r.room_number || (c.room_id ? 'ห้อง ' + c.room_id : 'N/A'),
    tenant_name: t.fullname || 'N/A',
    tenant_phone: t.phone || '-',
    promptpay_payload: promptpayPayload,
    dorm_settings: settings
  };
}

/**
 * สร้างบิลประจำงวดแบบอัตโนมัติ (1-Click Generate Bills)
 */
function generateBills(monthYear, dueDate) {
  const ss = getSpreadsheet();
  const billSheet = ss.getSheetByName(SHEETS.BILLS);
  const settings = getSettings();
  const targetMonth = normalizeMonthYear(monthYear);

  const contracts = getSheetData(SHEETS.CONTRACTS).filter(c => c.status === 'active');
  const meterReadings = getMeterReadings(targetMonth);
  const meterMap = {};
  meterReadings.forEach(m => { meterMap[String(m.room_id).trim()] = m; });

  const existingBills = getSheetData(SHEETS.BILLS).filter(b => normalizeMonthYear(b.month_year) === targetMonth);
  const billedContractIds = new Set(existingBills.map(b => String(b.contract_id).trim()));

  const defaultWaterRate = Number(settings.water_rate_default || 18);
  const defaultElecRate = Number(settings.elec_rate_default || 8);
  const defaultCommonFee = Number(settings.common_fee_default || 0);

  let generatedCount = 0;
  const currentCount = existingBills.length;

  contracts.forEach((c) => {
    if (billedContractIds.has(String(c.id).trim())) return; // ข้ามสัญญาที่สร้างบิลเดือนนี้แล้ว

    const meter = meterMap[String(c.room_id).trim()] || { water_units: 0, elec_units: 0 };
    const rentAmount = Number(c.monthly_rent || 0);
    const waterRate = Number(c.water_rate || defaultWaterRate);
    const elecRate = Number(c.elec_rate || defaultElecRate);
    const waterUnits = Number(meter.water_units || 0);
    const elecUnits = Number(meter.elec_units || 0);

    const waterAmount = waterUnits * waterRate;
    const elecAmount = elecUnits * elecRate;
    const commonFee = defaultCommonFee;
    const totalAmount = rentAmount + waterAmount + elecAmount + commonFee;

    const billNumber = 'INV-' + targetMonth.replace('-', '') + '-' + String(currentCount + generatedCount + 1).padStart(3, '0');
    const billId = String(Date.now() + generatedCount);

    billSheet.appendRow([
      billId,
      billNumber,
      String(c.id),
      "'" + targetMonth, // เก็บเป็นข้อความชัดเจน
      rentAmount,
      waterUnits,
      waterRate,
      waterAmount,
      elecUnits,
      elecRate,
      elecAmount,
      commonFee,
      0, // other_amount
      '', // other_desc
      0, // discount_amount
      totalAmount,
      dueDate || '',
      'unpaid',
      '', // payment_method
      '', // paid_at
      '', // slip_url
      new Date()
    ]);

    generatedCount++;
  });

  return {
    success: true,
    message: `สร้างบิลประจำงวด ${targetMonth} เรียบร้อยแล้วจำนวน ${generatedCount} ใบ`,
    generated_count: generatedCount
  };
}

/**
 * บันทึกรับชำระเงินและแนบสลิป
 */
function recordPayment(billId, paymentMethod, slipBase64, slipName, paidAt) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.BILLS);
  const data = sheet.getDataRange().getValues();

  let slipUrl = '';
  if (slipBase64) {
    try {
      const folder = getOrCreateUploadsFolder();
      const contentType = slipBase64.substring(slipBase64.indexOf(':') + 1, slipBase64.indexOf(';')) || 'image/jpeg';
      const base64Data = slipBase64.split(',')[1] || slipBase64;
      const decodedBlob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, slipName || `slip_${billId}.jpg`);
      const file = folder.createFile(decodedBlob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      slipUrl = file.getUrl();
    } catch (e) {
      slipUrl = 'error_upload: ' + e.message;
    }
  }

  const paymentTime = paidAt || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(billId).trim()) {
      sheet.getRange(i + 1, 18).setValue('paid');
      sheet.getRange(i + 1, 19).setValue(paymentMethod || 'โอนเงิน');
      sheet.getRange(i + 1, 20).setValue(paymentTime);
      if (slipUrl) {
        sheet.getRange(i + 1, 21).setValue(slipUrl);
      }
      return { success: true, message: 'บันทึกรับเงินเรียบร้อยแล้ว บิลเปลี่ยนสถานะเป็นชำระแล้ว', slip_url: slipUrl };
    }
  }

  throw new Error('ไม่พบบิลที่ต้องการบันทึกรับเงิน');
}

/**
 * ลบบิล
 */
function deleteBill(billId) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.BILLS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(billId).trim()) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบบิลเรียบร้อยแล้ว' };
    }
  }
  throw new Error('ไม่พบบิลที่ต้องการลบ');
}

/**
 * ==============================================================================
 * BOT EMVCo PromptPay QR Code Generator
 * ==============================================================================
 */
function calculateCRC16(data) {
  let crc = 0xFFFF;
  for (let i = 0; i < data.length; i++) {
    let x = ((crc >> 8) ^ data.charCodeAt(i)) & 0xFF;
    x ^= x >> 4;
    crc = ((crc << 8) ^ (x << 12) ^ (x << 5) ^ x) & 0xFFFF;
  }
  return crc.toString(16).toUpperCase().padStart(4, '0');
}

function formatPromptPayTarget(target) {
  const cleaned = String(target).replace(/[^0-9]/g, '');
  if (cleaned.length === 13) {
    return { type: '02', value: cleaned };
  } else {
    if (cleaned.startsWith('0')) {
      return { type: '01', value: '0066' + cleaned.substring(1) };
    }
    return { type: '01', value: '0066' + cleaned };
  }
}

function generatePromptPayPayload(target, amount) {
  const targetInfo = formatPromptPayTarget(target);
  const targetType = targetInfo.type;
  const formattedTarget = targetInfo.value;

  const subTag00 = '0016A000000677010111';
  const subTagTarget = targetType + String(formattedTarget.length).padStart(2, '0') + formattedTarget;
  const targetSubField = subTag00 + subTagTarget;

  const tag29 = '29' + String(targetSubField.length).padStart(2, '0') + targetSubField;
  const numAmount = Number(amount || 0);

  let payload = '000201' + (numAmount > 0 ? '010212' : '010211') + tag29 + '5303764';

  if (numAmount > 0) {
    const amountStr = numAmount.toFixed(2);
    payload += '54' + String(amountStr.length).padStart(2, '0') + amountStr;
  }

  payload += '5802TH';
  const payloadDataForCRC = payload + '6304';
  const checksum = calculateCRC16(payloadDataForCRC);

  return payloadDataForCRC + checksum;
}
