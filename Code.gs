/**
 * Champion Petshop — ระบบจองคิวอาบน้ำ/ตัดขนสัตว์เลี้ยง
 * Google Apps Script Web App (แบ็กเอนด์ API) ทำงานคู่กับ Google Sheet
 *
 * ============================================================
 * วิธีติดตั้ง (สรุปสั้น ๆ — รายละเอียดเต็มอยู่ในคู่มือแยกต่างหาก)
 * ============================================================
 * 1. สร้าง Google Sheet ใหม่ แล้วสร้าง 4 แท็บชื่อ (ตัวพิมพ์เล็ก-ใหญ่ต้องตรง):
 *      Bookings, Services, Staff, Settings
 *    ใส่หัวตารางตามที่กำหนดไว้ใน ensureSheets() ด้านล่าง — หรือรัน
 *    ฟังก์ชัน setupSheets() หนึ่งครั้งจาก Apps Script editor เพื่อให้ระบบ
 *    สร้างแท็บ/หัวตาราง/ข้อมูลตัวอย่างให้อัตโนมัติ
 * 2. เปิด Extensions > Apps Script บน Sheet นี้ แล้ววางไฟล์นี้ทับ Code.gs
 * 3. รัน setupSheets() หนึ่งครั้ง (เมนู Run ด้านบน เลือกฟังก์ชัน setupSheets)
 *    เพื่อสร้างหัวตาราง + ข้อมูลตั้งต้น
 * 4. Deploy > New deployment > Web app
 *      Execute as: Me
 *      Who has access: Anyone
 *    กด Deploy แล้วคัดลอก URL ที่ลงท้ายด้วย /exec
 * 5. นำ URL ไปวางในไฟล์เว็บแอพ (ตัวแปร APPS_SCRIPT_URL ในหน้าเว็บ)
 * ============================================================
 */

// ---------- ค่าตั้งต้น (ใช้เมื่อไม่มีข้อมูลในแท็บ Settings/Staff/Services) ----------
var DEFAULTS = {
  shopName: 'Champion Petshop',
  tagline: 'จองคิวอาบน้ำ-ตัดขน',
  hoursText: '10:00–20:00 น. ทุกวัน',
  lineOaId: '@championpetshop',
  openTime: '10:00',
  closeTime: '20:00',
  slotMinutes: 30,
  daysAhead: 14,
  staffList: ['ช่างพิมพ์', 'พี่ใหม่', 'เต้น', 'ว่ายน้ำ'],
  services: [
    { name: 'อาบน้ำ-เป่าขน', durationMinutes: 60, price: 300 },
    { name: 'อาบน้ำ-ตัดขน', durationMinutes: 120, price: 600 },
    { name: 'ตัดเล็บ-ทำความสะอาดหู', durationMinutes: 30, price: 150 },
    { name: 'สปาบำรุงขน', durationMinutes: 90, price: 450 }
  ]
};

var SHEET_NAMES = {
  bookings: 'Bookings',
  services: 'Services',
  staff: 'Staff',
  settings: 'Settings'
};

var BOOKING_HEADERS = ['BookingID', 'Timestamp', 'Date', 'StartTime', 'EndTime', 'Staff', 'Status', 'CustomerName', 'Phone', 'PetName', 'PetType', 'Service', 'Notes'];
var SERVICE_HEADERS = ['ServiceName', 'DurationMinutes', 'Price'];
var STAFF_HEADERS = ['StaffName', 'Active'];
var SETTINGS_HEADERS = ['Key', 'Value'];

var STATUS_ACTIVE = 'จอง';       // นัดจริง มีผลกันคิว
var STATUS_CANCELLED = 'ยกเลิก'; // ยกเลิกแล้ว ไม่กันคิว
var STATUS_BLOCKED = 'ปิด';      // ช่างหยุด/ปิดคิวช่วงนั้น มีผลกันคิว

// ================================================================
// ตั้งค่าเริ่มต้น — รันครั้งเดียวจาก Apps Script editor
// ================================================================
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var bookings = ss.getSheetByName(SHEET_NAMES.bookings) || ss.insertSheet(SHEET_NAMES.bookings);
  if (bookings.getLastRow() === 0) {
    bookings.appendRow(BOOKING_HEADERS);
    bookings.setFrozenRows(1);
  }

  var services = ss.getSheetByName(SHEET_NAMES.services) || ss.insertSheet(SHEET_NAMES.services);
  if (services.getLastRow() === 0) {
    services.appendRow(SERVICE_HEADERS);
    DEFAULTS.services.forEach(function (s) {
      services.appendRow([s.name, s.durationMinutes, s.price]);
    });
    services.setFrozenRows(1);
  }

  var staff = ss.getSheetByName(SHEET_NAMES.staff) || ss.insertSheet(SHEET_NAMES.staff);
  if (staff.getLastRow() === 0) {
    staff.appendRow(STAFF_HEADERS);
    DEFAULTS.staffList.forEach(function (name) {
      staff.appendRow([name, true]);
    });
    staff.setFrozenRows(1);
  }

  var settings = ss.getSheetByName(SHEET_NAMES.settings) || ss.insertSheet(SHEET_NAMES.settings);
  if (settings.getLastRow() === 0) {
    settings.appendRow(SETTINGS_HEADERS);
    var rows = [
      ['ShopName', DEFAULTS.shopName],
      ['Tagline', DEFAULTS.tagline],
      ['HoursText', DEFAULTS.hoursText],
      ['LineOaId', DEFAULTS.lineOaId],
      ['OpenTime', DEFAULTS.openTime],
      ['CloseTime', DEFAULTS.closeTime],
      ['SlotMinutes', DEFAULTS.slotMinutes],
      ['DaysAhead', DEFAULTS.daysAhead]
    ];
    rows.forEach(function (r) { settings.appendRow(r); });
    settings.setFrozenRows(1);
  }

  SpreadsheetApp.flush();
  return 'ตั้งค่าเรียบร้อย: ' + [SHEET_NAMES.bookings, SHEET_NAMES.services, SHEET_NAMES.staff, SHEET_NAMES.settings].join(', ');
}

// ================================================================
// Web App entry points
// ================================================================
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'schedule';
  try {
    if (action === 'schedule') {
      return jsonOut(buildScheduleResponse());
    }
    return jsonOut({ error: 'unknown action: ' + action });
  } catch (err) {
    return jsonOut({ error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    var raw = e && e.postData && e.postData.contents;
    if (!raw) return jsonOut({ error: 'missing request body' });
    var body = JSON.parse(raw);
    var action = body.action || 'book';

    if (action === 'book') {
      return jsonOut(createBooking(body));
    }
    if (action === 'cancel') {
      return jsonOut(cancelBooking(body));
    }
    return jsonOut({ error: 'unknown action: ' + action });
  } catch (err) {
    return jsonOut({ error: String(err && err.message || err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ================================================================
// อ่านการตั้งค่า / รายชื่อช่าง / บริการ จาก Sheet
// ================================================================
function getSettings() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.settings);
  var out = {
    shopName: DEFAULTS.shopName, tagline: DEFAULTS.tagline, hoursText: DEFAULTS.hoursText,
    lineOaId: DEFAULTS.lineOaId, openTime: DEFAULTS.openTime, closeTime: DEFAULTS.closeTime,
    slotMinutes: DEFAULTS.slotMinutes, daysAhead: DEFAULTS.daysAhead
  };
  if (!sh || sh.getLastRow() < 2) return out;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var map = {};
  rows.forEach(function (r) { if (r[0]) map[String(r[0]).trim()] = r[1]; });
  if (map.ShopName) out.shopName = map.ShopName;
  if (map.Tagline) out.tagline = map.Tagline;
  if (map.HoursText) out.hoursText = map.HoursText;
  if (map.LineOaId) out.lineOaId = map.LineOaId;
  if (map.OpenTime) out.openTime = map.OpenTime;
  if (map.CloseTime) out.closeTime = map.CloseTime;
  if (map.SlotMinutes) out.slotMinutes = Number(map.SlotMinutes);
  if (map.DaysAhead) out.daysAhead = Number(map.DaysAhead);
  return out;
}

function getStaffList() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.staff);
  if (!sh || sh.getLastRow() < 2) return DEFAULTS.staffList.slice();
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var list = rows.filter(function (r) { return r[0] && (r[1] === true || r[1] === 'TRUE' || r[1] === 1); })
    .map(function (r) { return String(r[0]).trim(); });
  return list.length ? list : DEFAULTS.staffList.slice();
}

function getServices() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.services);
  if (!sh || sh.getLastRow() < 2) return DEFAULTS.services.slice();
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  var list = rows.filter(function (r) { return r[0]; })
    .map(function (r) { return { name: String(r[0]).trim(), durationMinutes: Number(r[1]) || 30, price: Number(r[2]) || 0 }; });
  return list.length ? list : DEFAULTS.services.slice();
}

// ================================================================
// อ่านรายการจอง (เฉพาะที่ยังมีผลกันคิว) ในช่วงวันที่ที่ต้องการ
// ================================================================
function getActiveBookingsInRange(startDate, endDate) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.bookings);
  if (!sh || sh.getLastRow() < 2) return [];
  var numRows = sh.getLastRow() - 1;
  var data = sh.getRange(2, 1, numRows, BOOKING_HEADERS.length).getValues();
  var idx = headerIndexMap();
  var out = [];
  data.forEach(function (row) {
    var status = row[idx.Status];
    if (status !== STATUS_ACTIVE && status !== STATUS_BLOCKED) return;
    var d = row[idx.Date];
    var dateObj = (d instanceof Date) ? d : new Date(d);
    if (dateObj < startDate || dateObj > endDate) return;
    out.push({
      date: formatDateYMD(dateObj),
      startTime: normalizeTime(row[idx.StartTime]),
      endTime: normalizeTime(row[idx.EndTime]),
      staff: row[idx.Staff],
      status: status
    });
  });
  return out;
}

function headerIndexMap() {
  var m = {};
  BOOKING_HEADERS.forEach(function (h, i) { m[h] = i; });
  return m;
}

// ================================================================
// สร้าง JSON ตารางคิวว่าง (ใช้ทั้งหน้าเช็กคิว และคำนวณช่วงว่างสำหรับจอง)
// ================================================================
function buildScheduleResponse() {
  var settings = getSettings();
  var staffList = getStaffList();
  var services = getServices();

  var openMinutes = timeToMinutes(settings.openTime);
  var closeMinutes = timeToMinutes(settings.closeTime);
  var slotMinutes = settings.slotMinutes;
  var daysAhead = settings.daysAhead;

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var rangeEnd = new Date(today);
  rangeEnd.setDate(rangeEnd.getDate() + daysAhead);

  var bookings = getActiveBookingsInRange(today, rangeEnd);
  // จัดกลุ่มตาม date+staff เพื่อค้นหาทับซ้อนได้เร็ว
  var byDateStaff = {};
  bookings.forEach(function (b) {
    var key = b.date + '|' + b.staff;
    (byDateStaff[key] = byDateStaff[key] || []).push(b);
  });

  var days = [];
  for (var d = 0; d < daysAhead; d++) {
    var date = new Date(today);
    date.setDate(date.getDate() + d);
    var dateStr = formatDateYMD(date);

    var slots = [];
    for (var m = openMinutes; m < closeMinutes; m += slotMinutes) {
      var slotStart = m, slotEnd = m + slotMinutes;
      var staffState = {};
      staffList.forEach(function (name) {
        var list = byDateStaff[dateStr + '|' + name] || [];
        var state = null; // null = ว่าง
        for (var i = 0; i < list.length; i++) {
          var b = list[i];
          var bStart = timeToMinutes(b.startTime), bEnd = timeToMinutes(b.endTime);
          if (slotStart < bEnd && slotEnd > bStart) {
            state = (b.status === STATUS_BLOCKED) ? 'closed' : 'booked';
            if (state === 'booked') break; // booked ชนะ closed ถ้าทับกันแปลก ๆ
          }
        }
        staffState[name] = state;
      });
      slots.push({ start: minutesToTime(slotStart), staff: staffState });
    }
    days.push({ date: dateStr, slots: slots });
  }

  return {
    shopName: settings.shopName,
    tagline: settings.tagline,
    hoursText: settings.hoursText,
    lineOaId: settings.lineOaId,
    openTime: settings.openTime,
    closeTime: settings.closeTime,
    slotMinutes: slotMinutes,
    staffList: staffList,
    services: services,
    days: days,
    generatedAt: new Date().toISOString()
  };
}

// ================================================================
// สร้างการจองใหม่ (มี lock กันจองชนกัน)
// ================================================================
function createBooking(body) {
  var customerName = trimStr(body.customerName);
  var phone = trimStr(body.phone);
  var petName = trimStr(body.petName);
  var petType = trimStr(body.petType);
  var serviceName = trimStr(body.service);
  var date = trimStr(body.date);       // 'YYYY-MM-DD'
  var startTime = trimStr(body.startTime); // 'HH:MM'
  var requestedStaff = trimStr(body.staff); // อาจว่าง = ไม่ระบุ/สุ่มช่างว่าง
  var notes = trimStr(body.notes);

  if (!customerName || !phone || !serviceName || !date || !startTime) {
    return { ok: false, error: 'กรอกข้อมูลไม่ครบ (ชื่อ, เบอร์โทร, บริการ, วันที่, เวลา)' };
  }

  var services = getServices();
  var service = services.filter(function (s) { return s.name === serviceName; })[0];
  if (!service) return { ok: false, error: 'ไม่พบบริการที่เลือก' };

  var staffList = getStaffList();
  var candidates = requestedStaff ? [requestedStaff] : staffList;
  if (requestedStaff && staffList.indexOf(requestedStaff) === -1) {
    return { ok: false, error: 'ไม่พบช่างที่เลือก' };
  }

  var settings = getSettings();
  var openMinutes = timeToMinutes(settings.openTime);
  var closeMinutes = timeToMinutes(settings.closeTime);
  var startMin = timeToMinutes(startTime);
  var endMin = startMin + service.durationMinutes;
  if (startMin < openMinutes || endMin > closeMinutes) {
    return { ok: false, error: 'ช่วงเวลานี้อยู่นอกเวลาทำการของร้าน (' + settings.openTime + '–' + settings.closeTime + ')' };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.bookings);
    var existing = readAllBookings(sh);

    var chosenStaff = null;
    for (var i = 0; i < candidates.length; i++) {
      var name = candidates[i];
      var conflict = existing.some(function (b) {
        if (b.staff !== name) return false;
        if (b.status !== STATUS_ACTIVE && b.status !== STATUS_BLOCKED) return false;
        if (b.date !== date) return false;
        var bStart = timeToMinutes(b.startTime), bEnd = timeToMinutes(b.endTime);
        return startMin < bEnd && endMin > bStart;
      });
      if (!conflict) { chosenStaff = name; break; }
    }

    if (!chosenStaff) {
      return { ok: false, error: requestedStaff ? 'ช่างคนนี้ไม่ว่างในช่วงเวลาที่เลือกแล้ว กรุณาเลือกเวลาอื่น' : 'ไม่มีช่างว่างในช่วงเวลาที่เลือกแล้ว กรุณาเลือกเวลาอื่น' };
    }

    var bookingId = 'CP' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyMMdd') + '-' + Utilities.getUuid().slice(0, 4).toUpperCase();

    sh.appendRow([
      bookingId,
      new Date(),
      date,
      startTime,
      minutesToTime(endMin),
      chosenStaff,
      STATUS_ACTIVE,
      customerName,
      phone,
      petName,
      petType,
      serviceName,
      notes
    ]);

    return {
      ok: true,
      bookingId: bookingId,
      staff: chosenStaff,
      date: date,
      startTime: startTime,
      endTime: minutesToTime(endMin),
      service: serviceName
    };
  } finally {
    lock.releaseLock();
  }
}

function cancelBooking(body) {
  var bookingId = trimStr(body.bookingId);
  var phone = trimStr(body.phone);
  if (!bookingId) return { ok: false, error: 'ไม่พบเลขที่การจอง' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.bookings);
    var idx = headerIndexMap();
    var numRows = sh.getLastRow() - 1;
    if (numRows <= 0) return { ok: false, error: 'ไม่พบการจอง' };
    var data = sh.getRange(2, 1, numRows, BOOKING_HEADERS.length).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][idx.BookingID] === bookingId) {
        if (phone && String(data[i][idx.Phone]) !== phone) {
          return { ok: false, error: 'เบอร์โทรไม่ตรงกับการจองนี้' };
        }
        sh.getRange(i + 2, idx.Status + 1).setValue(STATUS_CANCELLED);
        return { ok: true, bookingId: bookingId };
      }
    }
    return { ok: false, error: 'ไม่พบการจองนี้ หรือถูกยกเลิกไปแล้ว' };
  } finally {
    lock.releaseLock();
  }
}

function readAllBookings(sh) {
  var idx = headerIndexMap();
  var numRows = sh.getLastRow() - 1;
  if (numRows <= 0) return [];
  var data = sh.getRange(2, 1, numRows, BOOKING_HEADERS.length).getValues();
  return data.map(function (row) {
    var d = row[idx.Date];
    var dateObj = (d instanceof Date) ? d : new Date(d);
    return {
      bookingId: row[idx.BookingID],
      date: formatDateYMD(dateObj),
      startTime: normalizeTime(row[idx.StartTime]),
      endTime: normalizeTime(row[idx.EndTime]),
      staff: row[idx.Staff],
      status: row[idx.Status],
      phone: row[idx.Phone]
    };
  });
}

// ================================================================
// Helpers
// ================================================================
function trimStr(v) { return (v === undefined || v === null) ? '' : String(v).trim(); }

function timeToMinutes(t) {
  var parts = String(t).split(':');
  return Number(parts[0]) * 60 + Number(parts[1] || 0);
}

function minutesToTime(m) {
  var h = Math.floor(m / 60), mm = m % 60;
  return (h < 10 ? '0' + h : h) + ':' + (mm < 10 ? '0' + mm : mm);
}

function normalizeTime(v) {
  if (v instanceof Date) {
    return (v.getHours() < 10 ? '0' + v.getHours() : v.getHours()) + ':' + (v.getMinutes() < 10 ? '0' + v.getMinutes() : v.getMinutes());
  }
  return String(v);
}

function formatDateYMD(date) {
  var y = date.getFullYear(), m = date.getMonth() + 1, d = date.getDate();
  return y + '-' + (m < 10 ? '0' + m : m) + '-' + (d < 10 ? '0' + d : d);
}
