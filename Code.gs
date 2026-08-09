/**
 * Champion Petshop — ระบบจองคิวอาบน้ำ/ตัดขนสัตว์เลี้ยง
 * Google Apps Script Web App (backend API) ทำงานคู่กับ Google Sheet
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
  lineOaId: '@413utlzb',
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
    if (action === 'lookup') {
      return jsonOut(lookupBooking(e.parameter.bookingId, e.parameter.phone));
    }
    if (action === 'staff') {
      return renderStaffPage();
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
  if (map.OpenTime) out.openTime = normalizeTimeSetting(map.OpenTime, DEFAULTS.openTime);
  if (map.CloseTime) out.closeTime = normalizeTimeSetting(map.CloseTime, DEFAULTS.closeTime);
  if (map.SlotMinutes) out.slotMinutes = Number(map.SlotMinutes);
  if (map.DaysAhead) out.daysAhead = Number(map.DaysAhead);
  return out;
}

// แปลงค่าเวลาจาก Settings ให้เป็นสตริง 'HH:MM' เสมอ: ถ้า Sheets แปลงเซลล์เป็นเวลา/วันที่อัตโนมัติ (Date object)
// จะดึง HH:MM ออกมาแทนที่จะเอามาตีความตรงๆ ถ้าแปลงไม่ได้ จะใช้ค่า default แทน เพื่อกันการคำนวณช่วงเวลาพังเกิด NaN
function normalizeTimeSetting(v, fallback) {
  if (v && typeof v.getHours === 'function' && typeof v.getMinutes === 'function') {
    var h = v.getHours(), m = v.getMinutes();
    return (h < 10 ? '0' + h : h) + ':' + (m < 10 ? '0' + m : m);
  }
  var s = String(v).trim();
  if (/^\d{1,2}:\d{2}$/.test(s)) return s;
  return fallback;
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
    var dateObj = toDateObj(d);
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
    var dateObj = toDateObj(d);
    return {
      bookingId: row[idx.BookingID],
      timestamp: row[idx.Timestamp] ? String(toDateObj(row[idx.Timestamp])) : '',
      date: formatDateYMD(dateObj),
      startTime: normalizeTime(row[idx.StartTime]),
      endTime: normalizeTime(row[idx.EndTime]),
      staff: row[idx.Staff],
      status: row[idx.Status],
      customerName: row[idx.CustomerName],
      phone: row[idx.Phone],
      petName: row[idx.PetName],
      petType: row[idx.PetType],
      service: row[idx.Service],
      notes: row[idx.Notes]
    };
  });
}

// ================================================================
// ค้นหาการจองเดียว (สำหรับลูกค้าเช็ก/ยกเลิก/เปลี่ยนวัน-เวลาเอง) — ต้องรู้ทั้งเลขที่จองและเบอร์โทร
// ================================================================
function lookupBooking(bookingId, phone) {
  bookingId = trimStr(bookingId);
  phone = trimStr(phone);
  if (!bookingId || !phone) return { ok: false, error: 'กรอกเลขที่การจองและเบอร์โทรให้ครบ' };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.bookings);
  if (!sh) return { ok: false, error: 'ไม่พบการจอง' };
  var all = readAllBookings(sh);
  var found = all.filter(function (b) {
    return String(b.bookingId) === bookingId && String(b.phone) === phone;
  })[0];
  if (!found) return { ok: false, error: 'ไม่พบการจอง หรือเบอร์โทรไม่ตรงกับการจองนี้' };
  return { ok: true, booking: found };
}

// ================================================================
// หน้าเจ้าหน้าที่ (Staff dashboard) — ดู/ยกเลิกการจองทั้งหมดโดยไม่ต้องเปิด Google Sheet
// เข้าถึงได้เฉพาะอีเมล Google ที่อยู่ใน Settings!StaffEmails เท่านั้น (เช็คผ่าน Session.getActiveUser())
// หมายเหตุ: ค่านี้จะมีผลจริงเฉพาะเมื่อเปิดผ่าน deployment ที่ตั้ง Access เป็น
// "Anyone with Google account" เท่านั้น — deployment แบบ "Anyone" (ที่ลูกค้าใช้จองคิว)
// จะไม่มีอีเมลผู้ใช้ให้เช็ค (Session.getActiveUser().getEmail() จะว่างเปล่า) จึงเข้าไม่ได้เสมอ
// ================================================================
function getStaffEmails() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.settings);
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var found = rows.filter(function (r) { return String(r[0]).trim() === 'StaffEmails'; })[0];
  if (!found || !found[1]) return [];
  return String(found[1]).split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(Boolean);
}

function getActiveUserEmail() {
  try {
    return (Session.getActiveUser().getEmail() || '').toLowerCase();
  } catch (e) {
    return '';
  }
}

function isAuthorizedStaff() {
  var email = getActiveUserEmail();
  if (!email) return false;
  return getStaffEmails().indexOf(email) !== -1;
}

function staffGetBookings() {
  if (!isAuthorizedStaff()) {
    throw new Error('ไม่มีสิทธิ์เข้าถึง (' + (getActiveUserEmail() || 'ไม่พบอีเมลผู้ใช้') + ')');
  }
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.bookings);
  var list = readAllBookings(sh);
  list.sort(function (a, b) {
    var ka = a.date + ' ' + a.startTime, kb = b.date + ' ' + b.startTime;
    return ka < kb ? -1 : ka > kb ? 1 : 0;
  });
  return list;
}

function staffCancelBooking(bookingId) {
  if (!isAuthorizedStaff()) {
    throw new Error('ไม่มีสิทธิ์เข้าถึง');
  }
  bookingId = trimStr(bookingId);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.bookings);
    var idx = headerIndexMap();
    var numRows = sh.getLastRow() - 1;
    if (numRows <= 0) return { ok: false, error: 'ไม่พบการจอง' };
    var data = sh.getRange(2, 1, numRows, BOOKING_HEADERS.length).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][idx.BookingID]) === bookingId) {
        sh.getRange(i + 2, idx.Status + 1).setValue(STATUS_CANCELLED);
        return { ok: true, bookingId: bookingId };
      }
    }
    return { ok: false, error: 'ไม่พบการจองนี้' };
  } finally {
    lock.releaseLock();
  }
}

function escapeHtmlGs(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function renderStaffPage() {
  var email = getActiveUserEmail();
  var authorized = isAuthorizedStaff();
  return HtmlService.createHtmlOutput(buildStaffHtml(email, authorized))
    .setTitle('Champion Petshop — หน้าเจ้าหน้าที่')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function buildStaffHtml(email, authorized) {
  if (!authorized) {
    return '<!DOCTYPE html><html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"></head>' +
      '<body style="font-family:sans-serif;padding:48px 24px;text-align:center;background:#F5F4F2;color:#2C2A27;">' +
      '<h2 style="margin-bottom:12px;">ไม่มีสิทธิ์เข้าถึงหน้านี้</h2>' +
      '<p style="color:#5A5650;">บัญชีที่เข้าสู่ระบบ: <b>' + escapeHtmlGs(email || 'ไม่พบอีเมล (กรุณาเข้าสู่ระบบด้วยบัญชี Google)') + '</b></p>' +
      '<p style="color:#5A5650;">กรุณาแจ้งผู้ดูแลระบบให้เพิ่มอีเมลนี้ในชีต Settings แถว <code>StaffEmails</code></p>' +
      '</body></html>';
  }
  return '<!DOCTYPE html><html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>Champion Petshop — หน้าเจ้าหน้าที่</title>' +
    '<style>' +
    'body{font-family:-apple-system,"Segoe UI",Kanit,sans-serif;margin:0;background:#F5F4F2;color:#2C2A27;}' +
    '.wrap{max-width:1100px;margin:0 auto;padding:20px 16px 60px;}' +
    'h1{font-size:20px;margin:0 0 4px;}' +
    '.sub{color:#5A5650;font-size:13px;margin-bottom:16px;}' +
    '.toolbar{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px;align-items:center;}' +
    'input[type=text]{padding:9px 12px;border:1.5px solid #D9D7D2;border-radius:10px;font-size:14px;min-width:220px;}' +
    'button{cursor:pointer;border:none;border-radius:999px;padding:9px 16px;font-weight:700;font-size:13.5px;background:#3E5A47;color:#fff;}' +
    'button.ghost{background:#fff;color:#3E5A47;border:1.5px solid #D9D7D2;}' +
    'button.danger{background:#C0553F;}' +
    'button:disabled{opacity:.5;cursor:not-allowed;}' +
    'table{width:100%;border-collapse:collapse;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 6px rgba(27,38,30,.07);}' +
    'th,td{padding:10px 12px;text-align:left;font-size:13px;border-bottom:1px solid #EFEEEC;white-space:nowrap;}' +
    'th{background:#EEF3EF;color:#403D38;font-weight:700;position:sticky;top:0;}' +
    'tr:last-child td{border-bottom:none;}' +
    '.status-active{color:#3E7A52;font-weight:700;}' +
    '.status-cancelled{color:#9A948B;}' +
    '.status-blocked{color:#B07C1E;font-weight:700;}' +
    '.tablewrap{overflow:auto;max-height:70vh;border-radius:12px;}' +
    '.empty{padding:40px;text-align:center;color:#5A5650;}' +
    '.me{font-size:12px;color:#5A5650;}' +
    '</style></head><body><div class="wrap">' +
    '<h1>Champion Petshop — หน้าเจ้าหน้าที่</h1>' +
    '<div class="sub">รายการจองทั้งหมด (ไม่ต้องเปิด Google Sheet) · เข้าสู่ระบบเป็น <span class="me">' + escapeHtmlGs(email) + '</span></div>' +
    '<div class="toolbar">' +
    '<input type="text" id="q" placeholder="ค้นหา: ชื่อ / เบอร์โทร / เลขที่การจอง">' +
    '<button class="ghost" id="btn-refresh">รีเฟรช</button>' +
    '<span id="count" style="font-size:12.5px;color:#5A5650;"></span>' +
    '</div>' +
    '<div id="tablewrap" class="tablewrap"><div class="empty">กำลังโหลด...</div></div>' +
    '</div>' +
    '<script>' +
    'var ALL = [];' +
    'function load(){ google.script.run.withSuccessHandler(onData).withFailureHandler(onError).staffGetBookings(); }' +
    'function onError(e){ document.getElementById("tablewrap").innerHTML = "<div class=empty>โหลดไม่สำเร็จ: " + (e && e.message ? e.message : e) + "</div>"; }' +
    'function onData(list){ ALL = list || []; renderTable(); }' +
    'function statusClass(s){ if (s === "จอง") return "status-active"; if (s === "ปิด") return "status-blocked"; return "status-cancelled"; }' +
    'function esc(s){ return String(s == null ? "" : s).replace(/[&<>"\x27]/g, function(c){ return {"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","\x27":"&#39;"}[c]; }); }' +
    'function renderTable(){' +
    '  var q = (document.getElementById("q").value || "").trim().toLowerCase();' +
    '  var rows = ALL.filter(function(b){ if (!q) return true; return [b.bookingId,b.customerName,b.phone,b.petName,b.staff].join(" ").toLowerCase().indexOf(q) !== -1; });' +
    '  document.getElementById("count").textContent = rows.length + " / " + ALL.length + " รายการ";' +
    '  if (!rows.length){ document.getElementById("tablewrap").innerHTML = "<div class=empty>ไม่พบรายการ</div>"; return; }' +
    '  var html = "<table><thead><tr><th>วันที่</th><th>เวลา</th><th>ช่าง</th><th>สถานะ</th><th>ลูกค้า</th><th>เบอร์โทร</th><th>สัตว์เลี้ยง</th><th>บริการ</th><th>เลขที่จอง</th><th>หมายเหตุ</th><th></th></tr></thead><tbody>";' +
    '  rows.forEach(function(b){' +
    '    html += "<tr><td>" + esc(b.date) + "</td><td>" + esc(b.startTime) + "–" + esc(b.endTime) + "</td><td>" + esc(b.staff) + "</td>" +' +
    '      "<td class=" + statusClass(b.status) + ">" + esc(b.status) + "</td><td>" + esc(b.customerName) + "</td><td>" + esc(b.phone) + "</td>" +' +
    '      "<td>" + esc(b.petName) + (b.petType ? " (" + esc(b.petType) + ")" : "") + "</td><td>" + esc(b.service) + "</td><td>" + esc(b.bookingId) + "</td>" +' +
    '      "<td>" + esc(b.notes) + "</td><td>" + (b.status === "จอง" ? "<button class=danger data-id=\"" + esc(b.bookingId) + "\">ยกเลิก</button>" : "") + "</td></tr>";' +
    '  });' +
    '  html += "</tbody></table>";' +
    '  document.getElementById("tablewrap").innerHTML = html;' +
    '  document.querySelectorAll("button[data-id]").forEach(function(btn){' +
    '    btn.onclick = function(){' +
    '      if (!confirm("ยืนยันยกเลิกการจอง " + btn.getAttribute("data-id") + " ?")) return;' +
    '      btn.disabled = true; btn.textContent = "กำลังยกเลิก...";' +
    '      google.script.run.withSuccessHandler(function(res){' +
    '        if (res && res.ok) { load(); } else { alert((res && res.error) || "ยกเลิกไม่สำเร็จ"); btn.disabled = false; btn.textContent = "ยกเลิก"; }' +
    '      }).withFailureHandler(function(e){ alert(e && e.message ? e.message : e); btn.disabled = false; btn.textContent = "ยกเลิก"; })' +
    '        .staffCancelBooking(btn.getAttribute("data-id"));' +
    '    };' +
    '  });' +
    '}' +
    'window.addEventListener("DOMContentLoaded", function(){' +
    '  document.getElementById("btn-refresh").onclick = load;' +
    '  document.getElementById("q").oninput = renderTable;' +
    '  load();' +
    '});' +
    '</script></body></html>';
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

// แปลงค่าจาก Sheet ให้เป็น Date object เสมอ: ใช้ duck-typing (มีเมธอด getFullYear ไหม)
// แทน instanceof Date ด้วยเหตุผลเดียวกับ normalizeTime — ถ้าไม่ใช่ ให้ลองสร้าง new Date(v) ใหม่
function toDateObj(v) {
  if (v && typeof v.getFullYear === 'function') return v;
  return new Date(v);
}

function normalizeTime(v) {
  // ใช้ duck-typing (เช็คว่ามีเมธอด getHours/getMinutes) แทน instanceof Date
  // เพราะค่าที่ได้จาก Range#getValues() ของ Apps Script บางครั้งเป็นอ็อบเจกต์แบบวันที่
  // แต่ไม่ผ่าน instanceof Date (มาจากคนละ realm) ทำให้ String(v) ออกมาเป็น
  // "Sat Dec 30 1899 10:00:00 GMT+xxxx" แทนที่จะเป็น "10:00" และพังตอนคำนวณช่วงเวลา
  if (v && typeof v.getHours === 'function' && typeof v.getMinutes === 'function') {
    var h = v.getHours(), m = v.getMinutes();
    return (h < 10 ? '0' + h : h) + ':' + (m < 10 ? '0' + m : m);
  }
  var s = String(v).trim();
  if (/^\d{1,2}:\d{2}$/.test(s)) return s;
  return s;
}

function formatDateYMD(date) {
  var y = date.getFullYear(), m = date.getMonth() + 1, d = date.getDate();
  return y + '-' + (m < 10 ? '0' + m : m) + '-' + (d < 10 ? '0' + d : d);
}
