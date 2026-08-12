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
  // บังคับคอลัมน์เบอร์โทรทั้งคอลัมน์ให้เป็นข้อความ กันเลข 0 หน้าเบอร์หายเวลา Sheets ตีความเป็นตัวเลข
  bookings.getRange(2, BOOKING_HEADERS.indexOf('Phone') + 1, 998, 1).setNumberFormat('@');

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

    // ปิดการจอง/ยกเลิกด้วยตนเองของลูกค้าผ่าน API สาธารณะแล้ว — ให้เจ้าหน้าที่ทำผ่านหน้าเจ้าหน้าที่ (staff dashboard) แทน
    if (action === 'book' || action === 'cancel') {
      return jsonOut({ ok: false, error: 'ปิดการจอง/ยกเลิกด้วยตนเองแล้ว กรุณาติดต่อร้านผ่าน LINE' });
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
var SCHEDULE_CACHE_KEY = 'schedule_v1';
// เก็บ cache ไว้ 6 นาที ให้ยาวกว่ารอบ keepWarm (ทุก 5 นาที) เล็กน้อย จะได้มีของพร้อมเสิร์ฟตลอดเวลา
// ความสดของข้อมูลไม่เสีย เพราะล้าง cache ทันทีทั้งเมื่อจอง/ยกเลิก/แก้ไขผ่านระบบ และเมื่อมีคนแก้ Google Sheet เอง
var SCHEDULE_CACHE_SECONDS = 360;

// ล้าง cache ตารางคิว — เรียกทุกครั้งที่มีการจอง/ยกเลิก/แก้ไขคิว เพื่อให้ลูกค้าเห็นข้อมูลล่าสุดทันที
function invalidateScheduleCache() {
  try { CacheService.getScriptCache().remove(SCHEDULE_CACHE_KEY); } catch (e) {}
}

// ==================================================================
// กันไม่ให้ลูกค้าเจอ "cold start" (Google ปิดเครื่องทิ้งเมื่อไม่มีคนใช้ แล้วคนแรกต้องรอบูตใหม่ 10-20 วินาที)
// ตั้งเป็น trigger รายเวลา ทุก 5 นาที ผ่านฟังก์ชัน installPerformanceTriggers() ด้านล่าง
// ทำ 2 อย่างพร้อมกัน: (1) ให้สคริปต์ถูกเรียกใช้สม่ำเสมอจนไม่ถูกปิดทิ้ง (2) เตรียมข้อมูลใส่ cache ไว้ล่วงหน้า
// ==================================================================
function keepWarm() {
  invalidateScheduleCache();   // บังคับให้อ่านของจริงใหม่ ไม่ใช่คืนของเก่าใน cache
  buildScheduleResponse();     // สร้างใหม่แล้วเก็บลง cache ให้ลูกค้าคนถัดไปได้ของทันที
}

// ถ้าเจ้าหน้าที่ไปแก้คิวใน Google Sheet โดยตรง (ไม่ผ่านหน้าเว็บ) ให้ล้าง cache ทันทีเช่นกัน
// กันข้อมูลค้าง — ติดตั้งเป็น trigger แบบ on edit ผ่าน installPerformanceTriggers()
function onEditInvalidate(e) {
  invalidateScheduleCache();
}

// รันฟังก์ชันนี้ "ครั้งเดียว" จาก Apps Script editor เพื่อติดตั้ง trigger ทั้งสองตัว
// (รันซ้ำได้ ระบบจะลบ trigger เดิมของสองฟังก์ชันนี้ก่อน ไม่เกิดตัวซ้ำ)
function installPerformanceTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existing = ScriptApp.getProjectTriggers();
  existing.forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'keepWarm' || fn === 'onEditInvalidate') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('keepWarm').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('onEditInvalidate').forSpreadsheet(ss).onEdit().create();

  var names = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  return 'ติดตั้ง trigger เรียบร้อย: ' + names.join(', ');
}

function buildScheduleResponse() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(SCHEDULE_CACHE_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

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

  var result = {
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

  try { cache.put(SCHEDULE_CACHE_KEY, JSON.stringify(result), SCHEDULE_CACHE_SECONDS); } catch (e) {}

  return result;
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

    // บังคับให้เซลล์เบอร์โทรของแถวนี้เป็นข้อความเสมอ กัน Sheets ตัดเลข 0 หน้าเบอร์ทิ้ง (ตีความเป็นตัวเลข)
    var newRow = sh.getLastRow();
    sh.getRange(newRow, headerIndexMap().Phone + 1).setNumberFormat('@').setValue(phone);

    invalidateScheduleCache();
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
        invalidateScheduleCache();
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
        invalidateScheduleCache();
        return { ok: true, bookingId: bookingId };
      }
    }
    return { ok: false, error: 'ไม่พบการจองนี้' };
  } finally {
    lock.releaseLock();
  }
}

// เจ้าหน้าที่จองคิวใหม่แทนลูกค้า (ใช้ฟังก์ชัน createBooking เดิมทุกอย่าง แค่เช็คสิทธิ์ก่อน)
function staffCreateBooking(body) {
  if (!isAuthorizedStaff()) {
    throw new Error('ไม่มีสิทธิ์เข้าถึง');
  }
  return createBooking(body);
}

// เจ้าหน้าที่แก้ไขข้อมูลการจองที่มีอยู่แล้ว (เปลี่ยนวัน/เวลา/ช่าง/ข้อมูลลูกค้าได้ในที่เดียว ไม่ต้องยกเลิกแล้วจองใหม่)
function staffUpdateBooking(bookingId, fields) {
  if (!isAuthorizedStaff()) {
    throw new Error('ไม่มีสิทธิ์เข้าถึง');
  }
  bookingId = trimStr(bookingId);
  fields = fields || {};
  var customerName = trimStr(fields.customerName);
  var phone = trimStr(fields.phone);
  var petName = trimStr(fields.petName);
  var petType = trimStr(fields.petType);
  var serviceName = trimStr(fields.service);
  var date = trimStr(fields.date);
  var startTime = trimStr(fields.startTime);
  var staffName = trimStr(fields.staff);
  var notes = trimStr(fields.notes);

  if (!bookingId) return { ok: false, error: 'ไม่พบเลขที่การจอง' };
  if (!customerName || !phone || !serviceName || !date || !startTime || !staffName) {
    return { ok: false, error: 'กรอกข้อมูลไม่ครบ (ชื่อ, เบอร์โทร, บริการ, วันที่, เวลา, ช่าง)' };
  }

  var services = getServices();
  var service = services.filter(function (s) { return s.name === serviceName; })[0];
  if (!service) return { ok: false, error: 'ไม่พบบริการที่เลือก' };

  var staffList = getStaffList();
  if (staffList.indexOf(staffName) === -1) return { ok: false, error: 'ไม่พบช่างที่เลือก' };

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
    var idx = headerIndexMap();
    var numRows = sh.getLastRow() - 1;
    if (numRows <= 0) return { ok: false, error: 'ไม่พบการจอง' };
    var data = sh.getRange(2, 1, numRows, BOOKING_HEADERS.length).getValues();
    var rowIndex = -1;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][idx.BookingID]) === bookingId) { rowIndex = i; break; }
    }
    if (rowIndex === -1) return { ok: false, error: 'ไม่พบการจองนี้' };

    var conflict = data.some(function (row, i) {
      if (i === rowIndex) return false;
      if (row[idx.Staff] !== staffName) return false;
      var st = row[idx.Status];
      if (st !== STATUS_ACTIVE && st !== STATUS_BLOCKED) return false;
      var rowDate = formatDateYMD(toDateObj(row[idx.Date]));
      if (rowDate !== date) return false;
      var bStart = timeToMinutes(normalizeTime(row[idx.StartTime]));
      var bEnd = timeToMinutes(normalizeTime(row[idx.EndTime]));
      return startMin < bEnd && endMin > bStart;
    });
    if (conflict) return { ok: false, error: 'ช่างคนนี้ไม่ว่างในช่วงเวลาที่เลือก กรุณาเลือกเวลาหรือช่างอื่น' };

    var sheetRow = rowIndex + 2;
    sh.getRange(sheetRow, idx.Date + 1).setValue(date);
    sh.getRange(sheetRow, idx.StartTime + 1).setValue(startTime);
    sh.getRange(sheetRow, idx.EndTime + 1).setValue(minutesToTime(endMin));
    sh.getRange(sheetRow, idx.Staff + 1).setValue(staffName);
    sh.getRange(sheetRow, idx.CustomerName + 1).setValue(customerName);
    sh.getRange(sheetRow, idx.Phone + 1).setNumberFormat('@').setValue(phone);
    sh.getRange(sheetRow, idx.PetName + 1).setValue(petName);
    sh.getRange(sheetRow, idx.PetType + 1).setValue(petType);
    sh.getRange(sheetRow, idx.Service + 1).setValue(serviceName);
    sh.getRange(sheetRow, idx.Notes + 1).setValue(notes);

    invalidateScheduleCache();
    return { ok: true, bookingId: bookingId };
  } finally {
    lock.releaseLock();
  }
}

// ข้อมูลตัวเลือกสำหรับฟอร์มจองคิวของเจ้าหน้าที่ (รายการบริการ/ช่าง/เวลาเปิด-ปิดร้าน)
function staffGetFormOptions() {
  if (!isAuthorizedStaff()) {
    throw new Error('ไม่มีสิทธิ์เข้าถึง');
  }
  var settings = getSettings();
  return {
    services: getServices(),
    staffList: getStaffList(),
    openTime: settings.openTime,
    closeTime: settings.closeTime,
    slotMinutes: settings.slotMinutes
  };
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
    '.modal-overlay{position:fixed;inset:0;background:rgba(28,27,25,.45);display:flex;align-items:center;justify-content:center;z-index:50;padding:20px;}' +
    '.modal{background:#fff;border-radius:16px;max-width:480px;width:100%;max-height:90vh;overflow:auto;padding:22px;}' +
    '.modal h2{font-size:17px;margin:0 0 14px;}' +
    '.modal .field{margin-bottom:12px;}' +
    '.modal label{display:block;font-size:12.5px;font-weight:700;color:#5A5650;margin-bottom:5px;}' +
    '.modal input,.modal select,.modal textarea{width:100%;padding:9px 11px;border:1.5px solid #D9D7D2;border-radius:8px;font-size:14px;box-sizing:border-box;font-family:inherit;}' +
    '.modal textarea{resize:vertical;}' +
    '.modal .row2{display:flex;gap:10px;}' +
    '.modal .row2>.field{flex:1;}' +
    '.modal .hint{font-size:11.5px;color:#9A948B;margin-top:8px;}' +
    '.modal .err{color:#C0553F;font-size:12.5px;margin-bottom:10px;display:none;}' +
    '.modal .actions{display:flex;gap:10px;margin-top:6px;}' +
    '.modal .actions button{flex:1;}' +
    '.gtabs{display:flex;gap:8px;margin-bottom:14px;}' +
    '.gtab-btn{background:#fff;color:#3E5A47;border:1.5px solid #D9D7D2;border-radius:999px;padding:8px 16px;font-weight:700;font-size:13px;cursor:pointer;}' +
    '.gtab-btn.active{background:#3E5A47;color:#fff;border-color:#3E5A47;}' +
    '.daynav{display:flex;align-items:center;gap:12px;margin-bottom:10px;}' +
    '.daynav .lbl{font-weight:700;font-size:14px;flex:1;text-align:center;}' +
    '.daynav button:disabled{opacity:.4;cursor:not-allowed;}' +
    '.glegend{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:10px;font-size:12px;color:#5A5650;}' +
    '.gswatch{width:12px;height:12px;border-radius:3px;display:inline-block;margin-right:5px;vertical-align:-1px;}' +
    '.gridwrap-outer{position:relative;}' +
    '.gridwrap{overflow:auto;max-height:65vh;border-radius:12px;background:#fff;box-shadow:0 2px 6px rgba(27,38,30,.07);}' +
    '.g-grid{display:flex;flex-direction:column;gap:3px;width:fit-content;padding:10px;}' +
    '.g-row{display:flex;gap:3px;}' +
    '.g-corner{width:88px;flex:0 0 88px;}' +
    '.g-colhead{width:46px;flex:0 0 46px;font-size:10.5px;text-align:center;color:#5A5650;font-weight:700;padding-top:4px;}' +
    '.g-rowlabel{width:88px;flex:0 0 88px;font-size:12.5px;font-weight:700;display:flex;align-items:center;padding-right:6px;position:sticky;left:0;background:#fff;}' +
    '.g-cell{width:46px;flex:0 0 46px;height:34px;border-radius:6px;border:1.5px solid #D9D7D2;background:#fff;cursor:pointer;font-size:9.5px;padding:0 3px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;color:#fff;font-family:inherit;}' +
    '.g-cell.g-free:hover{background:#EEF3EF;border-color:#3E5A47;}' +
    '.g-cell.g-booked{background:#C0553F;border-color:#C0553F;}' +
    '.g-cell.g-blocked{background:#9A948B;border-color:#9A948B;}' +
    '</style></head><body><div class="wrap">' +
    '<h1>Champion Petshop — หน้าเจ้าหน้าที่</h1>' +
    '<div class="sub">จัดการคิวทั้งหมด (ไม่ต้องเปิด Google Sheet) · เข้าสู่ระบบเป็น <span class="me">' + escapeHtmlGs(email) + '</span></div>' +
    '<div class="gtabs">' +
    '<button type="button" class="gtab-btn active" id="tab-btn-list">รายการทั้งหมด</button>' +
    '<button type="button" class="gtab-btn" id="tab-btn-grid">ตารางคิว 10 วัน</button>' +
    '</div>' +
    '<div id="panel-list">' +
    '<div class="toolbar">' +
    '<input type="text" id="q" placeholder="ค้นหา: ชื่อ / เบอร์โทร / เลขที่การจอง">' +
    '<button class="ghost" id="btn-refresh">รีเฟรช</button>' +
    '<button id="btn-new-booking">+ จองคิวใหม่</button>' +
    '<span id="count" style="font-size:12.5px;color:#5A5650;"></span>' +
    '</div>' +
    '<div id="tablewrap" class="tablewrap"><div class="empty">กำลังโหลด...</div></div>' +
    '</div>' +
    '<div id="panel-grid" style="display:none;">' +
    '<div class="daynav">' +
    '<button type="button" class="ghost" id="btn-day-prev">← วันก่อนหน้า</button>' +
    '<div class="lbl" id="grid-date-label"></div>' +
    '<button type="button" class="ghost" id="btn-day-next">วันถัดไป →</button>' +
    '</div>' +
    '<div class="glegend">' +
    '<span><span class="gswatch" style="background:#fff;border:1.5px solid #D9D7D2;"></span>ว่าง (กดเพื่อจองคิว)</span>' +
    '<span><span class="gswatch" style="background:#C0553F;"></span>มีคิวจอง (กดเพื่อดู/แก้ไข)</span>' +
    '<span><span class="gswatch" style="background:#9A948B;"></span>ปิดคิว</span>' +
    '</div>' +
    '<div class="gridwrap-outer"><div class="gridwrap" id="grid-wrap"><div class="empty">กำลังโหลด...</div></div></div>' +
    '</div>' +
    '</div>' +
    '<div id="booking-modal" class="modal-overlay" style="display:none;">' +
    '<div class="modal">' +
    '<h2 id="modal-title">จองคิวใหม่</h2>' +
    '<div id="modal-err" class="err"></div>' +
    '<div class="field"><label>บริการ</label><select id="m-service"></select></div>' +
    '<div class="row2">' +
    '<div class="field"><label>วันที่</label><input type="date" id="m-date"></div>' +
    '<div class="field"><label>เวลาเริ่ม</label><input type="time" id="m-start"></div>' +
    '</div>' +
    '<div class="field"><label>ช่าง</label><select id="m-staff"></select></div>' +
    '<div class="field"><label>ชื่อลูกค้า</label><input type="text" id="m-name" placeholder="เช่น คุณสมชาย"></div>' +
    '<div class="field"><label>เบอร์โทร</label><input type="text" id="m-phone" inputmode="tel" placeholder="08x-xxx-xxxx"></div>' +
    '<div class="row2">' +
    '<div class="field"><label>ชื่อสัตว์เลี้ยง</label><input type="text" id="m-petname"></div>' +
    '<div class="field"><label>ประเภทสัตว์เลี้ยง</label><input type="text" id="m-pettype"></div>' +
    '</div>' +
    '<div class="field"><label>หมายเหตุ</label><textarea id="m-notes" rows="2"></textarea></div>' +
    '<div id="modal-hint" class="hint"></div>' +
    '<div class="actions">' +
    '<button class="ghost" id="btn-modal-cancel">ยกเลิก</button>' +
    '<button id="btn-modal-submit">บันทึกการจอง</button>' +
    '</div>' +
    '</div>' +
    '</div>' +
    '<div id="detail-modal" class="modal-overlay" style="display:none;">' +
    '<div class="modal">' +
    '<h2>รายละเอียดคิว</h2>' +
    '<div id="detail-err" class="err"></div>' +
    '<div id="detail-blocked-view" style="display:none;">' +
    '<p style="color:#5A5650;font-size:13.5px;">ช่วงเวลานี้ถูกปิดไว้ (ไม่ใช่การจองของลูกค้า)</p>' +
    '<div class="actions">' +
    '<button type="button" class="ghost" id="btn-detail-close2">ปิดหน้าต่าง</button>' +
    '<button type="button" class="danger" id="btn-detail-unblock">เปิดคิวนี้อีกครั้ง</button>' +
    '</div>' +
    '</div>' +
    '<div id="detail-edit-view">' +
    '<div class="field"><label>เลขที่การจอง</label><input type="text" id="d-bookingid" disabled></div>' +
    '<div class="field"><label>บริการ</label><select id="d-service"></select></div>' +
    '<div class="row2">' +
    '<div class="field"><label>วันที่</label><input type="date" id="d-date"></div>' +
    '<div class="field"><label>เวลาเริ่ม</label><input type="time" id="d-start"></div>' +
    '</div>' +
    '<div class="field"><label>ช่าง</label><select id="d-staff"></select></div>' +
    '<div class="field"><label>ชื่อลูกค้า</label><input type="text" id="d-name"></div>' +
    '<div class="field"><label>เบอร์โทร</label><input type="text" id="d-phone" inputmode="tel"></div>' +
    '<div class="row2">' +
    '<div class="field"><label>ชื่อสัตว์เลี้ยง</label><input type="text" id="d-petname"></div>' +
    '<div class="field"><label>ประเภทสัตว์เลี้ยง</label><input type="text" id="d-pettype"></div>' +
    '</div>' +
    '<div class="field"><label>หมายเหตุ</label><textarea id="d-notes" rows="2"></textarea></div>' +
    '<div class="actions">' +
    '<button type="button" class="ghost" id="btn-detail-close">ปิด</button>' +
    '<button type="button" class="danger" id="btn-detail-cancel">ยกเลิกคิวนี้</button>' +
    '<button type="button" id="btn-detail-save">บันทึกการแก้ไข</button>' +
    '</div>' +
    '</div>' +
    '</div>' +
    '</div>' +
    '<script>' +
    'var ALL = [];' +
    'var FORM = { services: [], staffList: [], openTime: "", closeTime: "", slotMinutes: 30 };' +
    'var gridDayIndex = 0;' +
    'var GRID_DAYS = 10;' +
    'var currentDetailId = null;' +
    'function load(){ google.script.run.withSuccessHandler(onData).withFailureHandler(onError).staffGetBookings(); }' +
    'function onError(e){ document.getElementById("tablewrap").innerHTML = "<div class=empty>โหลดไม่สำเร็จ: " + (e && e.message ? e.message : e) + "</div>"; }' +
    'function onData(list){ ALL = list || []; renderTable(); if (document.getElementById("panel-grid").style.display !== "none") renderGrid(); }' +
    'function statusClass(s){ if (s === "จอง") return "status-active"; if (s === "ปิด") return "status-blocked"; return "status-cancelled"; }' +
    'function esc(s){ return String(s == null ? "" : s).replace(/[&<>"\\x27]/g, function(c){ return {"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;","\\x27":"&#39;"}[c]; }); }' +
    'function timeToMin(t){ var p = String(t).split(":"); return Number(p[0]) * 60 + Number(p[1] || 0); }' +
    'function minToTime(m){ var h = Math.floor(m / 60), mm = m % 60; return (h < 10 ? "0" + h : h) + ":" + (mm < 10 ? "0" + mm : mm); }' +
    'function todayLocal(){ var d = new Date(); d.setHours(0,0,0,0); return d; }' +
    'function ymdLocal(d){ var m = d.getMonth() + 1, day = d.getDate(); return d.getFullYear() + "-" + (m < 10 ? "0" + m : m) + "-" + (day < 10 ? "0" + day : day); }' +
    'function dateStrForIndex(idx){ var d = todayLocal(); d.setDate(d.getDate() + idx); return ymdLocal(d); }' +
    'function parseYMDLocal(s){ var p = s.split("-"); return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])); }' +
    'function labelForDate(dateStr, idx){' +
    '  var d = parseYMDLocal(dateStr);' +
    '  var weekdays = ["อาทิตย์","จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์","เสาร์"];' +
    '  var months = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];' +
    '  var txt = "วัน" + weekdays[d.getDay()] + " " + d.getDate() + " " + months[d.getMonth()] + " " + (d.getFullYear() + 543);' +
    '  if (idx === 0) txt += " (วันนี้)";' +
    '  return txt;' +
    '}' +
    'function renderTable(){' +
    '  var q = (document.getElementById("q").value || "").trim().toLowerCase();' +
    '  var rows = ALL.filter(function(b){ if (!q) return true; return [b.bookingId,b.customerName,b.phone,b.petName,b.staff].join(" ").toLowerCase().indexOf(q) !== -1; });' +
    '  document.getElementById("count").textContent = rows.length + " / " + ALL.length + " รายการ";' +
    '  if (!rows.length){ document.getElementById("tablewrap").innerHTML = "<div class=empty>ไม่พบรายการ</div>"; return; }' +
    '  var html = "<table><thead><tr><th>วันที่</th><th>เวลา</th><th>ช่าง</th><th>สถานะ</th><th>ลูกค้า</th><th>เบอร์โทร</th><th>สัตว์เลี้ยง</th><th>บริการ</th><th>เลขที่จอง</th><th>หมายเหตุ</th><th></th></tr></thead><tbody>";' +
    '  rows.forEach(function(b){' +
    '    html += "<tr><td>" + esc(b.date) + "</td><td>" + esc(b.startTime) + "\u2013" + esc(b.endTime) + "</td><td>" + esc(b.staff) + "</td>" +' +
    '      "<td class=" + statusClass(b.status) + ">" + esc(b.status) + "</td><td>" + esc(b.customerName) + "</td><td>" + esc(b.phone) + "</td>" +' +
    '      "<td>" + esc(b.petName) + (b.petType ? " (" + esc(b.petType) + ")" : "") + "</td><td>" + esc(b.service) + "</td><td>" + esc(b.bookingId) + "</td>" +' +
    '      "<td>" + esc(b.notes) + "</td><td>" + (b.status === "จอง" ? ("<button class=ghost data-reschedule=\\"" + esc(b.bookingId) + "\\">เปลี่ยน</button> <button class=danger data-id=\\"" + esc(b.bookingId) + "\\">ยกเลิก</button>") : "") + "</td></tr>";' +
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
    '  document.querySelectorAll("button[data-reschedule]").forEach(function(btn){' +
    '    btn.onclick = function(){' +
    '      var id = btn.getAttribute("data-reschedule");' +
    '      var rec = ALL.filter(function(b){ return String(b.bookingId) === id; })[0];' +
    '      if (!rec) return;' +
    '      if (!confirm("จะยกเลิกคิวเดิมของ " + rec.customerName + " แล้วเปิดฟอร์มจองใหม่ให้ ยืนยันหรือไม่?")) return;' +
    '      btn.disabled = true; btn.textContent = "กำลังยกเลิก...";' +
    '      google.script.run.withSuccessHandler(function(res){' +
    '        if (res && res.ok) { load(); openNewBookingModal(rec); } else { alert((res && res.error) || "ยกเลิกไม่สำเร็จ"); btn.disabled = false; btn.textContent = "เปลี่ยน"; }' +
    '      }).withFailureHandler(function(e){ alert(e && e.message ? e.message : e); btn.disabled = false; btn.textContent = "เปลี่ยน"; })' +
    '        .staffCancelBooking(id);' +
    '    };' +
    '  });' +
    '}' +
    'function loadFormOptions(){ google.script.run.withSuccessHandler(onFormOptions).withFailureHandler(function(){}).staffGetFormOptions(); }' +
    'function onFormOptions(opts){' +
    '  FORM = opts || FORM;' +
    '  var svcSel = document.getElementById("m-service");' +
    '  svcSel.innerHTML = FORM.services.map(function(s){ return "<option value=\\"" + esc(s.name) + "\\">" + esc(s.name) + " (" + s.durationMinutes + " นาที, ฿" + s.price + ")</option>"; }).join("");' +
    '  var staffSel = document.getElementById("m-staff");' +
    '  staffSel.innerHTML = "<option value=\\"\\">ไม่ระบุ (ระบบเลือกช่างว่างให้)</option>" + FORM.staffList.map(function(n){ return "<option value=\\"" + esc(n) + "\\">" + esc(n) + "</option>"; }).join("");' +
    '  document.getElementById("modal-hint").textContent = "เวลาทำการ " + FORM.openTime + "–" + FORM.closeTime + " น. (ทีละ " + FORM.slotMinutes + " นาที)";' +
    '  var dSvcSel = document.getElementById("d-service");' +
    '  dSvcSel.innerHTML = FORM.services.map(function(s){ return "<option value=\\"" + esc(s.name) + "\\">" + esc(s.name) + " (" + s.durationMinutes + " นาที, ฿" + s.price + ")</option>"; }).join("");' +
    '  var dStaffSel = document.getElementById("d-staff");' +
    '  dStaffSel.innerHTML = FORM.staffList.map(function(n){ return "<option value=\\"" + esc(n) + "\\">" + esc(n) + "</option>"; }).join("");' +
    '  if (document.getElementById("panel-grid").style.display !== "none") renderGrid();' +
    '}' +
    'function openNewBookingModal(prefill){' +
    '  document.getElementById("modal-title").textContent = (prefill && prefill.customerName) ? "จองคิวใหม่ (เปลี่ยนวัน-เวลา)" : "จองคิวใหม่";' +
    '  document.getElementById("modal-err").style.display = "none";' +
    '  document.getElementById("m-date").value = (prefill && prefill.date) ? prefill.date : "";' +
    '  document.getElementById("m-start").value = (prefill && prefill.startTime) ? prefill.startTime : "";' +
    '  document.getElementById("m-service").value = (prefill && prefill.service) ? prefill.service : "";' +
    '  document.getElementById("m-staff").value = (prefill && prefill.staff) ? prefill.staff : "";' +
    '  document.getElementById("m-name").value = (prefill && prefill.customerName) ? prefill.customerName : "";' +
    '  document.getElementById("m-phone").value = (prefill && prefill.phone) ? prefill.phone : "";' +
    '  document.getElementById("m-petname").value = (prefill && prefill.petName) ? prefill.petName : "";' +
    '  document.getElementById("m-pettype").value = (prefill && prefill.petType) ? prefill.petType : "";' +
    '  document.getElementById("m-notes").value = (prefill && prefill.notes) ? prefill.notes : "";' +
    '  document.getElementById("booking-modal").style.display = "flex";' +
    '}' +
    'function closeModal(){ document.getElementById("booking-modal").style.display = "none"; }' +
    'function submitNewBooking(){' +
    '  var payload = {' +
    '    action: "book",' +
    '    service: document.getElementById("m-service").value,' +
    '    date: document.getElementById("m-date").value,' +
    '    startTime: document.getElementById("m-start").value,' +
    '    staff: document.getElementById("m-staff").value,' +
    '    customerName: document.getElementById("m-name").value.trim(),' +
    '    phone: document.getElementById("m-phone").value.trim(),' +
    '    petName: document.getElementById("m-petname").value.trim(),' +
    '    petType: document.getElementById("m-pettype").value.trim(),' +
    '    notes: document.getElementById("m-notes").value.trim()' +
    '  };' +
    '  var errEl = document.getElementById("modal-err");' +
    '  if (!payload.service || !payload.date || !payload.startTime || !payload.customerName || !payload.phone) {' +
    '    errEl.textContent = "กรอกข้อมูลให้ครบ: บริการ, วันที่, เวลา, ชื่อลูกค้า, เบอร์โทร"; errEl.style.display = "block"; return;' +
    '  }' +
    '  var btn = document.getElementById("btn-modal-submit");' +
    '  btn.disabled = true; btn.textContent = "กำลังบันทึก...";' +
    '  google.script.run.withSuccessHandler(function(res){' +
    '    btn.disabled = false; btn.textContent = "บันทึกการจอง";' +
    '    if (res && res.ok) { closeModal(); load(); }' +
    '    else { errEl.textContent = (res && res.error) || "จองไม่สำเร็จ"; errEl.style.display = "block"; }' +
    '  }).withFailureHandler(function(e){' +
    '    btn.disabled = false; btn.textContent = "บันทึกการจอง";' +
    '    errEl.textContent = e && e.message ? e.message : String(e); errEl.style.display = "block";' +
    '  }).staffCreateBooking(payload);' +
    '}' +
    'function switchTab(tab){' +
    '  var isList = tab === "list";' +
    '  document.getElementById("panel-list").style.display = isList ? "" : "none";' +
    '  document.getElementById("panel-grid").style.display = isList ? "none" : "";' +
    '  document.getElementById("tab-btn-list").className = "gtab-btn" + (isList ? " active" : "");' +
    '  document.getElementById("tab-btn-grid").className = "gtab-btn" + (isList ? "" : " active");' +
    '  if (!isList) renderGrid();' +
    '}' +
    'function renderGrid(){' +
    '  var wrap = document.getElementById("grid-wrap");' +
    '  if (!FORM.staffList || !FORM.staffList.length) { wrap.innerHTML = "<div class=empty>ยังโหลดข้อมูลไม่เสร็จ...</div>"; return; }' +
    '  var dateStr = dateStrForIndex(gridDayIndex);' +
    '  document.getElementById("grid-date-label").textContent = labelForDate(dateStr, gridDayIndex);' +
    '  document.getElementById("btn-day-prev").disabled = gridDayIndex <= 0;' +
    '  document.getElementById("btn-day-next").disabled = gridDayIndex >= (GRID_DAYS - 1);' +
    '  var openM = timeToMin(FORM.openTime), closeM = timeToMin(FORM.closeTime), slotM = FORM.slotMinutes || 30;' +
    '  var bookingsForDate = ALL.filter(function(b){ return b.date === dateStr && (b.status === "จอง" || b.status === "ปิด"); });' +
    '  wrap.innerHTML = "";' +
    '  var grid = document.createElement("div");' +
    '  grid.className = "g-grid";' +
    '  var headRow = document.createElement("div");' +
    '  headRow.className = "g-row";' +
    '  var corner = document.createElement("div");' +
    '  corner.className = "g-corner";' +
    '  headRow.appendChild(corner);' +
    '  var mm;' +
    '  for (mm = openM; mm < closeM; mm += slotM) {' +
    '    var colhead = document.createElement("div");' +
    '    colhead.className = "g-colhead";' +
    '    colhead.textContent = minToTime(mm);' +
    '    headRow.appendChild(colhead);' +
    '  }' +
    '  grid.appendChild(headRow);' +
    '  FORM.staffList.forEach(function(staffName){' +
    '    var row = document.createElement("div");' +
    '    row.className = "g-row";' +
    '    var rowLabel = document.createElement("div");' +
    '    rowLabel.className = "g-rowlabel";' +
    '    rowLabel.textContent = staffName;' +
    '    row.appendChild(rowLabel);' +
    '    for (var m2 = openM; m2 < closeM; m2 += slotM) {' +
    '      var slotStart = m2, slotEnd = m2 + slotM;' +
    '      var match = null;' +
    '      for (var i = 0; i < bookingsForDate.length; i++) {' +
    '        var b = bookingsForDate[i];' +
    '        if (b.staff !== staffName) continue;' +
    '        var bStart = timeToMin(b.startTime), bEnd = timeToMin(b.endTime);' +
    '        if (slotStart < bEnd && slotEnd > bStart) { match = b; break; }' +
    '      }' +
    '      var cell = document.createElement("button");' +
    '      cell.type = "button";' +
    '      if (match) {' +
    '        var isBlocked = match.status === "ปิด";' +
    '        var isStart = timeToMin(match.startTime) === slotStart;' +
    '        cell.className = "g-cell " + (isBlocked ? "g-blocked" : "g-booked");' +
    '        cell.title = (isBlocked ? "ปิดคิว" : (match.customerName + " · " + match.service)) + " · " + match.startTime + "-" + match.endTime;' +
    '        if (isStart) cell.textContent = (isBlocked ? "ปิด" : (match.customerName || "")).slice(0, 6);' +
    '        cell.onclick = (function(rec){ return function(){ openDetailModal(rec); }; })(match);' +
    '      } else {' +
    '        cell.className = "g-cell g-free";' +
    '        cell.title = "ว่าง " + minToTime(m2) + " น. — กดเพื่อจอง";' +
    '        cell.onclick = (function(d, t, s){ return function(){ openNewBookingModal({ date: d, startTime: t, staff: s }); }; })(dateStr, minToTime(m2), staffName);' +
    '      }' +
    '      row.appendChild(cell);' +
    '    }' +
    '    grid.appendChild(row);' +
    '  });' +
    '  wrap.appendChild(grid);' +
    '}' +
    'function openDetailModal(rec){' +
    '  currentDetailId = rec.bookingId;' +
    '  document.getElementById("detail-err").style.display = "none";' +
    '  if (rec.status === "ปิด") {' +
    '    document.getElementById("detail-blocked-view").style.display = "block";' +
    '    document.getElementById("detail-edit-view").style.display = "none";' +
    '  } else {' +
    '    document.getElementById("detail-blocked-view").style.display = "none";' +
    '    document.getElementById("detail-edit-view").style.display = "block";' +
    '    document.getElementById("d-bookingid").value = rec.bookingId;' +
    '    document.getElementById("d-service").value = rec.service;' +
    '    document.getElementById("d-date").value = rec.date;' +
    '    document.getElementById("d-start").value = rec.startTime;' +
    '    document.getElementById("d-staff").value = rec.staff;' +
    '    document.getElementById("d-name").value = rec.customerName;' +
    '    document.getElementById("d-phone").value = rec.phone;' +
    '    document.getElementById("d-petname").value = rec.petName;' +
    '    document.getElementById("d-pettype").value = rec.petType;' +
    '    document.getElementById("d-notes").value = rec.notes;' +
    '  }' +
    '  document.getElementById("detail-modal").style.display = "flex";' +
    '}' +
    'function closeDetailModal(){ document.getElementById("detail-modal").style.display = "none"; currentDetailId = null; }' +
    'function saveDetailEdit(){' +
    '  var payload = {' +
    '    service: document.getElementById("d-service").value,' +
    '    date: document.getElementById("d-date").value,' +
    '    startTime: document.getElementById("d-start").value,' +
    '    staff: document.getElementById("d-staff").value,' +
    '    customerName: document.getElementById("d-name").value.trim(),' +
    '    phone: document.getElementById("d-phone").value.trim(),' +
    '    petName: document.getElementById("d-petname").value.trim(),' +
    '    petType: document.getElementById("d-pettype").value.trim(),' +
    '    notes: document.getElementById("d-notes").value.trim()' +
    '  };' +
    '  var errEl = document.getElementById("detail-err");' +
    '  if (!payload.service || !payload.date || !payload.startTime || !payload.staff || !payload.customerName || !payload.phone) {' +
    '    errEl.textContent = "กรอกข้อมูลให้ครบ: บริการ, วันที่, เวลา, ช่าง, ชื่อลูกค้า, เบอร์โทร"; errEl.style.display = "block"; return;' +
    '  }' +
    '  var btn = document.getElementById("btn-detail-save");' +
    '  btn.disabled = true; btn.textContent = "กำลังบันทึก...";' +
    '  google.script.run.withSuccessHandler(function(res){' +
    '    btn.disabled = false; btn.textContent = "บันทึกการแก้ไข";' +
    '    if (res && res.ok) { closeDetailModal(); load(); }' +
    '    else { errEl.textContent = (res && res.error) || "บันทึกไม่สำเร็จ"; errEl.style.display = "block"; }' +
    '  }).withFailureHandler(function(e){' +
    '    btn.disabled = false; btn.textContent = "บันทึกการแก้ไข";' +
    '    errEl.textContent = e && e.message ? e.message : String(e); errEl.style.display = "block";' +
    '  }).staffUpdateBooking(currentDetailId, payload);' +
    '}' +
    'function cancelFromDetail(){' +
    '  if (!currentDetailId) return;' +
    '  if (!confirm("ยืนยันยกเลิกคิวนี้?")) return;' +
    '  var btn = document.getElementById("btn-detail-cancel");' +
    '  btn.disabled = true; btn.textContent = "กำลังยกเลิก...";' +
    '  google.script.run.withSuccessHandler(function(res){' +
    '    btn.disabled = false; btn.textContent = "ยกเลิกคิวนี้";' +
    '    if (res && res.ok) { closeDetailModal(); load(); }' +
    '    else { alert((res && res.error) || "ยกเลิกไม่สำเร็จ"); }' +
    '  }).withFailureHandler(function(e){' +
    '    btn.disabled = false; btn.textContent = "ยกเลิกคิวนี้";' +
    '    alert(e && e.message ? e.message : e);' +
    '  }).staffCancelBooking(currentDetailId);' +
    '}' +
    'function unblockFromDetail(){' +
    '  if (!currentDetailId) return;' +
    '  if (!confirm("ยืนยันเปิดคิวช่วงเวลานี้อีกครั้ง?")) return;' +
    '  var btn = document.getElementById("btn-detail-unblock");' +
    '  btn.disabled = true; btn.textContent = "กำลังเปิดคิว...";' +
    '  google.script.run.withSuccessHandler(function(res){' +
    '    btn.disabled = false; btn.textContent = "เปิดคิวนี้อีกครั้ง";' +
    '    if (res && res.ok) { closeDetailModal(); load(); }' +
    '    else { alert((res && res.error) || "ทำรายการไม่สำเร็จ"); }' +
    '  }).withFailureHandler(function(e){' +
    '    btn.disabled = false; btn.textContent = "เปิดคิวนี้อีกครั้ง";' +
    '    alert(e && e.message ? e.message : e);' +
    '  }).staffCancelBooking(currentDetailId);' +
    '}' +
    'window.addEventListener("DOMContentLoaded", function(){' +
    '  document.getElementById("btn-refresh").onclick = load;' +
    '  document.getElementById("q").oninput = renderTable;' +
    '  document.getElementById("btn-new-booking").onclick = function(){ openNewBookingModal(null); };' +
    '  document.getElementById("btn-modal-cancel").onclick = closeModal;' +
    '  document.getElementById("btn-modal-submit").onclick = submitNewBooking;' +
    '  document.getElementById("tab-btn-list").onclick = function(){ switchTab("list"); };' +
    '  document.getElementById("tab-btn-grid").onclick = function(){ switchTab("grid"); };' +
    '  document.getElementById("btn-day-prev").onclick = function(){ if (gridDayIndex > 0){ gridDayIndex--; renderGrid(); } };' +
    '  document.getElementById("btn-day-next").onclick = function(){ if (gridDayIndex < GRID_DAYS - 1){ gridDayIndex++; renderGrid(); } };' +
    '  document.getElementById("btn-detail-close").onclick = closeDetailModal;' +
    '  document.getElementById("btn-detail-close2").onclick = closeDetailModal;' +
    '  document.getElementById("btn-detail-cancel").onclick = cancelFromDetail;' +
    '  document.getElementById("btn-detail-save").onclick = saveDetailEdit;' +
    '  document.getElementById("btn-detail-unblock").onclick = unblockFromDetail;' +
    '  loadFormOptions();' +
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
