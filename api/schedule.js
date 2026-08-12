/**
 * /api/schedule — ตัวกลางระหว่างหน้าเว็บลูกค้ากับ Google Apps Script
 *
 * ทำไมต้องมีไฟล์นี้:
 * เดิมเบราว์เซอร์ของลูกค้ายิงตรงไปที่ script.google.com ซึ่งมีปัญหา 2 อย่าง
 *   1. Google ตอบกลับด้วย redirect ไปอีกโดเมน (script.googleusercontent.com)
 *      เท่ากับต้องต่อเน็ตใหม่ทั้ง DNS + TLS สองรอบ บนมือถือ 4G ช้ามาก บางทีหลุดไปเลย
 *   2. เป็นคนละโดเมนกับเว็บเรา (cross-origin) มือถือบางเครื่อง/บางเครือข่ายบล็อก
 *
 * พอเปลี่ยนมาผ่านไฟล์นี้:
 *   - เบราว์เซอร์ยิงมาที่ cpsbooking.vercel.app เหมือนกับตอนโหลดหน้าเว็บ (โดเมนเดียวกัน)
 *     ใช้การเชื่อมต่อเดิมที่เปิดค้างอยู่แล้ว ไม่ต้อง DNS/TLS ใหม่ ไม่มี redirect
 *   - Vercel มีเซิร์ฟเวอร์ใกล้ผู้ใช้ และเก็บ cache ไว้ให้ ลูกค้าส่วนใหญ่จึงได้ของทันที
 *     โดยไม่ต้องรอ Google เลย
 */

const APPS_SCRIPT_URL =
  'https://script.google.com/macros/s/AKfycbzIoxyPC9tX9dJrymqOUwM6l_BzEv79gZXqyIAtCE-fZIt69hb7UEFlimmSOAbf0W61Ig/exec?action=schedule';

module.exports = async (req, res) => {
  try {
    const upstream = await fetch(APPS_SCRIPT_URL, { redirect: 'follow' });
    const text = await upstream.text();

    // กันกรณี Apps Script ตอบเป็นหน้า HTML error แทน JSON (เคยเกิดตอนสิทธิ์มีปัญหา)
    // ถ้าไม่ใช่ JSON จริง อย่าเอาไป cache เด็ดขาด ไม่งั้นลูกค้าจะเจอค้างนาน
    if (!upstream.ok || text.charAt(0) !== '{') {
      res.setHeader('Cache-Control', 'no-store');
      res.status(502).json({ error: 'ต้นทางไม่พร้อมใช้งานชั่วคราว' });
      return;
    }

    // s-maxage=30        → Vercel เก็บ cache ไว้ 30 วินาที ลูกค้าที่เข้ามาในช่วงนี้ได้ของทันที
    // stale-while-revalidate=300 → พ้น 30 วิ ยังส่งของเดิมให้ก่อนแบบไม่ต้องรอ แล้วค่อยไปดึงใหม่เบื้องหลัง
    //                              ลูกค้าจึงไม่มีทางต้องนั่งรอ Google เลยแม้แต่คนเดียว
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Cache-Control', 's-maxage=30, stale-while-revalidate=300');
    res.status(200).send(text);
  } catch (e) {
    res.setHeader('Cache-Control', 'no-store');
    res.status(502).json({ error: String((e && e.message) || e) });
  }
};
