# คู่มือติดตั้งระบบจองคิว Champion Petshop

ระบบนี้มี 2 ส่วน:

1. **เว็บแอพ** (`index.html`) — หน้าเช็กคิว + จองคิว ที่ลูกค้าเห็น
2. **แบ็กเอนด์** (`Code.gs`) — Google Apps Script ที่อ่าน/เขียนข้อมูลลง Google Sheet

ตอนนี้เว็บแอพยังไม่ได้ต่อกับ Google Sheet จริง (โชว์ "ข้อมูลจำลอง" อยู่) ทำตามขั้นตอนด้านล่างเพื่อเชื่อมให้ครบวงจร

---

## ขั้นตอนที่ 1 — สร้าง Google Sheet

1. ไปที่ sheets.google.com สร้างสเปรดชีตใหม่ ตั้งชื่อว่า "Champion Petshop - ระบบจองคิว"
2. ไม่ต้องสร้างแท็บ/หัวตารางเอง — ระบบจะสร้างให้อัตโนมัติในขั้นตอนถัดไป

## ขั้นตอนที่ 2 — วางโค้ด Apps Script

1. ในสเปรดชีต ไปที่เมนู Extensions > Apps Script
2. ลบโค้ดเดิมทั้งหมดในไฟล์ Code.gs แล้ววางโค้ดจากไฟล์ Code.gs ที่แนบมาให้ทับ
3. กดบันทึก (ไอคอนแผ่นดิสก์ หรือ Ctrl/Cmd+S)
4. ที่แถบด้านบน เลือกฟังก์ชัน setupSheets จากดรอปดาวน์ (ข้าง Debug) แล้วกด Run
   - ครั้งแรกจะขึ้นขอสิทธิ์ (Authorization) — กด Review permissions → เลือกบัญชี Google ของร้าน → กด Advanced → Go to (ชื่อโปรเจกต์) (unsafe) → Allow
5. กลับไปเช็กที่สเปรดชีต ควรเห็นแท็บใหม่ 4 แท็บ: Bookings, Services, Staff, Settings พร้อมข้อมูลตั้งต้น

## ขั้นตอนที่ 3 — Deploy เป็น Web App

1. ใน Apps Script editor กดปุ่ม Deploy > New deployment
2. เลือกประเภท Web app, ตั้ง Execute as: Me, Who has access: Anyone
3. กด Deploy แล้วคัดลอก URL ที่ลงท้ายด้วย /exec

## ขั้นตอนที่ 4 — เชื่อมเว็บแอพเข้ากับ Apps Script

เปิดไฟล์ index.html หาบรรทัด `const APPS_SCRIPT_URL = '';` แล้ววาง URL จากขั้นตอนที่ 3 ระหว่างเครื่องหมายคำพูด จากนั้น deploy ขึ้น Vercel ใหม่
