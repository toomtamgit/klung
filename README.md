# คลัง ⇄ จัดซื้อ

เว็บแอปจัดการสต็อกคลังและใบขอซื้อ (React + Vite)

## ข้อมูลและ Supabase
ข้อมูลสินค้า ใบขอซื้อ และประวัติ เก็บใน Supabase (ตาราง `kv`) ทุกเครื่องที่เปิดเว็บเห็นชุดเดียวกัน
- ต้องรันไฟล์ `supabase.sql` ใน Supabase → SQL Editor หนึ่งครั้งก่อนใช้งาน
- ถ้าสองคนแก้พร้อมกัน คนที่บันทึกทีหลังจะทับของคนก่อน และต้องรีเฟรชหน้าเพื่อดูสิ่งที่คนอื่นแก้
- ชื่อผู้ใช้เก็บในเครื่องแต่ละเครื่อง
- แอปไม่มีระบบล็อกอิน ห้ามใส่ข้อมูลผู้ป่วยหรือข้อมูลส่วนบุคคล
- กด "💾 สำรองข้อมูล" เป็นประจำ

## วิธี A: Cloudflare Pages (แนะนำ ง่ายสุด)
1. อัปโหลดไฟล์ทั้งหมดในโฟลเดอร์นี้ขึ้น GitHub repo ใหม่
2. เข้า dash.cloudflare.com → Workers & Pages → Create → Pages → Connect to Git → เลือก repo
3. ตั้งค่า
   - Framework preset: Vite
   - Build command: `npm run build`
   - Build output directory: `dist`
4. กด Save and Deploy รอ 1-2 นาที จะได้ลิงก์ `ชื่อ.pages.dev`

## วิธี B: GitHub Pages
1. อัปโหลดไฟล์ทั้งหมด (รวมโฟลเดอร์ `.github`) ขึ้น GitHub repo ใหม่ branch `main`
2. ใน repo ไปที่ Settings → Pages → Source เลือก **GitHub Actions**
3. ไปที่แท็บ Actions รอให้ขึ้นเครื่องหมายถูกสีเขียว
4. ลิงก์จะเป็น `https://ชื่อผู้ใช้.github.io/ชื่อ-repo/`

## อัปเดตครั้งต่อไป
แก้ไฟล์ `src/App.jsx` ใน GitHub แล้ว commit ระบบจะ build และ deploy ใหม่ให้อัตโนมัติ

## ทดสอบในเครื่อง (ไม่บังคับ)
    npm install
    npm run dev
