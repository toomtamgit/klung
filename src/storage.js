// ตัวแทน window.storage ที่ App.jsx เรียกใช้
// - ข้อมูลที่แชร์ทั้งทีม (shared = true): เก็บใน Supabase ตาราง kv
// - ข้อมูลเฉพาะเครื่อง (shared = false เช่น ชื่อผู้ใช้): เก็บใน localStorage
const SUPABASE_URL = "https://tyoozfkutslmmmrjicwt.supabase.co";
const SUPABASE_KEY = "sb_publishable_St5PhxFno3NRxx11MzNDww_L5PcZHHw"; // publishable key ใส่ในเว็บได้
const TABLE = `${SUPABASE_URL}/rest/v1/kv`;
const HEADERS = { apikey: SUPABASE_KEY, "Content-Type": "application/json" };
const LOCAL_PREFIX = "wh:";

// คีย์ที่โหลดจาก Supabase ไม่สำเร็จ → ห้ามเขียนทับ ไม่งั้นข้อมูลจริงอาจถูกแทนด้วยข้อมูลตัวอย่าง
const failedLoad = new Set();

const local = {
  get(key) {
    try {
      const v = localStorage.getItem(LOCAL_PREFIX + key);
      return v === null ? null : { key, value: v };
    } catch (e) {
      return null;
    }
  },
  set(key, value) {
    try {
      localStorage.setItem(LOCAL_PREFIX + key, value);
    } catch (e) {}
    return { key, value };
  },
};

window.storage = {
  async get(key, shared = false) {
    if (!shared) return local.get(key);
    try {
      const res = await fetch(`${TABLE}?key=eq.${encodeURIComponent(key)}&select=value`, { headers: HEADERS });
      if (!res.ok) throw new Error("HTTP " + res.status);
      const rows = await res.json();
      failedLoad.delete(key);
      return rows.length ? { key, value: rows[0].value } : null;
    } catch (e) {
      failedLoad.add(key);
      console.error("โหลดข้อมูลจาก Supabase ไม่สำเร็จ", key, e);
      return null;
    }
  },
  async set(key, value, shared = false) {
    if (!shared) return local.set(key, value);
    if (failedLoad.has(key)) return null; // กันเขียนทับข้อมูลจริงเมื่อโหลดไม่สำเร็จ
    const res = await fetch(TABLE, {
      method: "POST",
      headers: { ...HEADERS, Prefer: "resolution=merge-duplicates,return=minimal" },
      body: JSON.stringify({ key, value, updated_at: new Date().toISOString() }),
    });
    if (!res.ok) throw new Error("บันทึกไป Supabase ไม่สำเร็จ HTTP " + res.status);
    return { key, value };
  },
};
