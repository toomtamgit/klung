import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

const FONT_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;500;600;700&family=Noto+Sans+Thai:wght@400;500;600&family=IBM+Plex+Mono:wght@500;600&display=swap');`;

const COLORS = {
  bg: "#F4F5F8",
  surface: "#FFFFFF",
  ink: "#1B2233",
  inkSoft: "#5B6478",
  inkFaint: "#9AA1B2",
  amber: "#DD9A3D",
  amberSoft: "#FBF0DD",
  teal: "#1E7A6C",
  tealSoft: "#E4F3EF",
  danger: "#C94444",
  dangerSoft: "#FBE9E9",
  border: "#E3E6EC",
};

const UNITS = ["ชิ้น", "กล่อง", "ขวด", "ม้วน", "แพ็ค", "ชุด", "ซอง"];
const CATEGORIES = ["เวชภัณฑ์", "อุปกรณ์การแพทย์", "เครื่องเขียน", "ทำความสะอาด", "อื่นๆ"];
// รายชื่อหน่วยงานตั้งต้นสำหรับช่อง "หน่วยงานที่เบิก" (พิมพ์ชื่ออื่นเองได้ และระบบจะจำชื่อที่เคยใช้ให้อัตโนมัติ)
const DEFAULT_DEPARTMENTS = ["ห้องผ่าตัด", "ผู้ป่วยใน", "ผู้ป่วยนอก (OPD)", "ห้องฉุกเฉิน (ER)", "ICU", "ห้องคลอด", "ห้องปฏิบัติการ", "เภสัชกรรม", "ซ่อมบำรุง"];
const REQUEST_LOG_TYPES = ["สร้างใบขอซื้อ", "อนุมัติใบขอซื้อ", "รับเข้าคลังจากใบขอซื้อ"];
// ทิศทางที่แต่ละประเภทรายการกระทบสต็อก (+ เพิ่ม / − ลด)
const STOCK_SIGN = { "เพิ่มรายการใหม่": 1, "รับเข้าคลัง": 1, "รับเข้าคลังจากใบขอซื้อ": 1, "เบิกออก": -1 };
const EDITABLE_LOG_TYPES = ["เพิ่มรายการใหม่", "รับเข้าคลัง", "เบิกออก", "สร้างใบขอซื้อ", "อนุมัติใบขอซื้อ", "รับเข้าคลังจากใบขอซื้อ"];

const MAX_LOGS = 5000; // เก็บประวัติได้สูงสุดกี่รายการ (เดิม 200)
const EXPIRY_WARN_DAYS = 90; // เตือนล่วงหน้ากี่วันก่อนหมดอายุ
const BACKUP_VERSION = 1;

// ---------- Lot / วันหมดอายุ ----------
const daysUntil = (dateStr) => {
  if (!dateStr) return null;
  const d = new Date(dateStr + "T00:00:00");
  const t = new Date();
  t.setHours(0, 0, 0, 0);
  return Math.round((d - t) / 86400000);
};
const fmtDay = (dateStr) =>
  dateStr ? new Date(dateStr + "T12:00:00").toLocaleDateString("th-TH", { day: "2-digit", month: "short", year: "2-digit" }) : "-";
const lotStatus = (expiry) => {
  const d = daysUntil(expiry);
  if (d === null) return { kind: "none", label: "ไม่ระบุวันหมดอายุ", days: null };
  if (d < 0) return { kind: "expired", label: `หมดอายุแล้ว ${Math.abs(d)} วัน`, days: d };
  if (d <= EXPIRY_WARN_DAYS) return { kind: "soon", label: d === 0 ? "หมดอายุวันนี้" : `เหลือ ${d} วัน`, days: d };
  return { kind: "ok", label: `เหลือ ${d} วัน`, days: d };
};
const lotsSum = (lots) => (lots || []).reduce((s, l) => s + l.qty, 0);
// เรียงตามวันหมดอายุใกล้สุดก่อน (FEFO) · ไม่มีวันหมดอายุไว้ท้ายสุด
const sortFEFO = (lots) =>
  [...(lots || [])].sort((a, b) => {
    if (!a.expiry && !b.expiry) return 0;
    if (!a.expiry) return 1;
    if (!b.expiry) return -1;
    return a.expiry < b.expiry ? -1 : a.expiry > b.expiry ? 1 : 0;
  });
// เพิ่มของเข้า Lot · ถ้าไม่ระบุทั้ง Lot และวันหมดอายุ จะไม่สร้าง Lot (นับเป็นสต็อกทั่วไป)
const addLotEntry = (lots, { lot, expiry, qty }) => {
  const lotName = (lot || "").trim();
  if (!lotName && !expiry) return { lots: lots || [], id: null };
  const list = [...(lots || [])];
  const same = list.findIndex((l) => (l.lot || "").toLowerCase() === lotName.toLowerCase() && (l.expiry || "") === (expiry || ""));
  if (same >= 0) {
    list[same] = { ...list[same], qty: list[same].qty + qty };
    return { lots: list, id: list[same].id };
  }
  const entry = { id: uid(), lot: lotName, expiry: expiry || "", qty };
  list.push(entry);
  return { lots: list, id: entry.id };
};
// ตัดสต็อกตาม FEFO · คืนค่า Lot ที่เหลือ และรายการ Lot ที่ถูกตัด
const deductFEFO = (lots, qty) => {
  let remain = qty;
  const used = [];
  const out = sortFEFO(lots).map((l) => {
    if (remain <= 0) return l;
    const take = Math.min(l.qty, remain);
    remain -= take;
    if (take > 0) used.push({ lot: l.lot, expiry: l.expiry, qty: take });
    return { ...l, qty: l.qty - take };
  });
  return { lots: out.filter((l) => l.qty > 0), used };
};
// ให้ผลรวม Lot ไม่เกินสต็อกคงเหลือเสมอ
const trimLots = (lots, maxQty) => {
  const excess = lotsSum(lots) - maxQty;
  if (excess <= 0) return lots || [];
  return deductFEFO(lots, excess).lots;
};
const lotLabel = (l) => (l.lot ? Lot ${l.lot} : "ไม่ระบุ Lot");
const lotText = (l) => [l.lot ? Lot ${l.lot} : "", l.expiry ? หมดอายุ ${fmtDay(l.expiry)} : ""].filter(Boolean).join(" · ");
// ข้อความรายละเอียดที่ต่อท้ายแถวประวัติ (หมายเหตุ + Lot + ผลตรวจนับ)
const noteOf = (l) => {
  const parts = [];
  if (l.note) parts.push(l.note);
  if (l.type === "ตรวจนับสต็อก" && l.before !== undefined) {
    const diff = l.qty - l.before;
    parts.push(`นับได้ ${l.qty} (ระบบ ${l.before}, ${diff === 0 ? "ตรงกัน" : `ต่าง ${diff > 0 ? "+" : ""}${diff}`})`);
  }
  if (l.lotNote) parts.push(l.lotNote);
  return parts.join(" · ");
};
// หน่วยงานที่เบิก / ผู้จำหน่าย ของรายการประวัตินั้น
const partyOf = (l) => {
  const parts = [];
  if (l.dept) parts.push(`หน่วยงาน: ${l.dept}`);
  if (l.supplier) parts.push(`ผู้จำหน่าย: ${l.supplier}`);
  return parts.join(" · ");
};
// ข้อความรายละเอียดสำหรับแสดงในแถวประวัติ (หน่วยงาน/ผู้จำหน่าย + หมายเหตุ)
const detailOf = (l) => [partyOf(l), noteOf(l)].filter(Boolean).join(" · ");
// แถวข้อมูลสำหรับส่งออก Excel / PDF
const logRow = (l, withItem) => ({
  ...(withItem ? { รายการ: l.itemName } : {}),
  ประเภท: l.type,
  จำนวน: l.qty,
  หน่วยงานที่เบิก: l.dept || "",
  ผู้จำหน่าย: l.supplier || "",
  หมายเหตุ: noteOf(l),
  วันที่: fmtDate(l.at),
  ผู้ทำรายการ: l.by,
});

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const now = () => new Date().toISOString();
const today = () => new Date().toISOString().slice(0, 10);
const dateInputToISO = (dateStr) => (dateStr ? new Date(dateStr + "T12:00:00").toISOString() : now());
const defaultPeriodFilter = () => ({
  mode: "all",
  month: new Date().toISOString().slice(0, 7),
  year: String(new Date().getFullYear()),
});

const inPeriod = (iso, filter) => {
  const f = filter || defaultPeriodFilter();
  if (f.mode === "all") return true;
  const d = new Date(iso);
  if (f.mode === "month") {
    const [fy, fm] = (f.month || "").split("-").map(Number);
    return d.getFullYear() === fy && d.getMonth() + 1 === fm;
  }
  if (f.mode === "year") {
    return d.getFullYear() === Number(f.year);
  }
  return true;
};

const fmtDate = (iso) =>
  new Date(iso).toLocaleString("th-TH", { day: "2-digit", month: "short", year: "2-digit", hour: "2-digit", minute: "2-digit" });

const periodLabel = (filter) => {
  const f = filter || defaultPeriodFilter();
  if (f.mode === "month") {
    const [y, m] = (f.month || "").split("-");
    const d = new Date(Number(y), Number(m) - 1, 1);
    return d.toLocaleDateString("th-TH", { month: "long", year: "numeric" });
  }
  if (f.mode === "year") return `พ.ศ. ${Number(f.year) + 543}`;
  return "ทั้งหมด";
};

const exportRowsToExcel = (rows, filename, sheetName = "ข้อมูล") => {
  if (!rows || rows.length === 0) return false;
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
  return true;
};

const exportRowsToPDF = (rows, title) => {
  if (!rows || rows.length === 0) return false;
  const headers = Object.keys(rows[0]);
  const head = `<tr>${headers.map((h) => `<th style="padding:6px 10px;border:1px solid #ddd;background:#f4f4f4;text-align:left;font-size:12.5px;">${h}</th>`).join("")}</tr>`;
  const body = rows
    .map((r) => `<tr>${headers.map((h) => `<td style="padding:6px 10px;border:1px solid #ddd;font-size:12.5px;">${r[h] ?? ""}</td>`).join("")}</tr>`)
    .join("");
  const html = `<!DOCTYPE html><html lang="th"><head><meta charset="utf-8" /><title>${title}</title>
    <style>
      body { font-family: 'Sarabun','Noto Sans Thai',Tahoma,sans-serif; padding: 24px; color:#1B2233; }
      h1 { font-size: 18px; margin-bottom: 4px; }
      p.meta { color:#5B6478; font-size:12px; margin-top:0; margin-bottom:18px; }
      table { border-collapse: collapse; width: 100%; }
    </style></head>
    <body>
      <h1>${title}</h1>
      <p class="meta">สร้างเมื่อ ${new Date().toLocaleString("th-TH")}</p>
      <table>${head}${body}</table>
      <script>window.onload = function () { window.print(); };<\/script>
    </body></html>`;
  const win = window.open("", "_blank");
  if (!win) return false;
  win.document.write(html);
  win.document.close();
  return true;
};

// ---------- สำรอง / กู้คืนข้อมูล ----------
const downloadJSON = (obj, filename) => {
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
};

// ตรวจไฟล์สำรองก่อนกู้คืน · คืนค่า { error } หรือ { items, requests, logs, exportedAt }
const parseBackup = (text) => {
  let d;
  try {
    d = JSON.parse(text);
  } catch (e) {
    return { error: "ไฟล์ไม่ใช่รูปแบบ JSON ที่ถูกต้อง" };
  }
  if (!d || typeof d !== "object" || !Array.isArray(d.items)) {
    return { error: "ไม่ใช่ไฟล์สำรองข้อมูลของระบบนี้" };
  }
  const okItem = (it) => it && typeof it.id === "string" && typeof it.name === "string" && Number.isFinite(Number(it.qty));
  if (!d.items.every(okItem)) {
    return { error: "ข้อมูลรายการสินค้าในไฟล์ไม่สมบูรณ์" };
  }
  return {
    items: d.items,
    requests: Array.isArray(d.requests) ? d.requests : [],
    logs: Array.isArray(d.logs) ? d.logs : [],
    exportedAt: d.exportedAt || null,
  };
};

const SEED_ITEMS = [
  { id: uid(), code: "MED-001", name: "ถุงมือยางไนไตรล์ (M)", unit: "กล่อง", category: "เวชภัณฑ์", qty: 12, reorderPoint: 20 },
  { id: uid(), code: "MED-002", name: "หน้ากากอนามัย 3 ชั้น", unit: "กล่อง", category: "เวชภัณฑ์", qty: 45, reorderPoint: 25 },
  { id: uid(), code: "EQP-010", name: "แอลกอฮอล์ 70% ขนาด 5L", unit: "ขวด", category: "ทำความสะอาด", qty: 6, reorderPoint: 10 },
  { id:
