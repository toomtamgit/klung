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
const lotLabel = (l) => (l.lot ? `Lot ${l.lot}` : "ไม่ระบุ Lot");
const lotText = (l) => [l.lot ? `Lot ${l.lot}` : "", l.expiry ? `หมดอายุ ${fmtDay(l.expiry)}` : ""].filter(Boolean).join(" · ");
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
  { id: uid(), code: "STA-004", name: "กระดาษ A4", unit: "แพ็ค", category: "เครื่องเขียน", qty: 30, reorderPoint: 15 },
];

function Badge({ children, tone = "default" }) {
  const map = {
    default: { bg: COLORS.border, fg: COLORS.inkSoft },
    amber: { bg: COLORS.amberSoft, fg: COLORS.amber },
    teal: { bg: COLORS.tealSoft, fg: COLORS.teal },
    danger: { bg: COLORS.dangerSoft, fg: COLORS.danger },
  };
  const c = map[tone];
  return (
    <span
      style={{
        background: c.bg,
        color: c.fg,
        fontSize: 12,
        fontWeight: 600,
        padding: "3px 9px",
        borderRadius: 999,
        whiteSpace: "nowrap",
        fontFamily: "'Noto Sans Thai', sans-serif",
      }}
    >
      {children}
    </span>
  );
}

function Card({ children, style }) {
  return (
    <div
      style={{
        background: COLORS.surface,
        border: `1px solid ${COLORS.border}`,
        borderRadius: 14,
        padding: 16,
        ...style,
      }}
    >
      {children}
    </div>
  );
}

function Button({ children, onClick, variant = "primary", disabled, style, small }) {
  const base = {
    fontFamily: "'IBM Plex Sans Thai', sans-serif",
    fontWeight: 600,
    fontSize: small ? 13 : 14,
    padding: small ? "6px 12px" : "9px 16px",
    borderRadius: 10,
    border: "none",
    cursor: disabled ? "not-allowed" : "pointer",
    opacity: disabled ? 0.5 : 1,
    transition: "transform 0.08s ease, opacity 0.15s ease",
  };
  const variants = {
    primary: { background: COLORS.ink, color: "#fff" },
    amber: { background: COLORS.amber, color: "#fff" },
    ghost: { background: "transparent", color: COLORS.inkSoft, border: `1px solid ${COLORS.border}` },
    danger: { background: COLORS.dangerSoft, color: COLORS.danger },
    teal: { background: COLORS.teal, color: "#fff" },
  };
  return (
    <button
      disabled={disabled}
      onClick={onClick}
      onMouseDown={(e) => (e.currentTarget.style.transform = "scale(0.97)")}
      onMouseUp={(e) => (e.currentTarget.style.transform = "scale(1)")}
      style={{ ...base, ...variants[variant], ...style }}
    >
      {children}
    </button>
  );
}

function Input(props) {
  return (
    <input
      {...props}
      style={{
        fontFamily: "'Noto Sans Thai', sans-serif",
        fontSize: 14,
        padding: "8px 10px",
        borderRadius: 8,
        border: `1px solid ${COLORS.border}`,
        outline: "none",
        width: "100%",
        boxSizing: "border-box",
        ...(props.style || {}),
      }}
    />
  );
}

function Select({ options, ...props }) {
  return (
    <select
      {...props}
      style={{
        fontFamily: "'Noto Sans Thai', sans-serif",
        fontSize: 14,
        padding: "8px 10px",
        borderRadius: 8,
        border: `1px solid ${COLORS.border}`,
        outline: "none",
        width: "100%",
        boxSizing: "border-box",
        background: "#fff",
        ...(props.style || {}),
      }}
    >
      {options.map((o) => (
        <option key={o} value={o}>
          {o}
        </option>
      ))}
    </select>
  );
}

function PeriodFilter({ value, onChange }) {
  const filter = value || defaultPeriodFilter();
  const opts = [
    { id: "all", label: "ทั้งหมด" },
    { id: "month", label: "รายเดือน" },
    { id: "year", label: "รายปี" },
  ];
  const currentYear = new Date().getFullYear();
  const years = Array.from({ length: 6 }, (_, i) => currentYear - i);
  return (
    <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
      {opts.map((o) => (
        <button
          key={o.id}
          onClick={() => onChange({ ...filter, mode: o.id })}
          style={{
            fontFamily: "'IBM Plex Sans Thai', sans-serif",
            fontWeight: 600,
            fontSize: 12.5,
            padding: "5px 11px",
            borderRadius: 999,
            border: `1px solid ${filter.mode === o.id ? COLORS.ink : COLORS.border}`,
            background: filter.mode === o.id ? COLORS.ink : "#fff",
            color: filter.mode === o.id ? "#fff" : COLORS.inkSoft,
            cursor: "pointer",
          }}
        >
          {o.label}
        </button>
      ))}
      {filter.mode === "month" && (
        <input
          type="month"
          value={filter.month}
          onChange={(e) => onChange({ ...filter, mode: "month", month: e.target.value })}
          style={{
            fontFamily: "'Noto Sans Thai', sans-serif",
            fontSize: 12.5,
            padding: "5px 8px",
            borderRadius: 8,
            border: `1px solid ${COLORS.border}`,
            outline: "none",
          }}
        />
      )}
      {filter.mode === "year" && (
        <select
          value={filter.year}
          onChange={(e) => onChange({ ...filter, mode: "year", year: e.target.value })}
          style={{
            fontFamily: "'Noto Sans Thai', sans-serif",
            fontSize: 12.5,
            padding: "5px 8px",
            borderRadius: 8,
            border: `1px solid ${COLORS.border}`,
            outline: "none",
            background: "#fff",
          }}
        >
          {years.map((y) => (
            <option key={y} value={y}>
              {y + 543}
            </option>
          ))}
        </select>
      )}
    </div>
  );
}

function SaveMenu({ onExcel, onPdf }) {
  const [open, setOpen] = useState(false);
  return (
    <div style={{ position: "relative" }}>
      <Button
        small
        variant="ghost"
        onClick={(e) => {
          e.stopPropagation();
          setOpen((o) => !o);
        }}
      >
        💾 บันทึก
      </Button>
      {open && (
        <div
          style={{
            position: "absolute",
            right: 0,
            top: "115%",
            background: COLORS.surface,
            border: `1px solid ${COLORS.border}`,
            borderRadius: 10,
            boxShadow: "0 8px 20px rgba(0,0,0,0.14)",
            zIndex: 30,
            minWidth: 170,
            overflow: "hidden",
          }}
        >
          <button
            onClick={(e) => {
              e.stopPropagation();
              onExcel();
              setOpen(false);
            }}
            style={{ display: "block", width: "100%", textAlign: "left", padding: "10px 12px", background: "none", border: "none", fontSize: 13, cursor: "pointer", fontFamily: "'Noto Sans Thai', sans-serif", color: COLORS.ink }}
          >
            📊 บันทึกเป็น Excel
          </button>
          <button
            onClick={(e) => {
              e.stopPropagation();
              onPdf();
              setOpen(false);
            }}
            style={{ display: "block", width: "100%", textAlign: "left", padding: "10px 12px", background: "none", border: "none", borderTop: `1px solid ${COLORS.border}`, fontSize: 13, cursor: "pointer", fontFamily: "'Noto Sans Thai', sans-serif", color: COLORS.ink }}
          >
            🖨️ บันทึกเป็น PDF
          </button>
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [tab, setTab] = useState("overview");
  const [items, setItems] = useState(null);
  const [requests, setRequests] = useState([]);
  const [logs, setLogs] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [toast, setToast] = useState(null);
  const [userName, setUserName] = useState("");
  const [nameDraft, setNameDraft] = useState("");
  const [showNamePrompt, setShowNamePrompt] = useState(false);

  // form state
  const [newItem, setNewItem] = useState({ code: "", name: "", unit: UNITS[0], category: CATEGORIES[0], qty: "", reorderPoint: "", lot: "", expiry: "" });
  const [adjust, setAdjust] = useState({});
  const [adjustDate, setAdjustDate] = useState({});
  const [adjustNote, setAdjustNote] = useState({});
  const [requestQty, setRequestQty] = useState({});
  const [requestDate, setRequestDate] = useState({});
  const [requestPrice, setRequestPrice] = useState({});
  const [expanded, setExpanded] = useState({});
  const [stockSearch, setStockSearch] = useState("");
  const [itemHistoryPeriod, setItemHistoryPeriod] = useState({});
  const [purchasingPeriod, setPurchasingPeriod] = useState(defaultPeriodFilter());
  const [logPeriod, setLogPeriod] = useState(defaultPeriodFilter());
  const [logOptionsOpen, setLogOptionsOpen] = useState({});
  const [historyDetailOpen, setHistoryDetailOpen] = useState({});
  const [itemOptionsOpen, setItemOptionsOpen] = useState({});
  const [invoiceInput, setInvoiceInput] = useState({});

  // หน่วยงานที่เบิก / ผู้จำหน่าย
  const [adjustDept, setAdjustDept] = useState({});
  const [adjustSupplier, setAdjustSupplier] = useState({});
  const [requestSupplier, setRequestSupplier] = useState({});
  const [receiveSupplier, setReceiveSupplier] = useState({});

  // สำรอง / กู้คืนข้อมูล
  const [lastBackup, setLastBackup] = useState(null);
  const [restorePending, setRestorePending] = useState(null);
  const fileInputRef = useRef(null);

  // Lot / วันหมดอายุ
  const [adjustLot, setAdjustLot] = useState({});
  const [adjustExpiry, setAdjustExpiry] = useState({});
  const [receiveLot, setReceiveLot] = useState({});
  const [receiveExpiry, setReceiveExpiry] = useState({});
  const [disposeConfirm, setDisposeConfirm] = useState({});

  // ตรวจนับสต็อก
  const [countingItemId, setCountingItemId] = useState(null);
  const [countQty, setCountQty] = useState("");
  const [countNote, setCountNote] = useState("");
  const openCount = (item) => {
    setCountingItemId(item.id);
    setCountQty(String(item.qty));
    setCountNote("");
  };
  const closeCount = () => setCountingItemId(null);

  // แก้ไขข้อมูลสินค้า (รหัส / ชื่อ / หน่วย / หมวดหมู่ / จุดสั่งซื้อ)
  const [editingItemId, setEditingItemId] = useState(null);
  const [editItemCode, setEditItemCode] = useState("");
  const [editItemName, setEditItemName] = useState("");
  const [editItemUnit, setEditItemUnit] = useState("");
  const [editItemCategory, setEditItemCategory] = useState("");
  const [editItemReorder, setEditItemReorder] = useState("");
  const openItemEdit = (item) => {
    setEditingItemId(item.id);
    setEditItemCode(item.code);
    setEditItemName(item.name);
    setEditItemUnit(item.unit);
    setEditItemCategory(item.category);
    setEditItemReorder(String(item.reorderPoint));
  };
  const closeItemEdit = () => setEditingItemId(null);
  const applyItemEdit = () => {
    const item = items.find((i) => i.id === editingItemId);
    if (!item) return;
    const code = editItemCode.trim();
    const name = editItemName.trim();
    const unit = editItemUnit;
    const category = editItemCategory;
    const reorder = Number(editItemReorder);
    if (!code || !name) {
      showToast("กรอกรหัสสินค้าและชื่อสินค้าให้ครบ");
      return;
    }
    if (editItemReorder === "" || !Number.isFinite(reorder) || reorder < 0) {
      showToast("จุดสั่งซื้อต้องเป็นตัวเลข 0 ขึ้นไป");
      return;
    }
    if (items.some((i) => i.id !== item.id && (i.code || "").toLowerCase() === code.toLowerCase())) {
      showToast(`รหัส "${code}" ซ้ำกับสินค้าอื่น`);
      return;
    }
    const diffs = [];
    if (name !== item.name) diffs.push(`ชื่อ "${item.name}" → "${name}"`);
    if (code !== item.code) diffs.push(`รหัส ${item.code} → ${code}`);
    if (unit !== item.unit) diffs.push(`หน่วย ${item.unit} → ${unit}`);
    if (category !== item.category) diffs.push(`หมวดหมู่ ${item.category} → ${category}`);
    if (reorder !== item.reorderPoint) diffs.push(`จุดสั่งซื้อ ${item.reorderPoint} → ${reorder}`);
    if (diffs.length === 0) {
      closeItemEdit();
      return;
    }
    const id = item.id;
    setItems((prev) => prev.map((i) => (i.id === id ? { ...i, code, name, unit, category, reorderPoint: reorder } : i)));
    // อัปเดตชื่อ/รหัส/หน่วยในใบขอซื้อและประวัติเดิมด้วย ให้ตรงกับของใหม่
    setRequests((prev) => prev.map((r) => (r.itemId === id ? { ...r, itemCode: code, itemName: name, unit } : r)));
    setLogs((prev) => prev.map((l) => (l.itemId === id ? { ...l, itemName: name } : l)));
    addLog({ itemId: id, itemName: name, type: "แก้ไขข้อมูลสินค้า", qty: item.qty, by: userName, note: diffs.join(" · ") });
    setItemOptionsOpen((prev) => ({ ...prev, [id]: false }));
    closeItemEdit();
    showToast("แก้ไขข้อมูลสินค้าแล้ว");
  };

  // แก้ไขตัวเลขในประวัติ
  const [editingLog, setEditingLog] = useState(null);
  const [editQty, setEditQty] = useState("");
  const [editReason, setEditReason] = useState("");

  const openLogEditor = (log) => {
    setEditingLog(log);
    setEditQty(String(log.qty));
    setEditReason("");
  };
  const closeLogEditor = () => setEditingLog(null);

  // คำนวณผลกระทบต่อสต็อกก่อนกดบันทึกจริง
  const getEditImpact = (log, newQty) => {
    const delta = newQty - log.qty;
    const linkedReq = log.requestId ? requests.find((r) => r.id === log.requestId) : null;
    let stockDelta = 0;
    if (linkedReq) {
      if (linkedReq.status === "รับเข้าคลังแล้ว") stockDelta = delta;
    } else if (STOCK_SIGN[log.type]) {
      stockDelta = STOCK_SIGN[log.type] * delta;
    }
    const item = items.find((i) => i.id === log.itemId);
    const before = item ? item.qty : 0;
    return { delta, stockDelta, before, after: before + stockDelta, linkedReq };
  };

  const applyLogEdit = () => {
    const log = editingLog;
    const q = Number(editQty);
    if (!log || !Number.isFinite(q) || q <= 0) {
      showToast("กรอกจำนวนให้ถูกต้อง (มากกว่า 0)");
      return;
    }
    if (q === log.qty) {
      closeLogEditor();
      return;
    }
    const imp = getEditImpact(log, q);
    if (imp.after < 0) {
      showToast("แก้ไม่ได้ เพราะสต็อกจะติดลบ");
      return;
    }
    const editedAt = now();
    const reason = editReason.trim();
    if (imp.stockDelta !== 0) {
      const lotId = log.lotId || (imp.linkedReq && imp.linkedReq.lotId) || null;
      setItems((prev) =>
        prev.map((it) => {
          if (it.id !== log.itemId) return it;
          const newQty = Math.max(0, it.qty + imp.stockDelta);
          let lots = it.lots || [];
          // แก้จำนวนรับเข้า → ปรับ Lot นั้นตาม แล้วให้ผลรวม Lot ไม่เกินสต็อก
          if (lotId && (imp.linkedReq || STOCK_SIGN[log.type] > 0)) {
            lots = lots.map((l) => (l.id === lotId ? { ...l, qty: l.qty + imp.stockDelta } : l)).filter((l) => l.qty > 0);
          }
          return { ...it, qty: newQty, lots: trimLots(lots, newQty) };
        })
      );
    }
    if (imp.linkedReq) {
      setRequests((prev) => prev.map((r) => (r.id === imp.linkedReq.id ? { ...r, qty: q } : r)));
    }
    setLogs((prev) =>
      prev.map((l) => {
        const inGroup = imp.linkedReq ? l.requestId === imp.linkedReq.id && l.type !== "ยกเลิกใบขอซื้อ" : l.id === log.id;
        if (!inGroup) return l;
        return {
          ...l,
          qty: q,
          edited: true,
          origQty: l.origQty !== undefined ? l.origQty : l.qty,
          editedBy: userName || "ไม่ระบุชื่อ",
          editedAt,
          editNote: reason,
        };
      })
    );
    closeLogEditor();
    showToast("แก้ไขตัวเลขแล้ว");
  };

  const toggleHistoryDetail = (logId) => setHistoryDetailOpen((prev) => ({ ...prev, [logId]: !prev[logId] }));
  const toggleItemOptions = (itemId) => setItemOptionsOpen((prev) => ({ ...prev, [itemId]: !prev[itemId] }));
  const deleteItem = (itemId, itemName) => {
    setItems((prev) => prev.filter((i) => i.id !== itemId));
    setItemOptionsOpen((prev) => ({ ...prev, [itemId]: false }));
    showToast(`ลบ "${itemName}" ออกจากคลังแล้ว`);
  };

  const toggleLogOptions = (logId) => setLogOptionsOpen((prev) => ({ ...prev, [logId]: !prev[logId] }));
  const deleteLog = (logId) => {
    setLogs((prev) => prev.filter((l) => l.id !== logId));
    setLogOptionsOpen((prev) => ({ ...prev, [logId]: false }));
    showToast("ลบประวัติแล้ว");
  };

  const toggleExpanded = (itemId) => setExpanded((prev) => ({ ...prev, [itemId]: !prev[itemId] }));

  const showToast = useCallback((msg) => {
    setToast(msg);
    setTimeout(() => setToast(null), 2200);
  }, []);

  // load
  useEffect(() => {
    (async () => {
      try {
        const [i, r, l, u, b] = await Promise.all([
          window.storage.get("wh-items", true).catch(() => null),
          window.storage.get("wh-requests", true).catch(() => null),
          window.storage.get("wh-logs", true).catch(() => null),
          window.storage.get("wh-username", false).catch(() => null),
          window.storage.get("wh-lastbackup", false).catch(() => null),
        ]);
        setItems(i ? JSON.parse(i.value) : SEED_ITEMS);
        setRequests(r ? JSON.parse(r.value) : []);
        setLogs(l ? JSON.parse(l.value) : []);
        if (b && b.value) setLastBackup(b.value);
        if (u && u.value) {
          setUserName(u.value);
        } else {
          setShowNamePrompt(true);
        }
      } catch (e) {
        setItems(SEED_ITEMS);
        setShowNamePrompt(true);
      }
      setLoaded(true);
    })();
  }, []);

  useEffect(() => {
    if (loaded && items) window.storage.set("wh-items", JSON.stringify(items), true).catch(() => {});
  }, [items, loaded]);
  useEffect(() => {
    if (loaded) window.storage.set("wh-requests", JSON.stringify(requests), true).catch(() => {});
  }, [requests, loaded]);
  useEffect(() => {
    if (loaded) window.storage.set("wh-logs", JSON.stringify(logs), true).catch(() => {});
  }, [logs, loaded]);

  const saveUserName = (name) => {
    const trimmed = name.trim();
    if (!trimmed) return;
    setUserName(trimmed);
    window.storage.set("wh-username", trimmed, false).catch(() => {});
    setShowNamePrompt(false);
  };

  const lowStock = useMemo(() => (items || []).filter((it) => it.qty <= it.reorderPoint), [items]);
  // Lot ที่หมดอายุแล้ว หรือใกล้หมดอายุ (เรียงใกล้สุดก่อน)
  const expiryAlerts = useMemo(() => {
    const list = [];
    (items || []).forEach((it) =>
      (it.lots || []).forEach((l) => {
        const st = lotStatus(l.expiry);
        if (st.kind === "expired" || st.kind === "soon") list.push({ item: it, lot: l, st });
      })
    );
    return list.sort((a, b) => a.st.days - b.st.days);
  }, [items]);
  const pendingReqs = useMemo(() => requests.filter((r) => r.status === "รออนุมัติ" || r.status === "สั่งซื้อแล้ว"), [requests]);

  // รายชื่อที่เคยใช้ สำหรับช่องแนะนำ (datalist)
  const deptOptions = useMemo(() => {
    const set = new Set(DEFAULT_DEPARTMENTS);
    logs.forEach((l) => l.dept && set.add(l.dept));
    return [...set];
  }, [logs]);
  const supplierOptions = useMemo(() => {
    const set = new Set();
    requests.forEach((r) => r.supplier && set.add(r.supplier));
    logs.forEach((l) => l.supplier && set.add(l.supplier));
    return [...set];
  }, [logs, requests]);

  const addLog = (entry) =>
    setLogs((prev) =>
      [{ id: uid(), at: now(), ...entry, by: entry.by && entry.by.trim() ? entry.by.trim() : "ไม่ระบุชื่อ" }, ...prev].slice(0, MAX_LOGS)
    );

  // ---------- สำรอง / กู้คืนข้อมูล ----------
  const backupNow = () => {
    const at = now();
    downloadJSON(
      { app: "warehouse-inventory", version: BACKUP_VERSION, exportedAt: at, exportedBy: userName || "ไม่ระบุชื่อ", items: items || [], requests, logs },
      `warehouse-backup-${today()}.json`
    );
    setLastBackup(at);
    window.storage.set("wh-lastbackup", at, false).catch(() => {});
    showToast("สำรองข้อมูลแล้ว ดูไฟล์ในโฟลเดอร์ดาวน์โหลด");
  };

  const onPickRestoreFile = (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const res = parseBackup(String(reader.result || ""));
      if (res.error) {
        showToast(res.error);
        return;
      }
      setRestorePending({ ...res, fileName: file.name });
    };
    reader.onerror = () => showToast("อ่านไฟล์ไม่สำเร็จ");
    reader.readAsText(file);
  };

  const applyRestore = () => {
    const d = restorePending;
    if (!d) return;
    setItems(d.items);
    setRequests(d.requests);
    setLogs(d.logs.slice(0, MAX_LOGS));
    setRestorePending(null);
    showToast("กู้คืนข้อมูลเรียบร้อยแล้ว");
  };

  // สำรองเป็น Excel หลายชีต (สำหรับเปิดดู/พิมพ์ · กู้คืนกลับเข้าระบบไม่ได้ ให้ใช้ไฟล์ .json)
  const backupExcel = () => {
    const wb = XLSX.utils.book_new();
    const stockRows = items.map((it) => ({
      รหัสสินค้า: it.code,
      ชื่อสินค้า: it.name,
      หมวดหมู่: it.category,
      หน่วย: it.unit,
      คงเหลือ: it.qty,
      จุดสั่งซื้อ: it.reorderPoint,
      สถานะ: it.qty <= it.reorderPoint ? "ต่ำกว่าเกณฑ์" : "ปกติ",
    }));
    const lotRows = [];
    items.forEach((it) =>
      (it.lots || []).forEach((l) =>
        lotRows.push({
          รหัสสินค้า: it.code,
          ชื่อสินค้า: it.name,
          Lot: l.lot || "",
          วันหมดอายุ: l.expiry ? fmtDay(l.expiry) : "",
          จำนวน: l.qty,
          หน่วย: it.unit,
        })
      )
    );
    const reqRows = requests.map((r) => ({
      รายการ: r.itemName,
      รหัสสินค้า: r.itemCode,
      จำนวน: r.qty,
      หน่วย: r.unit,
      "ราคา/หน่วย": r.price || "",
      ราคารวม: r.price ? r.price * r.qty : "",
      ผู้จำหน่าย: r.supplier || "",
      สถานะ: r.status,
      "Invoice No.": r.invoiceNo || "",
      วันที่ขอ: fmtDate(r.createdAt),
      ขอโดย: r.requestedBy || "",
      อนุมัติโดย: r.approvedBy || "",
      รับเข้าโดย: r.receivedBy || "",
    }));
    const logRows = logs.map((l) => logRow(l, true));
    const sheets = [
      ["สต็อกคงเหลือ", stockRows],
      ["Lot และวันหมดอายุ", lotRows],
      ["ใบขอซื้อ", reqRows],
      ["ประวัติ", logRows],
    ];
    sheets.forEach(([name, rows]) => {
      if (rows.length > 0) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), name);
    });
    if (wb.SheetNames.length === 0) {
      showToast("ยังไม่มีข้อมูลให้บันทึก");
      return;
    }
    XLSX.writeFile(wb, `warehouse-backup-${today()}.xlsx`);
    showToast("บันทึก Excel แล้ว ดูไฟล์ในโฟลเดอร์ดาวน์โหลด");
  };

  const addItem = () => {
    if (!newItem.name.trim() || newItem.qty === "" || newItem.reorderPoint === "") {
      showToast("กรอกชื่อ จำนวน และจุดสั่งซื้อให้ครบ");
      return;
    }
    const qty = Number(newItem.qty);
    const added = qty > 0 ? addLotEntry([], { lot: newItem.lot, expiry: newItem.expiry, qty }) : { lots: [], id: null };
    const item = {
      id: uid(),
      code: newItem.code.trim() || `ITM-${(items.length + 1).toString().padStart(3, "0")}`,
      name: newItem.name.trim(),
      unit: newItem.unit,
      category: newItem.category,
      qty,
      reorderPoint: Number(newItem.reorderPoint),
      lots: added.lots,
    };
    const lotNote = added.lots[0] ? lotText(added.lots[0]) : "";
    setItems((prev) => [...prev, item]);
    addLog({ itemId: item.id, itemName: item.name, type: "เพิ่มรายการใหม่", qty: item.qty, by: userName, lotNote, lotId: added.id });
    setNewItem({ code: "", name: "", unit: UNITS[0], category: CATEGORIES[0], qty: "", reorderPoint: "", lot: "", expiry: "" });
    showToast(`เพิ่ม "${item.name}" แล้ว`);
  };

  // รับเข้าคลังเอง (ระบุ Lot/วันหมดอายุ/ผู้จำหน่ายได้) · เบิกออกตามหลัก FEFO (หมดอายุก่อน ออกก่อน) และไม่เบิกของที่หมดอายุแล้ว (ระบุหน่วยงานที่เบิกได้)
  const applyAdjust = (itemId, sign) => {
    const val = Number(adjust[itemId] || 0);
    if (!val || val < 0) return;
    const item = items.find((i) => i.id === itemId);
    if (!item) return;
    const at = dateInputToISO(adjustDate[itemId]);
    const note = (adjustNote[itemId] || "").trim();
    const dept = (adjustDept[itemId] || "").trim();
    const supplier = (adjustSupplier[itemId] || "").trim();
    let newLots = item.lots || [];
    let lotNote = "";
    let lotId = null;

    if (sign > 0) {
      const added = addLotEntry(newLots, { lot: adjustLot[itemId], expiry: adjustExpiry[itemId], qty: val });
      newLots = added.lots;
      lotId = added.id;
      const entry = added.lots.find((l) => l.id === added.id);
      if (entry) lotNote = lotText(entry);
    } else {
      const expiredQty = newLots.filter((l) => lotStatus(l.expiry).kind === "expired").reduce((s, l) => s + l.qty, 0);
      const usable = item.qty - expiredQty;
      if (val > item.qty) {
        showToast(`เบิกไม่ได้ สต็อกคงเหลือแค่ ${item.qty} ${item.unit}`);
        return;
      }
      if (val > usable) {
        showToast(`เบิกได้สูงสุด ${usable} ${item.unit} (อีก ${expiredQty} หมดอายุแล้ว กรุณาจำหน่ายออกก่อน)`);
        return;
      }
      const fresh = newLots.filter((l) => lotStatus(l.expiry).kind !== "expired");
      const expiredLots = newLots.filter((l) => lotStatus(l.expiry).kind === "expired");
      const fromLots = Math.min(val, lotsSum(fresh));
      const ded = deductFEFO(fresh, fromLots);
      newLots = [...expiredLots, ...ded.lots];
      lotNote = ded.used.map((u) => `${u.lot ? `Lot ${u.lot}` : "ไม่ระบุ Lot"}${u.expiry ? ` (หมด ${fmtDay(u.expiry)})` : ""} ×${u.qty}`).join(", ");
    }

    setItems((prev) => prev.map((it) => (it.id === itemId ? { ...it, qty: Math.max(0, it.qty + sign * val), lots: newLots } : it)));
    addLog({
      itemId,
      itemName: item.name,
      type: sign > 0 ? "รับเข้าคลัง" : "เบิกออก",
      qty: val,
      at,
      by: userName,
      note,
      lotNote,
      lotId,
      ...(sign > 0 ? { supplier } : { dept }),
    });
    setAdjust((prev) => ({ ...prev, [itemId]: "" }));
    setAdjustNote((prev) => ({ ...prev, [itemId]: "" }));
    setAdjustLot((prev) => ({ ...prev, [itemId]: "" }));
    setAdjustExpiry((prev) => ({ ...prev, [itemId]: "" }));
    if (sign > 0) setAdjustSupplier((prev) => ({ ...prev, [itemId]: "" }));
    else setAdjustDept((prev) => ({ ...prev, [itemId]: "" }));
  };

  // จำหน่ายออก Lot ที่หมดอายุ (ตัดสต็อก + บันทึกประวัติ)
  const disposeLot = (item, lot) => {
    setItems((prev) =>
      prev.map((it) =>
        it.id === item.id ? { ...it, qty: Math.max(0, it.qty - lot.qty), lots: (it.lots || []).filter((l) => l.id !== lot.id) } : it
      )
    );
    addLog({ itemId: item.id, itemName: item.name, type: "จำหน่ายหมดอายุ", qty: lot.qty, by: userName, lotNote: lotText(lot) });
    setDisposeConfirm((prev) => ({ ...prev, [lot.id]: false }));
    showToast(`จำหน่าย ${lotLabel(lot)} ออกแล้ว`);
  };

  // ตรวจนับสต็อก: ตั้งยอดตามที่นับได้จริง และบันทึกส่วนต่าง
  const applyCount = () => {
    const item = items.find((i) => i.id === countingItemId);
    const counted = Number(countQty);
    if (!item || countQty === "" || !Number.isFinite(counted) || counted < 0) {
      showToast("กรอกจำนวนที่นับได้ (0 ขึ้นไป)");
      return;
    }
    setItems((prev) => prev.map((it) => (it.id === item.id ? { ...it, qty: counted, lots: trimLots(it.lots, counted) } : it)));
    addLog({
      itemId: item.id,
      itemName: item.name,
      type: "ตรวจนับสต็อก",
      qty: counted,
      before: item.qty,
      by: userName,
      note: countNote.trim(),
    });
    closeCount();
    showToast(counted === item.qty ? "นับแล้ว ยอดตรงกับระบบ" : `ปรับยอด ${item.name} เป็น ${counted} แล้ว`);
  };

  const createRequest = (item, qty, dateStr, price, supplier) => {
    const at = dateInputToISO(dateStr);
    const unitPrice = Number(price) > 0 ? Number(price) : null;
    const supplierName = (supplier || "").trim();
    const req = {
      id: uid(),
      itemId: item.id,
      itemName: item.name,
      itemCode: item.code,
      unit: item.unit,
      qty,
      price: unitPrice,
      supplier: supplierName,
      status: "รออนุมัติ",
      createdAt: at,
      requestedBy: userName || "ไม่ระบุชื่อ",
    };
    setRequests((prev) => [req, ...prev]);
    addLog({ itemId: item.id, itemName: item.name, type: "สร้างใบขอซื้อ", qty, at, by: userName, requestId: req.id, supplier: supplierName });
    showToast(`สร้างใบขอซื้อ "${item.name}" แล้ว`);
  };

  const createAllLowStockRequests = () => {
    const existingPendingItemIds = new Set(pendingReqs.map((r) => r.itemId));
    const toRequest = lowStock.filter((it) => !existingPendingItemIds.has(it.id));
    if (toRequest.length === 0) {
      showToast("ไม่มีรายการใหม่ที่ต้องขอซื้อ");
      return;
    }
    toRequest.forEach((it) => createRequest(it, Math.max(it.reorderPoint * 2 - it.qty, it.reorderPoint)));
  };

  const updateRequestStatus = (reqId, status, invoiceNo, lotInfo) => {
    const req = requests.find((r) => r.id === reqId);
    const statusByField = status === "สั่งซื้อแล้ว" ? "approvedBy" : status === "รับเข้าคลังแล้ว" ? "receivedBy" : null;
    let lotId = null;
    let lotNote = "";
    let newLots = null;
    let supplierFinal = "";
    if (status === "รับเข้าคลังแล้ว" && req) {
      const target = items.find((i) => i.id === req.itemId);
      const added = addLotEntry(target ? target.lots : [], { lot: lotInfo && lotInfo.lot, expiry: lotInfo && lotInfo.expiry, qty: req.qty });
      newLots = added.lots;
      lotId = added.id;
      const entry = added.lots.find((l) => l.id === added.id);
      if (entry) lotNote = lotText(entry);
      // ผู้จำหน่าย: ใช้ที่กรอกตอนรับเข้า ถ้าไม่ได้แก้ให้ใช้ค่าจากใบขอซื้อ
      supplierFinal = ((lotInfo && lotInfo.supplier !== undefined ? lotInfo.supplier : req.supplier) || "").trim();
    }
    setRequests((prev) =>
      prev.map((r) =>
        r.id === reqId
          ? {
              ...r,
              status,
              ...(statusByField ? { [statusByField]: userName || "ไม่ระบุชื่อ" } : {}),
              ...(status === "รับเข้าคลังแล้ว" ? { invoiceNo: (invoiceNo || "").trim(), lotId, supplier: supplierFinal } : {}),
            }
          : r
      )
    );
    if (status === "รับเข้าคลังแล้ว" && req) {
      setItems((prev) => prev.map((it) => (it.id === req.itemId ? { ...it, qty: it.qty + req.qty, lots: newLots } : it)));
      addLog({ itemId: req.itemId, itemName: req.itemName, type: "รับเข้าคลังจากใบขอซื้อ", qty: req.qty, by: userName, requestId: req.id, lotNote, lotId, supplier: supplierFinal });
      setInvoiceInput((prev) => ({ ...prev, [reqId]: "" }));
      setReceiveLot((prev) => ({ ...prev, [reqId]: "" }));
      setReceiveExpiry((prev) => ({ ...prev, [reqId]: "" }));
      setReceiveSupplier((prev) => {
        const next = { ...prev };
        delete next[reqId];
        return next;
      });
      showToast(`รับ "${req.itemName}" เข้าคลังแล้ว`);
    } else if (status === "สั่งซื้อแล้ว" && req) {
      addLog({ itemId: req.itemId, itemName: req.itemName, type: "อนุมัติใบขอซื้อ", qty: req.qty, by: userName, requestId: req.id, supplier: req.supplier || "" });
    } else if (status === "ยกเลิก" && req) {
      addLog({ itemId: req.itemId, itemName: req.itemName, type: "ยกเลิกใบขอซื้อ", qty: req.qty, by: userName, requestId: req.id });
    }
  };

  // ผู้จำหน่ายที่แสดงในช่องรับเข้า (ค่าที่กำลังพิมพ์ ถ้ายังไม่พิมพ์ใช้ค่าจากใบขอซื้อ)
  const receiveSupplierOf = (r) => (receiveSupplier[r.id] !== undefined ? receiveSupplier[r.id] : r.supplier || "");

  if (!loaded || !items) {
    return (
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100vh", fontFamily: "'Noto Sans Thai', sans-serif", color: COLORS.inkSoft }}>
        <style>{FONT_IMPORT}</style>
        กำลังโหลดข้อมูล...
      </div>
    );
  }

  const TABS = [
    { id: "overview", label: "ภาพรวม" },
    { id: "stock", label: "คลังสินค้า" },
    { id: "purchasing", label: "จัดซื้อ" },
    { id: "log", label: "ประวัติ" },
  ];

  const backupAgeDays = lastBackup ? Math.floor((Date.now() - new Date(lastBackup).getTime()) / 86400000) : null;

  return (
    <div style={{ minHeight: "100vh", background: COLORS.bg, fontFamily: "'Noto Sans Thai', sans-serif", color: COLORS.ink, paddingBottom: 40 }}>
      <style>{`
        ${FONT_IMPORT}
        * { box-sizing: border-box; }
        ::selection { background: ${COLORS.amberSoft}; }
        @keyframes pulse-travel {
          0% { left: 0%; opacity: 0; }
          10% { opacity: 1; }
          90% { opacity: 1; }
          100% { left: 100%; opacity: 0; }
        }
        button:focus-visible, input:focus-visible, select:focus-visible {
          outline: 2px solid ${COLORS.amber};
          outline-offset: 1px;
        }
        @media (prefers-reduced-motion: reduce) {
          .pulse-dot { animation: none !important; }
        }
      `}</style>

      {/* รายชื่อแนะนำสำหรับช่องหน่วยงาน / ผู้จำหน่าย */}
      <datalist id="dept-list">
        {deptOptions.map((d) => (
          <option key={d} value={d} />
        ))}
      </datalist>
      <datalist id="supplier-list">
        {supplierOptions.map((s) => (
          <option key={s} value={s} />
        ))}
      </datalist>

      {/* Header */}
      <div style={{ padding: "20px 16px 8px", maxWidth: 720, margin: "0 auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
          <div>
            <div style={{ fontFamily: "'IBM Plex Sans Thai', sans-serif", fontWeight: 700, fontSize: 22, letterSpacing: -0.3 }}>
              คลัง ⇄ จัดซื้อ
            </div>
            <div style={{ color: COLORS.inkFaint, fontSize: 13, marginTop: 2 }}>เชื่อมสต็อกคงคลังกับใบขอซื้อไว้ในที่เดียว</div>
          </div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", justifyContent: "flex-end" }}>
            <button
              onClick={backupNow}
              style={{
                fontFamily: "'Noto Sans Thai', sans-serif",
                fontSize: 12.5,
                fontWeight: 600,
                background: COLORS.surface,
                border: `1px solid ${COLORS.border}`,
                borderRadius: 999,
                padding: "6px 12px",
                color: COLORS.ink,
                cursor: "pointer",
                whiteSpace: "nowrap",
              }}
            >
              💾 สำรองข้อมูล
            </button>
            <button
              onClick={() => {
                setNameDraft(userName);
                setShowNamePrompt(true);
              }}
              style={{
                fontFamily: "'Noto Sans Thai', sans-serif",
                fontSize: 12.5,
                fontWeight: 600,
                background: COLORS.surface,
                border: `1px solid ${COLORS.border}`,
                borderRadius: 999,
                padding: "6px 12px",
                color: userName ? COLORS.ink : COLORS.inkFaint,
                cursor: "pointer",
                whiteSpace: "nowrap",
              }}
            >
              👤 {userName || "ระบุชื่อคุณ"}
            </button>
          </div>
        </div>
      </div>

      {restorePending && (
        <div
          onClick={() => setRestorePending(null)}
          style={{ position: "fixed", inset: 0, background: "rgba(27,34,51,0.45)", zIndex: 70, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}
        >
          <div onClick={(e) => e.stopPropagation()} style={{ maxWidth: 400, width: "100%" }}>
            <Card>
              <div style={{ fontFamily: "'IBM Plex Sans Thai', sans-serif", fontWeight: 700, fontSize: 16 }}>กู้คืนข้อมูลจากไฟล์นี้?</div>
              <div style={{ fontSize: 12.5, color: COLORS.inkSoft, marginTop: 4, wordBreak: "break-all" }}>{restorePending.fileName}</div>
              {restorePending.exportedAt && (
                <div style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 2 }}>สำรองไว้เมื่อ {fmtDate(restorePending.exportedAt)}</div>
              )}
              <div style={{ display: "flex", gap: 8, margin: "12px 0", flexWrap: "wrap" }}>
                <Badge tone="teal">สินค้า {restorePending.items.length} รายการ</Badge>
                <Badge tone="amber">ใบขอซื้อ {restorePending.requests.length}</Badge>
                <Badge>ประวัติ {restorePending.logs.length}</Badge>
              </div>
              <div style={{ background: COLORS.dangerSoft, color: COLORS.danger, borderRadius: 10, padding: "10px 12px", fontSize: 12.5 }}>
                ข้อมูลปัจจุบันทั้งหมดจะถูกแทนที่ด้วยข้อมูลในไฟล์ และมีผลกับทุกคนในทีมที่ใช้ข้อมูลร่วมกัน
              </div>
              <div style={{ marginTop: 10 }}>
                <Button variant="ghost" small onClick={backupNow}>
                  💾 สำรองข้อมูลปัจจุบันก่อน
                </Button>
              </div>
              <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                <Button variant="ghost" onClick={() => setRestorePending(null)} style={{ flex: 1 }}>
                  ยกเลิก
                </Button>
                <Button onClick={applyRestore} style={{ flex: 2 }}>
                  ยืนยันกู้คืนข้อมูล
                </Button>
              </div>
            </Card>
          </div>
        </div>
      )}

      {countingItemId && (() => {
        const item = items.find((i) => i.id === countingItemId);
        if (!item) return null;
        const filled = countQty !== "";
        const c = Number(countQty);
        const valid = filled && Number.isFinite(c) && c >= 0;
        const diff = valid ? c - item.qty : 0;
        const lotShrink = valid && lotsSum(item.lots) > c;
        const stepBtn = {
          width: 52,
          height: 52,
          borderRadius: 14,
          border: `1px solid ${COLORS.border}`,
          background: COLORS.surface,
          fontSize: 24,
          fontWeight: 600,
          color: COLORS.ink,
          cursor: "pointer",
          fontFamily: "'IBM Plex Mono', monospace",
        };
        const bump = (n) => setCountQty(String(Math.max(0, (valid ? c : item.qty) + n)));
        return (
          <div
            onClick={closeCount}
            style={{ position: "fixed", inset: 0, background: "rgba(27,34,51,0.45)", zIndex: 60, display: "flex", alignItems: "flex-end", justifyContent: "center" }}
          >
            <div
              onClick={(e) => e.stopPropagation()}
              style={{
                background: COLORS.surface,
                width: "100%",
                maxWidth: 480,
                borderRadius: "20px 20px 0 0",
                padding: "10px 18px 22px",
                boxShadow: "0 -10px 40px rgba(0,0,0,0.18)",
                maxHeight: "92vh",
                overflowY: "auto",
              }}
            >
              <div style={{ width: 40, height: 4, borderRadius: 4, background: COLORS.border, margin: "0 auto 14px" }} />
              <div style={{ fontFamily: "'IBM Plex Sans Thai', sans-serif", fontWeight: 700, fontSize: 17 }}>ตรวจนับสต็อก</div>
              <div style={{ fontSize: 13, color: COLORS.inkSoft, marginTop: 3 }}>
                {item.name} · {item.code}
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 12, padding: "10px 12px", background: COLORS.bg, borderRadius: 12 }}>
                <span style={{ fontSize: 13, color: COLORS.inkSoft }}>ยอดในระบบ</span>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600, fontSize: 18 }}>
                  {item.qty} <span style={{ fontSize: 12, color: COLORS.inkFaint }}>{item.unit}</span>
                </span>
              </div>

              <div style={{ fontSize: 12.5, fontWeight: 600, color: COLORS.inkSoft, margin: "16px 0 6px", textAlign: "center" }}>จำนวนที่นับได้จริง</div>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 10, marginBottom: 8 }}>
                <button style={stepBtn} onClick={() => bump(-1)} aria-label="ลด 1">−</button>
                <input
                  type="number"
                  inputMode="numeric"
                  autoFocus
                  value={countQty}
                  onChange={(e) => setCountQty(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && applyCount()}
                  style={{
                    width: 150,
                    height: 56,
                    textAlign: "center",
                    fontFamily: "'IBM Plex Mono', monospace",
                    fontSize: 28,
                    fontWeight: 600,
                    color: COLORS.ink,
                    border: `2px solid ${valid || !filled ? COLORS.ink : COLORS.danger}`,
                    borderRadius: 14,
                    outline: "none",
                  }}
                />
                <button style={stepBtn} onClick={() => bump(1)} aria-label="เพิ่ม 1">+</button>
              </div>
              <div style={{ display: "flex", justifyContent: "center", gap: 6, marginBottom: 14 }}>
                {[-100, -10, 10, 100].map((n) => (
                  <button
                    key={n}
                    onClick={() => bump(n)}
                    style={{
                      fontFamily: "'IBM Plex Mono', monospace",
                      fontSize: 12.5,
                      fontWeight: 600,
                      padding: "5px 11px",
                      borderRadius: 999,
                      border: `1px solid ${COLORS.border}`,
                      background: COLORS.bg,
                      color: COLORS.inkSoft,
                      cursor: "pointer",
                    }}
                  >
                    {n > 0 ? `+${n}` : n}
                  </button>
                ))}
              </div>

              <div
                style={{
                  borderRadius: 12,
                  padding: "10px 12px",
                  fontSize: 13,
                  marginBottom: 12,
                  background: !valid ? COLORS.bg : diff === 0 ? COLORS.tealSoft : diff < 0 ? COLORS.dangerSoft : COLORS.amberSoft,
                  color: COLORS.ink,
                }}
              >
                {!valid ? (
                  <span style={{ color: COLORS.inkSoft }}>กรอกจำนวนที่นับได้ (0 ขึ้นไป)</span>
                ) : diff === 0 ? (
                  <span>✓ ยอดตรงกับระบบ ไม่ต้องปรับ</span>
                ) : (
                  <>
                    <div>
                      ส่วนต่าง: <b>{diff > 0 ? `+${diff}` : diff}</b> {item.unit} ({diff < 0 ? "ของขาด" : "ของเกิน"})
                    </div>
                    <div style={{ marginTop: 3 }}>
                      ยอดคงเหลือจะเปลี่ยน: <b>{item.qty}</b> → <b>{c}</b> {item.unit}
                    </div>
                    {lotShrink && (
                      <div style={{ marginTop: 3, color: COLORS.inkSoft }}>ยอดใน Lot จะถูกตัดตามให้ (ตัดของที่ใกล้หมดอายุก่อน)</div>
                    )}
                  </>
                )}
              </div>

              <Input placeholder="หมายเหตุ (ไม่บังคับ) เช่น นับรอบสิ้นเดือน" value={countNote} onChange={(e) => setCountNote(e.target.value)} />
              <div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginTop: 6 }}>
                บันทึกเป็น "ตรวจนับสต็อก" ในประวัติ พร้อมยอดเดิมและผู้นับ ({userName || "ไม่ระบุชื่อ"})
              </div>

              <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
                <Button variant="ghost" onClick={closeCount} style={{ flex: 1 }}>
                  ยกเลิก
                </Button>
                <Button onClick={applyCount} disabled={!valid} style={{ flex: 2 }}>
                  {valid && diff !== 0 ? "ปรับยอดตามที่นับ" : "บันทึกผลการนับ"}
                </Button>
              </div>
            </div>
          </div>
        );
      })()}

      {editingItemId && (() => {
        const item = items.find((i) => i.id === editingItemId);
        if (!item) return null;
        const ready = editItemCode.trim() && editItemName.trim() && editItemReorder !== "";
        const changed =
          editItemCode.trim() !== item.code ||
          editItemName.trim() !== item.name ||
          editItemUnit !== item.unit ||
          editItemCategory !== item.category ||
          Number(editItemReorder) !== item.reorderPoint;
        const unitOptions = [...new Set([item.unit, ...UNITS])];
        const categoryOptions = [...new Set([item.category, ...CATEGORIES])];
        const label = { fontSize: 12, fontWeight: 600, color: COLORS.inkSoft, margin: "12px 0 5px" };
        return (
          <div
            onClick={closeItemEdit}
            style={{ position: "fixed", inset: 0, background: "rgba(27,34,51,0.45)", zIndex: 60, display: "flex", alignItems: "flex-end", justifyContent: "center" }}
          >
            <div
              onClick={(e) => e.stopPropagation()}
              style={{
                background: COLORS.surface,
                width: "100%",
                maxWidth: 480,
                borderRadius: "20px 20px 0 0",
                padding: "10px 18px 22px",
                boxShadow: "0 -10px 40px rgba(0,0,0,0.18)",
                maxHeight: "92vh",
                overflowY: "auto",
              }}
            >
              <div style={{ width: 40, height: 4, borderRadius: 4, background: COLORS.border, margin: "0 auto 14px" }} />
              <div style={{ fontFamily: "'IBM Plex Sans Thai', sans-serif", fontWeight: 700, fontSize: 17 }}>แก้ไขข้อมูลสินค้า</div>
              <div style={{ fontSize: 12.5, color: COLORS.inkSoft, marginTop: 3 }}>
                จำนวนคงเหลือ ({item.qty} {item.unit}) และ Lot จะไม่เปลี่ยน
              </div>

              <div style={{ ...label, marginTop: 16 }}>รหัสสินค้า</div>
              <Input value={editItemCode} onChange={(e) => setEditItemCode(e.target.value)} placeholder="เช่น MED-001" />

              <div style={label}>ชื่อสินค้า</div>
              <Input value={editItemName} onChange={(e) => setEditItemName(e.target.value)} placeholder="ชื่อสินค้า" />

              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                <div>
                  <div style={label}>หน่วย</div>
                  <Select options={unitOptions} value={editItemUnit} onChange={(e) => setEditItemUnit(e.target.value)} />
                </div>
                <div>
                  <div style={label}>หมวดหมู่</div>
                  <Select options={categoryOptions} value={editItemCategory} onChange={(e) => setEditItemCategory(e.target.value)} />
                </div>
              </div>

              <div style={label}>จุดสั่งซื้อ (เตือนเมื่อคงเหลือน้อยกว่าหรือเท่ากับค่านี้)</div>
              <Input
                type="number"
                inputMode="numeric"
                value={editItemReorder}
                onChange={(e) => setEditItemReorder(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && applyItemEdit()}
              />

              <div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginTop: 10 }}>
                ชื่อ รหัส และหน่วยในใบขอซื้อและประวัติเดิมของสินค้านี้จะเปลี่ยนตามด้วย
                และระบบจะบันทึกการแก้ไขลงประวัติพร้อมค่าเดิม → ค่าใหม่ ({userName || "ไม่ระบุชื่อ"})
              </div>
              <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
                <Button variant="ghost" onClick={closeItemEdit} style={{ flex: 1 }}>
                  ยกเลิก
                </Button>
                <Button onClick={applyItemEdit} disabled={!ready || !changed} style={{ flex: 2 }}>
                  บันทึกการแก้ไข
                </Button>
              </div>
            </div>
          </div>
        );
      })()}

      {editingLog && (() => {
        const q = Number(editQty);
        const valid = Number.isFinite(q) && q > 0;
        const imp = valid ? getEditImpact(editingLog, q) : null;
        const unit = (items.find((i) => i.id === editingLog.itemId) || {}).unit || "";
        const negative = imp && imp.after < 0;
        const unchanged = valid && q === editingLog.qty;
        const stepBtn = {
          width: 52,
          height: 52,
          borderRadius: 14,
          border: `1px solid ${COLORS.border}`,
          background: COLORS.surface,
          fontSize: 24,
          fontWeight: 600,
          color: COLORS.ink,
          cursor: "pointer",
          fontFamily: "'IBM Plex Mono', monospace",
        };
        const bump = (n) => setEditQty(String(Math.max(1, (valid ? q : editingLog.qty) + n)));
        return (
          <div
            onClick={closeLogEditor}
            style={{ position: "fixed", inset: 0, background: "rgba(27,34,51,0.45)", zIndex: 60, display: "flex", alignItems: "flex-end", justifyContent: "center" }}
          >
            <div
              onClick={(e) => e.stopPropagation()}
              style={{
                background: COLORS.surface,
                width: "100%",
                maxWidth: 480,
                borderRadius: "20px 20px 0 0",
                padding: "10px 18px 22px",
                boxShadow: "0 -10px 40px rgba(0,0,0,0.18)",
                maxHeight: "92vh",
                overflowY: "auto",
              }}
            >
              <div style={{ width: 40, height: 4, borderRadius: 4, background: COLORS.border, margin: "0 auto 14px" }} />
              <div style={{ fontFamily: "'IBM Plex Sans Thai', sans-serif", fontWeight: 700, fontSize: 17 }}>แก้ไขจำนวน</div>
              <div style={{ fontSize: 13, color: COLORS.inkSoft, marginTop: 3 }}>
                {editingLog.type} · {editingLog.itemName}
              </div>
              <div style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 1 }}>
                {editingLog.by} · {fmtDate(editingLog.at)}
              </div>

              <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 10, margin: "20px 0 8px" }}>
                <button style={stepBtn} onClick={() => bump(-1)} aria-label="ลด 1">−</button>
                <input
                  type="number"
                  inputMode="numeric"
                  autoFocus
                  value={editQty}
                  onChange={(e) => setEditQty(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && applyLogEdit()}
                  style={{
                    width: 150,
                    height: 56,
                    textAlign: "center",
                    fontFamily: "'IBM Plex Mono', monospace",
                    fontSize: 28,
                    fontWeight: 600,
                    color: COLORS.ink,
                    border: `2px solid ${valid ? COLORS.ink : COLORS.danger}`,
                    borderRadius: 14,
                    outline: "none",
                  }}
                />
                <button style={stepBtn} onClick={() => bump(1)} aria-label="เพิ่ม 1">+</button>
              </div>
              <div style={{ display: "flex", justifyContent: "center", gap: 6, marginBottom: 14 }}>
                {[-100, -10, 10, 100].map((n) => (
                  <button
                    key={n}
                    onClick={() => bump(n)}
                    style={{
                      fontFamily: "'IBM Plex Mono', monospace",
                      fontSize: 12.5,
                      fontWeight: 600,
                      padding: "5px 11px",
                      borderRadius: 999,
                      border: `1px solid ${COLORS.border}`,
                      background: COLORS.bg,
                      color: COLORS.inkSoft,
                      cursor: "pointer",
                    }}
                  >
                    {n > 0 ? `+${n}` : n}
                  </button>
                ))}
              </div>

              <div
                style={{
                  borderRadius: 12,
                  padding: "10px 12px",
                  fontSize: 13,
                  background: negative ? COLORS.dangerSoft : unchanged || !imp ? COLORS.bg : COLORS.tealSoft,
                  color: negative ? COLORS.danger : COLORS.ink,
                  marginBottom: 12,
                }}
              >
                {!valid ? (
                  <span>กรอกจำนวนที่มากกว่า 0</span>
                ) : unchanged ? (
                  <span style={{ color: COLORS.inkSoft }}>ยังไม่ได้เปลี่ยนตัวเลข</span>
                ) : (
                  <>
                    <div>
                      รายการนี้: <b>{editingLog.qty}</b> → <b>{q}</b> {unit}
                    </div>
                    {imp.stockDelta !== 0 ? (
                      <div style={{ marginTop: 3 }}>
                        สต็อกคงเหลือ: <b>{imp.before}</b> → <b>{imp.after}</b> {unit}
                        {negative && " (ติดลบ แก้ไม่ได้)"}
                      </div>
                    ) : (
                      <div style={{ marginTop: 3, color: COLORS.inkSoft }}>
                        {imp.linkedReq ? "ใบขอซื้อจะอัปเดตตามด้วย · สต็อกไม่เปลี่ยน (ยังไม่รับเข้าคลัง)" : "สต็อกคงเหลือไม่เปลี่ยน"}
                      </div>
                    )}
                    {imp.linkedReq && imp.stockDelta !== 0 && (
                      <div style={{ marginTop: 3, color: COLORS.inkSoft }}>จำนวนในใบขอซื้อและประวัติที่เกี่ยวข้องจะอัปเดตตามด้วย</div>
                    )}
                  </>
                )}
              </div>

              <Input placeholder="เหตุผลที่แก้ไข (ไม่บังคับ) เช่น กรอกผิด" value={editReason} onChange={(e) => setEditReason(e.target.value)} />
              <div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginTop: 6 }}>
                ระบบจะเก็บตัวเลขเดิมและชื่อผู้แก้ไขไว้ในประวัติ ({userName || "ไม่ระบุชื่อ"})
              </div>

              <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
                <Button variant="ghost" onClick={closeLogEditor} style={{ flex: 1 }}>
                  ยกเลิก
                </Button>
                <Button onClick={applyLogEdit} disabled={!valid || unchanged || negative} style={{ flex: 2 }}>
                  บันทึกการแก้ไข
                </Button>
              </div>
            </div>
          </div>
        );
      })()}

      {showNamePrompt && (
        <div
          style={{
            position: "fixed",
            inset: 0,
            background: "rgba(27,34,51,0.4)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            zIndex: 50,
            padding: 16,
          }}
        >
          <Card style={{ maxWidth: 360, width: "100%" }}>
            <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 4 }}>คุณคือใคร</div>
            <div style={{ fontSize: 12.5, color: COLORS.inkFaint, marginBottom: 10 }}>
              ชื่อนี้จะถูกบันทึกไว้ในประวัติทุกครั้งที่คุณรับเข้า เบิกออก หรือขอซื้อ เพื่อให้ทีมรู้ว่าใครทำรายการ
            </div>
            <Input
              placeholder="เช่น ทำทาม, พยาบาลสมชาย"
              value={nameDraft}
              onChange={(e) => setNameDraft(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && saveUserName(nameDraft)}
            />
            <div style={{ display: "flex", gap: 8, marginTop: 10, justifyContent: "flex-end" }}>
              {userName && (
                <Button variant="ghost" small onClick={() => setShowNamePrompt(false)}>
                  ยกเลิก
                </Button>
              )}
              <Button small onClick={() => saveUserName(nameDraft)}>
                บันทึกชื่อ
              </Button>
            </div>
          </Card>
        </div>
      )}

      {/* Tabs */}
      <div style={{ maxWidth: 720, margin: "0 auto", padding: "10px 16px 0", display: "flex", gap: 6, overflowX: "auto" }}>
        {TABS.map((t) => (
          <button
            key={t.id}
            onClick={() => setTab(t.id)}
            style={{
              fontFamily: "'IBM Plex Sans Thai', sans-serif",
              fontWeight: 600,
              fontSize: 13.5,
              padding: "8px 14px",
              borderRadius: 999,
              border: `1px solid ${tab === t.id ? COLORS.ink : COLORS.border}`,
              background: tab === t.id ? COLORS.ink : "#fff",
              color: tab === t.id ? "#fff" : COLORS.inkSoft,
              cursor: "pointer",
              whiteSpace: "nowrap",
            }}
          >
            {t.label}
            {t.id === "purchasing" && pendingReqs.length > 0 && (
              <span style={{ marginLeft: 6, fontSize: 11, opacity: 0.85 }}>({pendingReqs.length})</span>
            )}
          </button>
        ))}
      </div>

      <div style={{ maxWidth: 720, margin: "0 auto", padding: "16px" }}>
        {tab === "overview" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            {/* Bridge signature element */}
            <Card style={{ padding: 20 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <div style={{ fontSize: 12, color: COLORS.inkFaint, fontWeight: 600 }}>คลังสินค้า</div>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 26, fontWeight: 600 }}>{items.length}</div>
                  <div style={{ fontSize: 12, color: COLORS.inkFaint }}>รายการทั้งหมด</div>
                </div>
                <div style={{ flex: 1, margin: "0 14px", position: "relative", height: 2, background: COLORS.border, borderRadius: 2 }}>
                  {lowStock.length > 0 && (
                    <div
                      className="pulse-dot"
                      style={{
                        position: "absolute",
                        top: -3,
                        width: 8,
                        height: 8,
                        borderRadius: "50%",
                        background: COLORS.amber,
                        animation: "pulse-travel 2.2s linear infinite",
                      }}
                    />
                  )}
                </div>
                <div style={{ textAlign: "right" }}>
                  <div style={{ fontSize: 12, color: COLORS.inkFaint, fontWeight: 600 }}>จัดซื้อ</div>
                  <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 26, fontWeight: 600, color: pendingReqs.length ? COLORS.amber : COLORS.ink }}>
                    {pendingReqs.length}
                  </div>
                  <div style={{ fontSize: 12, color: COLORS.inkFaint }}>รออนุมัติ/สั่งซื้อ</div>
                </div>
              </div>
            </Card>

            <Card>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                <div style={{ fontWeight: 700, fontSize: 15 }}>รายการต้องเติมสต็อก</div>
                {lowStock.length > 0 && (
                  <Button variant="amber" small onClick={createAllLowStockRequests}>
                    สร้างใบขอซื้อทั้งหมด
                  </Button>
                )}
              </div>
              {lowStock.length === 0 ? (
                <div style={{ color: COLORS.inkFaint, fontSize: 13.5, padding: "8px 0" }}>สต็อกทุกรายการอยู่เหนือจุดสั่งซื้อ ไม่มีอะไรต้องทำตอนนี้</div>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {lowStock.map((it) => (
                    <div key={it.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 10px", background: COLORS.dangerSoft, borderRadius: 10 }}>
                      <div>
                        <div style={{ fontWeight: 600, fontSize: 13.5 }}>{it.name}</div>
                        <div style={{ fontSize: 12, color: COLORS.inkSoft }}>
                          เหลือ {it.qty} {it.unit} · จุดสั่งซื้อ {it.reorderPoint}
                        </div>
                      </div>
                      <Badge tone="danger">ต่ำกว่าเกณฑ์</Badge>
                    </div>
                  ))}
                </div>
              )}
            </Card>

            <Card>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 10 }}>
                <div style={{ fontWeight: 700, fontSize: 15 }}>ของใกล้หมดอายุ / หมดอายุแล้ว</div>
                <div style={{ fontSize: 11.5, color: COLORS.inkFaint }}>เตือนล่วงหน้า {EXPIRY_WARN_DAYS} วัน</div>
              </div>
              {expiryAlerts.length === 0 ? (
                <div style={{ color: COLORS.inkFaint, fontSize: 13.5, padding: "8px 0" }}>
                  ไม่มี Lot ที่ใกล้หมดอายุ (ระบุ Lot และวันหมดอายุตอนรับเข้า ระบบจะเตือนให้อัตโนมัติ)
                </div>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {expiryAlerts.map(({ item, lot, st }) => (
                    <div
                      key={lot.id}
                      style={{
                        display: "flex",
                        justifyContent: "space-between",
                        alignItems: "center",
                        gap: 10,
                        padding: "8px 10px",
                        background: st.kind === "expired" ? COLORS.dangerSoft : COLORS.amberSoft,
                        borderRadius: 10,
                      }}
                    >
                      <div>
                        <div style={{ fontWeight: 600, fontSize: 13.5 }}>{item.name}</div>
                        <div style={{ fontSize: 12, color: COLORS.inkSoft }}>
                          {lotLabel(lot)} · {lot.qty} {item.unit} · หมด {fmtDay(lot.expiry)}
                        </div>
                      </div>
                      <Badge tone={st.kind === "expired" ? "danger" : "amber"}>{st.label}</Badge>
                    </div>
                  ))}
                </div>
              )}
            </Card>

            <Card>
              <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 4 }}>สำรองข้อมูล</div>
              <div style={{ fontSize: 12.5, color: COLORS.inkSoft, marginBottom: 10 }}>
                ดาวน์โหลดข้อมูลทั้งหมด (สินค้า · ใบขอซื้อ · ประวัติ) เป็นไฟล์ .json เก็บไว้เผื่อข้อมูลหาย หรือใช้ย้ายข้อมูลได้
                ส่วนไฟล์ Excel ใช้เปิดดู/พิมพ์/ส่งต่อได้ แต่กู้คืนกลับเข้าระบบไม่ได้
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <Button variant="teal" onClick={backupNow}>
                  💾 สำรองข้อมูล
                </Button>
                <Button variant="ghost" onClick={backupExcel}>
                  📊 บันทึกเป็น Excel
                </Button>
                <Button variant="ghost" onClick={() => fileInputRef.current && fileInputRef.current.click()}>
                  📂 กู้คืนจากไฟล์
                </Button>
                <input ref={fileInputRef} type="file" accept="application/json,.json" onChange={onPickRestoreFile} style={{ display: "none" }} />
              </div>
              <div style={{ fontSize: 12, marginTop: 10, color: backupAgeDays !== null && backupAgeDays >= 7 ? COLORS.amber : COLORS.inkFaint }}>
                {lastBackup
                  ? `สำรองล่าสุด: ${fmtDate(lastBackup)}${backupAgeDays > 0 ? ` (${backupAgeDays} วันที่แล้ว)` : ""}`
                  : "ยังไม่เคยสำรองข้อมูลจากเครื่องนี้"}
              </div>
            </Card>
          </div>
        )}

        {tab === "stock" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <Card>
              <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 10 }}>เพิ่มรายการสินค้าใหม่</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                <Input placeholder="รหัสสินค้า (ไม่บังคับ)" value={newItem.code} onChange={(e) => setNewItem({ ...newItem, code: e.target.value })} />
                <Select options={CATEGORIES} value={newItem.category} onChange={(e) => setNewItem({ ...newItem, category: e.target.value })} />
                <Input style={{ gridColumn: "1 / -1" }} placeholder="ชื่อสินค้า" value={newItem.name} onChange={(e) => setNewItem({ ...newItem, name: e.target.value })} />
                <Input type="number" placeholder="จำนวนคงเหลือ" value={newItem.qty} onChange={(e) => setNewItem({ ...newItem, qty: e.target.value })} />
                <Input type="number" placeholder="จุดสั่งซื้อ (min)" value={newItem.reorderPoint} onChange={(e) => setNewItem({ ...newItem, reorderPoint: e.target.value })} />
                <Select options={UNITS} value={newItem.unit} onChange={(e) => setNewItem({ ...newItem, unit: e.target.value })} />
                <div style={{ gridColumn: "1 / -1", fontSize: 11.5, color: COLORS.inkFaint, marginTop: 2 }}>
                  ถ้ามี Lot / วันหมดอายุ ใส่ไว้ได้เลย (ไม่บังคับ) ระบบจะเตือนก่อนหมดอายุ
                </div>
                <Input placeholder="เลข Lot (ถ้ามี)" value={newItem.lot} onChange={(e) => setNewItem({ ...newItem, lot: e.target.value })} />
                <Input
                  type="date"
                  aria-label="วันหมดอายุ"
                  value={newItem.expiry}
                  onChange={(e) => setNewItem({ ...newItem, expiry: e.target.value })}
                  style={{ minHeight: 38 }}
                />
                <Button onClick={addItem} style={{ gridColumn: "1 / -1" }}>+ เพิ่มรายการ</Button>
              </div>
            </Card>

            <Input
              placeholder="ค้นหารายการ (ชื่อ หรือ รหัสสินค้า)"
              value={stockSearch}
              onChange={(e) => setStockSearch(e.target.value)}
            />

            {items
              .filter((it) => {
                const q = stockSearch.trim().toLowerCase();
                if (!q) return true;
                return it.name.toLowerCase().includes(q) || it.code.toLowerCase().includes(q);
              })
              .map((it) => {
              const low = it.qty <= it.reorderPoint;
              const suggestedQty = Math.max(it.reorderPoint * 2 - it.qty, it.reorderPoint);
              const isOpen = !!expanded[it.id];
              const historyPeriod = itemHistoryPeriod[it.id] || defaultPeriodFilter();
              const itemHistory = logs.filter((l) => l.itemId === it.id && inPeriod(l.at, historyPeriod));
              const itemLots = sortFEFO(it.lots || []);
              const hasExpired = itemLots.some((l) => lotStatus(l.expiry).kind === "expired");
              const hasSoon = itemLots.some((l) => lotStatus(l.expiry).kind === "soon");
              const untracked = Math.max(0, it.qty - lotsSum(itemLots));
              return (
                <Card key={it.id}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                    <div style={{ display: "flex", gap: 8, alignItems: "flex-start" }}>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          toggleItemOptions(it.id);
                        }}
                        style={{
                          flexShrink: 0,
                          marginTop: 1,
                          background: itemOptionsOpen[it.id] ? COLORS.ink : "transparent",
                          color: itemOptionsOpen[it.id] ? "#fff" : COLORS.inkFaint,
                          border: `1px solid ${itemOptionsOpen[it.id] ? COLORS.ink : COLORS.border}`,
                          borderRadius: 8,
                          width: 26,
                          height: 26,
                          fontSize: 14,
                          lineHeight: "24px",
                          cursor: "pointer",
                        }}
                        aria-label="ตัวเลือก"
                      >
                        ⋮
                      </button>
                      <div style={{ cursor: "pointer" }} onClick={() => toggleExpanded(it.id)}>
                        <div style={{ fontWeight: 700, fontSize: 14.5 }}>{it.name}</div>
                        <div style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 2 }}>
                          {it.code} · {it.category}
                        </div>
                      </div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer" }} onClick={() => toggleExpanded(it.id)}>
                      {hasExpired && <Badge tone="danger">มีของหมดอายุ</Badge>}
                      {!hasExpired && hasSoon && <Badge tone="amber">ใกล้หมดอายุ</Badge>}
                      {low ? <Badge tone="danger">ต่ำ</Badge> : <Badge tone="teal">ปกติ</Badge>}
                      <span style={{ color: COLORS.inkFaint, fontSize: 13, transform: isOpen ? "rotate(180deg)" : "none", transition: "transform 0.15s ease" }}>
                        ▾
                      </span>
                    </div>
                  </div>
                  {itemOptionsOpen[it.id] && (
                    <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 8 }}>
                      <Button variant="ghost" small onClick={() => openItemEdit(it)} style={{ marginRight: 8 }}>
                        ✏️ แก้ไขข้อมูล
                      </Button>
                      <Button variant="danger" small onClick={() => deleteItem(it.id, it.name)}>
                        🗑️ ลบรายการ
                      </Button>
                    </div>
                  )}
                  <div style={{ display: "flex", alignItems: "baseline", gap: 6, margin: "10px 0" }}>
                    <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 22, fontWeight: 600 }}>{it.qty}</span>
                    <span style={{ fontSize: 13, color: COLORS.inkFaint }}>{it.unit} · จุดสั่งซื้อ {it.reorderPoint}</span>
                  </div>
                  {itemLots.length > 0 && (
                    <div style={{ display: "flex", flexDirection: "column", gap: 6, margin: "0 0 12px" }}>
                      <div style={{ fontSize: 11.5, fontWeight: 700, color: COLORS.inkFaint }}>แยกตาม Lot (หมดอายุก่อน เบิกก่อน)</div>
                      {itemLots.map((l) => {
                        const st = lotStatus(l.expiry);
                        const bg = st.kind === "expired" ? COLORS.dangerSoft : st.kind === "soon" ? COLORS.amberSoft : COLORS.bg;
                        return (
                          <div key={l.id} style={{ background: bg, borderRadius: 10, padding: "8px 10px" }}>
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
                              <div>
                                <div style={{ fontSize: 13, fontWeight: 600 }}>{lotLabel(l)}</div>
                                <div style={{ fontSize: 11.5, color: COLORS.inkSoft }}>
                                  {l.expiry ? `หมดอายุ ${fmtDay(l.expiry)}` : "ไม่ระบุวันหมดอายุ"}
                                </div>
                              </div>
                              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                                {st.kind !== "none" && st.kind !== "ok" && <Badge tone={st.kind === "expired" ? "danger" : "amber"}>{st.label}</Badge>}
                                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600, fontSize: 14 }}>
                                  {l.qty} <span style={{ fontSize: 11, color: COLORS.inkFaint }}>{it.unit}</span>
                                </span>
                              </div>
                            </div>
                            {st.kind === "expired" && (
                              <div style={{ marginTop: 6, display: "flex", justifyContent: "flex-end", gap: 6, alignItems: "center" }}>
                                {disposeConfirm[l.id] ? (
                                  <>
                                    <span style={{ fontSize: 12, color: COLORS.danger }}>ตัดออกจากสต็อก {l.qty} {it.unit}?</span>
                                    <Button variant="ghost" small onClick={() => setDisposeConfirm((p) => ({ ...p, [l.id]: false }))}>
                                      ไม่ใช่
                                    </Button>
                                    <Button variant="danger" small onClick={() => disposeLot(it, l)}>
                                      ยืนยันจำหน่าย
                                    </Button>
                                  </>
                                ) : (
                                  <Button variant="danger" small onClick={() => setDisposeConfirm((p) => ({ ...p, [l.id]: true }))}>
                                    จำหน่ายออก
                                  </Button>
                                )}
                              </div>
                            )}
                          </div>
                        );
                      })}
                      {untracked > 0 && (
                        <div style={{ fontSize: 11.5, color: COLORS.inkFaint }}>
                          ไม่ระบุ Lot อีก {untracked} {it.unit}
                        </div>
                      )}
                    </div>
                  )}
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                    <Input
                      type="number"
                      placeholder="จำนวน"
                      value={adjust[it.id] || ""}
                      onChange={(e) => setAdjust((prev) => ({ ...prev, [it.id]: e.target.value }))}
                      style={{ maxWidth: 90 }}
                    />
                    <Input
                      type="date"
                      value={adjustDate[it.id] || today()}
                      onChange={(e) => setAdjustDate((prev) => ({ ...prev, [it.id]: e.target.value }))}
                      style={{ maxWidth: 150 }}
                    />
                    <Button variant="teal" small onClick={() => applyAdjust(it.id, 1)}>
                      + รับเข้า
                    </Button>
                    <Button variant="ghost" small onClick={() => applyAdjust(it.id, -1)}>
                      − เบิกออก
                    </Button>
                    <Button variant="ghost" small onClick={() => openCount(it)}>
                      📋 ตรวจนับ
                    </Button>
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6, marginTop: 6 }}>
                    <div>
                      <div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginBottom: 3 }}>หน่วยงานที่เบิก (ใช้ตอนเบิกออก)</div>
                      <Input
                        list="dept-list"
                        placeholder="เช่น ห้องผ่าตัด, ICU"
                        value={adjustDept[it.id] || ""}
                        onChange={(e) => setAdjustDept((prev) => ({ ...prev, [it.id]: e.target.value }))}
                      />
                    </div>
                    <div>
                      <div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginBottom: 3 }}>ผู้จำหน่าย (ใช้ตอนรับเข้า)</div>
                      <Input
                        list="supplier-list"
                        placeholder="ชื่อผู้จำหน่าย (ถ้ามี)"
                        value={adjustSupplier[it.id] || ""}
                        onChange={(e) => setAdjustSupplier((prev) => ({ ...prev, [it.id]: e.target.value }))}
                      />
                    </div>
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6, marginTop: 6 }}>
                    <div>
                      <div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginBottom: 3 }}>Lot (ใช้ตอนรับเข้า)</div>
                      <Input
                        placeholder="เลข Lot (ถ้ามี)"
                        value={adjustLot[it.id] || ""}
                        onChange={(e) => setAdjustLot((prev) => ({ ...prev, [it.id]: e.target.value }))}
                      />
                    </div>
                    <div>
                      <div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginBottom: 3 }}>วันหมดอายุ (ถ้ามี)</div>
                      <Input
                        type="date"
                        value={adjustExpiry[it.id] || ""}
                        onChange={(e) => setAdjustExpiry((prev) => ({ ...prev, [it.id]: e.target.value }))}
                        style={{ minHeight: 38 }}
                      />
                    </div>
                  </div>
                  <div style={{ marginTop: 6 }}>
                    <div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginBottom: 3 }}>เพิ่มเติม</div>
                    <Input
                      placeholder="หมายเหตุเพิ่มเติม (ถ้ามี)"
                      value={adjustNote[it.id] || ""}
                      onChange={(e) => setAdjustNote((prev) => ({ ...prev, [it.id]: e.target.value }))}
                    />
                  </div>
                  {low && (
                    <div style={{ display: "flex", gap: 6, marginTop: 8, alignItems: "center", flexWrap: "wrap", paddingTop: 8, borderTop: `1px solid ${COLORS.border}` }}>
                      <Input
                        type="number"
                        placeholder={`จำนวนขอซื้อ (แนะนำ ${suggestedQty})`}
                        value={requestQty[it.id] || ""}
                        onChange={(e) => setRequestQty((prev) => ({ ...prev, [it.id]: e.target.value }))}
                        style={{ maxWidth: 180 }}
                      />
                      <Input
                        type="date"
                        value={requestDate[it.id] || today()}
                        onChange={(e) => setRequestDate((prev) => ({ ...prev, [it.id]: e.target.value }))}
                        style={{ maxWidth: 150 }}
                      />
                      <Input
                        type="number"
                        placeholder="ราคา/หน่วย (บาท)"
                        value={requestPrice[it.id] || ""}
                        onChange={(e) => setRequestPrice((prev) => ({ ...prev, [it.id]: e.target.value }))}
                        style={{ maxWidth: 150 }}
                      />
                      <Input
                        list="supplier-list"
                        placeholder="ผู้จำหน่าย (ถ้ามี)"
                        value={requestSupplier[it.id] || ""}
                        onChange={(e) => setRequestSupplier((prev) => ({ ...prev, [it.id]: e.target.value }))}
                        style={{ maxWidth: 170 }}
                      />
                      <Button
                        variant="amber"
                        small
                        onClick={() => {
                          const q = Number(requestQty[it.id]) > 0 ? Number(requestQty[it.id]) : suggestedQty;
                          createRequest(it, q, requestDate[it.id], requestPrice[it.id], requestSupplier[it.id]);
                          setRequestQty((prev) => ({ ...prev, [it.id]: "" }));
                          setRequestPrice((prev) => ({ ...prev, [it.id]: "" }));
                          setRequestSupplier((prev) => ({ ...prev, [it.id]: "" }));
                        }}
                      >
                        ขอซื้อ
                      </Button>
                    </div>
                  )}
                  {isOpen && (
                    <div style={{ marginTop: 12, paddingTop: 10, borderTop: `1px solid ${COLORS.border}` }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8, flexWrap: "wrap", gap: 6 }}>
                        <div style={{ fontSize: 12, fontWeight: 700, color: COLORS.inkFaint }}>ประวัติของรายการนี้</div>
                        <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                          <PeriodFilter
                            value={historyPeriod}
                            onChange={(p) => setItemHistoryPeriod((prev) => ({ ...prev, [it.id]: p }))}
                          />
                          <SaveMenu
                            onExcel={() => {
                              const rows = itemHistory.map((l) => logRow(l, false));
                              const ok = exportRowsToExcel(rows, `ประวัติ-${it.name}-${periodLabel(historyPeriod)}.xlsx`);
                              if (!ok) showToast("ไม่มีข้อมูลให้บันทึกในช่วงที่เลือก");
                            }}
                            onPdf={() => {
                              const rows = itemHistory.map((l) => logRow(l, false));
                              const ok = exportRowsToPDF(rows, `ประวัติ ${it.name} (${periodLabel(historyPeriod)})`);
                              if (!ok) showToast("ไม่มีข้อมูล หรือเบราว์เซอร์บล็อกป๊อปอัพ");
                            }}
                          />
                        </div>
                      </div>
                      {historyPeriod.mode === "month" && itemHistory.length > 0 && (
                        <div style={{ display: "flex", gap: 8, marginBottom: 10, flexWrap: "wrap" }}>
                          <div style={{ flex: 1, minWidth: 90, background: COLORS.tealSoft, borderRadius: 10, padding: "8px 10px" }}>
                            <div style={{ fontSize: 11, color: COLORS.teal, fontWeight: 700 }}>รับเข้ารวม</div>
                            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 600, color: COLORS.teal }}>
                              {itemHistory.filter((l) => l.type === "รับเข้าคลัง" || l.type === "รับเข้าคลังจากใบขอซื้อ").reduce((s, l) => s + l.qty, 0)}
                            </div>
                          </div>
                          <div style={{ flex: 1, minWidth: 90, background: COLORS.dangerSoft, borderRadius: 10, padding: "8px 10px" }}>
                            <div style={{ fontSize: 11, color: COLORS.danger, fontWeight: 700 }}>เบิกออกรวม</div>
                            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 600, color: COLORS.danger }}>
                              {itemHistory.filter((l) => l.type === "เบิกออก").reduce((s, l) => s + l.qty, 0)}
                            </div>
                          </div>
                          <div style={{ flex: 1, minWidth: 90, background: COLORS.amberSoft, borderRadius: 10, padding: "8px 10px" }}>
                            <div style={{ fontSize: 11, color: COLORS.amber, fontWeight: 700 }}>ขอซื้อรวม</div>
                            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 600, color: COLORS.amber }}>
                              {itemHistory.filter((l) => l.type === "สร้างใบขอซื้อ").reduce((s, l) => s + l.qty, 0)}
                            </div>
                          </div>
                        </div>
                      )}
                      {itemHistory.length === 0 ? (
                        <div style={{ fontSize: 13, color: COLORS.inkFaint }}>ยังไม่มีประวัติ</div>
                      ) : (
                        <div style={{ display: "flex", flexDirection: "column" }}>
                          {itemHistory.map((l) => {
                            const clickable = REQUEST_LOG_TYPES.includes(l.type);
                            const linkedReq = clickable ? requests.find((r) => r.id === l.requestId) : null;
                            const isDetailOpen = !!historyDetailOpen[l.id];
                            return (
                              <div key={l.id} style={{ borderBottom: `1px solid ${COLORS.border}`, padding: "7px 0" }}>
                                <div
                                  onClick={clickable ? () => toggleHistoryDetail(l.id) : undefined}
                                  style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, cursor: clickable ? "pointer" : "default" }}
                                >
                                  <span>
                                    <span style={{ fontWeight: 600 }}>{l.type}</span>
                                    <span style={{ color: COLORS.inkFaint }}> · {l.by}{detailOf(l) ? ` · ${detailOf(l)}` : ""}</span>
                                    {clickable && <span style={{ color: COLORS.inkFaint, marginLeft: 4 }}>{isDetailOpen ? "▴" : "▾"}</span>}
                                  </span>
                                  <span style={{ display: "flex", gap: 8, alignItems: "center" }}>
                                    {EDITABLE_LOG_TYPES.includes(l.type) ? (
                                      <button
                                        onClick={(e) => {
                                          e.stopPropagation();
                                          openLogEditor(l);
                                        }}
                                        aria-label="แก้ไขจำนวน"
                                        style={{
                                          display: "inline-flex",
                                          alignItems: "center",
                                          gap: 5,
                                          fontFamily: "'IBM Plex Mono', monospace",
                                          fontWeight: 600,
                                          fontSize: 12.5,
                                          color: l.edited ? COLORS.amber : COLORS.ink,
                                          background: l.edited ? COLORS.amberSoft : COLORS.bg,
                                          border: `1px solid ${l.edited ? COLORS.amber : COLORS.border}`,
                                          borderRadius: 8,
                                          padding: "3px 8px",
                                          cursor: "pointer",
                                        }}
                                      >
                                        {l.qty}
                                        <span style={{ fontSize: 11, opacity: 0.7 }}>✎</span>
                                      </button>
                                    ) : (
                                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLORS.inkSoft }}>{l.qty}</span>
                                    )}
                                    <span style={{ color: COLORS.inkFaint }}>{fmtDate(l.at)}</span>
                                  </span>
                                </div>
                                {l.edited && (
                                  <div style={{ marginTop: 4, fontSize: 11.5, color: COLORS.amber }}>
                                    แก้ไขแล้ว · เดิม {l.origQty} → {l.qty} · โดย {l.editedBy} · {fmtDate(l.editedAt)}
                                    {l.editNote ? ` · ${l.editNote}` : ""}
                                  </div>
                                )}
                                {clickable && isDetailOpen && (
                                  <div style={{ marginTop: 6, padding: "8px 10px", background: COLORS.bg, borderRadius: 8, fontSize: 12 }}>
                                    {linkedReq && linkedReq.price ? (
                                      <>
                                        <div>ราคา/หน่วย: {linkedReq.price.toLocaleString("th-TH")} บาท</div>
                                        <div>ราคารวม: {(linkedReq.price * l.qty).toLocaleString("th-TH")} บาท</div>
                                      </>
                                    ) : (
                                      <div style={{ color: COLORS.inkFaint }}>ไม่มีข้อมูลราคาสำหรับรายการนี้</div>
                                    )}
                                    {linkedReq && linkedReq.supplier && (
                                      <div style={{ marginTop: 4 }}>ผู้จำหน่าย: {linkedReq.supplier}</div>
                                    )}
                                    {l.type === "รับเข้าคลังจากใบขอซื้อ" && (
                                      <div style={{ marginTop: linkedReq && linkedReq.price ? 4 : 0 }}>
                                        Invoice No.: {linkedReq && linkedReq.invoiceNo ? linkedReq.invoiceNo : "ไม่มีข้อมูล"}
                                      </div>
                                    )}
                                  </div>
                                )}
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  )}
                </Card>
              );
            })}
          </div>
        )}

        {tab === "purchasing" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <div style={{ display: "flex", justifyContent: "flex-end" }}>
              <PeriodFilter value={purchasingPeriod} onChange={setPurchasingPeriod} />
            </div>
            {requests.filter((r) => inPeriod(r.createdAt, purchasingPeriod)).length === 0 && (
              <div style={{ color: COLORS.inkFaint, fontSize: 13.5, padding: "20px 0", textAlign: "center" }}>
                {requests.length === 0 ? 'ยังไม่มีใบขอซื้อ ไปที่แท็บ "คลังสินค้า" เพื่อสร้างใบขอซื้อ' : "ไม่มีใบขอซื้อในช่วงที่เลือก"}
              </div>
            )}
            {requests.filter((r) => inPeriod(r.createdAt, purchasingPeriod)).map((r) => (
              <Card key={r.id}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                  <div>
                    <div style={{ fontWeight: 700, fontSize: 14.5 }}>{r.itemName}</div>
                    <div style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 2 }}>
                      {r.itemCode} · ขอ {r.qty} {r.unit} · {fmtDate(r.createdAt)}
                      {r.price ? ` · ราคา ${r.price.toLocaleString("th-TH")} บาท/หน่วย · รวม ${(r.price * r.qty).toLocaleString("th-TH")} บาท` : ""}
                    </div>
                    <div style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 2 }}>
                      ขอโดย {r.requestedBy || "ไม่ระบุชื่อ"}
                      {r.approvedBy && ` · อนุมัติโดย ${r.approvedBy}`}
                      {r.receivedBy && ` · รับเข้าโดย ${r.receivedBy}`}
                      {r.supplier && ` · ผู้จำหน่าย ${r.supplier}`}
                      {r.invoiceNo && ` · Invoice No. ${r.invoiceNo}`}
                    </div>
                  </div>
                  <Badge tone={r.status === "รับเข้าคลังแล้ว" ? "teal" : r.status === "สั่งซื้อแล้ว" ? "amber" : "default"}>{r.status}</Badge>
                </div>
                {r.status === "สั่งซื้อแล้ว" && (
                  <div style={{ marginTop: 10, paddingTop: 10, borderTop: `1px solid ${COLORS.border}` }}>
                    <div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginBottom: 3 }}>Invoice No.</div>
                    <Input
                      placeholder="เลขที่ใบแจ้งหนี้ (ถ้ามี)"
                      value={invoiceInput[r.id] || ""}
                      onChange={(e) => setInvoiceInput((prev) => ({ ...prev, [r.id]: e.target.value }))}
                    />
                    <div style={{ marginTop: 8 }}>
                      <div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginBottom: 3 }}>ผู้จำหน่าย</div>
                      <Input
                        list="supplier-list"
                        placeholder="ชื่อผู้จำหน่าย / บริษัท (ถ้ามี)"
                        value={receiveSupplierOf(r)}
                        onChange={(e) => setReceiveSupplier((prev) => ({ ...prev, [r.id]: e.target.value }))}
                      />
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6, marginTop: 8 }}>
                      <div>
                        <div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginBottom: 3 }}>Lot ที่ได้รับ</div>
                        <Input
                          placeholder="เลข Lot (ถ้ามี)"
                          value={receiveLot[r.id] || ""}
                          onChange={(e) => setReceiveLot((prev) => ({ ...prev, [r.id]: e.target.value }))}
                        />
                      </div>
                      <div>
                        <div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginBottom: 3 }}>วันหมดอายุ (ถ้ามี)</div>
                        <Input
                          type="date"
                          value={receiveExpiry[r.id] || ""}
                          onChange={(e) => setReceiveExpiry((prev) => ({ ...prev, [r.id]: e.target.value }))}
                          style={{ minHeight: 38 }}
                        />
                      </div>
                    </div>
                  </div>
                )}
                <div style={{ display: "flex", gap: 6, marginTop: 10 }}>
                  {r.status === "รออนุมัติ" && (
                    <Button small variant="amber" onClick={() => updateRequestStatus(r.id, "สั่งซื้อแล้ว")}>
                      อนุมัติ → สั่งซื้อ
                    </Button>
                  )}
                  {r.status === "สั่งซื้อแล้ว" && (
                    <Button
                      small
                      variant="teal"
                      onClick={() =>
                        updateRequestStatus(r.id, "รับเข้าคลังแล้ว", invoiceInput[r.id], {
                          lot: receiveLot[r.id],
                          expiry: receiveExpiry[r.id],
                          supplier: receiveSupplierOf(r),
                        })
                      }
                    >
                      รับเข้าคลังแล้ว
                    </Button>
                  )}
                  {r.status !== "รับเข้าคลังแล้ว" && r.status !== "ยกเลิก" && (
                    <Button small variant="ghost" onClick={() => updateRequestStatus(r.id, "ยกเลิก")}>
                      ยกเลิก
                    </Button>
                  )}
                </div>
              </Card>
            ))}
          </div>
        )}

        {tab === "log" && (
          <Card>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
              <div style={{ fontWeight: 700, fontSize: 15 }}>ประวัติการเคลื่อนไหว</div>
              <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                <PeriodFilter value={logPeriod} onChange={setLogPeriod} />
                <SaveMenu
                  onExcel={() => {
                    const rows = logs.filter((l) => inPeriod(l.at, logPeriod)).map((l) => logRow(l, true));
                    const ok = exportRowsToExcel(rows, `ประวัติการเคลื่อนไหว-${periodLabel(logPeriod)}.xlsx`);
                    if (!ok) showToast("ไม่มีข้อมูลให้บันทึกในช่วงที่เลือก");
                  }}
                  onPdf={() => {
                    const rows = logs.filter((l) => inPeriod(l.at, logPeriod)).map((l) => logRow(l, true));
                    const ok = exportRowsToPDF(rows, `ประวัติการเคลื่อนไหว (${periodLabel(logPeriod)})`);
                    if (!ok) showToast("ไม่มีข้อมูล หรือเบราว์เซอร์บล็อกป๊อปอัพ");
                  }}
                />
              </div>
            </div>
            {logs.filter((l) => inPeriod(l.at, logPeriod)).length === 0 && (
              <div style={{ color: COLORS.inkFaint, fontSize: 13.5 }}>{logs.length === 0 ? "ยังไม่มีประวัติ" : "ไม่มีประวัติในช่วงที่เลือก"}</div>
            )}
            <div style={{ display: "flex", flexDirection: "column" }}>
              {logs.filter((l) => inPeriod(l.at, logPeriod)).map((l) => (
                <div key={l.id} style={{ padding: "9px 0", borderBottom: `1px solid ${COLORS.border}` }}>
                  <div style={{ display: "flex", alignItems: "flex-start", gap: 8 }}>
                    <button
                      onClick={() => toggleLogOptions(l.id)}
                      style={{
                        flexShrink: 0,
                        marginTop: 1,
                        background: logOptionsOpen[l.id] ? COLORS.ink : "transparent",
                        color: logOptionsOpen[l.id] ? "#fff" : COLORS.inkFaint,
                        border: `1px solid ${logOptionsOpen[l.id] ? COLORS.ink : COLORS.border}`,
                        borderRadius: 8,
                        width: 26,
                        height: 26,
                        fontSize: 14,
                        lineHeight: "24px",
                        cursor: "pointer",
                      }}
                      aria-label="ตัวเลือก"
                    >
                      ⋮
                    </button>
                    <div style={{ flex: 1, display: "flex", justifyContent: "space-between", fontSize: 13, gap: 10, flexWrap: "wrap" }}>
                      <div>
                        <span style={{ fontWeight: 600 }}>{l.type}</span>
                        <span style={{ color: COLORS.inkFaint }}> · {l.itemName}</span>
                        <div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginTop: 1 }}>
                          โดย {l.by}{detailOf(l) ? ` · ${detailOf(l)}` : ""}
                        </div>
                      </div>
                      <div style={{ display: "flex", gap: 10 }}>
                        <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLORS.inkSoft }}>{l.qty}</span>
                        <span style={{ color: COLORS.inkFaint }}>{fmtDate(l.at)}</span>
                      </div>
                    </div>
                  </div>
                  {logOptionsOpen[l.id] && (
                    <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 6 }}>
                      <Button variant="danger" small onClick={() => deleteLog(l.id)}>
                        🗑️ ลบประวัติ
                      </Button>
                    </div>
                  )}
                </div>
              ))}
            </div>
          </Card>
        )}
      </div>

      {toast && (
        <div
          style={{
            position: "fixed",
            bottom: 20,
            left: "50%",
            transform: "translateX(-50%)",
            background: COLORS.ink,
            color: "#fff",
            padding: "10px 18px",
            borderRadius: 10,
            fontSize: 13.5,
            fontWeight: 500,
            boxShadow: "0 8px 24px rgba(0,0,0,0.18)",
            zIndex: 80,
          }}
        >
          {toast}
        </div>
      )}
    </div>
  );
}
