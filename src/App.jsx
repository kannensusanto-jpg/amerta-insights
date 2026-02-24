import { useState, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  AreaChart, Area, BarChart, Bar, XAxis, YAxis,
  CartesianGrid, Tooltip, ResponsiveContainer, Cell, PieChart, Pie, Legend,
} from "recharts";

// ─── STYLES ───────────────────────────────────────────────────────────────────
const css = `
  @import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Mono:wght@400;500&display=swap');

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: #0c1017; color: #ddd9cf; font-family: 'DM Mono', monospace; }

  .app {
    min-height: 100vh;
    background: #0c1017;
    background-image:
      radial-gradient(ellipse at 15% 0%, rgba(255,172,0,0.07) 0%, transparent 55%),
      radial-gradient(ellipse at 85% 100%, rgba(79,209,165,0.04) 0%, transparent 55%);
  }

  /* ── HEADER ── */
  .hdr {
    padding: 28px 40px 22px;
    border-bottom: 1px solid rgba(255,196,40,0.1);
    display: flex; align-items: center; justify-content: space-between;
    gap: 16px; flex-wrap: wrap;
  }
  .hdr-left h1 {
    font-family: 'Syne', sans-serif; font-size: 22px; font-weight: 800;
    color: #FFC828; letter-spacing: -0.3px;
  }
  .hdr-left p { font-size: 10px; color: rgba(255,196,40,0.38); margin-top: 5px; letter-spacing: 2.5px; text-transform: uppercase; }
  .hdr-right { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }

  .period-pill {
    background: rgba(255,200,40,0.08); border: 1px solid rgba(255,200,40,0.2);
    border-radius: 5px; padding: 7px 14px; font-size: 11px; color: #FFC828; letter-spacing: 1px;
  }
  .swap-btn {
    background: rgba(255,200,40,0.07); border: 1px solid rgba(255,200,40,0.22);
    color: #FFC828; border-radius: 5px; padding: 7px 16px;
    font-family: 'DM Mono', monospace; font-size: 11px; letter-spacing: 1px;
    cursor: pointer; position: relative; transition: background 0.15s;
  }
  .swap-btn:hover { background: rgba(255,200,40,0.15); }
  .swap-btn input { position: absolute; inset: 0; opacity: 0; cursor: pointer; }

  /* ── UPLOAD SCREEN ── */
  .upload-screen {
    display: flex; flex-direction: column; align-items: center;
    justify-content: center; min-height: 80vh; gap: 20px; padding: 40px;
  }
  .upload-title { font-family: 'Syne', sans-serif; font-size: 32px; font-weight: 800; color: #FFC828; text-align: center; }
  .upload-sub { font-size: 12px; color: rgba(221,217,207,0.38); text-align: center; line-height: 2; }
  .drop-box {
    width: 360px; height: 190px; border: 2px dashed rgba(255,200,40,0.28);
    border-radius: 10px; display: flex; flex-direction: column;
    align-items: center; justify-content: center; gap: 14px;
    cursor: pointer; transition: all 0.2s; position: relative;
    background: rgba(255,200,40,0.03);
  }
  .drop-box:hover, .drop-box.over { border-color: rgba(255,200,40,0.65); background: rgba(255,200,40,0.07); }
  .drop-box input { position: absolute; inset: 0; opacity: 0; cursor: pointer; }
  .drop-icon { font-size: 36px; }
  .drop-lbl { font-size: 12px; color: rgba(255,200,40,0.65); text-align: center; line-height: 1.7; }
  .drop-hint { font-size: 10px; color: rgba(221,217,207,0.28); }

  /* ── DASHBOARD ── */
  .dash { padding: 28px 40px 64px; display: flex; flex-direction: column; gap: 28px; }

  /* ── KPI GRID ── */
  .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 14px; }
  .kpi {
    background: rgba(255,255,255,0.025); border: 1px solid rgba(255,200,40,0.09);
    border-radius: 9px; padding: 20px 22px; position: relative; overflow: hidden;
    transition: border-color 0.2s;
  }
  .kpi:hover { border-color: rgba(255,200,40,0.25); }
  .kpi::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 2px; }
  .kpi.yellow::before { background: linear-gradient(90deg, #FFC828, transparent); }
  .kpi.orange::before { background: linear-gradient(90deg, #FF8C3A, transparent); }
  .kpi.teal::before   { background: linear-gradient(90deg, #4FD1A5, transparent); }
  .kpi.blue::before   { background: linear-gradient(90deg, #7B9FFF, transparent); }
  .kpi.red::before    { background: linear-gradient(90deg, #FF5757, transparent); }

  .kpi-lbl { font-size: 9px; letter-spacing: 2.5px; text-transform: uppercase; margin-bottom: 10px; }
  .kpi.yellow .kpi-lbl { color: rgba(255,200,40,0.45); }
  .kpi.orange .kpi-lbl { color: rgba(255,140,58,0.45); }
  .kpi.teal .kpi-lbl   { color: rgba(79,209,165,0.45); }
  .kpi.blue .kpi-lbl   { color: rgba(123,159,255,0.45); }
  .kpi.red .kpi-lbl    { color: rgba(255,87,87,0.45); }

  .kpi-val { font-family: 'Syne', sans-serif; font-size: 24px; font-weight: 700; line-height: 1; }
  .kpi.yellow .kpi-val { color: #FFC828; }
  .kpi.orange .kpi-val { color: #FF8C3A; }
  .kpi.teal .kpi-val   { color: #4FD1A5; }
  .kpi.blue .kpi-val   { color: #7B9FFF; }
  .kpi.red .kpi-val    { color: #FF5757; }

  .kpi-sub { font-size: 10px; color: rgba(221,217,207,0.3); margin-top: 6px; }
  .badges { display: flex; gap: 5px; flex-wrap: wrap; margin-top: 8px; }
  .bdg { padding: 2px 7px; border-radius: 20px; font-size: 9px; letter-spacing: 1px; }
  .bdg-sale  { background: rgba(255,200,40,0.1);  color: #FFC828; border: 1px solid rgba(255,200,40,0.2); }
  .bdg-dp    { background: rgba(123,159,255,0.1); color: #7B9FFF; border: 1px solid rgba(123,159,255,0.2); }
  .bdg-um    { background: rgba(79,209,165,0.1);  color: #4FD1A5; border: 1px solid rgba(79,209,165,0.2); }
  .bdg-batal { background: rgba(255,87,87,0.1);   color: #FF5757; border: 1px solid rgba(255,87,87,0.2); }

  /* ── CHARTS ── */
  .card { background: rgba(255,255,255,0.025); border: 1px solid rgba(255,200,40,0.09); border-radius: 9px; padding: 22px; }
  .card-title { font-family: 'Syne', sans-serif; font-size: 13px; font-weight: 700; color: #ddd9cf; margin-bottom: 18px; }
  .card-title span { font-family: 'DM Mono', monospace; font-size: 10px; font-weight: 400; color: rgba(221,217,207,0.3); margin-left: 8px; }
  .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; }

  .no-data { height: 200px; display: flex; align-items: center; justify-content: center; color: rgba(221,217,207,0.2); font-size: 12px; letter-spacing: 1px; }

  .tooltip-box { background: #18202e; border: 1px solid rgba(255,200,40,0.18); border-radius: 7px; padding: 9px 13px; font-size: 11px; font-family: 'DM Mono', monospace; color: #ddd9cf; }

  /* ── TYPE BREAKDOWN TABLE ── */
  .type-table { width: 100%; border-collapse: collapse; font-size: 11px; }
  .type-table th { padding: 8px 12px; text-align: left; font-size: 9px; letter-spacing: 2px; text-transform: uppercase; color: rgba(255,200,40,0.4); border-bottom: 1px solid rgba(255,255,255,0.05); }
  .type-table td { padding: 9px 12px; border-bottom: 1px solid rgba(255,255,255,0.03); color: rgba(221,217,207,0.75); }
  .type-table tr:last-child td { border-bottom: none; }
  .type-table tr:hover td { background: rgba(255,200,40,0.03); }

  @media (max-width: 860px) {
    .two-col { grid-template-columns: 1fr; }
    .dash { padding: 18px 14px 60px; }
    .hdr { padding: 18px 14px; }
    .drop-box { width: 88vw; }
  }
`;

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const fmt = (n) =>
  n >= 1e9 ? `Rp ${(n / 1e9).toFixed(2)}M`
  : n >= 1e6 ? `Rp ${(n / 1e6).toFixed(1)}Jt`
  : `Rp ${Math.round(n).toLocaleString("id-ID")}`;

const fmtShort = (n) =>
  n >= 1e9 ? `${(n / 1e9).toFixed(1)}M`
  : n >= 1e6 ? `${(n / 1e6).toFixed(0)}Jt`
  : n;

const fmtKg = (n) => `${Math.round(n).toLocaleString("id-ID")} kg`;

const COLORS = ["#FFC828","#FF8C3A","#4FD1A5","#7B9FFF","#FF5757","#C084FC","#38BDF8","#F472B6","#A3E635","#FB923C"];

function extractPeriod(filename) {
  const months = ["JANUARI","FEBRUARI","MARET","APRIL","MEI","JUNI","JULI","AGUSTUS","SEPTEMBER","OKTOBER","NOVEMBER","DESEMBER"];
  const base = filename.replace(/\.xlsx?$/i, "").toUpperCase();
  const parts = base.split("_").filter(Boolean).reverse();
  let month = "", year = "";
  for (const p of parts) {
    if (!year && /^20\d{2}$/.test(p)) year = p;
    if (!month && months.includes(p)) month = p;
    if (month && year) break;
  }
  return [month, year].filter(Boolean).join(" ") || base.replace(/_+/g, " ").trim();
}

// ─── PARSER ─────────────────────────────────────────────────────────────────
// Use header:1 (raw arrays by index) to avoid column-name lookup failures.
// Column layout: A=invNo B=customer C=date D=qty E=total F=ppn G=totalInv H=type
// PPN (F) and TOTAL INVOICE (G) are formulas — compute from TOTAL (E) instead.
function parsePPN(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  return rows
    .slice(1) // skip header row
    .filter(r => r[0] && String(r[0]).startsWith("FK"))
    .map(r => {
      // Date: col C (index 2) — may be JS Date, serial number, or string
      const dr = r[2];
      let date = dr instanceof Date ? dr
               : typeof dr === "number" ? new Date(Math.round((dr - 25569) * 86400000))
               : new Date(dr);
      const total    = parseFloat(r[4]) || 0; // col E: TOTAL (plain number, always reliable)
      const ppn      = total * 0.11;
      const totalInv = total * 1.11;
      return {
        invNo:    String(r[0]),
        customer: String(r[1] || "").trim(),
        date:     isNaN(date?.getTime()) ? null : date,
        qty:      parseFloat(r[3]) || 0, // col D: QTY
        total, ppn, totalInv,
        txType:   String(r[7] || "Sale").trim(), // col H: Type
      };
    });
}

// ─── TOOLTIP ──────────────────────────────────────────────────────────────────
const CT = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div className="tooltip-box">
      <div style={{ color: "#FFC828", marginBottom: 4, fontSize: 10 }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color || "#ddd9cf" }}>
          {p.name}: {typeof p.value === "number" && p.value > 10000 ? fmt(p.value) : p.value?.toLocaleString?.("id-ID") ?? p.value}
        </div>
      ))}
    </div>
  );
};

// ─── APP ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [data, setData]   = useState(null);
  const [drag, setDrag]   = useState(false);

  const load = useCallback((file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
      let rows = [];
      wb.SheetNames.forEach(n => {
        if (n.toUpperCase() === "PPN") rows = parsePPN(wb.Sheets[n]);
      });
      setData({ rows, period: extractPeriod(file.name) });
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const onDrop = useCallback((e) => {
    e.preventDefault(); setDrag(false);
    load(e.dataTransfer?.files[0] || e.target?.files[0]);
  }, [load]);

  // ── COMPUTED ────────────────────────────────────────────────────────────────
  const s = useMemo(() => {
    if (!data) return null;
    const { rows } = data;

    const sales    = rows.filter(r => r.txType === "Sale");
    const nonSales = rows.filter(r => r.txType !== "Sale");

    const totalRev = sales.reduce((a, r) => a + r.totalInv, 0);
    const totalQty = sales.reduce((a, r) => a + r.qty, 0);
    const totalPPN = sales.reduce((a, r) => a + r.ppn, 0);
    const avgInv   = totalRev / (sales.length || 1);

    // Type breakdown
    const typeCounts = {};
    rows.forEach(r => { typeCounts[r.txType] = (typeCounts[r.txType] || 0) + 1; });

    // Customer map (sales only)
    const custMap = {};
    sales.forEach(r => {
      if (!custMap[r.customer]) custMap[r.customer] = { revenue: 0, qty: 0, count: 0 };
      custMap[r.customer].revenue += r.totalInv;
      custMap[r.customer].qty     += r.qty;
      custMap[r.customer].count   += 1;
    });
    const topByRevenue = Object.entries(custMap)
      .map(([name, v]) => ({ name, ...v }))
      .sort((a, b) => b.revenue - a.revenue)
      .slice(0, 10);

    const topByQty = Object.entries(custMap)
      .map(([name, v]) => ({ name, ...v }))
      .sort((a, b) => b.qty - a.qty)
      .slice(0, 10);

    // Daily trend
    const dayMap = {};
    sales.forEach(r => {
      if (!r.date) return;
      const d = r.date.getDate();
      const k = String(d).padStart(2, "0");
      if (!dayMap[k]) dayMap[k] = { date: k, revenue: 0, qty: 0, _d: d };
      dayMap[k].revenue += r.totalInv;
      dayMap[k].qty     += r.qty;
    });
    const dailyTrend = Object.values(dayMap).sort((a, b) => a._d - b._d);

    // Type breakdown for pie
    const typeSlices = Object.entries(typeCounts).map(([name, value]) => ({ name, value }));

    return {
      totalRev, totalQty, totalPPN, avgInv,
      invoiceCount: sales.length,
      custCount: Object.keys(custMap).length,
      typeCounts, typeSlices,
      topByRevenue, topByQty, dailyTrend,
      nonSalesCount: nonSales.length,
    };
  }, [data]);

  // ── RENDER ──────────────────────────────────────────────────────────────────
  return (
    <>
      <style>{css}</style>
      <div className="app">

        {/* HEADER */}
        <header className="hdr">
          <div className="hdr-left">
            <h1>⬡ PT AMERTA NIAGATAMA</h1>
            <p>Sales Intelligence Dashboard</p>
          </div>
          <div className="hdr-right">
            {data && <div className="period-pill">📅 {data.period}</div>}
            {data && (
              <button className="swap-btn">
                ↑ GANTI FILE
                <input type="file" accept=".xlsx,.xls" onChange={onDrop} />
              </button>
            )}
          </div>
        </header>

        {/* UPLOAD */}
        {!data ? (
          <div className="upload-screen">
            <div className="upload-title">DROP YOUR INVOICE FILE</div>
            <p className="upload-sub">Upload file Excel invoice bulanan<br />Dashboard otomatis terbentuk seketika</p>
            <div
              className={`drop-box ${drag ? "over" : ""}`}
              onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
              onDragLeave={() => setDrag(false)}
              onDrop={onDrop}
            >
              <input type="file" accept=".xlsx,.xls" onChange={onDrop} />
              <div className="drop-icon">📊</div>
              <div className="drop-lbl">Drag & drop file Excel di sini<br />atau klik untuk pilih file</div>
              <div className="drop-hint">.xlsx · .xls</div>
            </div>
          </div>

        /* DASHBOARD */
        ) : (
          <div className="dash">

            {/* ── KPI CARDS ── */}
            <div className="kpi-grid">
              <div className="kpi yellow">
                <div className="kpi-lbl">Total Pendapatan</div>
                <div className="kpi-val">{fmt(s.totalRev)}</div>
                <div className="kpi-sub">Invoice Sale · incl. PPN</div>
              </div>
              <div className="kpi orange">
                <div className="kpi-lbl">Total PPN Dipungut</div>
                <div className="kpi-val">{fmt(s.totalPPN)}</div>
                <div className="kpi-sub">PPN 11% dari semua Sale</div>
              </div>
              <div className="kpi teal">
                <div className="kpi-lbl">Total Volume</div>
                <div className="kpi-val">{fmtKg(s.totalQty)}</div>
                <div className="kpi-sub">Total berat (Sale only)</div>
              </div>
              <div className="kpi blue">
                <div className="kpi-lbl">Jumlah Invoice Sale</div>
                <div className="kpi-val">{s.invoiceCount.toLocaleString()}</div>
                <div className="badges">
                  {Object.entries(s.typeCounts).map(([t, n]) => (
                    <span key={t} className={`bdg ${t === "Sale" ? "bdg-sale" : t === "DP" ? "bdg-dp" : t.trim() === "UANG MUKA" ? "bdg-um" : "bdg-batal"}`}>
                      {t.trim()} {n}
                    </span>
                  ))}
                </div>
              </div>
              <div className="kpi yellow">
                <div className="kpi-lbl">Jumlah Customer</div>
                <div className="kpi-val">{s.custCount}</div>
                <div className="kpi-sub">Customer unik periode ini</div>
              </div>
              <div className="kpi teal">
                <div className="kpi-lbl">Rata-rata per Invoice</div>
                <div className="kpi-val">{fmt(s.avgInv)}</div>
                <div className="kpi-sub">Nilai rata-rata per Sale</div>
              </div>
            </div>

            {/* ── DAILY TREND ── */}
            <div className="card">
              <div className="card-title">
                Tren Pendapatan Harian — {data.period}
                <span>per tanggal · Sale only</span>
              </div>
              {s.dailyTrend.length > 0 ? (
                <ResponsiveContainer width="100%" height={220}>
                  <AreaChart data={s.dailyTrend} margin={{ top: 5, right: 5, bottom: 5, left: 10 }}>
                    <defs>
                      <linearGradient id="ag" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="5%"  stopColor="#FFC828" stopOpacity={0.22} />
                        <stop offset="95%" stopColor="#FFC828" stopOpacity={0} />
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" />
                    <XAxis dataKey="date" tick={{ fill: "rgba(221,217,207,0.35)", fontSize: 11 }} axisLine={false} tickLine={false}
                      label={{ value: "Tanggal", position: "insideBottomRight", offset: -4, fill: "rgba(221,217,207,0.18)", fontSize: 9 }} />
                    <YAxis tick={{ fill: "rgba(221,217,207,0.35)", fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort} />
                    <Tooltip content={<CT />} />
                    <Area type="monotone" dataKey="revenue" name="Pendapatan" stroke="#FFC828" strokeWidth={2}
                      fill="url(#ag)" dot={{ fill: "#FFC828", r: 3 }} activeDot={{ r: 5 }} />
                  </AreaChart>
                </ResponsiveContainer>
              ) : <div className="no-data">— DATA TIDAK TERSEDIA —</div>}
            </div>

            {/* ── TOP CUSTOMERS ── */}
            <div className="two-col">
              <div className="card">
                <div className="card-title">Top 10 Customer <span>by Revenue (Rp)</span></div>
                <ResponsiveContainer width="100%" height={300}>
                  <BarChart data={s.topByRevenue} layout="vertical" margin={{ top: 0, right: 8, left: 8, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" horizontal={false} />
                    <XAxis type="number" tick={{ fill: "rgba(221,217,207,0.35)", fontSize: 10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort} />
                    <YAxis type="category" dataKey="name" width={148} tick={{ fill: "rgba(221,217,207,0.6)", fontSize: 10 }} axisLine={false} tickLine={false} />
                    <Tooltip content={<CT />} />
                    <Bar dataKey="revenue" name="Revenue" radius={[0, 4, 4, 0]}>
                      {s.topByRevenue.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>

              <div className="card">
                <div className="card-title">Top 10 Customer <span>by Volume (KG)</span></div>
                <ResponsiveContainer width="100%" height={300}>
                  <BarChart data={s.topByQty} layout="vertical" margin={{ top: 0, right: 8, left: 8, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" horizontal={false} />
                    <XAxis type="number" tick={{ fill: "rgba(221,217,207,0.35)", fontSize: 10 }} axisLine={false} tickLine={false}
                      tickFormatter={v => v >= 1000 ? `${(v/1000).toFixed(0)}k` : v} />
                    <YAxis type="category" dataKey="name" width={148} tick={{ fill: "rgba(221,217,207,0.6)", fontSize: 10 }} axisLine={false} tickLine={false} />
                    <Tooltip content={<CT />} />
                    <Bar dataKey="qty" name="Qty (kg)" radius={[0, 4, 4, 0]}>
                      {s.topByQty.map((_, i) => <Cell key={i} fill={COLORS[(i + 2) % COLORS.length]} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* ── TYPE BREAKDOWN + DAILY QTY ── */}
            <div className="two-col">
              <div className="card">
                <div className="card-title">Komposisi Tipe Invoice <span>semua transaksi</span></div>
                <ResponsiveContainer width="100%" height={200}>
                  <PieChart>
                    <Pie data={s.typeSlices} cx="50%" cy="50%" innerRadius={58} outerRadius={85} paddingAngle={4} dataKey="value">
                      {s.typeSlices.map((_, i) => <Cell key={i} fill={COLORS[i]} />)}
                    </Pie>
                    <Tooltip content={<CT />} />
                    <Legend formatter={v => <span style={{ color: "rgba(221,217,207,0.55)", fontSize: 11 }}>{v}</span>} />
                  </PieChart>
                </ResponsiveContainer>
                <table className="type-table" style={{ marginTop: 12 }}>
                  <thead><tr><th>Tipe</th><th>Jumlah</th><th>%</th></tr></thead>
                  <tbody>
                    {s.typeSlices.map((t, i) => (
                      <tr key={i}>
                        <td style={{ color: COLORS[i] }}>{t.name}</td>
                        <td>{t.value}</td>
                        <td>{((t.value / data.rows.length) * 100).toFixed(1)}%</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="card">
                <div className="card-title">Tren Volume Harian <span>KG · Sale only</span></div>
                {s.dailyTrend.length > 0 ? (
                  <ResponsiveContainer width="100%" height={200}>
                    <AreaChart data={s.dailyTrend} margin={{ top: 5, right: 5, bottom: 5, left: 10 }}>
                      <defs>
                        <linearGradient id="qg" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="5%"  stopColor="#4FD1A5" stopOpacity={0.22} />
                          <stop offset="95%" stopColor="#4FD1A5" stopOpacity={0} />
                        </linearGradient>
                      </defs>
                      <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" />
                      <XAxis dataKey="date" tick={{ fill: "rgba(221,217,207,0.35)", fontSize: 10 }} axisLine={false} tickLine={false} />
                      <YAxis tick={{ fill: "rgba(221,217,207,0.35)", fontSize: 10 }} axisLine={false} tickLine={false}
                        tickFormatter={v => v >= 1000 ? `${(v/1000).toFixed(0)}k` : v} />
                      <Tooltip content={<CT />} />
                      <Area type="monotone" dataKey="qty" name="Volume (kg)" stroke="#4FD1A5" strokeWidth={2}
                        fill="url(#qg)" dot={{ fill: "#4FD1A5", r: 3 }} activeDot={{ r: 5 }} />
                    </AreaChart>
                  </ResponsiveContainer>
                ) : <div className="no-data">— DATA TIDAK TERSEDIA —</div>}
              </div>
            </div>

          </div>
        )}
      </div>
    </>
  );
}
