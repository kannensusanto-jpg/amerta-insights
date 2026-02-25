import { useState, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  AreaChart, Area, BarChart, Bar, XAxis, YAxis,
  CartesianGrid, Tooltip, ResponsiveContainer, Cell,
  PieChart, Pie, Legend,
} from "recharts";

const css = `
  @import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Mono:wght@400;500&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: #0c1017; color: #ddd9cf; font-family: 'DM Mono', monospace; }
  .app {
    min-height: 100vh; background: #0c1017;
    background-image:
      radial-gradient(ellipse at 15% 0%, rgba(255,172,0,0.07) 0%, transparent 55%),
      radial-gradient(ellipse at 85% 100%, rgba(79,209,165,0.04) 0%, transparent 55%);
  }
  .hdr { padding: 24px 40px 0; border-bottom: 1px solid rgba(255,196,40,0.1); display: flex; flex-direction: column; gap: 0; }
  .hdr-top { display: flex; align-items: center; justify-content: space-between; gap: 16px; flex-wrap: wrap; padding-bottom: 18px; }
  .hdr-left h1 { font-family: 'Syne', sans-serif; font-size: 22px; font-weight: 800; color: #FFC828; letter-spacing: -0.3px; }
  .hdr-left p { font-size: 10px; color: rgba(255,196,40,0.38); margin-top: 5px; letter-spacing: 2.5px; text-transform: uppercase; }
  .hdr-right { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }
  .period-pill { background: rgba(255,200,40,0.08); border: 1px solid rgba(255,200,40,0.2); border-radius: 5px; padding: 7px 14px; font-size: 11px; color: #FFC828; letter-spacing: 1px; }
  .swap-btn { background: rgba(255,200,40,0.07); border: 1px solid rgba(255,200,40,0.22); color: #FFC828; border-radius: 5px; padding: 7px 16px; font-family: 'DM Mono', monospace; font-size: 11px; letter-spacing: 1px; cursor: pointer; position: relative; transition: background 0.15s; }
  .swap-btn:hover { background: rgba(255,200,40,0.15); }
  .swap-btn input { position: absolute; inset: 0; opacity: 0; cursor: pointer; }
  .tabs { display: flex; gap: 0; }
  .tab { padding: 12px 22px; font-family: 'DM Mono', monospace; font-size: 11px; letter-spacing: 1.5px; text-transform: uppercase; cursor: pointer; border-bottom: 2px solid transparent; color: rgba(221,217,207,0.35); transition: all 0.15s; background: none; border-top: none; border-left: none; border-right: none; }
  .tab:hover { color: rgba(221,217,207,0.7); }
  .tab.active { color: #FFC828; border-bottom-color: #FFC828; }
  .upload-screen { display: flex; flex-direction: column; align-items: center; justify-content: center; min-height: 80vh; gap: 20px; padding: 40px; }
  .upload-title { font-family: 'Syne', sans-serif; font-size: 32px; font-weight: 800; color: #FFC828; text-align: center; }
  .upload-sub { font-size: 12px; color: rgba(221,217,207,0.38); text-align: center; line-height: 2; }
  .drop-box { width: 360px; height: 190px; border: 2px dashed rgba(255,200,40,0.28); border-radius: 10px; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 14px; cursor: pointer; transition: all 0.2s; position: relative; background: rgba(255,200,40,0.03); }
  .drop-box:hover, .drop-box.over { border-color: rgba(255,200,40,0.65); background: rgba(255,200,40,0.07); }
  .drop-box input { position: absolute; inset: 0; opacity: 0; cursor: pointer; }
  .drop-icon { font-size: 36px; }
  .drop-lbl { font-size: 12px; color: rgba(255,200,40,0.65); text-align: center; line-height: 1.7; }
  .drop-hint { font-size: 10px; color: rgba(221,217,207,0.28); }
  .dash { padding: 28px 40px 64px; display: flex; flex-direction: column; gap: 24px; }
  .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(185px, 1fr)); gap: 14px; }
  .kpi { background: rgba(255,255,255,0.025); border: 1px solid rgba(255,200,40,0.09); border-radius: 9px; padding: 20px 22px; position: relative; overflow: hidden; transition: border-color 0.2s; }
  .kpi:hover { border-color: rgba(255,200,40,0.25); }
  .kpi::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 2px; }
  .kpi.yellow::before { background: linear-gradient(90deg, #FFC828, transparent); }
  .kpi.orange::before { background: linear-gradient(90deg, #FF8C3A, transparent); }
  .kpi.teal::before   { background: linear-gradient(90deg, #4FD1A5, transparent); }
  .kpi.blue::before   { background: linear-gradient(90deg, #7B9FFF, transparent); }
  .kpi.red::before    { background: linear-gradient(90deg, #FF5757, transparent); }
  .kpi.purple::before { background: linear-gradient(90deg, #C084FC, transparent); }
  .kpi-lbl { font-size: 9px; letter-spacing: 2.5px; text-transform: uppercase; margin-bottom: 10px; }
  .kpi.yellow .kpi-lbl { color: rgba(255,200,40,0.45); }
  .kpi.orange .kpi-lbl { color: rgba(255,140,58,0.45); }
  .kpi.teal .kpi-lbl   { color: rgba(79,209,165,0.45); }
  .kpi.blue .kpi-lbl   { color: rgba(123,159,255,0.45); }
  .kpi.red .kpi-lbl    { color: rgba(255,87,87,0.45); }
  .kpi.purple .kpi-lbl { color: rgba(192,132,252,0.45); }
  .kpi-val { font-family: 'Syne', sans-serif; font-size: 24px; font-weight: 700; line-height: 1; }
  .kpi.yellow .kpi-val { color: #FFC828; }
  .kpi.orange .kpi-val { color: #FF8C3A; }
  .kpi.teal .kpi-val   { color: #4FD1A5; }
  .kpi.blue .kpi-val   { color: #7B9FFF; }
  .kpi.red .kpi-val    { color: #FF5757; }
  .kpi.purple .kpi-val { color: #C084FC; }
  .kpi-sub { font-size: 10px; color: rgba(221,217,207,0.3); margin-top: 6px; }
  .badges { display: flex; gap: 5px; flex-wrap: wrap; margin-top: 8px; }
  .bdg { padding: 2px 7px; border-radius: 20px; font-size: 9px; letter-spacing: 1px; }
  .bdg-sale  { background: rgba(255,200,40,0.1);  color: #FFC828; border: 1px solid rgba(255,200,40,0.2); }
  .bdg-dp    { background: rgba(123,159,255,0.1); color: #7B9FFF; border: 1px solid rgba(123,159,255,0.2); }
  .bdg-um    { background: rgba(79,209,165,0.1);  color: #4FD1A5; border: 1px solid rgba(79,209,165,0.2); }
  .bdg-batal { background: rgba(255,87,87,0.1);   color: #FF5757; border: 1px solid rgba(255,87,87,0.2); }
  .card { background: rgba(255,255,255,0.025); border: 1px solid rgba(255,200,40,0.09); border-radius: 9px; padding: 22px; }
  .card-title { font-family: 'Syne', sans-serif; font-size: 13px; font-weight: 700; color: #ddd9cf; margin-bottom: 18px; display: flex; align-items: baseline; gap: 8px; flex-wrap: wrap; }
  .card-title span { font-family: 'DM Mono', monospace; font-size: 10px; font-weight: 400; color: rgba(221,217,207,0.3); }
  .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; }
  .no-data { height: 200px; display: flex; align-items: center; justify-content: center; color: rgba(221,217,207,0.2); font-size: 12px; letter-spacing: 1px; }
  .tooltip-box { background: #18202e; border: 1px solid rgba(255,200,40,0.18); border-radius: 7px; padding: 9px 13px; font-size: 11px; font-family: 'DM Mono', monospace; color: #ddd9cf; }
  .tbl { width: 100%; border-collapse: collapse; font-size: 11px; }
  .tbl th { padding: 8px 12px; text-align: left; font-size: 9px; letter-spacing: 2px; text-transform: uppercase; color: rgba(255,200,40,0.4); border-bottom: 1px solid rgba(255,255,255,0.07); white-space: nowrap; }
  .tbl td { padding: 9px 12px; border-bottom: 1px solid rgba(255,255,255,0.03); color: rgba(221,217,207,0.75); }
  .tbl tr:last-child td { border-bottom: none; }
  .tbl tr:hover td { background: rgba(255,200,40,0.03); }
  .tbl .rank { color: rgba(255,200,40,0.35); font-size: 10px; width: 28px; }
  .tbl .num { text-align: right; font-variant-numeric: tabular-nums; }
  .tbl-wrap { max-height: 440px; overflow-y: auto; }
  .tbl-wrap::-webkit-scrollbar { width: 4px; }
  .tbl-wrap::-webkit-scrollbar-track { background: transparent; }
  .tbl-wrap::-webkit-scrollbar-thumb { background: rgba(255,200,40,0.2); border-radius: 2px; }
  .prog-wrap { display: flex; align-items: center; gap: 10px; }
  .prog-bar { flex: 1; height: 5px; background: rgba(255,255,255,0.06); border-radius: 3px; overflow: hidden; }
  .prog-fill { height: 100%; border-radius: 3px; }
  .prog-pct { font-size: 10px; color: rgba(221,217,207,0.4); width: 40px; text-align: right; }
  .type-tag { display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: 9px; letter-spacing: 1px; }
  .type-dp    { background: rgba(123,159,255,0.12); color: #7B9FFF; }
  .type-um    { background: rgba(79,209,165,0.12);  color: #4FD1A5; }
  .type-batal { background: rgba(255,87,87,0.12);   color: #FF5757; }
  .insight { background: rgba(255,255,255,0.02); border: 1px solid rgba(255,255,255,0.06); border-radius: 8px; padding: 18px 20px; }
  .insight-label { font-size: 9px; letter-spacing: 2px; text-transform: uppercase; color: rgba(221,217,207,0.3); margin-bottom: 8px; }
  .insight-val { font-family: 'Syne', sans-serif; font-size: 20px; font-weight: 700; margin-bottom: 4px; }
  .insight-sub { font-size: 10px; color: rgba(221,217,207,0.35); line-height: 1.6; }
  .warn-box { margin-bottom: 14px; padding: 10px 14px; background: rgba(255,87,87,0.05); border: 1px solid rgba(255,87,87,0.15); border-radius: 7px; font-size: 11px; color: rgba(255,87,87,0.85); line-height: 1.7; }
  @media (max-width: 900px) {
    .two-col { grid-template-columns: 1fr; }
    .dash { padding: 18px 14px 60px; }
    .hdr { padding: 18px 14px 0; }
    .drop-box { width: 88vw; }
    .tab { padding: 10px 14px; font-size: 10px; }
  }
`;

const fmt = (n) =>
  n >= 1e9 ? `Rp ${(n/1e9).toFixed(2)}M`
  : n >= 1e6 ? `Rp ${(n/1e6).toFixed(1)}Jt`
  : `Rp ${Math.round(n).toLocaleString("id-ID")}`;

const fmtShort = (n) =>
  n >= 1e9 ? `${(n/1e9).toFixed(1)}M`
  : n >= 1e6 ? `${(n/1e6).toFixed(0)}Jt`
  : n >= 1e3 ? `${(n/1e3).toFixed(0)}k`
  : n;

const fmtKg = (n) => `${Math.round(n).toLocaleString("id-ID")} kg`;
const COLORS = ["#FFC828","#FF8C3A","#4FD1A5","#7B9FFF","#FF5757","#C084FC","#38BDF8","#F472B6","#A3E635","#FB923C"];

function extractPeriod(filename) {
  const months = ["JANUARI","FEBRUARI","MARET","APRIL","MEI","JUNI","JULI","AGUSTUS","SEPTEMBER","OKTOBER","NOVEMBER","DESEMBER"];
  const base = filename.replace(/\.xlsx?$/i,"").toUpperCase();
  const parts = base.split("_").filter(Boolean).reverse();
  let month = "", year = "";
  for (const p of parts) {
    if (!year && /^20\d{2}$/.test(p)) year = p;
    if (!month && months.includes(p)) month = p;
    if (month && year) break;
  }
  return [month, year].filter(Boolean).join(" ") || base.replace(/_+/g," ").trim();
}

function parsePPN(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header:1, defval:null });
  return rows.slice(1)
    .filter(r => r[0] && String(r[0]).startsWith("FK"))
    .map(r => {
      const dr = r[2];
      let date = dr instanceof Date ? dr
               : typeof dr === "number" ? new Date(Math.round((dr-25569)*86400000))
               : new Date(dr);
      const total    = parseFloat(r[4]) || 0;
      const ppn      = total * 0.11;
      const totalInv = total * 1.11;
      return {
        invNo: String(r[0]),
        customer: String(r[1]||"").trim(),
        date: isNaN(date?.getTime()) ? null : date,
        qty: parseFloat(r[3]) || 0,
        total, ppn, totalInv,
        txType: String(r[7]||"Sale").trim(),
      };
    });
}

const CT = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div className="tooltip-box">
      <div style={{ color:"#FFC828", marginBottom:4, fontSize:10 }}>{label}</div>
      {payload.map((p,i) => (
        <div key={i} style={{ color:p.color||"#ddd9cf" }}>
          {p.name}: {typeof p.value==="number" && p.value>10000 ? fmt(p.value) : (p.value?.toLocaleString?.("id-ID")??p.value)}
        </div>
      ))}
    </div>
  );
};

const ProgRow = ({ label, value, max, color, fmtFn }) => (
  <div style={{ marginBottom:10 }}>
    <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4 }}>
      <span style={{ fontSize:11, color:"rgba(221,217,207,0.7)", maxWidth:"60%", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{label}</span>
      <span style={{ fontSize:11, color }}>{fmtFn ? fmtFn(value) : value}</span>
    </div>
    <div className="prog-wrap">
      <div className="prog-bar"><div className="prog-fill" style={{ width:`${Math.min(100,(value/max)*100)}%`, background:color }} /></div>
      <span className="prog-pct">{((value/max)*100).toFixed(1)}%</span>
    </div>
  </div>
);

export default function App() {
  const [data, setData] = useState(null);
  const [drag, setDrag] = useState(false);
  const [tab,  setTab]  = useState("overview");

  const load = useCallback((file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type:"array", cellDates:true });
      let rows = [];
      wb.SheetNames.forEach(n => { if (n.toUpperCase()==="PPN") rows = parsePPN(wb.Sheets[n]); });
      setData({ rows, period: extractPeriod(file.name) });
      setTab("overview");
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const onDrop = useCallback((e) => {
    e.preventDefault(); setDrag(false);
    load(e.dataTransfer?.files[0] || e.target?.files[0]);
  }, [load]);

  const s = useMemo(() => {
    if (!data) return null;
    const { rows } = data;
    const sales    = rows.filter(r => r.txType==="Sale");
    const nonSales = rows.filter(r => r.txType!=="Sale");

    const totalRev  = sales.reduce((a,r) => a+r.totalInv, 0);
    const totalBase = sales.reduce((a,r) => a+r.total, 0);
    const totalQty  = sales.reduce((a,r) => a+r.qty, 0);
    const totalPPN  = sales.reduce((a,r) => a+r.ppn, 0);
    const avgInv    = totalRev / (sales.length||1);
    const revenuePerKg = totalQty>0 ? totalBase/totalQty : 0;

    const typeCounts = {};
    rows.forEach(r => { typeCounts[r.txType]=(typeCounts[r.txType]||0)+1; });
    const typeSlices = Object.entries(typeCounts).map(([name,value])=>({ name, value }));

    const custMap = {};
    sales.forEach(r => {
      if (!custMap[r.customer]) custMap[r.customer]={ revenue:0, qty:0, count:0 };
      custMap[r.customer].revenue += r.totalInv;
      custMap[r.customer].qty     += r.qty;
      custMap[r.customer].count   += 1;
    });
    const custArr = Object.entries(custMap).map(([name,v])=>({ name,...v }));
    const topByRevenue = [...custArr].sort((a,b)=>b.revenue-a.revenue).slice(0,10);
    const topByQty     = [...custArr].sort((a,b)=>b.qty-a.qty).slice(0,10);
    const topByOrders  = [...custArr].sort((a,b)=>b.count-a.count).slice(0,10);

    const dayMap = {};
    sales.forEach(r => {
      if (!r.date) return;
      const k = String(r.date.getDate()).padStart(2,"0");
      if (!dayMap[k]) dayMap[k]={ date:k, revenue:0, qty:0, _d:r.date.getDate() };
      dayMap[k].revenue += r.totalInv;
      dayMap[k].qty     += r.qty;
    });
    const dailyTrend = Object.values(dayMap).sort((a,b)=>a._d-b._d);

    const top5rev = topByRevenue.slice(0,5).reduce((a,r)=>a+r.revenue,0);
    const top5pct = totalRev>0 ? (top5rev/totalRev)*100 : 0;
    const top1pct = totalRev>0 ? ((topByRevenue[0]?.revenue||0)/totalRev)*100 : 0;
    const biggestInv = [...sales].sort((a,b)=>b.totalInv-a.totalInv)[0];
    const activeDays = new Set(sales.filter(r=>r.date).map(r=>r.date.getDate())).size;
    const dailyAvg   = activeDays>0 ? totalRev/activeDays : 0;
    const bestDay    = dailyTrend.reduce((a,b)=>b.revenue>a.revenue?b:a, dailyTrend[0]||{});

    const buckets = { "<10Jt":0,"10–50Jt":0,"50–100Jt":0,"100–500Jt":0,">500Jt":0 };
    sales.forEach(r => {
      const v=r.totalInv;
      if (v<10e6) buckets["<10Jt"]++;
      else if (v<50e6) buckets["10–50Jt"]++;
      else if (v<100e6) buckets["50–100Jt"]++;
      else if (v<500e6) buckets["100–500Jt"]++;
      else buckets[">500Jt"]++;
    });
    const bucketData = Object.entries(buckets).map(([name,value])=>({ name, value }));

    return {
      totalRev, totalBase, totalQty, totalPPN, avgInv, revenuePerKg,
      invoiceCount:sales.length, custCount:custArr.length,
      typeCounts, typeSlices,
      topByRevenue, topByQty, topByOrders,
      dailyTrend, activeDays, dailyAvg, bestDay,
      top5pct, top1pct, biggestInv, bucketData,
      nonSales, custArr,
    };
  }, [data]);

  const renderOverview = () => (
    <>
      <div className="kpi-grid">
        <div className="kpi yellow"><div className="kpi-lbl">Total Pendapatan</div><div className="kpi-val">{fmt(s.totalRev)}</div><div className="kpi-sub">Invoice Sale · incl. PPN</div></div>
        <div className="kpi orange"><div className="kpi-lbl">Total PPN Dipungut</div><div className="kpi-val">{fmt(s.totalPPN)}</div><div className="kpi-sub">PPN 11% dari semua Sale</div></div>
        <div className="kpi teal"><div className="kpi-lbl">Total Volume</div><div className="kpi-val">{fmtKg(s.totalQty)}</div><div className="kpi-sub">Total berat (Sale only)</div></div>
        <div className="kpi blue">
          <div className="kpi-lbl">Jumlah Invoice Sale</div>
          <div className="kpi-val">{s.invoiceCount.toLocaleString()}</div>
          <div className="badges">
            {Object.entries(s.typeCounts).map(([t,n]) => (
              <span key={t} className={`bdg ${t==="Sale"?"bdg-sale":t==="DP"?"bdg-dp":t.includes("UANG")?"bdg-um":"bdg-batal"}`}>{t.trim()} {n}</span>
            ))}
          </div>
        </div>
        <div className="kpi yellow"><div className="kpi-lbl">Jumlah Customer</div><div className="kpi-val">{s.custCount}</div><div className="kpi-sub">Customer unik periode ini</div></div>
        <div className="kpi teal"><div className="kpi-lbl">Rata-rata per Invoice</div><div className="kpi-val">{fmt(s.avgInv)}</div><div className="kpi-sub">Nilai rata-rata per Sale</div></div>
      </div>

      <div className="card">
        <div className="card-title">Tren Pendapatan Harian — {data.period} <span>per tanggal · Sale only</span></div>
        <ResponsiveContainer width="100%" height={220}>
          <AreaChart data={s.dailyTrend} margin={{ top:5, right:5, bottom:5, left:10 }}>
            <defs><linearGradient id="ag" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#FFC828" stopOpacity={0.22}/><stop offset="95%" stopColor="#FFC828" stopOpacity={0}/></linearGradient></defs>
            <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)"/>
            <XAxis dataKey="date" tick={{ fill:"rgba(221,217,207,0.35)", fontSize:11 }} axisLine={false} tickLine={false}/>
            <YAxis tick={{ fill:"rgba(221,217,207,0.35)", fontSize:10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort}/>
            <Tooltip content={<CT/>}/>
            <Area type="monotone" dataKey="revenue" name="Pendapatan" stroke="#FFC828" strokeWidth={2} fill="url(#ag)" dot={{ fill:"#FFC828", r:3 }} activeDot={{ r:5 }}/>
          </AreaChart>
        </ResponsiveContainer>
      </div>

      <div className="two-col">
        <div className="card">
          <div className="card-title">Top 10 Customer <span>by Revenue (Rp)</span></div>
          <ResponsiveContainer width="100%" height={300}>
            <BarChart data={s.topByRevenue} layout="vertical" margin={{ top:0, right:8, left:8, bottom:0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" horizontal={false}/>
              <XAxis type="number" tick={{ fill:"rgba(221,217,207,0.35)", fontSize:10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort}/>
              <YAxis type="category" dataKey="name" width={148} tick={{ fill:"rgba(221,217,207,0.6)", fontSize:10 }} axisLine={false} tickLine={false}/>
              <Tooltip content={<CT/>}/>
              <Bar dataKey="revenue" name="Revenue" radius={[0,4,4,0]}>{s.topByRevenue.map((_,i)=><Cell key={i} fill={COLORS[i%COLORS.length]}/>)}</Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div className="card">
          <div className="card-title">Top 10 Customer <span>by Volume (KG)</span></div>
          <ResponsiveContainer width="100%" height={300}>
            <BarChart data={s.topByQty} layout="vertical" margin={{ top:0, right:8, left:8, bottom:0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" horizontal={false}/>
              <XAxis type="number" tick={{ fill:"rgba(221,217,207,0.35)", fontSize:10 }} axisLine={false} tickLine={false} tickFormatter={v=>v>=1000?`${(v/1000).toFixed(0)}k`:v}/>
              <YAxis type="category" dataKey="name" width={148} tick={{ fill:"rgba(221,217,207,0.6)", fontSize:10 }} axisLine={false} tickLine={false}/>
              <Tooltip content={<CT/>}/>
              <Bar dataKey="qty" name="Qty (kg)" radius={[0,4,4,0]}>{s.topByQty.map((_,i)=><Cell key={i} fill={COLORS[(i+2)%COLORS.length]}/>)}</Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="two-col">
        <div className="card">
          <div className="card-title">Komposisi Tipe Invoice <span>semua transaksi</span></div>
          <ResponsiveContainer width="100%" height={180}>
            <PieChart>
              <Pie data={s.typeSlices} cx="50%" cy="50%" innerRadius={52} outerRadius={78} paddingAngle={4} dataKey="value">
                {s.typeSlices.map((_,i)=><Cell key={i} fill={COLORS[i]}/>)}
              </Pie>
              <Tooltip content={<CT/>}/>
              <Legend formatter={v=><span style={{ color:"rgba(221,217,207,0.55)", fontSize:11 }}>{v}</span>}/>
            </PieChart>
          </ResponsiveContainer>
          <table className="tbl" style={{ marginTop:10 }}>
            <thead><tr><th>Tipe</th><th>Jumlah</th><th>%</th></tr></thead>
            <tbody>{s.typeSlices.map((t,i)=><tr key={i}><td style={{ color:COLORS[i] }}>{t.name}</td><td>{t.value}</td><td>{((t.value/data.rows.length)*100).toFixed(1)}%</td></tr>)}</tbody>
          </table>
        </div>
        <div className="card">
          <div className="card-title">Tren Volume Harian <span>KG · Sale only</span></div>
          <ResponsiveContainer width="100%" height={200}>
            <AreaChart data={s.dailyTrend} margin={{ top:5, right:5, bottom:5, left:10 }}>
              <defs><linearGradient id="qg" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#4FD1A5" stopOpacity={0.22}/><stop offset="95%" stopColor="#4FD1A5" stopOpacity={0}/></linearGradient></defs>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)"/>
              <XAxis dataKey="date" tick={{ fill:"rgba(221,217,207,0.35)", fontSize:10 }} axisLine={false} tickLine={false}/>
              <YAxis tick={{ fill:"rgba(221,217,207,0.35)", fontSize:10 }} axisLine={false} tickLine={false} tickFormatter={v=>v>=1000?`${(v/1000).toFixed(0)}k`:v}/>
              <Tooltip content={<CT/>}/>
              <Area type="monotone" dataKey="qty" name="Volume (kg)" stroke="#4FD1A5" strokeWidth={2} fill="url(#qg)" dot={{ fill:"#4FD1A5", r:3 }} activeDot={{ r:5 }}/>
            </AreaChart>
          </ResponsiveContainer>
        </div>
      </div>
    </>
  );

  const renderCustomers = () => (
    <>
      <div className="card">
        <div className="card-title">Top 10 Customer <span>by Jumlah Order (Sale only)</span></div>
        <div className="tbl-wrap">
          <table className="tbl">
            <thead><tr><th>#</th><th>Customer</th><th className="num">Jumlah Order</th><th className="num">Total Revenue</th><th className="num">Total Volume (kg)</th><th className="num">Avg per Order</th></tr></thead>
            <tbody>
              {s.topByOrders.map((c,i) => (
                <tr key={i}>
                  <td className="rank">{i+1}</td>
                  <td style={{ color:COLORS[i%COLORS.length] }}>{c.name}</td>
                  <td className="num">{c.count}</td>
                  <td className="num">{fmt(c.revenue)}</td>
                  <td className="num">{Math.round(c.qty).toLocaleString("id-ID")} kg</td>
                  <td className="num">{fmt(c.revenue/c.count)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div className="card">
        <div className="card-title">Order Frequency Chart <span>top 10 customer · frekuensi pembelian</span></div>
        <ResponsiveContainer width="100%" height={280}>
          <BarChart data={s.topByOrders} layout="vertical" margin={{ top:0, right:8, left:8, bottom:0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)" horizontal={false}/>
            <XAxis type="number" tick={{ fill:"rgba(221,217,207,0.35)", fontSize:10 }} axisLine={false} tickLine={false} allowDecimals={false}/>
            <YAxis type="category" dataKey="name" width={155} tick={{ fill:"rgba(221,217,207,0.6)", fontSize:10 }} axisLine={false} tickLine={false}/>
            <Tooltip content={<CT/>}/>
            <Bar dataKey="count" name="Jumlah Order" radius={[0,4,4,0]}>{s.topByOrders.map((_,i)=><Cell key={i} fill={COLORS[i%COLORS.length]}/>)}</Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className="two-col">
        <div className="card">
          <div className="card-title">Top 10 by Revenue <span>Rp · incl. PPN</span></div>
          <div className="tbl-wrap">
            <table className="tbl">
              <thead><tr><th>#</th><th>Customer</th><th className="num">Revenue</th><th className="num">Share</th></tr></thead>
              <tbody>
                {s.topByRevenue.map((c,i) => (
                  <tr key={i}>
                    <td className="rank">{i+1}</td>
                    <td style={{ color:COLORS[i%COLORS.length] }}>{c.name}</td>
                    <td className="num">{fmt(c.revenue)}</td>
                    <td className="num">
                      <div className="prog-wrap">
                        <div className="prog-bar"><div className="prog-fill" style={{ width:`${(c.revenue/s.totalRev)*100}%`, background:COLORS[i%COLORS.length] }}/></div>
                        <span className="prog-pct">{((c.revenue/s.totalRev)*100).toFixed(1)}%</span>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
        <div className="card">
          <div className="card-title">Top 10 by Volume <span>KG</span></div>
          <div className="tbl-wrap">
            <table className="tbl">
              <thead><tr><th>#</th><th>Customer</th><th className="num">Volume (kg)</th><th className="num">Share</th></tr></thead>
              <tbody>
                {s.topByQty.map((c,i) => (
                  <tr key={i}>
                    <td className="rank">{i+1}</td>
                    <td style={{ color:COLORS[(i+2)%COLORS.length] }}>{c.name}</td>
                    <td className="num">{Math.round(c.qty).toLocaleString("id-ID")}</td>
                    <td className="num">
                      <div className="prog-wrap">
                        <div className="prog-bar"><div className="prog-fill" style={{ width:`${(c.qty/s.totalQty)*100}%`, background:COLORS[(i+2)%COLORS.length] }}/></div>
                        <span className="prog-pct">{((c.qty/s.totalQty)*100).toFixed(1)}%</span>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </>
  );

  const renderNonSale = () => {
    const batal = s.nonSales.filter(r=>r.txType==="BATAL");
    const dp    = s.nonSales.filter(r=>r.txType==="DP");
    const um    = s.nonSales.filter(r=>r.txType.includes("UANG"));
    const totalDPval = dp.reduce((a,r)=>a+r.total,0);
    const totalUMval = um.reduce((a,r)=>a+r.total,0);

    const fmtDate = (d) => d ? `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}` : "—";

    const TypeTable = ({ rows, color, cls }) => (
      <div className="tbl-wrap">
        <table className="tbl">
          <thead><tr><th>No. Invoice</th><th>Customer</th><th>Tipe</th><th className="num">Total (Rp)</th><th>Tanggal</th></tr></thead>
          <tbody>
            {rows.length===0
              ? <tr><td colSpan={5} style={{ textAlign:"center", color:"rgba(221,217,207,0.2)", padding:24 }}>— Tidak ada data —</td></tr>
              : rows.map((r,i)=>(
                <tr key={i}>
                  <td style={{ color, fontFamily:"'DM Mono',monospace", fontSize:10 }}>{r.invNo}</td>
                  <td>{r.customer}</td>
                  <td><span className={`type-tag ${cls}`}>{r.txType}</span></td>
                  <td className="num">{r.total>0?fmt(r.total):"—"}</td>
                  <td>{fmtDate(r.date)}</td>
                </tr>
              ))
            }
          </tbody>
        </table>
      </div>
    );

    return (
      <>
        <div className="kpi-grid">
          <div className="kpi red"><div className="kpi-lbl">Invoice BATAL</div><div className="kpi-val">{batal.length}</div><div className="kpi-sub">Transaksi dibatalkan</div></div>
          <div className="kpi blue"><div className="kpi-lbl">Down Payment (DP)</div><div className="kpi-val">{dp.length}</div><div className="kpi-sub">{totalDPval>0?fmt(totalDPval):"Nilai belum tercatat"}</div></div>
          <div className="kpi teal"><div className="kpi-lbl">Uang Muka</div><div className="kpi-val">{um.length}</div><div className="kpi-sub">{totalUMval>0?fmt(totalUMval):"Nilai belum tercatat"}</div></div>
        </div>
        <div className="card"><div className="card-title">Down Payment (DP) <span>{dp.length} transaksi</span></div><TypeTable rows={dp} color="#7B9FFF" cls="type-dp"/></div>
        <div className="card"><div className="card-title">Uang Muka <span>{um.length} transaksi</span></div><TypeTable rows={um} color="#4FD1A5" cls="type-um"/></div>
        <div className="card"><div className="card-title">Invoice BATAL <span>{batal.length} transaksi</span></div><TypeTable rows={batal} color="#FF5757" cls="type-batal"/></div>
      </>
    );
  };

  const renderAnalytics = () => {
    const one   = s.custArr.filter(c=>c.count===1).length;
    const two5  = s.custArr.filter(c=>c.count>=2&&c.count<=5).length;
    const six10 = s.custArr.filter(c=>c.count>=6&&c.count<=10).length;
    const over10= s.custArr.filter(c=>c.count>10).length;

    return (
      <>
        <div className="kpi-grid">
          <div className="kpi yellow"><div className="kpi-lbl">Revenue / Hari Aktif</div><div className="kpi-val">{fmt(s.dailyAvg)}</div><div className="kpi-sub">{s.activeDays} hari aktif bulan ini</div></div>
          <div className="kpi purple"><div className="kpi-lbl">Harga Rata-rata / kg</div><div className="kpi-val">Rp {Math.round(s.revenuePerKg).toLocaleString("id-ID")}</div><div className="kpi-sub">Sebelum PPN · pricing efficiency</div></div>
          <div className="kpi orange"><div className="kpi-lbl">Konsentrasi Top 5</div><div className="kpi-val">{s.top5pct.toFixed(1)}%</div><div className="kpi-sub">Share revenue dari top 5 customer</div></div>
          <div className="kpi red"><div className="kpi-lbl">Konsentrasi #1 Customer</div><div className="kpi-val">{s.top1pct.toFixed(1)}%</div><div className="kpi-sub">{s.topByRevenue[0]?.name}</div></div>
          <div className="kpi teal"><div className="kpi-lbl">Invoice Terbesar</div><div className="kpi-val">{fmt(s.biggestInv?.totalInv||0)}</div><div className="kpi-sub">{s.biggestInv?.customer}</div></div>
          <div className="kpi blue"><div className="kpi-lbl">Hari Terbaik</div><div className="kpi-val">Tgl {s.bestDay?.date}</div><div className="kpi-sub">{fmt(s.bestDay?.revenue||0)}</div></div>
        </div>

        {/* Revenue concentration */}
        <div className="card">
          <div className="card-title">Revenue Concentration Risk <span>% share tiap customer · CEO/CFO view</span></div>
          <div className="warn-box">
            ⚠ Top 5 customer menyumbang <strong style={{ color:"#FF5757" }}>{s.top5pct.toFixed(1)}%</strong> dari total revenue.
            {s.top5pct>60?" Risiko konsentrasi TINGGI — ketergantungan berlebihan pada segelintir pelanggan.":s.top5pct>40?" Risiko konsentrasi SEDANG — pertimbangkan diversifikasi.":"  Risiko konsentrasi RENDAH — distribusi customer cukup sehat."}
          </div>
          {s.topByRevenue.slice(0,8).map((c,i) => (
            <ProgRow key={i} label={c.name} value={c.revenue} max={s.totalRev} color={COLORS[i%COLORS.length]} fmtFn={fmt}/>
          ))}
        </div>

        <div className="two-col">
          {/* Invoice size distribution */}
          <div className="card">
            <div className="card-title">Distribusi Ukuran Invoice <span>jumlah invoice per bracket nilai</span></div>
            <ResponsiveContainer width="100%" height={230}>
              <BarChart data={s.bucketData} margin={{ top:5, right:5, bottom:5, left:5 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)"/>
                <XAxis dataKey="name" tick={{ fill:"rgba(221,217,207,0.45)", fontSize:10 }} axisLine={false} tickLine={false}/>
                <YAxis tick={{ fill:"rgba(221,217,207,0.35)", fontSize:10 }} axisLine={false} tickLine={false} allowDecimals={false}/>
                <Tooltip content={<CT/>}/>
                <Bar dataKey="value" name="Jumlah Invoice" radius={[4,4,0,0]}>{s.bucketData.map((_,i)=><Cell key={i} fill={COLORS[i%COLORS.length]}/>)}</Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>

          {/* Revenue per day bar */}
          <div className="card">
            <div className="card-title">Revenue per Hari <span>hari terbaik disorot kuning</span></div>
            <ResponsiveContainer width="100%" height={230}>
              <BarChart data={s.dailyTrend} margin={{ top:5, right:5, bottom:5, left:5 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.04)"/>
                <XAxis dataKey="date" tick={{ fill:"rgba(221,217,207,0.35)", fontSize:10 }} axisLine={false} tickLine={false}/>
                <YAxis tick={{ fill:"rgba(221,217,207,0.35)", fontSize:10 }} axisLine={false} tickLine={false} tickFormatter={fmtShort}/>
                <Tooltip content={<CT/>}/>
                <Bar dataKey="revenue" name="Revenue" radius={[3,3,0,0]}>
                  {s.dailyTrend.map((d,i)=><Cell key={i} fill={d.date===s.bestDay?.date?"#FFC828":"rgba(255,200,40,0.3)"}/>)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Order frequency analysis */}
        <div className="card">
          <div className="card-title">Customer Loyalty Analysis <span>seberapa sering customer repeat order</span></div>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))", gap:14, marginBottom:22 }}>
            {[
              { label:"One-time (1x order)", val:one,   color:"#FF5757" },
              { label:"Repeat (2–5x)",       val:two5,  color:"#FF8C3A" },
              { label:"Loyal (6–10x)",       val:six10, color:"#FFC828" },
              { label:"VIP (>10x)",          val:over10,color:"#4FD1A5" },
            ].map((item,i) => (
              <div key={i} className="insight">
                <div className="insight-label">{item.label}</div>
                <div className="insight-val" style={{ color:item.color }}>{item.val}</div>
                <div className="insight-sub">customer · {s.custArr.length>0?((item.val/s.custArr.length)*100).toFixed(1):"0"}% dari total</div>
              </div>
            ))}
          </div>
          <div className="tbl-wrap" style={{ maxHeight:300 }}>
            <table className="tbl">
              <thead><tr><th>#</th><th>Customer</th><th className="num">Order</th><th className="num">Revenue</th><th className="num">Avg/Order</th><th className="num">Rp/kg (excl. PPN)</th></tr></thead>
              <tbody>
                {s.topByOrders.map((c,i) => (
                  <tr key={i}>
                    <td className="rank">{i+1}</td>
                    <td style={{ color:COLORS[i%COLORS.length] }}>{c.name}</td>
                    <td className="num">{c.count}</td>
                    <td className="num">{fmt(c.revenue)}</td>
                    <td className="num">{fmt(c.revenue/c.count)}</td>
                    <td className="num">Rp {c.qty>0?Math.round((c.revenue/1.11)/c.qty).toLocaleString("id-ID"):"—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </>
    );
  };

  return (
    <>
      <style>{css}</style>
      <div className="app">
        <header className="hdr">
          <div className="hdr-top">
            <div className="hdr-left">
              <h1>⬡ PT AMERTA NIAGATAMA</h1>
              <p>Sales Intelligence Dashboard</p>
            </div>
            <div className="hdr-right">
              {data && <div className="period-pill">📅 {data.period}</div>}
              {data && (
                <button className="swap-btn">↑ GANTI FILE
                  <input type="file" accept=".xlsx,.xls" onChange={onDrop}/>
                </button>
              )}
            </div>
          </div>
          {data && (
            <div className="tabs">
              {[
                { id:"overview",   label:"Overview"   },
                { id:"customers",  label:"Customers"  },
                { id:"nonsale",    label:"DP / Batal" },
                { id:"analytics",  label:"Analytics"  },
              ].map(t => (
                <button key={t.id} className={`tab ${tab===t.id?"active":""}`} onClick={()=>setTab(t.id)}>{t.label}</button>
              ))}
            </div>
          )}
        </header>

        {!data ? (
          <div className="upload-screen">
            <div className="upload-title">DROP YOUR INVOICE FILE</div>
            <p className="upload-sub">Upload file Excel invoice bulanan<br/>Dashboard otomatis terbentuk seketika</p>
            <div className={`drop-box ${drag?"over":""}`}
              onDragOver={(e)=>{ e.preventDefault(); setDrag(true); }}
              onDragLeave={()=>setDrag(false)}
              onDrop={onDrop}>
              <input type="file" accept=".xlsx,.xls" onChange={onDrop}/>
              <div className="drop-icon">📊</div>
              <div className="drop-lbl">Drag & drop file Excel di sini<br/>atau klik untuk pilih file</div>
              <div className="drop-hint">.xlsx · .xls</div>
            </div>
          </div>
        ) : (
          <div className="dash">
            {tab==="overview"  && renderOverview()}
            {tab==="customers" && renderCustomers()}
            {tab==="nonsale"   && renderNonSale()}
            {tab==="analytics" && renderAnalytics()}
          </div>
        )}
      </div>
    </>
  );
}
