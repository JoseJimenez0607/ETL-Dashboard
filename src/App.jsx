import { useState, useRef } from "react";
import * as XLSX from "xlsx";
import {
  AreaChart, Area, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer
} from "recharts";

// ─── TOKENS ───────────────────────────────────────────────────────────────────
const T = {
  bg:"#080C14", surface:"#0F1623", surf2:"#162032",
  border:"#1E3050", accent:"#3B82F6", success:"#10B981",
  warning:"#F59E0B", danger:"#EF4444", purple:"#8B5CF6", cyan:"#06B6D4",
  text:"#E2E8F0", muted:"#4B6A8A", sub:"#7A9CBC",
};
const PAL = ["#3B82F6","#10B981","#F59E0B","#8B5CF6","#EF4444","#06B6D4","#F97316","#EC4899","#14B8A6","#A855F7"];
const tip = { background:"#0B0F19", border:"1px solid #1E3050", color:"#E2E8F0", fontSize:12, borderRadius:6, padding:"8px 12px" };
const card = (x={}) => ({ background:T.surface, border:`1px solid ${T.border}`, borderRadius:10, ...x });

// ─── EXCEL READER ─────────────────────────────────────────────────────────────
function readSheets(workbook) {
  const result = [];
  workbook.SheetNames.forEach(sheetName => {
    const ws = workbook.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
    if (!raw || raw.length < 2) return;

    const row0 = raw[0] || [];
    const nullRatio = row0.filter(v => v == null).length / Math.max(row0.length, 1);
    const looksLikeTitle = nullRatio > 0.4 ||
      (typeof row0[0] === "string" && row0[0].length > 20 &&
       row0[0] === row0[0].toUpperCase() && /[A-ZÁÉÍÓÚ]{5,}/.test(row0[0]));

    const headerRowIdx = looksLikeTitle ? 1 : 0;
    const headers = (raw[headerRowIdx] || []).map(h => h != null ? String(h).trim() : null);
    if (headers.filter(Boolean).length < 2) return;
    if (sheetName.toLowerCase().includes("kpi")) return;

    const rows = [];
    for (let i = headerRowIdx + 1; i < raw.length; i++) {
      const r = raw[i] || [];
      if (r.every(v => v == null || v === "")) continue;
      const obj = {};
      headers.forEach((h, ci) => { if (h) obj[h] = r[ci] != null ? r[ci] : null; });
      if (Object.values(obj).every(v => v == null)) continue;
      rows.push(obj);
    }
    if (!rows.length) return;
    const titleText = looksLikeTitle && row0[0] ? String(row0[0]) : null;
    result.push({ name:sheetName, rows, headers:headers.filter(Boolean), titleText, looksLikeTitle });
  });
  return result;
}

// ─── COMPANY NAME ─────────────────────────────────────────────────────────────
function extractCompany(sheets, fileName) {
  for (const s of sheets) {
    if (!s.titleText) continue;
    const m = s.titleText.match(/dashboard\s+kpis?\s*[—–-]\s*(.+)/i);
    if (m) return m[1].replace(/\s+s\.a\..*$/i,"").replace(/\s+ltda\..*$/i,"").trim();
    const m2 = s.titleText.match(/[—–|]\s*(.{4,}?)(?:\s+s\.a\.|\s+ltda\.?|\s+spa\.?)?$/i);
    if (m2) return m2[1].trim();
  }
  // From filename: "base_datos_seguros_vida" → "Seguros Vida"
  return fileName
    .replace(/\.(xlsx?|csv)$/i, "")
    .replace(/^(base[_\s]datos?[_\s]|datos?[_\s]|reporte[_\s]|ventas[_\s])/i, "")
    .replace(/_/g, " ")
    .replace(/\b\w/g, c => c.toUpperCase())
    .trim();
}

function pickMain(sheets) {
  return sheets.reduce((best, s) =>
    (s.rows.length * s.headers.length) > (best.rows.length * best.headers.length) ? s : best
  , sheets[0]);
}

// ─── VALUE HELPERS ────────────────────────────────────────────────────────────
function isDateString(v) {
  if (v == null) return false;
  const s = String(v).trim();
  // ISO: 2025-09-07 or 2025-09-07T00:00:00
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return true;
  // DD/MM/YYYY or MM/DD/YYYY
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) return true;
  // DD-MM-YYYY
  if (/^\d{1,2}-\d{2}-\d{4}$/.test(s)) return true;
  return false;
}

function looksNumeric(v) {
  if (typeof v === "number") return true;
  if (typeof v === "string") return !isNaN(parseFloat(v)) && isFinite(v.trim());
  return false;
}

// ─── SMART CLASSIFIER ─────────────────────────────────────────────────────────
// Priority: values first → name pattern → category fallback
function classifyColumn(name, samples) {
  const n = String(name).toLowerCase().trim();
  const nonNull = samples.filter(v => v != null && v !== "");
  if (!nonNull.length) return "text";

  // ── 1. HARD ID rules by name ──
  if (/^(policy_id|id\b|n[°º]|folio|vin\b|rut\b|patente|chasis|cod[_\s]|#\s)/i.test(name)) return "id";
  if (/^(policy_id|claim_id|order_id|customer_id|employee_id|product_id)/i.test(n)) return "id";

  // ── 2. Detect dates by VALUE PATTERN (handles ISO, dd/mm/yyyy, etc.) ──
  const dateSamples = nonNull.slice(0, 20).filter(v => isDateString(v));
  if (dateSamples.length / Math.min(nonNull.length, 20) > 0.7) return "date";

  // ── 3. Name columns — detect by name pattern BEFORE category check ──
  // because names have low cardinality but should not be used as categories
  if (/(^(first|last|full)[_\s]?name$|^(nombre|apellido)$|^sales[_\s]agent$|^(agente|vendedor|ejecutivo|corredor)[_\s]?(nombre|name)?$|^nombre[_\s](agente|vendedor|cliente|ejecutivo))/i.test(n)) return "name";

  // ── 4. Check numeric ratio ──
  const numericVals = nonNull.filter(looksNumeric);
  const numericRatio = numericVals.length / nonNull.length;
  const uniqueVals = new Set(nonNull.map(v => String(v).trim()));
  const uniqueCount = uniqueVals.size;

  // ── 5. Clearly categorical: no numbers, low cardinality, not a name/date ──
  if (numericRatio < 0.1 && uniqueCount <= 30 && nonNull.length >= 3) {
    // Extra guard: if all values look like proper names (2 words, capitalized) → name
    const namePattern = nonNull.slice(0,10).filter(v => /^[A-ZÁÉÍÓÚ][a-záéíóú]+ [A-ZÁÉÍÓÚ][a-záéíóú]+$/.test(String(v)));
    if (namePattern.length / Math.min(nonNull.length,10) > 0.6) return "name";
    return "category";
  }

  // ── 6. Numeric columns — classify by name ──
  if (numericRatio > 0.7) {
    // Money — bilingual patterns
    if (/(_clp|_usd|_eur|\(clp\)|\(usd\))/i.test(n)) return "money";
    if (/(amount|revenue|income|salary|price|cost|fee|payment|premium|commission|margin|profit|total|monto|valor|prima|precio|comisi|margen|costo|ingreso|saldo|cuota)/i.test(n)) return "money";
    // Percentage
    if (/(rate|ratio|percent|_%|porcentaje|tasa|relaci)/i.test(n)) return "percent";
    // Score / rating
    if (/(score|rating|satisfaction|satisfac|nps\b|punt)/i.test(n)) return "score";
    // Count / age / plain
    if (/(age\b|edad\b|plazo|días|days|calls|llamadas|cuotas|visits|count\b|cantidad)/i.test(n)) return "count";
    return "numeric";
  }

  // ── 7. Remaining text ──
  if (/(name|nombre)\b/i.test(n)) return "name";
  return "text";
}

// ─── ANALYZE ──────────────────────────────────────────────────────────────────
function analyze(rows, headers) {
  const SAMPLE = Math.min(100, rows.length);
  const colTypes = {};
  headers.forEach(h => {
    colTypes[h] = classifyColumn(h, rows.slice(0, SAMPLE).map(r => r[h]));
  });

  const dateCol   = headers.find(h => colTypes[h]==="date");
  const moneyCols = headers.filter(h => colTypes[h]==="money");
  const catCols   = headers.filter(h => colTypes[h]==="category");
  const nameCols  = headers.filter(h => colTypes[h]==="name");
  const scoreCols = headers.filter(h => colTypes[h]==="score");
  const numCols   = headers.filter(h => ["numeric","count","money","score","percent"].includes(colTypes[h]));

  // Best primary metric: prefer _clp columns, then large-sum money cols
  const bestMoney = moneyCols.reduce((best, col) => {
    const sum = rows.slice(0,50).reduce((s,r) => s+(parseFloat(r[col])||0), 0);
    return sum > (best.sum||0) ? {col, sum} : best;
  }, {});
  const primaryMetric = bestMoney.col || moneyCols[0] || numCols.find(h=>colTypes[h]==="count") || null;

  const primaryCat   = catCols[0] || null;
  const secondaryCat = catCols[1] || null;
  const thirdCat     = catCols[2] || null;
  const primaryName  = nameCols[0] || null;

  return { headers, colTypes, dateCol, moneyCols, catCols, nameCols, scoreCols, numCols,
           primaryMetric, primaryCat, secondaryCat, thirdCat, primaryName, rows };
}

// ─── AGGREGATORS ──────────────────────────────────────────────────────────────
function aggBy(rows, groupCol, valueCol, mode="sum", limit=12) {
  const map = {};
  rows.forEach(r => {
    const key = r[groupCol] != null ? String(r[groupCol]).trim() : "N/A";
    const val = parseFloat(r[valueCol]) || 0;
    if (!map[key]) map[key] = { name:key, total:0, count:0 };
    map[key].total += val; map[key].count++;
  });
  return Object.values(map)
    .map(d => ({ name:d.name, value:mode==="avg" ? +(d.total/d.count).toFixed(2) : +d.total.toFixed(2), count:d.count }))
    .sort((a,b) => b.value-a.value).slice(0,limit);
}

function aggByDate(rows, dateCol, valueCol) {
  const map = {};
  rows.forEach(r => {
    let raw = r[dateCol]; if (!raw) return;
    let label = String(raw).trim();
    // ISO: 2025-09-07 → 2025-09
    const iso = label.match(/^(\d{4})-(\d{2})/);
    if (iso) { label = `${iso[1]}-${iso[2]}`; }
    else {
      const slash = label.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (slash) label = `${slash[3]}-${slash[2].padStart(2,"0")}`;
      else { const d = new Date(raw); if (!isNaN(d)) label=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`; }
    }
    const val = parseFloat(r[valueCol]) || 0;
    if (!map[label]) map[label] = { fecha:label, total:0, count:0 };
    map[label].total += val; map[label].count++;
  });
  return Object.values(map).sort((a,b)=>a.fecha.localeCompare(b.fecha)).map(d=>({...d,total:+d.total.toFixed(2)}));
}

// ─── FORMATTER ────────────────────────────────────────────────────────────────
function fmt(val, col="") {
  if (val==null) return "—";
  const n = String(col).toLowerCase();
  const isMoney = /(_clp|_usd|_eur|\(clp\)|\(usd\))|(amount|premium|price|cost|revenue|monto|prima|precio|comisi|valor|saldo|cuota|total\b|ingreso|margen)/i.test(n);
  if (isMoney && typeof val==="number") {
    if (Math.abs(val)>=1e9) return `$${(val/1e9).toFixed(1)}B`;
    if (Math.abs(val)>=1e6) return `$${(val/1e6).toFixed(1)}M`;
    if (Math.abs(val)>=1e3) return `$${(val/1e3).toFixed(0)}K`;
    return `$${Number(val).toLocaleString()}`;
  }
  if (typeof val==="number") return val.toLocaleString();
  return String(val);
}
const short = col => String(col).replace(/\s*\(.*?\)/g,"").replace(/_/g," ").replace(/\b\w/g,c=>c.toUpperCase()).trim();

// ─── ETL STEPS ────────────────────────────────────────────────────────────────
const ETL_STEPS = [
  {label:"EXTRACT",   desc:"Leyendo hojas del archivo",           color:"#3B82F6"},
  {label:"DETECT",    desc:"Detectando estructura y encabezados", color:"#F59E0B"},
  {label:"CLEAN",     desc:"Normalizando tipos y valores",        color:"#EF4444"},
  {label:"TRANSFORM", desc:"Calculando métricas adaptativas",     color:"#10B981"},
  {label:"LOAD",      desc:"Generando dashboard personalizado",   color:"#8B5CF6"},
];
const TABS = ["📊 Dashboard","🔄 Pipeline ETL","📋 Datos","📈 Análisis"];

// ─── UI ATOMS ─────────────────────────────────────────────────────────────────
function KPI({label, value, sub, color, icon}) {
  return (
    <div style={{...card(), padding:"20px 22px", position:"relative", overflow:"hidden"}}>
      <div style={{position:"absolute",top:0,right:0,width:52,height:52,background:`${color}18`,borderRadius:"0 10px 0 52px"}}/>
      <div style={{fontSize:20,marginBottom:6}}>{icon}</div>
      <div style={{color:T.muted,fontSize:10,fontFamily:"monospace",letterSpacing:2,textTransform:"uppercase",marginBottom:4}}>{label}</div>
      <div style={{color,fontSize:22,fontWeight:700,letterSpacing:-0.5,marginBottom:2}}>{value}</div>
      {sub&&<div style={{color:T.muted,fontSize:11}}>{sub}</div>}
    </div>
  );
}
function SLabel({children}) {
  return (
    <div style={{display:"flex",alignItems:"center",gap:10,margin:"22px 0 12px"}}>
      <div style={{width:3,height:16,background:T.accent,borderRadius:2}}/>
      <span style={{color:T.text,fontSize:14,fontWeight:600}}>{children}</span>
    </div>
  );
}
function Badge({text, color}) {
  return <span style={{background:`${color}22`,color,borderRadius:4,padding:"2px 9px",fontSize:11,fontWeight:600}}>{text}</span>;
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const [tab,        setTab]        = useState(0);
  const [mode,       setMode]       = useState("home");
  const [etlStep,    setEtlStep]    = useState(0);
  const [etlDone,    setEtlDone]    = useState(false);
  const [fileName,   setFileName]   = useState("");
  const [company,    setCompany]    = useState("");
  const [dateRange,  setDateRange]  = useState("");
  const [an,         setAn]         = useState(null);
  const [kpis,       setKpis]       = useState(null);
  const [charts,     setCharts]     = useState(null);
  const [etlLog,     setEtlLog]     = useState([]);
  const [sheets,     setSheets]     = useState([]);
  const [activeSheet,setActiveSheet]= useState(0);
  const [dragging,   setDragging]   = useState(false);
  const inputRef = useRef();

  const animateETL = (log, cb) => {
    setEtlStep(0); setEtlDone(false); setEtlLog(log); setMode("loading");
    let s=0;
    const iv=setInterval(()=>{ s++; setEtlStep(s); if(s>=ETL_STEPS.length){clearInterval(iv);setTimeout(()=>{setEtlDone(true);cb();},400);}},600);
  };

  const buildDash = (mainSheet, allSheets, comp) => {
    const {rows, headers} = mainSheet;
    const analysis = analyze(rows, headers);

    // Date range
    let range = "";
    if (analysis.dateCol) {
      const dates = rows.map(r => {
        const raw = r[analysis.dateCol]; if(!raw) return null;
        const s = String(raw).trim();
        const iso = s.match(/^(\d{4}-\d{2})/);
        if (iso) return iso[1];
        const sl = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (sl) return `${sl[3]}-${sl[2].padStart(2,"0")}`;
        return null;
      }).filter(Boolean).sort();
      if (dates.length>=2) range = `${dates[0]} → ${dates[dates.length-1]}`;
    }

    // KPIs
    const totalRows = rows.length;
    const totalMain = analysis.primaryMetric
      ? +rows.reduce((s,r)=>s+(parseFloat(r[analysis.primaryMetric])||0),0).toFixed(2) : null;
    const totalExtra = analysis.moneyCols[1] && analysis.moneyCols[1]!==analysis.primaryMetric
      ? +rows.reduce((s,r)=>s+(parseFloat(r[analysis.moneyCols[1]])||0),0).toFixed(2) : null;
    const avgScore = analysis.scoreCols[0]
      ? +(rows.reduce((s,r)=>s+(parseFloat(r[analysis.scoreCols[0]])||0),0)/rows.length).toFixed(1) : null;

    // Charts
    const timeSeries = analysis.dateCol && analysis.primaryMetric
      ? aggByDate(rows, analysis.dateCol, analysis.primaryMetric) : null;
    const byCat1 = analysis.primaryCat && analysis.primaryMetric
      ? aggBy(rows, analysis.primaryCat, analysis.primaryMetric) : null;
    const byCat2 = analysis.secondaryCat && analysis.primaryMetric
      ? aggBy(rows, analysis.secondaryCat, analysis.primaryMetric, 10) : null;
    const byCat3 = analysis.thirdCat && analysis.primaryMetric
      ? aggBy(rows, analysis.thirdCat, analysis.primaryMetric, 8) : null;
    const byName = analysis.primaryName && analysis.primaryMetric
      ? aggBy(rows, analysis.primaryName, analysis.primaryMetric, 10) : null;
    const byScore = analysis.primaryCat && analysis.scoreCols[0]
      ? aggBy(rows, analysis.primaryCat, analysis.scoreCols[0], "avg", 10) : null;

    setAn(analysis);
    setKpis({totalRows, totalMain, totalExtra, avgScore,
      mainCol:analysis.primaryMetric, extraCol:analysis.moneyCols[1]||null, scoreCol:analysis.scoreCols[0]||null});
    setCharts({timeSeries, byCat1, byCat2, byCat3, byName, byScore});
    setSheets(allSheets);
    setCompany(comp);
    setDateRange(range);
    setMode("done");
    setTab(0);
  };

  const processFile = (file) => {
    if (!file) return;
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      const log = [];
      try {
        const wb = XLSX.read(e.target.result, {type:"binary", cellDates:false});
        log.push(`✓ Archivo leído: ${file.name}`);
        log.push(`✓ ${wb.SheetNames.length} hoja(s): ${wb.SheetNames.join(", ")}`);
        const allSheets = readSheets(wb);
        log.push(`✓ ${allSheets.length} hoja(s) con datos válidos`);
        if (!allSheets.length) { log.push("⚠ Sin hojas válidas"); setEtlLog(log); setMode("home"); return; }
        const main = pickMain(allSheets);
        log.push(`✓ Hoja principal: "${main.name}" — ${main.rows.length} filas, ${main.headers.length} columnas`);
        const titled = allSheets.filter(s=>s.looksLikeTitle);
        if (titled.length) log.push(`✓ Fila título omitida en: ${titled.map(s=>s.name).join(", ")}`);
        const a = analyze(main.rows, main.headers);
        const detected = Object.values(a.colTypes).filter(v=>v!=="text"&&v!=="id").length;
        log.push(`✓ ${detected}/${main.headers.length} columnas clasificadas`);
        log.push(`✓ Métrica principal: "${a.primaryMetric||"no detectada"}"`);
        log.push(`✓ Categoría principal: "${a.primaryCat||"no detectada"}"`);
        log.push(`✓ Columna de fecha: "${a.dateCol||"no detectada"}"`);
        log.push(`✓ Columna de nombre: "${a.primaryName||"no detectada"}"`);
        const nulls = main.rows.reduce((s,r)=>s+Object.values(r).filter(v=>v==null||v==="").length,0);
        if (nulls>0) log.push(`✓ ${nulls} valor(es) nulo(s) detectado(s)`);
        const comp = extractCompany(allSheets, file.name);
        log.push(`✓ Empresa identificada: "${comp}"`);
        animateETL(log, ()=>buildDash(main, allSheets, comp));
      } catch(err) {
        log.push(`⚠ Error: ${err.message}`); setEtlLog(log); setMode("home");
      }
    };
    reader.readAsBinaryString(file);
  };

  const reset = () => {
    setMode("home"); setAn(null); setKpis(null); setCharts(null);
    setEtlDone(false); setEtlStep(0); setFileName(""); setCompany("");
    setDateRange(""); setSheets([]); setActiveSheet(0);
  };

  const dot = etlDone?T.success:mode==="loading"?T.warning:T.muted;

  return (
    <div style={{minHeight:"100vh",background:T.bg,color:T.text,fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}@keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}*{box-sizing:border-box}::-webkit-scrollbar{width:6px;height:6px}::-webkit-scrollbar-track{background:#080C14}::-webkit-scrollbar-thumb{background:#1E3050;border-radius:3px}select option{background:#162032}`}</style>

      {/* HEADER */}
      <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"0 28px",display:"flex",alignItems:"center",justifyContent:"space-between",height:52,position:"sticky",top:0,zIndex:50}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:28,height:28,background:`linear-gradient(135deg,${T.accent},${T.purple})`,borderRadius:7,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14}}>📊</div>
          <span style={{fontWeight:700,fontSize:15}}>ETL Dashboard</span>
          {company&&<><span style={{color:T.border,fontSize:18,margin:"0 2px"}}>|</span><span style={{color:T.accent,fontSize:13,fontWeight:600}}>{company}</span></>}
          {dateRange&&<span style={{color:T.muted,fontSize:11,fontFamily:"monospace",marginLeft:6,background:T.surf2,padding:"2px 8px",borderRadius:4}}>{dateRange}</span>}
        </div>
        <div style={{display:"flex",alignItems:"center",gap:14}}>
          {fileName&&<span style={{color:T.muted,fontSize:12,fontFamily:"monospace",background:T.surf2,padding:"3px 10px",borderRadius:4}}>📄 {fileName}</span>}
          <div style={{display:"flex",alignItems:"center",gap:6}}>
            <div style={{width:7,height:7,borderRadius:"50%",background:dot,boxShadow:etlDone?`0 0 8px ${dot}`:"none",transition:"all 0.3s"}}/>
            <span style={{fontSize:11,fontFamily:"monospace",color:dot}}>{etlDone?"ACTIVO":mode==="loading"?"PROCESANDO":"EN ESPERA"}</span>
          </div>
        </div>
      </div>

      {/* HOME */}
      {mode==="home"&&(
        <div style={{maxWidth:600,margin:"64px auto",padding:"0 24px"}}>
          <div style={{textAlign:"center",marginBottom:36}}>
            <div style={{display:"inline-block",background:`${T.accent}22`,color:T.accent,borderRadius:20,padding:"4px 16px",fontSize:11,fontFamily:"monospace",letterSpacing:2,marginBottom:14}}>SISTEMA ETL ADAPTATIVO</div>
            <h1 style={{color:T.text,fontSize:28,fontWeight:700,margin:"0 0 12px",letterSpacing:-0.5}}>Sube cualquier Excel.<br/>El dashboard se adapta solo.</h1>
            <p style={{color:T.muted,fontSize:14,lineHeight:1.7,margin:0}}>Compatible con columnas en español e inglés.<br/>Detecta fechas ISO, montos, categorías y nombres automáticamente.</p>
          </div>
          <div onDragOver={e=>{e.preventDefault();setDragging(true)}} onDragLeave={()=>setDragging(false)}
            onDrop={e=>{e.preventDefault();setDragging(false);processFile(e.dataTransfer.files[0])}}
            onClick={()=>inputRef.current.click()}
            style={{border:`2px dashed ${dragging?T.accent:T.border}`,borderRadius:14,padding:"44px 32px",textAlign:"center",cursor:"pointer",background:dragging?`${T.accent}08`:T.surface,transition:"all 0.2s",boxShadow:dragging?`0 0 0 4px ${T.accent}22`:"none"}}>
            <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>processFile(e.target.files[0])}/>
            <div style={{fontSize:40,marginBottom:12}}>📂</div>
            <div style={{color:T.text,fontSize:15,fontWeight:600,marginBottom:6}}>Arrastra tu archivo Excel aquí</div>
            <div style={{color:T.muted,fontSize:13,marginBottom:20}}>Español · Inglés · .xlsx · .xls · .csv · múltiples hojas</div>
            <span style={{background:T.accent,color:"#fff",borderRadius:8,padding:"9px 24px",fontSize:13,fontWeight:600}}>Seleccionar archivo</span>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginTop:20}}>
            {[{icon:"🌐",t:"ES + EN",d:"Detecta columnas en español e inglés"},{icon:"📅",t:"Fechas ISO",d:"Soporta 2025-09-07, dd/mm/yyyy y más formatos"},{icon:"🧠",t:"Valores primero",d:"Clasifica por datos reales, no solo por el nombre"}].map((f,i)=>(
              <div key={i} style={{...card(),padding:18}}>
                <div style={{fontSize:22,marginBottom:8}}>{f.icon}</div>
                <div style={{color:T.text,fontSize:12,fontWeight:600,marginBottom:4}}>{f.t}</div>
                <div style={{color:T.muted,fontSize:11,lineHeight:1.5}}>{f.d}</div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* LOADING */}
      {mode==="loading"&&(
        <div style={{maxWidth:540,margin:"80px auto",padding:"0 24px"}}>
          <div style={{...card(),padding:32}}>
            <div style={{color:T.accent,fontSize:11,fontFamily:"monospace",letterSpacing:2,marginBottom:20}}>EJECUTANDO PIPELINE ETL</div>
            {ETL_STEPS.map((s,i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",gap:14,marginBottom:14}}>
                <div style={{width:32,height:32,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",background:etlStep>i?`${s.color}22`:T.surf2,border:`2px solid ${etlStep>i?s.color:T.border}`,color:etlStep>i?s.color:T.muted,fontSize:13,transition:"all 0.3s"}}>{etlStep>i?"✓":i+1}</div>
                <div style={{flex:1}}>
                  <div style={{fontSize:12,fontWeight:600,color:etlStep>i?T.text:T.muted,fontFamily:"monospace",letterSpacing:1}}>{s.label}</div>
                  <div style={{fontSize:11,color:T.muted}}>{s.desc}</div>
                </div>
                {etlStep===i+1&&<div style={{width:14,height:14,borderRadius:"50%",border:`2px solid ${s.color}`,borderTopColor:"transparent",animation:"spin 0.6s linear infinite"}}/>}
              </div>
            ))}
            {etlLog.length>0&&(
              <div style={{marginTop:20,background:T.surf2,borderRadius:8,padding:14,border:`1px solid ${T.border}`,maxHeight:180,overflowY:"auto"}}>
                {etlLog.map((l,i)=><div key={i} style={{color:l.startsWith("⚠")?T.warning:T.success,fontSize:11,fontFamily:"monospace",marginBottom:4}}>{l}</div>)}
              </div>
            )}
            <div style={{height:4,background:T.surf2,borderRadius:4,marginTop:16}}>
              <div style={{height:"100%",borderRadius:4,background:`linear-gradient(90deg,${T.accent},${T.purple})`,width:`${(etlStep/ETL_STEPS.length)*100}%`,transition:"width 0.55s"}}/>
            </div>
            <div style={{color:T.muted,fontSize:11,textAlign:"right",marginTop:6,fontFamily:"monospace"}}>{Math.round((etlStep/ETL_STEPS.length)*100)}%</div>
          </div>
        </div>
      )}

      {/* DASHBOARD */}
      {mode==="done"&&an&&kpis&&charts&&(
        <>
          <div style={{display:"flex",borderBottom:`1px solid ${T.border}`,paddingLeft:28,background:T.surface}}>
            {TABS.map((t,i)=>(
              <button key={i} onClick={()=>setTab(i)} style={{background:"none",border:"none",cursor:"pointer",padding:"12px 18px",color:tab===i?T.accent:T.muted,borderBottom:`2px solid ${tab===i?T.accent:"transparent"}`,fontSize:13,fontWeight:tab===i?600:400,transition:"all 0.15s"}}>{t}</button>
            ))}
            <div style={{flex:1}}/>
            <button onClick={reset} style={{background:"none",border:`1px solid ${T.border}`,borderRadius:6,margin:"9px 16px",padding:"4px 14px",color:T.muted,cursor:"pointer",fontSize:12}}>+ Nuevo archivo</button>
          </div>

          <div style={{padding:"22px 28px",maxWidth:1320,margin:"0 auto"}}>

            {/* ── DASHBOARD ── */}
            {tab===0&&(
              <div style={{animation:"fadeIn 0.4s ease"}}>
                <div style={{display:"grid",gridTemplateColumns:`repeat(${2+[kpis.totalExtra,kpis.avgScore].filter(v=>v!=null).length},1fr)`,gap:12,marginBottom:4}}>
                  <KPI icon="📋" label="Total registros" value={kpis.totalRows.toLocaleString()} sub="filas procesadas" color={T.accent}/>
                  {kpis.totalMain!=null&&<KPI icon="💰" label={short(kpis.mainCol)} value={fmt(kpis.totalMain,kpis.mainCol)} sub="total acumulado" color={T.success}/>}
                  {kpis.totalExtra!=null&&<KPI icon="📊" label={short(kpis.extraCol)} value={fmt(kpis.totalExtra,kpis.extraCol)} sub="total acumulado" color={T.purple}/>}
                  {kpis.avgScore!=null&&<KPI icon="⭐" label={short(kpis.scoreCol)} value={`${kpis.avgScore}`} sub="promedio" color={T.warning}/>}
                </div>

                {charts.timeSeries&&charts.timeSeries.length>1&&(
                  <>
                    <SLabel>Tendencia — {short(an.primaryMetric)} por período</SLabel>
                    <div style={{...card(),padding:"18px 12px"}}>
                      <ResponsiveContainer width="100%" height={210}>
                        <AreaChart data={charts.timeSeries}>
                          <defs><linearGradient id="g1" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={T.accent} stopOpacity={0.35}/><stop offset="95%" stopColor={T.accent} stopOpacity={0}/></linearGradient></defs>
                          <CartesianGrid strokeDasharray="3 3" stroke={T.border}/>
                          <XAxis dataKey="fecha" stroke={T.muted} tick={{fontSize:11}}/>
                          <YAxis stroke={T.muted} tick={{fontSize:11}} tickFormatter={v=>fmt(v,an.primaryMetric)}/>
                          <Tooltip contentStyle={tip} formatter={v=>[fmt(v,an.primaryMetric),"Total"]}/>
                          <Area type="monotone" dataKey="total" stroke={T.accent} fill="url(#g1)" strokeWidth={2} dot={{fill:T.accent,r:3}}/>
                        </AreaChart>
                      </ResponsiveContainer>
                    </div>
                  </>
                )}

                {(charts.byCat1||charts.byCat2)&&(
                  <div style={{display:"grid",gridTemplateColumns:charts.byCat1&&charts.byCat2?"1fr 1fr":"1fr",gap:16}}>
                    {charts.byCat1&&(
                      <div>
                        <SLabel>{short(an.primaryMetric)} por {short(an.primaryCat)}</SLabel>
                        <div style={{...card(),padding:"18px 12px"}}>
                          <ResponsiveContainer width="100%" height={Math.max(200,charts.byCat1.length*30)}>
                            <BarChart data={charts.byCat1} layout="vertical">
                              <CartesianGrid strokeDasharray="3 3" stroke={T.border}/>
                              <XAxis type="number" stroke={T.muted} tick={{fontSize:10}} tickFormatter={v=>fmt(v,an.primaryMetric)}/>
                              <YAxis type="category" dataKey="name" stroke={T.muted} tick={{fontSize:10}} width={130}/>
                              <Tooltip contentStyle={tip} formatter={v=>[fmt(v,an.primaryMetric)]}/>
                              <Bar dataKey="value" radius={[0,4,4,0]}>{charts.byCat1.map((_,i)=><Cell key={i} fill={PAL[i%PAL.length]}/>)}</Bar>
                            </BarChart>
                          </ResponsiveContainer>
                        </div>
                      </div>
                    )}
                    {charts.byCat2&&(
                      <div>
                        <SLabel>{short(an.primaryMetric)} por {short(an.secondaryCat)}</SLabel>
                        <div style={{...card(),padding:"18px 12px"}}>
                          <ResponsiveContainer width="100%" height={Math.max(200,charts.byCat1?.length*30||200)}>
                            <PieChart>
                              <Pie data={charts.byCat2} cx="50%" cy="50%" outerRadius={85} dataKey="value" label={({name,percent})=>`${name} ${(percent*100).toFixed(0)}%`} labelLine={{stroke:T.muted}}>
                                {charts.byCat2.map((_,i)=><Cell key={i} fill={PAL[i%PAL.length]}/>)}
                              </Pie>
                              <Tooltip contentStyle={tip} formatter={v=>[fmt(v,an.primaryMetric)]}/>
                            </PieChart>
                          </ResponsiveContainer>
                        </div>
                      </div>
                    )}
                  </div>
                )}

                {charts.byCat3&&(
                  <>
                    <SLabel>{short(an.primaryMetric)} por {short(an.thirdCat)}</SLabel>
                    <div style={{...card(),padding:"18px 12px"}}>
                      <ResponsiveContainer width="100%" height={180}>
                        <BarChart data={charts.byCat3}>
                          <CartesianGrid strokeDasharray="3 3" stroke={T.border}/>
                          <XAxis dataKey="name" stroke={T.muted} tick={{fontSize:11}}/>
                          <YAxis stroke={T.muted} tick={{fontSize:11}} tickFormatter={v=>fmt(v,an.primaryMetric)}/>
                          <Tooltip contentStyle={tip} formatter={v=>[fmt(v,an.primaryMetric)]}/>
                          <Bar dataKey="value" radius={[4,4,0,0]}>{charts.byCat3.map((_,i)=><Cell key={i} fill={PAL[(i+4)%PAL.length]}/>)}</Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </>
                )}

                {charts.byName&&(
                  <>
                    <SLabel>Ranking — {short(an.primaryName)} por {short(an.primaryMetric)}</SLabel>
                    <div style={{...card(),padding:"18px 12px"}}>
                      <ResponsiveContainer width="100%" height={200}>
                        <BarChart data={charts.byName}>
                          <CartesianGrid strokeDasharray="3 3" stroke={T.border}/>
                          <XAxis dataKey="name" stroke={T.muted} tick={{fontSize:10}}/>
                          <YAxis stroke={T.muted} tick={{fontSize:10}} tickFormatter={v=>fmt(v,an.primaryMetric)}/>
                          <Tooltip contentStyle={tip} formatter={v=>[fmt(v,an.primaryMetric)]}/>
                          <Bar dataKey="value" radius={[4,4,0,0]}>{charts.byName.map((_,i)=><Cell key={i} fill={PAL[(i+2)%PAL.length]}/>)}</Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </>
                )}

                {charts.byScore&&(
                  <>
                    <SLabel>Promedio {short(an.scoreCols[0])} por {short(an.primaryCat)}</SLabel>
                    <div style={{...card(),padding:"18px 12px"}}>
                      <ResponsiveContainer width="100%" height={170}>
                        <BarChart data={charts.byScore}>
                          <CartesianGrid strokeDasharray="3 3" stroke={T.border}/>
                          <XAxis dataKey="name" stroke={T.muted} tick={{fontSize:11}}/>
                          <YAxis stroke={T.muted} tick={{fontSize:11}}/>
                          <Tooltip contentStyle={tip}/>
                          <Bar dataKey="value" fill={T.warning} radius={[4,4,0,0]}/>
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </>
                )}
              </div>
            )}

            {/* ── ETL ── */}
            {tab===1&&(
              <div style={{animation:"fadeIn 0.4s ease"}}>
                <div style={{marginBottom:20}}>
                  <div style={{color:T.accent,fontSize:11,fontFamily:"monospace",letterSpacing:2,marginBottom:4}}>REGISTRO DE EJECUCIÓN</div>
                  <h2 style={{color:T.text,fontSize:20,fontWeight:700,margin:0}}>Pipeline ETL</h2>
                </div>
                <div style={{display:"flex",gap:8,marginBottom:24}}>
                  {ETL_STEPS.map((s,i)=>(
                    <div key={i} style={{flex:1,...card(),padding:16,textAlign:"center",borderColor:etlStep>i?s.color:T.border,background:etlStep>i?`${s.color}0D`:T.surface}}>
                      <div style={{color:etlStep>i?s.color:T.muted,fontFamily:"monospace",fontSize:10,fontWeight:700,letterSpacing:1}}>{s.label}</div>
                      <div style={{color:etlStep>i?T.sub:T.muted,fontSize:10,marginTop:4}}>{s.desc}</div>
                      <div style={{marginTop:10,color:etlStep>i?s.color:T.muted,fontSize:18}}>{etlStep>i?"✓":"○"}</div>
                    </div>
                  ))}
                </div>
                <SLabel>Log del proceso</SLabel>
                <div style={{...card(),padding:20,fontFamily:"monospace",fontSize:12}}>
                  {etlLog.map((l,i)=><div key={i} style={{color:l.startsWith("⚠")?T.warning:T.success,marginBottom:8}}><span style={{color:T.muted}}>[ETL] </span>{l}</div>)}
                </div>
                <SLabel>Columnas detectadas ({an.headers.length})</SLabel>
                <div style={{...card(),overflow:"hidden"}}>
                  <table style={{width:"100%",borderCollapse:"collapse"}}>
                    <thead><tr style={{background:T.surf2}}>{["Columna original","Tipo detectado","Muestra"].map(h=><th key={h} style={{padding:"10px 16px",color:T.muted,fontSize:11,fontFamily:"monospace",letterSpacing:1,textAlign:"left",borderBottom:`1px solid ${T.border}`}}>{h}</th>)}</tr></thead>
                    <tbody>
                      {an.headers.map((h,i)=>{
                        const type=an.colTypes[h];
                        const cc={date:T.warning,money:T.success,category:T.accent,name:T.purple,score:T.cyan,id:T.muted,numeric:T.success,count:T.accent,percent:T.warning,text:T.muted}[type]||T.muted;
                        return(
                          <tr key={i} style={{borderBottom:`1px solid ${T.border}`}}>
                            <td style={{padding:"10px 16px",color:T.text,fontSize:13,fontFamily:"monospace"}}>{h}</td>
                            <td style={{padding:"10px 16px"}}><Badge text={type} color={cc}/></td>
                            <td style={{padding:"10px 16px",color:T.muted,fontSize:12,maxWidth:220,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{an.rows[0]?.[h]!=null?String(an.rows[0][h]):"—"}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* ── DATOS ── */}
            {tab===2&&(
              <div style={{animation:"fadeIn 0.4s ease"}}>
                {sheets.length>1&&(
                  <div style={{display:"flex",gap:8,marginBottom:16,flexWrap:"wrap"}}>
                    {sheets.map((s,i)=>(
                      <button key={i} onClick={()=>setActiveSheet(i)} style={{background:activeSheet===i?T.accent:"none",border:`1px solid ${activeSheet===i?T.accent:T.border}`,borderRadius:6,padding:"6px 14px",color:activeSheet===i?"#fff":T.muted,cursor:"pointer",fontSize:12,fontWeight:activeSheet===i?600:400}}>
                        {s.name} <span style={{opacity:0.6}}>({s.rows.length})</span>
                      </button>
                    ))}
                  </div>
                )}
                <SLabel>{sheets[activeSheet]?.name} — {sheets[activeSheet]?.rows.length} filas <span style={{color:T.muted,fontWeight:400,marginLeft:8}}>(primeras 50)</span></SLabel>
                <div style={{...card(),overflow:"auto",maxHeight:540}}>
                  {sheets[activeSheet]?.rows.length>0&&(()=>{
                    const rows=sheets[activeSheet].rows.slice(0,50);
                    const cols=Object.keys(rows[0]);
                    return(
                      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                        <thead><tr style={{background:T.surf2,position:"sticky",top:0}}>{cols.map(c=><th key={c} style={{padding:"10px 14px",color:T.muted,fontFamily:"monospace",fontSize:10,letterSpacing:1,textAlign:"left",borderBottom:`1px solid ${T.border}`,whiteSpace:"nowrap"}}>{c}</th>)}</tr></thead>
                        <tbody>{rows.map((row,i)=><tr key={i} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.surface:T.surf2}}>{cols.map(c=><td key={c} style={{padding:"8px 14px",color:T.sub,whiteSpace:"nowrap",maxWidth:200,overflow:"hidden",textOverflow:"ellipsis"}}>{row[c]!=null?String(row[c]):"—"}</td>)}</tr>)}</tbody>
                      </table>
                    );
                  })()}
                </div>
              </div>
            )}

            {/* ── ANÁLISIS ── */}
            {tab===3&&(
              <div style={{animation:"fadeIn 0.4s ease"}}>
                <SLabel>Estadísticas numéricas</SLabel>
                <div style={{...card(),overflow:"hidden",marginBottom:20}}>
                  <table style={{width:"100%",borderCollapse:"collapse"}}>
                    <thead><tr style={{background:T.surf2}}>{["Columna","Tipo","Mínimo","Máximo","Promedio","Total"].map(h=><th key={h} style={{padding:"10px 16px",color:T.muted,fontSize:11,fontFamily:"monospace",letterSpacing:1,textAlign:"left",borderBottom:`1px solid ${T.border}`}}>{h}</th>)}</tr></thead>
                    <tbody>
                      {an.numCols.slice(0,15).map((col,i)=>{
                        const vals=an.rows.map(r=>parseFloat(r[col])).filter(v=>!isNaN(v));
                        if(!vals.length) return null;
                        const sum=vals.reduce((a,b)=>a+b,0);
                        return(
                          <tr key={i} style={{borderBottom:`1px solid ${T.border}`}}>
                            <td style={{padding:"10px 16px",color:T.text,fontWeight:500,fontFamily:"monospace",fontSize:12}}>{col}</td>
                            <td style={{padding:"10px 16px"}}><Badge text={an.colTypes[col]} color={T.accent}/></td>
                            <td style={{padding:"10px 16px",color:T.sub,fontFamily:"monospace",fontSize:12}}>{fmt(Math.min(...vals),col)}</td>
                            <td style={{padding:"10px 16px",color:T.sub,fontFamily:"monospace",fontSize:12}}>{fmt(Math.max(...vals),col)}</td>
                            <td style={{padding:"10px 16px",color:T.accent,fontFamily:"monospace",fontSize:12}}>{fmt(+(sum/vals.length).toFixed(2),col)}</td>
                            <td style={{padding:"10px 16px",color:T.success,fontFamily:"monospace",fontSize:12}}>{fmt(+sum.toFixed(2),col)}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
                {an.catCols.slice(0,5).map((cat,ci)=>{
                  const counts={};
                  an.rows.forEach(r=>{const k=String(r[cat]||"N/A").trim();counts[k]=(counts[k]||0)+1;});
                  const data=Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,10).map(([name,value])=>({name,value}));
                  return(
                    <div key={ci}>
                      <SLabel>Distribución — {short(cat)}</SLabel>
                      <div style={{...card(),padding:"18px 12px"}}>
                        <ResponsiveContainer width="100%" height={160}>
                          <BarChart data={data}>
                            <CartesianGrid strokeDasharray="3 3" stroke={T.border}/>
                            <XAxis dataKey="name" stroke={T.muted} tick={{fontSize:10}}/>
                            <YAxis stroke={T.muted} tick={{fontSize:10}}/>
                            <Tooltip contentStyle={tip} formatter={v=>[v,"registros"]}/>
                            <Bar dataKey="value" fill={PAL[ci%PAL.length]} radius={[4,4,0,0]}/>
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
}