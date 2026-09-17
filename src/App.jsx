import * as XLSX from "xlsx";
import { useState, useMemo, useEffect, useLayoutEffect, useRef } from "react";

// ---------------------------------------------
// PALETTE & CONSTANTS
// ---------------------------------------------
const THEMES = {
  red:   { red:"#C41E1E", dark:"#8B0000", light_bg:"#FEF2F2", light_border:"#FECACA", light_accent:"#FDEAEA" },
  pink:  { red:"#D4457A", dark:"#8B1A4A", light_bg:"#FDF2F7", light_border:"#F9A8C9", light_accent:"#FDE8F0" },
  black: { red:"#1C1C1E", dark:"#000000", light_bg:"#F5F5F5", light_border:"#D1D5DB", light_accent:"#F2F2F2" },
  white: { red:"#1A1A1A", dark:"#FFFFFF", light_bg:"#FAFAFA", light_border:"#E5E5E5", light_accent:"#F5F5F5" },
  ocean: { red:"#0077B6", dark:"#023E8A", light_bg:"#EFF8FF", light_border:"#BAE0FD", light_accent:"#DBEEFF" },
};

const getTheme = () => {
  try { return localStorage.getItem('kb_theme') || 'red'; } catch { return 'red'; }
};

const C = {
  red:"#C41E1E", bg:"#FAFAFA", card:"#FFFFFF",
  green:"#0F7A3C", danger:"#B91C1C",
  border:"#E5E5E5", soft:"#F5F5F5", light:"#F0F0F0",
  text:"#1A1A1A", muted:"#8A8A8A", faint:"#BFBFBF",
  ink:"#262626",
};
const GRAYS = ["#171717","#404040","#595959","#737373","#8C8C8C","#BFBFBF"];
const shadow = "0 1px 3px rgba(0,0,0,0.06)";

// ---------------------------------------------
// THEME-AWARE MONOCHROME RAMP (for Donut slices etc.)
// Generates a light->dark ramp from the active theme's accent color, so
// charts stay in the same one-hue-by-lightness style already used across
// the app instead of an unrelated rainbow. Themes with a near-neutral
// accent (Black, White) fall back to a plain grayscale ramp — forcing a
// hue onto a monochrome theme would break its own look.
// ---------------------------------------------
const hexToHsl = hex => {
  const r=parseInt(hex.slice(1,3),16)/255, g=parseInt(hex.slice(3,5),16)/255, b=parseInt(hex.slice(5,7),16)/255;
  const max=Math.max(r,g,b), min=Math.min(r,g,b);
  let h=0, s=0; const l=(max+min)/2;
  if (max!==min) {
    const d=max-min;
    s = l>0.5 ? d/(2-max-min) : d/(max+min);
    if (max===r) h=(g-b)/d+(g<b?6:0);
    else if (max===g) h=(b-r)/d+2;
    else h=(r-g)/d+4;
    h*=60;
  }
  return [h, s*100, l*100];
};
const hslToHex = (h,s,l) => {
  s/=100; l/=100;
  const k = n => (n+h/30)%12;
  const a = s*Math.min(l,1-l);
  const f = n => l - a*Math.max(-1,Math.min(k(n)-3,Math.min(9-k(n),1)));
  const toHex = x => Math.round(255*x).toString(16).padStart(2,"0");
  return `#${toHex(f(0))}${toHex(f(8))}${toHex(f(4))}`;
};
const themeRamp = (baseHex, steps) => {
  const [h,s] = hexToHsl(baseHex);
  const useHue = s > 8;
  return Array.from({length:Math.max(steps,1)}, (_,i) => {
    const t = steps>1 ? i/(steps-1) : 0;
    const l = 24 + t*52; // dark -> light
    return useHue ? hslToHex(h, Math.max(s,45), l) : hslToHex(0, 0, l);
  });
};

// ---------------------------------------------
// HELPERS
// ---------------------------------------------
const num = v => { const s = String(v ?? ""); const neg = s.trim().startsWith("-"); const n = Number(s.replace(/[^0-9]/g,"")); return isFinite(n) ? (neg ? -n : n) : 0; };
const fmt = n => { const v = Number(n); if (!isFinite(v)) return "Rp0"; if (v === 0) return "Rp0"; return (v<0?"-":"")+"Rp"+Math.abs(v).toLocaleString("id-ID"); };
const toRp = raw => { const s = String(raw ?? ""); const neg = s.trim().startsWith("-"); const n = String(s).replace(/[^0-9]/g,""); return n ? (neg?"-":"")+"Rp"+Number(n).toLocaleString("id-ID") : ""; };
const uid = () => Math.random().toString(36).slice(2,9);
const isoToDisplay = iso => { if (!iso) return ""; const [y,m,d] = iso.split("-"); return `${d}/${m}/${y?.slice(2)}`; };

// ---------------------------------------------
// STORAGE (localStorage-backed, browser-safe)
// window.storage isn't a real browser API — it only exists inside certain
// sandboxed preview tools. On an actual deployed site (Vercel, etc.) it's
// undefined, so every load/save used to silently fail and the app reset
// on every refresh. This shim keeps the same get/set(key,value) shape but
// persists to the browser's own localStorage instead.
// ---------------------------------------------
const storage = {
  async get(key) {
    try {
      const v = localStorage.getItem(key);
      return v !== null ? { value: v } : null;
    } catch { return null; }
  },
  async set(key, value) {
    try { localStorage.setItem(key, value); } catch {}
  },
};

// ---------------------------------------------
// DEFAULT DATA
// ---------------------------------------------
const emptyRows = () => [
  {desc:"",budget:"",actual:"",notes:[]},
  {desc:"",budget:"",actual:"",notes:[]},
  {desc:"",budget:"",actual:"",notes:[]},
];
const defaultKats = () => [
  { id:uid(), nama:"", hint:"primer", rows:emptyRows() },
  { id:uid(), nama:"", hint:"sekunder", rows:emptyRows() },
  { id:uid(), nama:"", hint:"tersier", rows:emptyRows() },
  { id:uid(), nama:"", hint:"debt", rows:emptyRows() },
  { id:uid(), nama:"", hint:"", rows:emptyRows() },
];

const DEFAULT_HINTS = ["primer","sekunder","tersier","debt"];

const injectHints = kats => kats.map((k,i) => ({
  ...k,
  hint: k.hint || DEFAULT_HINTS[i] || "",
}));

const defaultWallets = () => [
  { name:"Saldo Utama", amount:"", isMain:true },
];

// Goal savings bawaan yang selalu ada, namanya gak bisa diganti user.
const DANA_DARURAT_ID = "dana-darurat";
const ensureDanaDarurat = (list) => {
  const arr = Array.isArray(list) ? list : [];
  if (arr.some(g => g?.id === DANA_DARURAT_ID)) return arr;
  return [{ id:DANA_DARURAT_ID, nama:"DANA DARURAT", target:"", saved:"", log:[], walletPenyimpanan:"", walletLocked:false }, ...arr];
};

const blankPeriod = label => ({
  periode: label || new Date().toLocaleString("id-ID",{month:"long",year:"numeric"}),
  income: { main:"", mainWallet:"", lastPeriod:"", mainLocked:false },
  kategoriList: defaultKats(),
  sideIncome: { nama:"", rows:[] },
  log: [],
});

// ---------------------------------------------
// ATOMS
// ---------------------------------------------
const baseInput = {
  border:`1px solid ${C.border}`, borderRadius:4,
  padding:"7px 10px", fontSize:13, fontFamily:"inherit",
  background:C.card, boxSizing:"border-box",
  outline:"none", width:"100%", minWidth:0, color:C.text,
  WebkitAppearance:"none", appearance:"none",
};

const RpInput = ({value, onChange, style={}}) => (
  <input
    value={value} placeholder="Rp0"
    className="m2os-input"
    onFocus={e => e.target.select()}
    onChange={e => {
      const raw = e.target.value.replace(/[^0-9]/g,"");
      onChange(raw ? "Rp"+Number(raw).toLocaleString("id-ID") : "");
    }}
    style={{...baseInput, textAlign:"right", color: (!value || value==="") ? "#BFBFBF" : "#1A1A1A", ...style}}
  />
);

const TxtInput = ({value, onChange, placeholder, style={}}) => (
  <input
    value={value} placeholder={placeholder}
    className="m2os-input"
    onChange={e => onChange(e.target.value)}
    style={{...baseInput, ...style}}
  />
);

const Sel = ({value, onChange, options, style={}}) => (
  <select value={value} onChange={e => onChange(e.target.value)}
    style={{...baseInput, ...style}}>
    {options.map(o => {
      const [v,l] = Array.isArray(o) ? o : [o,o];
      return <option key={v} value={v}>{l}</option>;
    })}
  </select>
);

const CalIcon = ({size=14, color=C.muted}) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none"
    stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="3" y="5" width="18" height="16" rx="2"/>
    <line x1="3" y1="9" x2="21" y2="9"/>
    <line x1="8" y1="3" x2="8" y2="7"/>
    <line x1="16" y1="3" x2="16" y2="7"/>
  </svg>
);

const DateBtn = ({value, onChange, style={}}) => (
  <div style={{position:"relative", height:36, width:"100%", overflow:"hidden", borderRadius:4, ...style}}>
    <input
      type="date"
      className="m2os-date-input"
      value={value}
      onChange={e => onChange(e.target.value)}
      style={{
        border:"none", borderRadius:4,
        padding:"0", fontSize:11, fontFamily:"inherit",
        background:"transparent", boxSizing:"border-box",
        outline:"none", width:"100%", height:"100%",
        color:"transparent", cursor:"pointer", colorScheme:"light",
        WebkitAppearance:"none", appearance:"none",
        position:"absolute", inset:0, opacity:1,
      }}
    />
    <div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center",pointerEvents:"none",borderRadius:4,background:"transparent"}}>
      {value
        ? <span style={{fontSize:9,color:C.text,fontWeight:600}}>{isoToDisplay(value)}</span>
        : <CalIcon size={13} color={C.muted}/>
      }
    </div>
  </div>
);

const Btn = ({children, onClick, color=C.red, outline=false, style={}}) => (
  <button onClick={onClick} style={{
    background: outline ? "transparent" : color,
    color: outline ? color : "#fff",
    border:`1.5px solid ${color}`,
    borderRadius:8, padding:"7px 16px",
    fontSize:12, fontWeight:600, cursor:"pointer",
    fontFamily:"inherit", whiteSpace:"nowrap", ...style
  }}>{children}</button>
);

const Card = ({children, style={}}) => (
  <div style={{background:C.card, borderRadius:4, boxShadow:shadow, border:`1px solid ${C.border}`, overflow:"hidden", ...style}}>
    {children}
  </div>
);

const Lbl = ({children}) => (
  <div style={{fontSize:10,color:C.muted,fontWeight:600,letterSpacing:0.4,textTransform:"uppercase",marginBottom:4}}>
    {children}
  </div>
);

const Tag = ({children, color}) => (
  <span style={{color,fontSize:10,fontWeight:700,whiteSpace:"nowrap"}}>
    {children}
  </span>
);

const Modal = ({children}) => (
  <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.35)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center"}}>
    <div style={{background:C.card,borderRadius:4,padding:24,width:320,maxWidth:"92vw",boxShadow:"0 16px 40px rgba(0,0,0,0.2)"}}>
      {children}
    </div>
  </div>
);

// ---------------------------------------------
// LOGO
// ---------------------------------------------
const LOGO_B64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAMgAAABECAYAAADEKno9AAAACXBIWXMAAAsTAAALEwEAmpwYAAAGuGlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgMTAuMC1jMDAwIDI1LkcuZWY3MmU0ZSwgMjAyNS8wNi8yNy0xODo1NDowNSAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bWxuczpkYz0iaHR0cDovL3B1cmwub3JnL2RjL2VsZW1lbnRzLzEuMS8iIHhtbG5zOnBob3Rvc2hvcD0iaHR0cDovL25zLmFkb2JlLmNvbS9waG90b3Nob3AvMS4wLyIgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iIHhtbG5zOnN0RXZ0PSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VFdmVudCMiIHhtcDpDcmVhdG9yVG9vbD0iQWRvYmUgUGhvdG9zaG9wIDI3LjMgKFdpbmRvd3MpIiB4bXA6Q3JlYXRlRGF0ZT0iMjAyNi0wOC0yM1QwOTo0NzowNSswNzowMCIgeG1wOk1vZGlmeURhdGU9IjIwMjYtMDgtMjNUMDk6NDg6NTUrMDc6MDAiIHhtcDpNZXRhZGF0YURhdGU9IjIwMjYtMDgtMjNUMDk6NDg6NTUrMDc6MDAiIGRjOmZvcm1hdD0iaW1hZ2UvcG5nIiBwaG90b3Nob3A6Q29sb3JNb2RlPSIzIiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOjhmYzRlNzYzLTFjNmUtYzg0Ny05MDQ4LWZkNzdlYzNmYTQ3YSIgeG1wTU06RG9jdW1lbnRJRD0iYWRvYmU6ZG9jaWQ6cGhvdG9zaG9wOmZlMDNkODQzLTZlYjEtNTE0Zi1hOTViLWMyNDBkZDMyZjJmZCIgeG1wTU06T3JpZ2luYWxEb2N1bWVudElEPSJ4bXAuZGlkOmMwMzg4ODNlLTEwOTItM2M0OS1iNmI2LTBkNzc3ZWYwMTNlMCI+IDxwaG90b3Nob3A6VGV4dExheWVycz4gPHJkZjpCYWc+IDxyZGY6bGkgcGhvdG9zaG9wOkxheWVyTmFtZT0ibTIwcyIgcGhvdG9zaG9wOkxheWVyVGV4dD0ibTIwcyIvPiA8L3JkZjpCYWc+IDwvcGhvdG9zaG9wOlRleHRMYXllcnM+IDx4bXBNTTpIaXN0b3J5PiA8cmRmOlNlcT4gPHJkZjpsaSBzdEV2dDphY3Rpb249ImNyZWF0ZWQiIHN0RXZ0Omluc3RhbmNlSUQ9InhtcC5paWQ6YzAzODg4M2UtMTA5Mi0zYzQ5LWI2YjYtMGQ3NzdlZjAxM2UwIiBzdEV2dDp3aGVuPSIyMDI2LTA4LTIzVDA5OjQ3OjA1KzA3OjAwIiBzdEV2dDpzb2Z0d2FyZUFnZW50PSJBZG9iZSBQaG90b3Nob3AgMjcuMyAoV2luZG93cykiLz4gPHJkZjpsaSBzdEV2dDphY3Rpb249ImNvbnZlcnRlZCIgc3RFdnQ6cGFyYW1ldGVycz0iZnJvbSBhcHBsaWNhdGlvbi92bmQuYWRvYmUucGhvdG9zaG9wIHRvIGltYWdlL3BuZyIvPiA8cmRmOmxpIHN0RXZ0OmFjdGlvbj0ic2F2ZWQiIHN0RXZ0Omluc3RhbmNlSUQ9InhtcC5paWQ6OGZjNGU3NjMtMWM2ZS1jODQ3LTkwNDgtZmQ3N2VjM2ZhNDdhIiBzdEV2dDp3aGVuPSIyMDI2LTA4LTIzVDA5OjQ4OjU1KzA3OjAwIiBzdEV2dDpzb2Z0d2FyZUFnZW50PSJBZG9iZSBQaG90b3Nob3AgMjcuMyAoV2luZG93cykiIHN0RXZ0OmNoYW5nZWQ9Ii8iLz4gPC9yZGY6U2VxPiA8L3htcE1NOkhpc3Rvcnk+IDwvcmRmOkRlc2NyaXB0aW9uPiA8L3JkZjpSREY+IDwveDp4bXBtZXRhPiA8P3hwYWNrZXQgZW5kPSJyIj8+tslhOAAAA35JREFUeJztnb2RFDEQhXuuIAQCwCEAbIoI8JQADgYhkAU+NqYCIA0cfDLAhipRW4VxbO327cxIrdez3+ftXd3qqdVPLWl+bmmtGQBc5uHKzwEAgwD4YBAABwwC4IBBABwwCIADBgFwwCAADhgEwAGDADhgEAAHDALggEEAHDAIgMMz75fgU5dl9bMCpbXFRDSN1rI1Tmt11WV5Z2bvzeybmb02sy9m9qG09tF2svR6HmRLsowarFFa9n7vmrZupaemXmNQO8fpKV11WU4roa9m9snMPpvZLzN7WVp7u7dtKsgNjDDG+XdvmDWbihaRGP35l88/zOyNmX3v0bZMBek1g/XWMXLgvXavoaYnWtMlXXVZTp9fmdlPM3thZs/N7Hdp7fT5OAbZa5KeGqIH/bz9Sz+foak446EYo95wipWEWclYJ5pAQdfDUTreM2Czk+K8fTU9CpqiNMgZBP4ffIVEPNehoilCi6RB1nZaacDgWMge856SPmojpnr6pGb8ukOP2gld2lOsbEFdY+IorWpxKwHXeEZNppJLLJXk7z0oERXx1jaiqnPZ0I7SykHaINkNFI2iYbMjuwd5ai8y0hwZE2er5tkXRdXH4OEeOq8SbMVEjOhzTdxv+QpyqYpkDLhyxYuoIvWJ71eZxGQMsnZQth77jgx8RqOqUp1YzjRPigryGO5Juj/qWewjDTP1FGv0+Xiv2+cxhxaRY5L6mHc0GEN7z3AXNyuOCvTe50pmmUMh8TJprfdws6JCoBXIGIfS2jJb90iTpNuk38LsATu63ktkP4aXN4jqFd0RHMEQao8Mj7r7W8YgmZOvd5tR14ge/33UbS3WWctoJPYgWWdWhfX30ShiMT1UBYkMrNIg7qkiIx/kKjtipLLklqogiomXCR5V7s9hKki0sVST69ZKEqG/JnpsOpVBVMprVpRiV4NerTrKiJIGWUv2WSrLRFF26FHqR+o9yNGTHnIha5BbTXJ0I6n0r/zToaLnMSM1yS+xFAfk3pZa5WwMZuuJzA/pCgLzJ4pypV2FiStCAwZJhNrLLspEk0S1Lb/EgsuJofISiBKg51qbEWCQpIxIzL23hthAo8yqVt3ezQsaqP2X27rRMAp7nBMYBMCBTTqAAwYBcMAgAA4YBMABgwA4YBAABwwC4IBBABwwCIADBgFwwCAADhgEwAGDANh1/gJNehQY/OQ0UwAAAABJRU5ErkJggg==";
const Logo = ({height=20}) => (
  <img src={LOGO_B64} style={{height:height*2,width:"auto",display:"block"}} alt="M2OS"/>
);

// ---------------------------------------------
// DONUT
// ---------------------------------------------
const Donut = ({data, size=116}) => {
  const total = data.reduce((s,d) => s+d.value, 0);
  if (!total) return <div style={{width:size,height:size,background:C.soft,borderRadius:"50%",flexShrink:0}}/>;
  let angle = -90;
  const polar = (cx,cy,r,deg) => { const rad=deg*Math.PI/180; return [cx+r*Math.cos(rad), cy+r*Math.sin(rad)]; };
  const slices = data.filter(d=>d.value>0).map(d => {
    const s = angle; angle += (d.value/total)*360; return {...d,s,e:angle};
  });
  const cx=size/2, cy=size/2, r=size/2-6, ir=size/2-24;
  const scale = Math.max(1, size/116);
  const totalStr = fmt(total);
  const amountFontSize = (totalStr.length>13?8:totalStr.length>10?9.5:11) * scale * 0.9;
  const labelFontSize = 9 * scale * 0.9;
  return (
    <svg width={size} height={size} style={{flexShrink:0}}>
      {slices.map((s,i) => {
        const [sx,sy]=polar(cx,cy,r,s.s), [ex,ey]=polar(cx,cy,r,s.e);
        const large = s.e-s.s>180?1:0;
        return <path key={i} d={`M${cx},${cy} L${sx},${sy} A${r},${r},0,${large},1,${ex},${ey} Z`} fill={s.color}/>;
      })}
      <circle cx={cx} cy={cy} r={ir} fill={C.card}/>
      <text x={cx} y={cy-4*scale} textAnchor="middle" fontSize={labelFontSize} fill={C.muted}>Total</text>
      <text x={cx} y={cy+10*scale} textAnchor="middle" fontSize={amountFontSize} fontWeight="700" fill={C.text}>{totalStr}</text>
    </svg>
  );
};

// ---------------------------------------------
// BAR CHART
// ---------------------------------------------
const BarChart = ({data, height=140}) => {
  const max = Math.max(...data.map(d => Math.max(d.in,d.out)), 1);
  return (
    <div style={{display:"flex",alignItems:"flex-end",gap:12,height,paddingBottom:4}}>
      {data.map((d,i) => (
        <div key={i} style={{display:"flex",flexDirection:"column",alignItems:"center",flex:1,gap:4}}>
          <div style={{display:"flex",gap:3,alignItems:"flex-end",height:height-20,width:"100%",justifyContent:"center"}}>
            <div style={{width:8,background:"linear-gradient(180deg, #C9EBD6 0%, #8FD1AC 100%)",height:`${(d.in/max)*100}%`,minHeight:d.in>0?3:0}}/>
            <div style={{width:8,background:"linear-gradient(180deg, #F6D2D2 0%, #E8A3A3 100%)",height:`${(d.out/max)*100}%`,minHeight:d.out>0?3:0}}/>
          </div>
          <span style={{fontSize:10,color:C.muted,fontWeight:600}}>{d.label}</span>
        </div>
      ))}
    </div>
  );
};

// ---------------------------------------------
// NOTES POPUP
// ---------------------------------------------
const NotesPopup = ({item,onClose,onSave,onTogglePaid,onAdjustBudget,wallets,isMobile,liveBudget,liveActual,isDebt,sisaSaldo,danaBebas}) => {
  const [notes, setNotes] = useState(item.notes || []);
  const [isPiutang, setIsPiutang] = useState(item.isPiutang || false);
  const [expandedCicilan, setExpandedCicilan] = useState({}); // {noteIdx: bool}
  const [cicilanMode, setCicilanMode] = useState(false); // toggle cicilan di form tambah
  const [cicilanTenor, setCicilanTenor] = useState(3);
  const [filterFrom, setFilterFrom] = useState("");
  const [filterTo, setFilterTo] = useState("");

  useEffect(() => {
    setNotes(item.notes || []);
    setIsPiutang(item.isPiutang || false);
    setFilterFrom("");
    setFilterTo("");
  }, [item.notes, item.isPiutang]);

  // Generate cicilan array dari total + tenor
  const generateCicilan = (total, tenor) => {
    const base = Math.floor(num(total) / tenor);
    const sisa = num(total) - base * tenor;
    return Array.from({length: tenor}, (_, i) => ({
      bulan: i + 1,
      amount: toRp(i === tenor-1 ? base + sisa : base),
      active: false,
      paid: false,
    }));
  };

  // Redistribute sisa ke bulan-bulan setelah index yang diubah
  const redistributeCicilan = (cicilan, changedIdx, newAmount) => {
    const total = cicilan.reduce((s,c)=>s+num(c.amount),0);
    const newTotal = newAmount;
    const before = cicilan.slice(0, changedIdx).reduce((s,c)=>s+num(c.amount),0);
    const sisa = total - before - newTotal;
    const afterCount = cicilan.length - changedIdx - 1;
    if (afterCount <= 0) return cicilan.map((c,i)=>i===changedIdx?{...c,amount:toRp(newTotal)}:c);
    const baseAfter = Math.floor(sisa / afterCount);
    const sisaAfter = sisa - baseAfter * afterCount;
    return cicilan.map((c,i)=>{
      if (i < changedIdx) return c;
      if (i === changedIdx) return {...c, amount: toRp(newTotal)};
      const afterIdx = i - changedIdx - 1;
      return {...c, amount: toRp(baseAfter + (afterIdx === afterCount-1 ? sisaAfter : 0))};
    });
  };

  // Budget (rincian & header) buat cicilan cuma ngitung bulan yang lagi "disentuh"
  // periode ini -- dipilih (active/merah) atau udah dibayar (paid/hijau, BUKAN locked).
  // Bulan abu-abu (belum disentuh) dan bulan locked (carry-over dari periode lalu)
  // otomatis kelewat karena keduanya punya active:false & paid:false.
  const cicilanEngagedSum = (cicilan) => (cicilan||[]).filter(c=>c.active||c.paid).reduce((s,c)=>s+num(c.amount),0);

  // Tambah slot bulan baru di akhir cicilan, otomatis keisi sisa kekurangan
  // (total utang dikurangi total semua slot yang udah ada) supaya gak ada dana yang "ngilang".
  const addCicilanBulan = (i) => {
    setNotes(prev => {
      const updated = prev.map((x,j) => {
        if (j!==i || !x.isCicilan) return x;
        const cicilan = x.cicilan||[];
        const sumAll = cicilan.reduce((s,c)=>s+num(c.amount),0);
        const shortfall = Math.max(0, num(x.totalUtang) - sumAll);
        const newCicilan = [...cicilan, { bulan: cicilan.length+1, amount: toRp(shortfall), active:true, paid:false }];
        const engaged = cicilanEngagedSum(newCicilan);
        return { ...x, cicilan: newCicilan, tenor: newCicilan.length, amount: engaged>0?toRp(engaged):"Rp0" };
      });
      onSave(updated, isPiutang);
      return updated;
    });
  };
  const [nd, setNd] = useState("");
  const [na, setNa] = useState("");
  const [nt, setNt] = useState("");
  const [nw, setNw] = useState("");
  const [confirmDelNote, setConfirmDelNote] = useState(null);

  const deleteNote = (i) => {
    const n = notes[i];
    if (n.paid) {
      setConfirmDelNote(i);
    } else {
      const u = notes.filter((_,j)=>j!==i);
      setNotes(u);
      onSave(u, isPiutang);
    }
  };

  const confirmDelete = () => {
    const noteToDelete = notes[confirmDelNote];
    if (noteToDelete?.paid && noteToDelete?.wallet) {
      // Buat notes tanpa item ini, lalu trigger applyNotesChange
      // Set paid=false dulu agar applyNotesChange detect sebagai uncheck
      const withUnpaid = notes.map((x,j)=>j===confirmDelNote?{...x,paid:false}:x);
      onTogglePaid(withUnpaid);
      // Tunggu sebentar lalu hapus dari list
      setTimeout(()=>{
        const u = notes.filter((_,j)=>j!==confirmDelNote);
        setNotes(u);
        onSave(u, isPiutang);
      }, 50);
    } else {
      const u = notes.filter((_,j)=>j!==confirmDelNote);
      setNotes(u);
      onSave(u, isPiutang);
    }
    setConfirmDelNote(null);
  };

  const handleTogglePiutang = (val) => {
    setIsPiutang(val);
    onSave(notes, val);
  };

  // Sama kayak notesBudgetTotal di App: cicilan cuma ngitung bulan yang lagi disentuh
  // periode ini (active/dipilih atau paid, bukan locked), biar "Rencana melebihi budget"
  // gak salah nyala gara-gara Total cicilan yang belum ada satupun bulan-nya dipilih.
  const total = notes.reduce((s,n) => {
    if (n.isCicilan) return s + cicilanEngagedSum(n.cicilan);
    return s+num(n.amount);
  }, 0);
  const totalPaid = notes.reduce((s,n) => s+(n.paid?num(n.amount):0), 0);
  const budgetNum = num(liveBudget);
  const actualNum = num(liveActual);
  const isPiutangEffective = isPiutang || item.isPiutangRow || item.isPiutang;
  const isActualOver = !isPiutangEffective && actualNum > budgetNum;
  const isPlanOver = !isPiutangEffective && total > budgetNum;
  const isOver = isActualOver || isPlanOver;

  // -- CEK PERIODE (filter tampilan rincian by tanggal) --
  const hasDateFilter = !!(filterFrom || filterTo);
  const inDateRange = d => (!filterFrom || d >= filterFrom) && (!filterTo || d <= filterTo);
  const visibleEntries = notes
    .map((n, i) => ({ n, i }))
    .filter(({ n }) => !hasDateFilter || !n.date || inDateRange(n.date));
  const filteredTotal = notes
    .filter(n => n.date && inDateRange(n.date))
    .reduce((s, n) => s + num(n.amount), 0);
  const noDateShown = hasDateFilter && notes.some(n => !n.date);

  const add = () => {
    if ((!nd && !na) || !nw) return;
    const entry = {
      id:uid(), desc:nd, amount:(cicilanMode && num(na)>0) ? "Rp0" : na, date:nt, wallet:nw, paid:false, _isNew:true,
      ...(cicilanMode && num(na)>0 ? {
        isCicilan: true,
        totalUtang: na,
        tenor: cicilanTenor,
        cicilan: generateCicilan(na, cicilanTenor),
      } : {})
    };
    const updated = [...notes, entry];
    setNotes(updated);
    onSave(updated, isPiutang);
    setNd(""); setNa(""); setNt(""); setNw("");
  };

  const togglePaid = i => {
    const n = notes[i];
    setNotes(prev => {
      const updated = prev.map((x,j) => {
        if (j!==i) return x;
        if (x.isCicilan) {
          if (!x.paid && !x.partial) {
            // Centang: bulan active → hijau. Lunas (hijau semua) cuma kalau total yang
            // udah settled (paid+locked) >= total utang -- kalau masih kurang, tetap partial (kuning)
            // walau semua bulan yang ADA udah dicentang (butuh + Tambah Bulan buat nutup sisanya).
            const newCicilan = (x.cicilan||[]).map(c=>c.active?{...c,paid:true,active:false}:c);
            const allSlotsSettled = newCicilan.filter(c=>!c.lockedPaid).every(c=>c.paid) && newCicilan.every(c=>c.paid||c.lockedPaid);
            const sumSettled = newCicilan.filter(c=>c.paid||c.lockedPaid).reduce((s,c)=>s+num(c.amount),0);
            const allPaid = allSlotsSettled && sumSettled >= num(x.totalUtang);
            const hasPaid = newCicilan.some(c=>c.paid);
            // Budget/amount rincian: cuma dari bulan yang dipilih/dibayar PERIODE INI
            // (active atau paid, bukan locked) -- beda dari sumSettled di atas yang buat nentuin
            // status Lunas (itu perlu tau history settled termasuk locked).
            const engaged = cicilanEngagedSum(newCicilan);
            return {
              ...x,
              cicilan: newCicilan,
              paid: allPaid,
              partial: !allPaid && hasPaid,
              amount: engaged>0 ? toRp(engaged) : "Rp0",
            };
          } else {
            // Uncheck dari kuning (partial) ATAU hijau (lunas): bulan yang paid (non-locked) kembali ke merah
            const newCicilan = (x.cicilan||[]).map(c=>
              c.paid && !c.lockedPaid ? {...c,paid:false,active:true} : c
            );
            const engaged = cicilanEngagedSum(newCicilan);
            return {
              ...x,
              cicilan: newCicilan,
              paid: false,
              partial: false,
              amount: engaged>0 ? toRp(engaged) : "Rp0",
            };
          }
        }
        return {...x, paid:!x.paid};
      });
      onTogglePaid(updated);
      return updated;
    });
  };

  const markAllPaid = () => {
    const updated = notes.map(n=>({...n,paid:true}));
    setNotes(updated);
    onSave(updated, isPiutang);
  };

  const wOpts = [["","Pilih wallet*"], ...(wallets||[]).map(w=>[w.name, w.name])];

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.35)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={onClose}>
      <div style={{background:C.card,borderRadius:4,padding:isMobile?14:20,width:isMobile?"92vw":520,maxWidth:isMobile?"92vw":"94vw",maxHeight:"85vh",overflow:"auto",boxShadow:"0 16px 40px rgba(0,0,0,0.18)"}} onClick={e=>e.stopPropagation()}>

        {/* HEADER */}
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
          <div style={{minWidth:0,paddingRight:9.4}}>
            <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
              <div style={{fontWeight:700,fontSize:14}}>{item.desc||"(belum ada nama)"}</div>
            </div>
            <div style={{fontSize:11,color:C.muted,display:"flex",alignItems:"center",gap:6}}>
              <span>{(item.isPiutangRow||item.isPiutang) ? "Rincian pemasukan" : "Rincian pengeluaran"}</span>
              {isDebt && <>
                <span>·</span>
                <span>{item.isPiutangRow||item.isPiutang ? "Piutang" : "Utang"}</span>
              </>}
              {budgetNum>0 && <>
                <span>·</span>
                <span>Budget: <strong>{fmt(budgetNum)}</strong></span>
              </>}
            </div>
          </div>
          <button onClick={onClose} style={{background:"none",border:"none",width:26,height:26,cursor:"pointer",fontSize:14,color:C.muted,flexShrink:0}}>x</button>
        </div>

        {/* FILTER PERIODE */}
        {notes.length > 0 && (
          <div style={{marginBottom:12}}>
            <div style={{display:"flex",gap:4,alignItems:"center",justifyContent:"flex-start",flexWrap:"wrap",rowGap:2}}>
              <span style={{fontSize:11,color:C.muted,fontWeight:600,flexShrink:0}}>Filter Periode:</span>
              <DateBtn value={filterFrom} onChange={setFilterFrom} style={{width:54,height:18,flexShrink:0}}/>
              <span style={{fontSize:11,color:C.muted,flexShrink:0}}>s/d</span>
              <DateBtn value={filterTo} onChange={setFilterTo} style={{width:54,height:18,flexShrink:0}}/>
              {hasDateFilter && (
                <span style={{fontSize:11,color:C.muted,marginLeft:3}}>
                  Total: <strong style={{color:C.text}}>{fmt(filteredTotal)}</strong>
                </span>
              )}
              {hasDateFilter && (
                <button onClick={()=>{setFilterFrom("");setFilterTo("");}}
                  style={{background:"none",border:"none",color:C.muted,cursor:"pointer",fontSize:11,flexShrink:0,textDecoration:"underline",marginLeft:5}}>
                  Clear
                </button>
              )}
            </div>
            {noDateShown && (
              <div style={{fontSize:10,color:C.muted,marginTop:0}}>*Terdapat rincian tanpa tanggal</div>
            )}
          </div>
        )}

        {/* TABLE */}
        <div style={{background:C.soft,borderRadius:4,overflow:"hidden",marginBottom:12,border:`1px solid ${C.border}`}}>
          {!isMobile && (
            <div style={{display:"grid",gridTemplateColumns:"22px 1fr 110px 40px 110px 28px",gap:6,padding:"7px 10px",fontSize:9,color:C.muted,fontWeight:700}}>
              <span></span><span>DETAIL</span><span style={{textAlign:"right"}}>RUPIAH</span><span></span><span>WALLET</span><span/>
            </div>
          )}

          {notes.length === 0 && (
            <div style={{padding:14,textAlign:"center",fontSize:12,color:C.muted}}>
              {(item.isPiutangRow||item.isPiutang)
                ? "Belum ada rincian. Tambah di bawah untuk mencatat piutang yang belum diterima."
                : "Belum ada rincian. Tambah di bawah untuk mulai rencanakan pengeluaran ini."
              }
            </div>
          )}
          {notes.length > 0 && visibleEntries.length === 0 && (
            <div style={{padding:14,textAlign:"center",fontSize:12,color:C.muted}}>
              Tidak ada rincian bertanggal di periode ini.
            </div>
          )}

          {visibleEntries.map(({n,i}) => isMobile ? (
            <div key={i} style={{borderTop:i>0?`1px solid ${C.border}`:"none",background:C.card}}>
              <div style={{padding:"8px 7px 8px 10px"}} className="m2os-note-row">
                <div style={{display:"flex",alignItems:"center",gap:10}}>
                  {/* Lingkaran */}
                  <button onClick={()=>togglePaid(i)} style={{width:7,height:7,borderRadius:"50%",border:"none",background:n.paid?"#22C55E":n.partial?"#EAB308":"#E5E7EB",boxShadow:n.paid?"0 0 5px 2px rgba(34,197,94,0.55)":n.partial?"0 0 5px 2px rgba(234,179,8,0.55)":"none",cursor:"pointer",padding:0,flexShrink:0,transition:"background 0.2s, box-shadow 0.2s"}}/>
                  {/* Konten: nama + tanggal */}
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{display:"flex",alignItems:"center",gap:7}}>
                      <input value={n.desc||""} onChange={e=>setNotes(prev=>{ const u=prev.map((x,j)=>j===i?{...x,desc:e.target.value}:x); onSave(u,isPiutang); return u; })}
                        placeholder="detail..." className="m2os-input"
                        style={{fontSize:12,fontWeight:600,flex:1,border:"none",background:"transparent",padding:0,outline:"none",color:n.paid?C.muted:C.text,fontFamily:"inherit"}}/>
                      <input value={n.amount||""} onChange={e=>{ const raw=e.target.value.replace(/[^0-9]/g,""); const v=raw?"Rp"+Number(raw).toLocaleString("id-ID"):""; setNotes(prev=>{ const u=prev.map((x,j)=>j===i?{...x,amount:v}:x); if(n.paid) onTogglePaid(u); else onSave(u,isPiutang); return u; }); }}
                        placeholder="Rp0" className="m2os-input"
                        style={{fontSize:12,fontWeight:700,width:90,textAlign:"right",border:"none",background:"transparent",padding:0,outline:"none",color:n.paid?"#22C55E":n.partial?"#EAB308":C.danger,fontFamily:"inherit"}}/>
                      <button onClick={()=>deleteNote(i)} className="m2os-note-del" style={{background:"none",border:"none",cursor:"pointer",flexShrink:0,opacity:0,transition:"opacity 0.15s",padding:0,display:"flex",alignItems:"center"}}><svg width="10" height="13" viewBox="0 0 10 13" fill="none" xmlns="http://www.w3.org/2000/svg">
  <polygon points="1,2 9,2 8,12 2,12" stroke="#E5E7EB" strokeWidth="1.2" fill="none" strokeLinejoin="round"/>
  <line x1="0" y1="2" x2="10" y2="2" stroke="#E5E7EB" strokeWidth="1.2" strokeLinecap="round"/>
  <line x1="3" y1="4.5" x2="3" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
  <line x1="5" y1="4.5" x2="5" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
  <line x1="7" y1="4.5" x2="7" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
</svg></button>
                    </div>
                    <div style={{fontSize:10,color:C.muted,marginTop:2,marginBottom:n.isCicilan&&expandedCicilan[i]?6:0,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                      <div style={{display:"flex",alignItems:"center",gap:4,flex:1,minWidth:0}}>
                        <span>{n.date?isoToDisplay(n.date):"tanpa tanggal"} ·</span>
                        {(n.paid||n.partial) ? (
                          <span style={{color:C.muted}}>{n.wallet||" - "}</span>
                        ) : (
                          <select value={n.wallet||""} onChange={e=>{
                            const u=notes.map((x,j)=>j===i?{...x,wallet:e.target.value}:x);
                            setNotes(u); onSave(u,isPiutang);
                          }} style={{fontSize:10,border:"none",background:"transparent",color:C.muted,padding:0,outline:"none",fontFamily:"inherit",cursor:"pointer",maxWidth:100}}>
                            {wOpts.map(([v,l])=><option key={v} value={v}>{l}</option>)}
                          </select>
                        )}
                      </div>
                      {n.isCicilan && (
                        <div style={{width:90,display:"flex",justifyContent:"flex-end",paddingRight:14,flexShrink:0}}>
                          <button onClick={()=>setExpandedCicilan(p=>({...p,[i]:!p[i]}))}
                            style={{background:"none",border:"none",color:C.muted,cursor:"pointer",fontSize:9,padding:0,display:"flex",alignItems:"center",gap:2,fontFamily:"inherit",lineHeight:1}}>
                            <span style={{fontWeight:700}}>Cicilan {n.tenor}×</span>
                            <span style={{fontSize:8,color:C.faint,transform:expandedCicilan[i]?"rotate(180deg)":"none",transition:"transform 0.2s",display:"inline-block"}}>▼</span>
                          </button>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              {/* EXPAND CICILAN */}
              {n.isCicilan && expandedCicilan[i] && (
                <div style={{background:"#FAFAFA",borderTop:`1px solid ${C.border}`,padding:"12px 10px 8px 36px",marginTop:4}}>
                  <div style={{color:C.muted,marginBottom:6}}>
                    <div style={{fontSize:11,fontWeight:400}}>Total {fmt(num(n.totalUtang))}</div>
                    <div style={{fontSize:10,fontWeight:700}}>Tersisa {fmt(Math.max(0, num(n.totalUtang)-(n.cicilan||[]).filter(c=>c.paid||c.lockedPaid).reduce((s,c)=>s+num(c.amount),0)))}</div>
                    {(() => {
                      const cic = n.cicilan||[];
                      const shortfall = Math.max(0, num(n.totalUtang)-cic.filter(c=>c.paid||c.lockedPaid).reduce((s,c)=>s+num(c.amount),0));
                      const allSlotsSettled = cic.length>0 && cic.every(c=>c.paid||c.lockedPaid);
                      return (allSlotsSettled && shortfall>0) ? (
                        <div style={{fontSize:9,color:C.danger,marginTop:2}}>*Masih kurang {fmt(shortfall)}. Uncheck dulu, lalu tambah bulan buat nutup sisanya.</div>
                      ) : null;
                    })()}
                  </div>
                  {(n.cicilan||[]).map((c,ci)=>(
                    <div key={ci} style={{display:"flex",alignItems:"center",gap:8,padding:"4px 0",borderBottom:ci<n.cicilan.length-1?`1px solid ${C.light}`:"none"}}>
                      {/* Lingkaran status */}
                      <button onClick={()=>{
                        if(n.paid||c.lockedPaid||n.partial) return;
                        const newCicilan = n.cicilan.map((x,xi)=>xi===ci&&!x.paid?{...x,active:!x.active}:x);
                        const engaged = cicilanEngagedSum(newCicilan);
                        const u = notes.map((x,j)=>j===i?{...x,cicilan:newCicilan,amount:engaged>0?toRp(engaged):"Rp0"}:x);
                        setNotes(u); onSave(u, isPiutang);
                      }} style={{
  width:7,height:7,borderRadius:"50%",border:"none",
  background:c.lockedPaid?"#22C55E":c.paid?"#22C55E":c.active?"#EF4444":"#D1D5DB",
  boxShadow:c.lockedPaid?"0 0 5px 2px rgba(34,197,94,0.3)":c.paid?"0 0 5px 2px rgba(34,197,94,0.55)":c.active?"0 0 5px 2px rgba(239,68,68,0.55)":"none",
  cursor:(c.lockedPaid||n.paid||n.partial)?"default":"pointer",flexShrink:0,padding:0
}}/>
                      <span style={{fontSize:11,color:C.muted,minWidth:52}}>Bulan {c.bulan}</span>
                      <input value={c.amount||""} onChange={e=>{
                        if(n.paid||c.paid||c.lockedPaid||n.partial) return;
                        const raw=e.target.value.replace(/[^0-9]/g,"");
                        const v=raw?"Rp"+Number(raw).toLocaleString("id-ID"):"";
                        const newCicilan = redistributeCicilan(n.cicilan, ci, num(v)||0);
                        const engaged = cicilanEngagedSum(newCicilan);
                        const u = notes.map((x,j)=>j===i?{...x,cicilan:newCicilan,amount:engaged>0?toRp(engaged):"Rp0"}:x);
                        setNotes(u); onSave(u, isPiutang);
                      }} className="m2os-input" style={{fontSize:11,fontWeight:600,flex:1,textAlign:"right",border:"none",background:"transparent",padding:0,outline:"none",color:c.lockedPaid?"#22C55E":c.paid?"#22C55E":c.active?C.danger:C.muted,fontFamily:"inherit",textDecoration:"none"}}/>

                    </div>
                  ))}
                  {!n.paid && !n.partial && (
                    <button onClick={()=>addCicilanBulan(i)} style={{background:"none",border:"none",padding:"6px 0 0",fontSize:10,fontWeight:600,color:C.faint,cursor:"pointer",fontFamily:"inherit"}}>+ Tambah Bulan</button>
                  )}
                </div>
              )}
            </div>
          </div>
          ) : (
            <div key={i} style={{borderTop:`1px solid ${C.border}`,background:C.card}}>
              <div style={{display:"grid",gridTemplateColumns:"16px 1fr 110px 40px 110px 36px",gap:8,padding:"7px 8px 7px 10px",alignItems:"center"}} className="m2os-note-row">
                <button onClick={()=>togglePaid(i)} style={{width:7,height:7,borderRadius:"50%",border:"none",background:n.paid?"#22C55E":n.partial?"#EAB308":"#E5E7EB",boxShadow:n.paid?"0 0 5px 2px rgba(34,197,94,0.55)":n.partial?"0 0 5px 2px rgba(234,179,8,0.55)":"none",cursor:"pointer",padding:0,transition:"background 0.2s, box-shadow 0.2s",justifySelf:"center"}}/>
                <div style={{minWidth:0,display:"flex",flexDirection:"column",justifyContent:"center"}}>
                  <input value={n.desc||""} onChange={e=>setNotes(prev=>{ const u=prev.map((x,j)=>j===i?{...x,desc:e.target.value}:x); onSave(u,isPiutang); return u; })}
                    placeholder="detail..." className="m2os-input"
                    style={{fontSize:12,border:"none",background:"transparent",padding:0,outline:"none",textDecoration:"none",color:n.paid?C.muted:C.text,fontFamily:"inherit",width:"100%"}}/>
                  {n.isCicilan && (
                    <button onClick={()=>setExpandedCicilan(p=>({...p,[i]:!p[i]}))}
                      style={{background:"none",border:"none",color:C.muted,cursor:"pointer",fontSize:9,padding:0,display:"flex",alignItems:"center",gap:2,fontFamily:"inherit",textAlign:"left"}}>
                      <span style={{fontWeight:700}}>Cicilan {n.tenor}×</span>
                      <span style={{fontSize:8,color:C.faint,transform:expandedCicilan[i]?"rotate(180deg)":"none",transition:"transform 0.2s",display:"inline-block"}}>▼</span>
                    </button>
                  )}
                </div>
                <input value={n.amount||""} onChange={e=>{ const raw=e.target.value.replace(/[^0-9]/g,""); const v=raw?"Rp"+Number(raw).toLocaleString("id-ID"):""; setNotes(prev=>{ const u=prev.map((x,j)=>j===i?{...x,amount:v}:x); if(n.paid) onTogglePaid(u); else onSave(u,isPiutang); return u; }); }}
                  placeholder="Rp0" className="m2os-input"
                  style={{fontSize:12,fontWeight:600,textAlign:"right",border:"none",background:"transparent",padding:0,outline:"none",color:n.paid?"#22C55E":n.partial?"#EAB308":C.danger,fontFamily:"inherit"}}/>
                <DateBtn value={n.date} onChange={v=>setNotes(prev=>{ const u=prev.map((x,j)=>j===i?{...x,date:v}:x); onSave(u,isPiutang); return u; })}/>
                {(n.paid||n.partial) ? (
                  <span style={{fontSize:11,color:C.muted,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{n.wallet||" - "}</span>
                ) : (
                  <select value={n.wallet||""} onChange={e=>{
                    const u=notes.map((x,j)=>j===i?{...x,wallet:e.target.value}:x);
                    setNotes(u); onSave(u,isPiutang);
                  }} style={{fontSize:11,border:`1px solid ${C.border}`,borderRadius:4,color:C.muted,padding:"4px 6px",outline:"none",fontFamily:"inherit",cursor:"pointer",background:"#fff"}}>
                    {wOpts.map(([v,l])=><option key={v} value={v}>{l}</option>)}
                  </select>
                )}
                <button onClick={()=>deleteNote(i)} className="m2os-note-del" style={{background:"none",border:"none",cursor:"pointer",opacity:0,transition:"opacity 0.15s",padding:0,justifySelf:"center",display:"flex",alignItems:"center"}}><svg width="10" height="13" viewBox="0 0 10 13" fill="none" xmlns="http://www.w3.org/2000/svg">
  <polygon points="1,2 9,2 8,12 2,12" stroke="#E5E7EB" strokeWidth="1.2" fill="none" strokeLinejoin="round"/>
  <line x1="0" y1="2" x2="10" y2="2" stroke="#E5E7EB" strokeWidth="1.2" strokeLinecap="round"/>
  <line x1="3" y1="4.5" x2="3" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
  <line x1="5" y1="4.5" x2="5" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
  <line x1="7" y1="4.5" x2="7" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
</svg></button>
              </div>
              {/* EXPAND CICILAN DESKTOP */}
              {n.isCicilan && expandedCicilan[i] && (
                <div style={{background:"#FAFAFA",borderTop:`1px solid ${C.border}`,padding:"8px 14px 8px 42px"}}>
                  <div style={{color:C.muted,marginBottom:6}}>
                    <div style={{fontSize:11,fontWeight:400}}>Total {fmt(num(n.totalUtang))}</div>
                    <div style={{fontSize:10,fontWeight:700}}>Tersisa {fmt(Math.max(0, num(n.totalUtang)-(n.cicilan||[]).filter(c=>c.paid||c.lockedPaid).reduce((s,c)=>s+num(c.amount),0)))}</div>
                    {(() => {
                      const cic = n.cicilan||[];
                      const shortfall = Math.max(0, num(n.totalUtang)-cic.filter(c=>c.paid||c.lockedPaid).reduce((s,c)=>s+num(c.amount),0));
                      const allSlotsSettled = cic.length>0 && cic.every(c=>c.paid||c.lockedPaid);
                      return (allSlotsSettled && shortfall>0) ? (
                        <div style={{fontSize:9,color:C.danger,marginTop:2}}>*Masih kurang {fmt(shortfall)}. Uncheck dulu, lalu tambah bulan buat nutup sisanya.</div>
                      ) : null;
                    })()}
                  </div>
                  {(n.cicilan||[]).map((c,ci)=>(
                    <div key={ci} style={{display:"flex",alignItems:"center",gap:10,padding:"4px 0",borderBottom:ci<n.cicilan.length-1?`1px solid ${C.light}`:"none"}}>
                      <button onClick={()=>{
                        if(n.paid||c.lockedPaid||n.partial) return;
                        const newCicilan = n.cicilan.map((x,xi)=>xi===ci&&!x.paid?{...x,active:!x.active}:x);
                        const engaged = cicilanEngagedSum(newCicilan);
                        const u = notes.map((x,j)=>j===i?{...x,cicilan:newCicilan,amount:engaged>0?toRp(engaged):"Rp0"}:x);
                        setNotes(u); onSave(u, isPiutang);
                      }} style={{
  width:7,height:7,borderRadius:"50%",border:"none",
  background:c.lockedPaid?"#22C55E":c.paid?"#22C55E":c.active?"#EF4444":"#D1D5DB",
  boxShadow:c.lockedPaid?"0 0 5px 2px rgba(34,197,94,0.3)":c.paid?"0 0 5px 2px rgba(34,197,94,0.55)":c.active?"0 0 5px 2px rgba(239,68,68,0.55)":"none",
  cursor:(c.lockedPaid||n.paid||n.partial)?"default":"pointer",flexShrink:0,padding:0
}}/>
                      <span style={{fontSize:11,color:C.muted,minWidth:60}}>Bulan {c.bulan}</span>
                      <input value={c.amount||""} onChange={e=>{
                        if(n.paid||c.paid||c.lockedPaid||n.partial) return;
                        const raw=e.target.value.replace(/[^0-9]/g,"");
                        const v=raw?"Rp"+Number(raw).toLocaleString("id-ID"):"";
                        const newCicilan = redistributeCicilan(n.cicilan, ci, num(v)||0);
                        const engaged = cicilanEngagedSum(newCicilan);
                        const u = notes.map((x,j)=>j===i?{...x,cicilan:newCicilan,amount:engaged>0?toRp(engaged):"Rp0"}:x);
                        setNotes(u); onSave(u, isPiutang);
                      }} className="m2os-input" style={{fontSize:11,fontWeight:600,width:100,textAlign:"right",border:"none",background:"transparent",padding:0,outline:"none",color:c.lockedPaid?"#22C55E":c.paid?"#22C55E":c.active?C.danger:C.muted,fontFamily:"inherit",textDecoration:"none"}}/>

                    </div>
                  ))}
                  {!n.paid && !n.partial && (
                    <button onClick={()=>addCicilanBulan(i)} style={{background:"none",border:"none",padding:"6px 0 0",fontSize:10,fontWeight:600,color:C.faint,cursor:"pointer",fontFamily:"inherit"}}>+ Tambah Bulan</button>
                  )}
                </div>
              )}
            </div>
          ))}

          {/* Add row */}
          {isMobile ? (
            <div style={{padding:"10px 14px",borderTop:`1px solid ${C.border}`}}>
              {isDebt ? (<>
                {/* DEBT Baris 1: detail + amount */}
                <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:8}}>
                  <TxtInput value={nd} onChange={setNd} placeholder="detail..." style={{flex:2}}/>
                  <RpInput value={na} onChange={setNa} style={{flex:1,flexShrink:0}}/>
                </div>
                {/* DEBT Baris 2: cicilan + wallet + tanggal */}
                <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:8}}>
                  <button onClick={()=>setCicilanMode(v=>!v)}
                    style={{background:cicilanMode?"#FEE2E2":"transparent",border:`1px solid ${cicilanMode?"#FECACA":C.border}`,borderRadius:6,padding:"4px 10px",fontSize:11,color:cicilanMode?C.red:C.muted,cursor:"pointer",fontFamily:"inherit",fontWeight:600,flexShrink:0}}>
                    {cicilanMode?"✓ Cicilan":"+ Cicilan"}
                  </button>
                  <Sel value={nw} onChange={setNw} options={wOpts} style={{flex:1}}/>
                  <DateBtn value={nt} onChange={v=>setNt(v)} style={{width:36,flexShrink:0}}/>
                </div>
                {cicilanMode && (
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                    <span style={{fontSize:11,color:C.muted}}>Tenor:</span>
                    <Sel value={String(cicilanTenor)} onChange={v=>setCicilanTenor(Number(v))}
                      options={Array.from({length:12},(_,i)=>[(i+1).toString(),`${i+1} bulan`])}
                      style={{flex:1,fontSize:11}}/>
                    {num(na)>0 && <span style={{fontSize:11,color:C.muted}}>= <strong style={{color:C.red}}>{fmt(Math.round(num(na)/cicilanTenor))}/bulan</strong></span>}
                  </div>
                )}
              </>) : (<>
                {/* NON-DEBT layout awal */}
                <TxtInput value={nd} onChange={setNd} placeholder="detail..." style={{marginBottom:8}}/>
                <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:8}}>
                  <RpInput value={na} onChange={setNa} style={{flex:1}}/>
                  <Sel value={nw} onChange={setNw} options={wOpts} style={{flex:1}}/>
                  <DateBtn value={nt} onChange={v=>setNt(v)} style={{width:36,flexShrink:0}}/>
                </div>
              </>)}
              <button onClick={add} style={{width:"100%",background:C.red,border:"none",color:"#fff",borderRadius:8,cursor:"pointer",fontSize:12,padding:"7px 0",fontFamily:"inherit",fontWeight:600}}>+ Tambah</button>
              {isOver && (
                <div style={{marginTop:8,display:"flex",gap:8}}>
                  <div style={{flex:1,fontSize:11,color:C.danger,background:"#FEF2F2",border:`1px solid #FECACA`,borderRadius:4,padding:"6px 10px",display:"flex",alignItems:"center"}}>
                    {isActualOver ? `Aktual melebihi budget sebesar ${fmt(actualNum-budgetNum)}` : `Rencana melebihi budget sebesar ${fmt(total-budgetNum)}`}
                  </div>
                  <button onClick={()=>onAdjustBudget(isActualOver?actualNum:total)}
                    style={{fontSize:11,color:C.danger,border:`1px solid #FECACA`,borderRadius:4,padding:"6px 10px",background:"none",cursor:"pointer",fontFamily:"inherit",fontWeight:600,whiteSpace:"nowrap"}}>
                    Sesuaikan Budget
                  </button>
                </div>
              )}
            </div>
          ) : (
            <div>
              <div style={{padding:"7px 8px 7px 10px",borderTop:`1px solid ${C.border}`,background:"#fff"}}>
                {isDebt ? (<>
                  {/* DEBT: grid kolomnya disamain persis sama baris rincian (16px 1fr 110px 40px 110px 36px)
                      biar kalender & wallet di form tambah lurus sama yang di list, dan tombol +
                      duduk di kolom yang sama kayak ikon sampah -- gak usah nyari-nyari lagi. */}
                  <div style={{display:"grid",gridTemplateColumns:"16px 1fr 110px 40px 110px 36px",gap:8,alignItems:"center"}}>
                    <TxtInput value={nd} onChange={setNd} placeholder="detail..." style={{gridColumn:"2",gridRow:"1"}}/>
                    <RpInput value={na} onChange={setNa} style={{gridColumn:"3",gridRow:"1"}}/>
                    <button onClick={()=>setCicilanMode(v=>!v)}
                      style={{gridColumn:"2",gridRow:"2",justifySelf:"start",background:cicilanMode?"#FEE2E2":"transparent",border:`1px solid ${cicilanMode?"#FECACA":C.border}`,borderRadius:6,padding:"3px 10px",fontSize:11,color:cicilanMode?C.red:C.muted,cursor:"pointer",fontFamily:"inherit",fontWeight:600}}>
                      {cicilanMode?"✓ Cicilan":"+ Cicilan"}
                    </button>
                    <DateBtn value={nt} onChange={v=>setNt(v)} style={{gridColumn:"4",gridRow:"2"}}/>
                    <Sel value={nw} onChange={setNw} options={wOpts} style={{gridColumn:"5",gridRow:"2",fontSize:11,padding:"6px"}}/>
                    <button onClick={add} style={{gridColumn:"6",gridRow:"2",justifySelf:"center",background:C.red,border:"none",color:"#fff",borderRadius:6,cursor:"pointer",fontSize:15,width:26,height:26,display:"flex",alignItems:"center",justifyContent:"center"}}>+</button>
                  </div>
                  {cicilanMode && (
                    <div style={{display:"flex",alignItems:"center",gap:8,marginTop:6}}>
                      <span style={{fontSize:11,color:C.muted}}>Tenor:</span>
                      <Sel value={String(cicilanTenor)} onChange={v=>setCicilanTenor(Number(v))}
                        options={Array.from({length:12},(_,i)=>[(i+1).toString(),`${i+1} bulan`])}
                        style={{fontSize:11,padding:"3px 6px"}}/>
                      {num(na)>0 && <span style={{fontSize:11,color:C.muted}}>= <strong style={{color:C.red}}>{fmt(Math.round(num(na)/cicilanTenor))}/bulan</strong></span>}
                    </div>
                  )}
                </>) : (
                  /* NON-DEBT layout awal -- grid disamain persis sama baris rincian
                     (16px 1fr 110px 40px 110px 36px) biar kalender/wallet/tombol lurus. */
                  <div style={{display:"grid",gridTemplateColumns:"16px 1fr 110px 40px 110px 36px",gap:8,alignItems:"center"}}>
                    <TxtInput value={nd} onChange={setNd} placeholder="detail..." style={{gridColumn:"2"}}/>
                    <RpInput value={na} onChange={setNa} style={{gridColumn:"3"}}/>
                    <DateBtn value={nt} onChange={v=>setNt(v)} style={{gridColumn:"4"}}/>
                    <Sel value={nw} onChange={setNw} options={wOpts} style={{gridColumn:"5",fontSize:11,padding:"6px"}}/>
                    <button onClick={add} style={{gridColumn:"6",justifySelf:"center",background:C.red,border:"none",color:"#fff",borderRadius:6,cursor:"pointer",fontSize:15,width:26,height:26,display:"flex",alignItems:"center",justifyContent:"center"}}>+</button>
                  </div>
                )}
              </div>
              {isOver && (
                <div style={{margin:"6px 10px",display:"flex",gap:8}}>
                  <div style={{flex:1,fontSize:11,color:C.danger,background:"#FEF2F2",border:`1px solid #FECACA`,borderRadius:4,padding:"6px 10px",display:"flex",alignItems:"center"}}>
                    {isActualOver ? `Aktual melebihi budget sebesar ${fmt(actualNum-budgetNum)}` : `Rencana melebihi budget sebesar ${fmt(total-budgetNum)}`}
                  </div>
                  <button onClick={()=>onAdjustBudget(isActualOver?actualNum:total)}
                    style={{fontSize:11,color:C.danger,border:`1px solid #FECACA`,borderRadius:4,padding:"6px 10px",background:"none",cursor:"pointer",fontFamily:"inherit",fontWeight:600,whiteSpace:"nowrap"}}>
                    Sesuaikan Budget
                  </button>
                </div>
              )}
            </div>
          )}
        </div>

        <div style={{fontSize:10,color:C.muted,marginBottom:8}}>
          {(item.isPiutangRow||item.isPiutang)
            ? "Rincian pemasukan · Wallet wajib dipilih · Centang saat uang sudah diterima - tersimpan otomatis."
            : "Rincian pengeluaran · Wallet wajib dipilih · Centang rincian yang sudah dibayar - tersimpan otomatis."
          }
        </div>

        {confirmDelNote !== null && (
          <div style={{fontSize:11,marginBottom:10,padding:"8px 12px",background:"#FFFBEB",borderRadius:8,border:"1px solid #FDE68A",display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,flexWrap:"wrap"}}>
            <span style={{color:C.ink}}>
              {(item.isPiutangRow||item.isPiutang)
                ? <>Hapus piutang yang sudah diterima? <strong>Saldo wallet akan berkurang.</strong></>
                : <>Hapus rincian yang sudah lunas? <strong>Saldo wallet akan dikembalikan.</strong></>
              }
            </span>
            <div style={{display:"flex",gap:8,flexShrink:0}}>
              <Btn color={C.muted} outline onClick={()=>setConfirmDelNote(null)} style={{padding:"4px 10px",fontSize:11}}>Batal</Btn>
              <Btn color={C.danger} onClick={confirmDelete} style={{padding:"4px 10px",fontSize:11}}>Hapus</Btn>
            </div>
          </div>
        )}

        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8}}>
          <div style={{fontSize:12,color:C.muted}}>
            Total: <strong style={{color:C.text}}>{fmt(total)}</strong>
            <span style={{marginLeft:10}}>Lunas: <strong style={{color:C.green}}>{fmt(totalPaid)}</strong></span>
          </div>
        </div>
      </div>
    </div>
  );
};

// ---------------------------------------------
// QUICK INPUT
// ---------------------------------------------
const QuickInput = ({wallets, kategoriList, onOut, onIn, onTransfer, isMobile, defaultAction=""}) => {
  const descInputRef = useRef(null);
  const descDropRef = useRef(null);
  const [showDescDrop, setShowDescDrop] = useState(false);
  const [action, setAction] = useState(defaultAction);
  const prevDefault = useRef(defaultAction);
  useEffect(()=>{
    if (defaultAction && defaultAction!==prevDefault.current) {
      setAction(defaultAction);
      prevDefault.current = defaultAction;
    }
  },[defaultAction]);
  const [katId, setKatId] = useState(kategoriList[0]?.id||"");
  const [desc, setDesc] = useState("");
  const [amount, setAmount] = useState("");
  const [wallet, setWallet] = useState(wallets[0]?.name||"");
  const [fromW, setFromW] = useState(wallets[0]?.name||"");
  const [toW, setToW] = useState(wallets[1]?.name||"");
  const [date, setDate] = useState("");
  const [note, setNote] = useState("");
  const [payNoteId, setPayNoteId] = useState("");
  const [done, setDone] = useState(false);

  useEffect(() => {
    if (!kategoriList.find(k=>k.id===katId) && kategoriList[0]) setKatId(kategoriList[0].id);
  }, [kategoriList]);

  const currentKat = kategoriList.find(k=>k.id===katId);
  const descs = action==="out" && currentKat ? currentKat.rows.map(r=>r.desc).filter(Boolean) : [];
  const wNames = (wallets||[]).map(w=>w.name);
  const matchedRow = action==="out" && currentKat ? currentKat.rows.find(r=>r.desc===desc) : null;
  const unpaidNotes = matchedRow ? (matchedRow.notes||[]).filter(n=>!n.paid) : [];

  const changeDesc = v => { setDesc(v); setPayNoteId(""); };
  const selectPayNote = id => {
    setPayNoteId(id);
    const n = unpaidNotes.find(x=>x.id===id);
    if (n) { setAmount(toRp(num(n.amount))); if (n.wallet) setWallet(n.wallet); }
  };

  const submit = () => {
    if (!amount) return;
    if (action==="out") onOut({katId,desc,amount,wallet,date,note,payNoteId:payNoteId||null});
    else if (action==="in") onIn({desc,amount,wallet,date,note});
    else if (action==="transfer") onTransfer({from:fromW,to:toW,amount,date,note});
    setAmount(""); setDesc(""); setNote(""); setDate(""); setPayNoteId("");
    setDone(true); setTimeout(()=>setDone(false),1800);
  };

  const gcol = {display:"grid",gridTemplateColumns:isMobile?"1fr":"repeat(auto-fit,minmax(140px,1fr))",gap:10};

  return (
    <Card style={{marginBottom:16}}>
      <div style={{padding:"16px 18px"}}>
        <Lbl>Quick Input</Lbl>
        <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
          {[["out","- Pengeluaran"],["in","+ Pemasukan"],["transfer","⇌ Transfer"]].map(([k,l])=>(
            <Btn key={k} color={C.red} outline={action!==k} onClick={()=>setAction(action===k?"":k)} style={{flex:1,minWidth:90,fontSize:11,padding:"6px 8px"}}>{l}</Btn>
          ))}
        </div>
        {action && (
          <div style={{background:C.soft,borderRadius:10,padding:14,border:`1px solid ${C.border}`,overflow:"visible"}}>
            <div style={{...gcol,overflow:"visible"}}>
              {action==="out" && <>
                <div><Lbl>Kategori</Lbl><Sel value={katId} onChange={v=>{setKatId(v);setDesc("");setPayNoteId("");}} options={kategoriList.map(k=>[k.id,k.nama||"(tanpa nama)"])}/></div>
                <div><Lbl>Deskripsi</Lbl>
                  <div style={{position:"relative"}}>
                    <input ref={descInputRef} value={desc}
                      onChange={e=>{changeDesc(e.target.value);setShowDescDrop(true);}}
                      onFocus={()=>setShowDescDrop(true)}
                      onBlur={()=>setTimeout(()=>setShowDescDrop(false),300)}
                      placeholder="pilih atau ketik..."
                      style={{...baseInput,paddingRight:28}}/>
                    <span
                      onMouseDown={e=>{e.preventDefault();setShowDescDrop(v=>!v);}}
                      style={{position:"absolute",right:14,top:"50%",transform:"translateY(-50%)",
                      fontSize:10,color:C.muted,cursor:"pointer",padding:"4px",pointerEvents:"auto",zIndex:1}}>&#9660;</span>
                    {showDescDrop && descs.filter(d=>!desc||d.toLowerCase().includes(desc.toLowerCase())).length>0 && (
                      <div ref={descDropRef} style={{position:"absolute",top:"100%",left:0,right:0,
                        background:"#fff", border:`1px solid ${C.border}`,borderRadius:8,zIndex:9999,
                        boxShadow:"0 4px 12px rgba(0,0,0,0.15)",maxHeight:160,overflowY:"auto",marginTop:2}}>
                        {descs.filter(d=>!desc||d.toLowerCase().includes(desc.toLowerCase())).map(d=>(
                          <div key={d}
                            onClick={()=>{changeDesc(d);setShowDescDrop(false);}}
                            style={{padding:"8px 12px",fontSize:12,cursor:"pointer",borderBottom:`1px solid ${C.light}`,userSelect:"none"}}
                            onMouseEnter={e=>e.currentTarget.style.background=C.soft}
                            onMouseLeave={e=>e.currentTarget.style.background="#fff"}>
                            {d}
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
                {unpaidNotes.length>0 && (
                  <div><Lbl>Bayar Rincian (opsional)</Lbl>
                    <Sel value={payNoteId} onChange={selectPayNote} options={[["","+ Transaksi baru"],...unpaidNotes.map(n=>[n.id,(n.desc||"(tanpa nama)") + " - " + fmt(num(n.amount))])]}/>
                  </div>
                )}
                <div><Lbl>Jumlah</Lbl><RpInput value={amount} onChange={setAmount}/></div>
                <div><Lbl>Dari Wallet</Lbl><Sel value={wallet} onChange={setWallet} options={wNames}/></div>
              </>}
              {action==="in" && <>
                <div><Lbl>Deskripsi</Lbl><TxtInput value={desc} onChange={setDesc} placeholder="freelance, bonus..."/></div>
                <div><Lbl>Jumlah</Lbl><RpInput value={amount} onChange={setAmount}/></div>
                <div><Lbl>Ke Wallet</Lbl><Sel value={wallet} onChange={setWallet} options={wNames}/></div>
              </>}
              {action==="transfer" && <>
                <div><Lbl>Dari</Lbl><Sel value={fromW} onChange={setFromW} options={wNames}/></div>
                <div><Lbl>Ke</Lbl><Sel value={toW} onChange={setToW} options={wNames}/></div>
                <div><Lbl>Jumlah</Lbl><RpInput value={amount} onChange={setAmount}/></div>
              </>}
              <div><Lbl>Catatan</Lbl>
                <div style={{display:"flex",gap:8,alignItems:"center"}}>
                  <DateBtn value={date} onChange={v=>setDate(v)} style={{width:36,flexShrink:0}}/>
                  <TxtInput value={note} onChange={setNote} placeholder="opsional..." style={{flex:1}}/>
                </div>
              </div>
            </div>
            <div style={{display:"flex",justifyContent:"flex-end",marginTop:12}}>
              <Btn color={done?C.green:C.red} onClick={submit}>
                {done?"v Tersimpan":action==="out"?"Submit Pengeluaran":action==="in"?"Submit Pemasukan":"Transfer Sekarang"}
              </Btn>
            </div>
          </div>
        )}
      </div>
    </Card>
  );
};

// ---------------------------------------------
// SAVINGS TAB
// ---------------------------------------------
const SavingsTab = ({goals, wallets, isMobile, onAdd, onDel, onUpd, onDeposit, onWithdrawW, onWithdrawU}) => {
  const [actionFor, setActionFor] = useState(null);
  const [amount, setAmount] = useState("");
  const [wallet, setWallet] = useState("");
  const [date, setDate] = useState("");
  const [note, setNote] = useState("");
  const [done, setDone] = useState(false);
  const [expandedFor, setExpandedFor] = useState(null); // goalId yang lagi nampilin 3 button aksi
  const toggleExpand = (goalId) => {
    setExpandedFor(prev => {
      if (prev===goalId) { setActionFor(af=>af?.goalId===goalId?null:af); return null; }
      return goalId;
    });
  };

  const openAction = (goalId, mode) => {
    const goal = (goals||[]).find(g=>g.id===goalId);
    const defaultWallet = mode==="deposit"
      ? (wallets||[]).find(w=>w.isMain)?.name||""  // setor default dari Wallet Saldo
      : goal?.walletPenyimpanan || (wallets||[])[0]?.name||""; // tarik default dari wallet penyimpanan
    setActionFor({goalId,mode}); setAmount(""); setNote(""); setDate(""); setWallet(defaultWallet);
  };

  const submit = () => {
    if (!amount || !actionFor) return;
    const goal = (goals||[]).find(g=>g.id===actionFor.goalId);
    if (actionFor.mode==="deposit") {
      if (!goal?.walletPenyimpanan) { alert("Pilih wallet penyimpanan di setting goal terlebih dahulu"); return; }
      if (!wallet) { alert("Pilih wallet sumber"); return; }
      onDeposit({goalId:actionFor.goalId,amount,wallet,date,note});
    } else if (actionFor.mode==="withdraw_wallet") {
      if (!wallet) { alert("Pilih wallet tujuan"); return; }
      onWithdrawW({goalId:actionFor.goalId,amount,wallet,date,note});
    } else if (actionFor.mode==="withdraw_urgent") {
      onWithdrawU({goalId:actionFor.goalId,amount,date,note});
    }
    setDone(true); setTimeout(()=>{setDone(false);setActionFor(null);},1200);
  };

  return (
    <div>
      {goals.map(g => {
        const pct = num(g.target)>0 ? Math.min((num(g.saved)/num(g.target))*100,100) : 0;
        const isOpen = actionFor?.goalId===g.id;
        const expanded = expandedFor===g.id || isOpen;
        return (
          <Card key={g.id} style={{marginBottom:12}}>
            <div style={{padding:"14px 18px"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,marginBottom:10}} className="m2os-note-row">
                <div style={{flex:1,minWidth:0}}>
                  {(g.walletLocked||g.id===DANA_DARURAT_ID)
                    ? <div style={{fontWeight:700,fontSize:13,padding:"2px 0"}}>{g.nama||"(tanpa nama)"}</div>
                    : <TxtInput value={g.nama} onChange={v=>onUpd(g.id,"nama",v)} placeholder="..."
                        style={{fontWeight:700,fontSize:13,border:"none",background:"transparent",padding:"2px 0"}}/>
                  }
                  <div style={{fontSize:11,color:C.muted,marginTop:2}}>{fmt(num(g.saved))} / {fmt(num(g.target))}</div>
                </div>
                <div style={{display:"flex",gap:6,alignItems:"center",flexShrink:0}}>
                  {g.walletLocked
                    ? <button onClick={()=>onUpd(g.id,"walletLocked",false)} style={{background:"none",border:"none",cursor:"pointer",fontSize:16,padding:"0 2px"}}>🔒</button>
                    : <button onClick={()=>onUpd(g.id,"walletLocked",true)} style={{background:"none",border:"none",cursor:"pointer",fontSize:16,padding:"0 2px"}}>🔓</button>
                  }
                  <button onClick={g.id===DANA_DARURAT_ID?undefined:()=>onDel(g.id,'confirm')} className={g.id===DANA_DARURAT_ID?"":"m2os-note-del"}
                    style={{background:"none",border:"none",cursor:g.id===DANA_DARURAT_ID?"default":"pointer",padding:0,opacity:g.id===DANA_DARURAT_ID?0:0.25,pointerEvents:g.id===DANA_DARURAT_ID?"none":"auto",display:"flex",alignItems:"center",justifySelf:"center"}}><svg width="10" height="13" viewBox="0 0 10 13" fill="none" xmlns="http://www.w3.org/2000/svg"><polygon points="1,2 9,2 8,12 2,12" stroke="#E5E7EB" strokeWidth="1.2" fill="none" strokeLinejoin="round"/><line x1="0" y1="2" x2="10" y2="2" stroke="#E5E7EB" strokeWidth="1.2" strokeLinecap="round"/><line x1="3" y1="4.5" x2="3" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/><line x1="5" y1="4.5" x2="5" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/><line x1="7" y1="4.5" x2="7" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/></svg></button>
                </div>
              </div>
              <div style={{marginBottom:12,display:"flex",gap:8,alignItems:"flex-end"}}>
                <div style={{flex:1}}>
                  <Lbl>Target</Lbl>
                  {g.walletLocked
                    ? <div style={{fontSize:13,color:C.muted,padding:"6px 2px",borderBottom:`1px solid ${C.light}`}}>{g.target||"Rp0"}</div>
                    : <RpInput value={g.target} onChange={v=>onUpd(g.id,"target",v)}/>
                  }
                </div>
              </div>
              <div style={{marginBottom:12}}>
                <Lbl>Wallet Penyimpanan</Lbl>
                {g.walletLocked
                  ? <div style={{fontSize:12,color:C.muted,padding:"6px 2px",borderBottom:`1px solid ${C.light}`}}>{g.walletPenyimpanan||"-"}</div>
                  : <Sel value={g.walletPenyimpanan||""} onChange={v=>onUpd(g.id,"walletPenyimpanan",v)}
                      options={[["","Pilih wallet..."],...(wallets||[]).filter(w=>!w.isMain).map(w=>w.name)]}/>
                }
              </div>
              <div style={{background:C.soft,borderRadius:6,height:4,overflow:"hidden",marginTop:8,marginBottom:12}}>
                <div style={{width:`${pct}%`,background:C.red,height:"100%",borderRadius:6,transition:"width 0.4s"}}/>
              </div>
              <div onClick={()=>toggleExpand(g.id)} className="m2os-note-row" style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:expanded?12:0,cursor:"pointer",width:"100%"}}>
                <span style={{fontSize:11,color:C.muted,lineHeight:1}}>{pct.toFixed(0)}% tercapai</span>
                <svg className="m2os-note-del" width="10" height="3" viewBox="0 0 10 3" style={{flexShrink:0,display:"block",opacity:0,transition:"opacity 0.15s"}}>
                  <circle cx="1.5" cy="1.5" r="1" fill={C.faint}/>
                  <circle cx="5" cy="1.5" r="1" fill={C.faint}/>
                  <circle cx="8.5" cy="1.5" r="1" fill={C.faint}/>
                </svg>
              </div>
              {expanded && (
              <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:isOpen?12:0}}>
                <Btn color={C.red} outline={!(isOpen&&actionFor.mode==="deposit")} onClick={()=>isOpen&&actionFor.mode==="deposit"?setActionFor(null):openAction(g.id,"deposit")} style={{flex:1,minWidth:70,fontSize:11,padding:"6px 8px"}}>+ Setor</Btn>
                <Btn color={C.ink} outline={!(isOpen&&actionFor.mode==="withdraw_wallet")} onClick={()=>isOpen&&actionFor.mode==="withdraw_wallet"?setActionFor(null):openAction(g.id,"withdraw_wallet")} style={{flex:1,minWidth:90,fontSize:11,padding:"6px 8px"}}>Tarik ke Wallet</Btn>
                <Btn color={C.danger} outline={!(isOpen&&actionFor.mode==="withdraw_urgent")} onClick={()=>isOpen&&actionFor.mode==="withdraw_urgent"?setActionFor(null):openAction(g.id,"withdraw_urgent")} style={{flex:1,minWidth:90,fontSize:11,padding:"6px 8px"}}>! Ambil Urgent</Btn>
              </div>
              )}
              {isOpen && (
                <div style={{background:C.soft,borderRadius:10,padding:14,border:`1px solid ${C.border}`}}>
                  {actionFor.mode==="deposit" && (
                    <div style={{marginBottom:10,padding:"6px 10px",background:C.card,borderRadius:8,fontSize:11,color:C.muted}}>
                      Dari <strong>{wallet||"..."}</strong> {'->'} disimpan ke <strong>{g.walletPenyimpanan||"(belum dipilih)"}</strong>
                    </div>
                  )}
                  <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"repeat(auto-fit,minmax(130px,1fr))",gap:10}}>
                    <div><Lbl>Jumlah</Lbl><RpInput value={amount} onChange={setAmount}/></div>
                    {actionFor.mode==="deposit" && (
                      <div><Lbl>Dari Wallet</Lbl>
                        <Sel value={wallet} onChange={setWallet} options={(wallets||[]).filter(w=>w.name!==g.walletPenyimpanan).map(w=>w.name)}/>
                      </div>
                    )}
                    {actionFor.mode==="withdraw_wallet" && (
                      <div><Lbl>Ke Wallet</Lbl>
                        <Sel value={wallet} onChange={setWallet} options={(wallets||[]).map(w=>w.name)}/>
                      </div>
                    )}
                    <div><Lbl>{actionFor.mode==="withdraw_urgent"?"Keperluan":"Catatan"}</Lbl>
                      <div style={{display:"flex",gap:8,alignItems:"center"}}>
                        <DateBtn value={date} onChange={v=>setDate(v)} style={{width:36,flexShrink:0}}/>
                        <TxtInput value={note} onChange={setNote} placeholder={actionFor.mode==="withdraw_urgent"?"untuk apa...":"opsional..."} style={{flex:1}}/>
                      </div>
                    </div>
                  </div>
                  <div style={{display:"flex",justifyContent:"flex-end",marginTop:12}}>
                    <Btn color={done?C.green:actionFor.mode==="withdraw_urgent"?C.danger:C.red} onClick={submit}>
                      {done?"v Tersimpan":actionFor.mode==="deposit"?"Setor Sekarang":actionFor.mode==="withdraw_wallet"?"Tarik ke Wallet":"Ambil Urgent"}
                    </Btn>
                  </div>
                </div>
              )}
              {expanded && g.log?.length>0 && (
                <div style={{marginTop:12,borderTop:`1px solid ${C.light}`,paddingTop:10}}>
                  <Lbl>Riwayat</Lbl>
                  {g.log.slice(0,4).map((l,i)=>(
                    <div key={i} style={{display:"flex",justifyContent:"space-between",fontSize:11,padding:"4px 0",color:C.muted}}>
                      <span>{l.type==="deposit"?"Setor":l.type==="withdraw_wallet"?"Tarik ke wallet":"Ambil urgent"}{l.date?` · ${isoToDisplay(l.date)}`:""}{l.note?` · ${l.note}`:""}</span>
                      <span style={{fontWeight:700,color:l.type==="deposit"?C.green:C.danger}}>{l.type==="deposit"?"+":"-"}{fmt(num(l.amount))}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </Card>
        );
      })}
      <div style={{marginBottom:14}}>
        <button onClick={onAdd} style={{width:"100%",background:"none",border:`1px solid ${C.border}`,borderRadius:4,padding:"9px 0",fontSize:11,color:C.muted,cursor:"pointer",fontFamily:"inherit",fontWeight:600,letterSpacing:0.2}}>
          + Tambah Goal Savings
        </button>
      </div>
    </div>
  );
};

// ---------------------------------------------
// SPENDING INSIGHTS (rule-based)
// ---------------------------------------------
const SpendingInsights = ({kats, totalIn, totalKat, saldo, uangDingin, periods, curIdx, goals}) => {
  const insights = [];

  // 1. Over budget per kategori
  kats.forEach(k => {
    const b = k.rows.reduce((s,r)=>s+num(r.budget),0);
    const a = k.rows.reduce((s,r)=>s+num(r.actual),0);
    if (b>0 && a>b) {
      const selisih = a - b;
      insights.push({
        type:"warning",
        text:`${k.nama||"Kategori tanpa nama"} melebihi anggaran sebesar ${fmt(selisih)} (aktual ${fmt(a)} dari budget ${fmt(b)}).`
      });
    }
  });

  // 2. Spending limit total
  const usedPct = totalIn>0 ? (totalKat.a/totalIn)*100 : 0;
  if (usedPct >= 90) {
    insights.push({type:"warning", text:`Pengeluaran telah mencapai ${usedPct.toFixed(0)}% dari total pemasukan. Saldo tersisa ${fmt(saldo)}.`});
  } else if (usedPct >= 75) {
    insights.push({type:"caution", text:`Pengeluaran berada di ${usedPct.toFixed(0)}% dari total pemasukan. Pantau pengeluaran hingga akhir periode.`});
  }

  // 3. Rincian belum dibayar (masih rencana)
  let unpaidCount = 0, unpaidTotal = 0;
  kats.forEach(k => k.rows.forEach(r => {
    (r.notes||[]).forEach(n => { if (!n.paid) { unpaidCount++; unpaidTotal += num(n.amount); } });
  }));
  if (unpaidCount > 0) {
    insights.push({type:"reminder", text:`${unpaidCount} rincian senilai ${fmt(unpaidTotal)} belum ditandai lunas.`});
  }

  // 4. Tren vs periode sebelumnya
  if (curIdx > 0 && periods[curIdx-1]) {
    const prev = periods[curIdx-1];
    let prevA = 0; (prev.kategoriList||[]).forEach(k=>k.rows.forEach(r=>{prevA+=num(r.actual);}));
    if (prevA > 0 && totalKat.a > 0) {
      const delta = ((totalKat.a - prevA) / prevA) * 100;
      if (delta > 10) {
        insights.push({type:"caution", text:`Pengeluaran naik ${delta.toFixed(0)}% dibanding ${prev.periode} (${fmt(prevA)} -> ${fmt(totalKat.a)}).`});
      } else if (delta < -10) {
        insights.push({type:"positive", text:`Pengeluaran turun ${Math.abs(delta).toFixed(0)}% dibanding ${prev.periode} (${fmt(prevA)} -> ${fmt(totalKat.a)}).`});
      }
    }
  }

  // 5. Kategori paling hemat (aktual < 70% budget, budget ada)
  const hemat = kats
    .map(k=>({nama:k.nama,b:k.rows.reduce((s,r)=>s+num(r.budget),0),a:k.rows.reduce((s,r)=>s+num(r.actual),0)}))
    .filter(k=>k.b>0 && k.a>0 && k.a/k.b < 0.7)
    .sort((a,b)=>a.a/a.b - b.a/b.b);
  if (hemat.length > 0) {
    const k = hemat[0];
    insights.push({type:"positive", text:`${k.nama||"Satu kategori"} tercatat efisien  -  realisasi ${fmt(k.a)} dari anggaran ${fmt(k.b)} (${((k.a/k.b)*100).toFixed(0)}%).`});
  }

  // 6. Uang dingin
  if (uangDingin > 0 && totalIn > 0) {
    insights.push({type:"positive", text:`Dana bebas periode ini ${fmt(uangDingin)} - selisih antara pemasukan efektif dan total anggaran.`});
  }

  // 7. Apresiasi savings
  const activeGoals = goals ? goals.filter(g=>num(g.saved)>0) : [];
  if (activeGoals.length > 0) {
    const totalSaved = activeGoals.reduce((s,g)=>s+num(g.saved),0);
    const names = activeGoals.map(g=>g.nama||"tanpa nama").join(", ");
    insights.push({type:"positive", text:`Dana tersimpan di ${activeGoals.length} goal savings (${names}) sebesar ${fmt(totalSaved)}.`});
  }

  if (insights.length === 0) {
    return null;
  }

  // Prioritas: warning -> caution -> reminder -> positive, max 4
  const order = {warning:0, caution:1, reminder:2, positive:3};
  const sorted = [...insights].sort((a,b)=>order[a.type]-order[b.type]).slice(0,4);

  const cfg = {
    warning: {color:C.danger,   bg:"#FEF2F2", icon:"!"},
    caution: {color:"#D97706",  bg:"#FFFBEB", icon:"o"},
    reminder:{color:C.ink,      bg:C.soft,    icon:"o"},
    positive:{color:C.green,    bg:"#F0FDF4", icon:"v"},
  };

  return (
    <Card style={{marginBottom:14}}>
      <div style={{padding:"14px 18px 6px"}}>
        <Lbl>Spending Insights</Lbl>
      </div>
      {sorted.map((ins,i)=>{
        const s = cfg[ins.type];
        return (
          <div key={i} style={{display:"flex",gap:10,alignItems:"flex-start",padding:"9px 18px",borderTop:`1px solid ${C.light}`}}>
            <span style={{fontSize:11,color:s.color,flexShrink:0,marginTop:1,fontWeight:700}}>{s.icon}</span>
            <span style={{fontSize:11,color:C.ink,lineHeight:1.55}}>{ins.text}</span>
          </div>
        );
      })}
      <div style={{height:8}}/>
    </Card>
  );
};

// ---------------------------------------------
// BUDGET BREAKDOWN (Overview)
// ---------------------------------------------
const BudgetBreakdown = ({kats, totalInBudget}) => {
  const [expanded, setExpanded] = useState({});
  const toggle = id => setExpanded(e=>({...e,[id]:!e[id]}));

  return (
    <Card style={{marginTop:14}}>
      <div style={{padding:"16px 18px 6px"}}>
        <Lbl>Budgeting per Kategori</Lbl>
      </div>
      {kats.length===0 && <div style={{padding:"0 18px 18px",fontSize:12,color:C.muted}}>Belum ada kategori. Mulai set budget di tab Budget.</div>}
      {kats.map(k=>{
        const b = k.rows.reduce((s,r)=>s+num(r.budget),0);
        const a = k.rows.reduce((s,r)=>s+num(r.actual),0);
        const over = a>b;
        const pct = b>0 ? Math.min((a/b)*100,100) : 0;
        const isOpen = !!expanded[k.id];
        const rowsWithData = k.rows.filter(r=>r.desc || num(r.budget) || num(r.actual));
        return (
          <div key={k.id} style={{borderTop:`1px solid ${C.light}`}}>
            <div onClick={()=>toggle(k.id)} style={{padding:"10px 18px",cursor:"pointer",display:"flex",alignItems:"center",gap:10}}>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:12,fontWeight:700,display:"flex",alignItems:"baseline",gap:4,minWidth:0}}>
                  <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{k.nama||"(tanpa nama)"}</span>
                  {b>0&&totalInBudget>0?<span style={{fontSize:10,fontWeight:400,color:"#C4C4C4",flexShrink:0}}>. {Math.round(b/totalInBudget*100)}% effective</span>:null}
                </div>
                <div style={{background:C.soft,borderRadius:6,height:4,overflow:"hidden",marginTop:8}}>
                  <div style={{width:`${pct}%`,background:over?C.danger:C.red,height:"100%",borderRadius:6}}/>
                </div>
              </div>
              <div style={{textAlign:"right",flexShrink:0}}>
                <div style={{fontSize:11,fontWeight:700,color:over?C.danger:C.text}}>{fmt(a)}</div>
                <div style={{fontSize:10,color:C.muted}}>dari {fmt(b)}</div>
              </div>
              <span style={{fontSize:9,color:C.muted,flexShrink:0,display:"inline-block",transform:isOpen?"rotate(180deg)":"none",transition:"transform 0.2s"}}>v</span>
            </div>
            {isOpen && (
              <div style={{background:C.soft,padding:"2px 18px 10px"}}>
                {rowsWithData.length===0 && <div style={{fontSize:11,color:C.muted,padding:"6px 0"}}>Belum ada baris. Isi di tab Budget.</div>}
                {rowsWithData.map((r,i)=>(
                  <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"5px 0",fontSize:11,borderTop:i>0?`1px solid ${C.border}`:"none",gap:8}}>
                    <span style={{color:C.text,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.desc||"(tanpa nama)"}</span>
                    <span style={{color:num(r.actual)>num(r.budget)?C.danger:C.muted,flexShrink:0,fontWeight:600}}>{fmt(num(r.actual))} / {fmt(num(r.budget))}</span>
                  </div>
                ))}
              </div>
            )}
          </div>
        );
      })}
    </Card>
  );
};

// ---------------------------------------------
// MAIN APP
// ---------------------------------------------
export default function App() {
  const [theme, setTheme] = useState(getTheme);
  const T = THEMES[theme] || THEMES.red;
  // Override C dengan theme aktif
  C.red = T.red;
  const headerText = theme==='white' ? '#1A1A1A' : '#ffffff';
  const headerTextMuted = theme==='white' ? 'rgba(0,0,0,0.35)' : 'rgba(255,255,255,0.35)';

  const saveTheme = t => { setTheme(t); try { localStorage.setItem('kb_theme', t); } catch {} };
  const [showTheme, setShowTheme] = useState(false);

  // -- STATUS BAR COLOR -- dibikin fixed hitam biar netral ke tema apapun yang
  // dipilih user (red/pink/black/white/ocean), gak usah ikut-ikutan ganti warna.
  // Sumber utama buat status bar app yang udah di-install (standalone/WebAPK) itu
  // theme_color di manifest.json + meta tag default di index.html (keduanya udah
  // di-set hitam juga) — effect ini cuma jaga-jaga kalau dibuka lewat browser tab biasa.
  useEffect(() => {
    const color = '#000000';
    let meta = document.querySelector('meta[name="theme-color"]');
    if (!meta) {
      meta = document.createElement('meta');
      meta.setAttribute('name','theme-color');
      document.head.appendChild(meta);
    }
    meta.setAttribute('content', color);
  }, []);

  // -- SPLASH SCREEN: logo M20S + "Kalibrasi Budget" pas app dibuka, ~3 detik,
  // ukuran kecil sama kayak yang di layar opening (Welcome), biar gak keliatan
  // layar hitam/kosong sebelum data ke-load.
  const [showSplash, setShowSplash] = useState(true);
  useEffect(() => {
    const t = setTimeout(() => setShowSplash(false), 2000);
    return () => clearTimeout(t);
  }, []);

  const [tab, setTab] = useState("overview");
  const [periods, setPeriods] = useState([]);
  const [curIdx, setCurIdx] = useState(0);
  const [wallets, setWallets] = useState(defaultWallets());
  const [goals, setGoals] = useState([]);
  const [notesOpen, setNotesOpen] = useState(null);
  const [headerMetric, setHeaderMetric] = useState("totalincome");
  const [quickAction, setQuickAction] = useState(""); // "saldo" | "effective" | "wallet" | "savings"
  const [search, setSearch] = useState("");
  const [logFilter, setLogFilter] = useState("all"); // all | out | in | transfer | void
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [showNew, setShowNew] = useState(false);
  const [newName, setNewName] = useState("");
  const [confirmDel, setConfirmDel] = useState(null);
  const [showExport, setShowExport] = useState(false);
  const [exportIdx, setExportIdx] = useState(0);
  const [exporting, setExporting] = useState(false);
  const [exportErr, setExportErr] = useState("");
  const [saving, setSaving] = useState(false);
  const [lastUpdated, setLastUpdated] = useState(null);
  const [scrolled, setScrolled] = useState(false);
  const scrollRef = useRef(null);
  const sentinelRef = useRef(null);
  const scrollRafRef = useRef(null);
  // Throttle ke requestAnimationFrame biar setScrolled (yang nge-trigger re-render + animasi
  // collapse header) gak numpuk bareng kerjaan scroll browser sendiri di frame yang sama —
  // ini yang bikin kerasa patah pas lagi aktif nge-scroll.
  const handleScroll = (e) => {
    const val = e.target.scrollTop > 80;
    if (scrollRafRef.current) return;
    scrollRafRef.current = requestAnimationFrame(() => {
      scrollRafRef.current = null;
      setScrolled(val);
    });
  };
  const [loaded, setLoaded] = useState(false);
  const [isMobile, setIsMobile] = useState(false);

  useEffect(() => {
    const check = () => setIsMobile(window.innerWidth < 820);
    check(); window.addEventListener("resize",check); return ()=>window.removeEventListener("resize",check);
  },[]);

  // -- Fix: paksa repaint tiap ganti tab --
  // Bug: di beberapa browser mobile (paling sering PWA "Add to Home Screen" di iOS,
  // kadang juga di web biasa), konten tab yang baru di-mount ("wallet" paling sering
  // kena, tapi bisa kena tab manapun) suka gagal ke-render/kosong sampai user pindah
  // tab lain lalu balik lagi. Ini bug repaint di WebKit, bukan masalah data/state React
  // (state-nya udah benar dari awal, cuma browser-nya telat/gagal repaint layer-nya).
  // Fix: tiap kali `tab` ganti, paksa browser reflow+repaint area konten secara manual,
  // biar gak nunggu trigger repaint dari luar (kayak switch tab lagi).
  useLayoutEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    el.style.transform = "translateZ(0)";
    void el.offsetHeight; // force reflow
    const raf = requestAnimationFrame(() => { el.style.transform = ""; });
    return () => cancelAnimationFrame(raf);
  }, [tab]);

  // -- 3 QUICK ACTIONS (header) hide/show --
  // 2 fase, DUA-duanya animasi beneran (bukan snap instan):
  // TUTUP: (1) opacity fade 1->0 dulu (box masih full height, gak ada yang kepotong) →
  //        (2) begitu transitionend opacity kelar (udah invisible), BARU maxHeight nyusut
  //        93->0 pelan-pelan — aman di-animate smooth karena kontennya udah gak keliatan
  //        sama sekali, jadi space-nya "jalan nutup" pelan, bukan snap.
  // BUKA: kebalikannya — box grow dulu (invisible, opacity masih 0), baru fade-in abis
  //       transitionend max-height kelar.
  const qaAutoFade = tab==="wallet"||tab==="savings"||tab==="log";
  const qaHide = scrolled || qaAutoFade;
  const [qaFadeIn, setQaFadeIn] = useState(!qaHide);
  const [qaExpanded, setQaExpanded] = useState(!qaHide);
  useEffect(() => {
    if (qaHide) {
      setQaFadeIn(prev => {
        if (!prev) setQaExpanded(false); // opacity udah 0 dari awal (interupsi) -> collapse langsung, gak ada transitionend yang bakal fire
        return false;
      });
    } else {
      setQaExpanded(prev => {
        if (prev) setQaFadeIn(true); // box udah full dari awal (interupsi) -> fade-in langsung, gak ada transitionend yang bakal fire
        return true;
      });
    }
  }, [qaHide]);

  const budget = periods[curIdx] || null;
  const setBudget = fn => setPeriods(ps => ps.map((p,i)=>i===curIdx?fn(p):p));

  // -- STORAGE LOAD --
  useEffect(()=>{(async()=>{
    try {
      const [r1,r2,r3,r4,r5] = await Promise.all([
        storage.get("bt_v2"),
        storage.get("bt_wallets"),
        storage.get("bt_cur"),
        storage.get("bt_goals"),
        storage.get("bt_last_updated"),
      ]);
      if (r5) { try { setLastUpdated(new Date(r5.value)); } catch(_) {} }
      if (r1) {
        const ps = JSON.parse(r1.value);
        setPeriods((Array.isArray(ps)?ps:[]).map(p=>({
          periode: p?.periode||"Periode",
          income: {...(p?.income||{}),main:toRp(num(p?.income?.main)),mainWallet:p?.income?.mainWallet||"",lastPeriod:toRp(num(p?.income?.lastPeriod)),mainLocked:p?.income?.mainLocked||false},
          kategoriList: (Array.isArray(p?.kategoriList)&&p.kategoriList.length)
            ? injectHints(p.kategoriList.map(k=>({id:k?.id||uid(),nama:k?.nama||"",hint:k?.hint||"",rows:(k?.rows||[]).map(r=>({...r,desc:r?.desc||"",budget:toRp(num(r?.budget)),actual:toRp(num(r?.actual)),notes:(r?.notes||[]).map(n=>({...n,id:n.id||uid(),paid:n.paid||false})),isPiutangRow:r?.isPiutangRow||false,isPiutang:r?.isPiutang||false,budgetManual:r?.budgetManual||false}))})))
            : defaultKats(),
          sideIncome: (() => {
            const si = p?.sideIncome;
            if (!si) return {nama:"",rows:[]};
            // Format baru: {nama, rows}
            if (si.rows) return {nama:si.nama||"",rows:si.rows.map(r=>({desc:r?.desc||"",amount:toRp(num(r?.amount)),date:r?.date||"",wallet:r?.wallet||""}))};
            // Format lama: array flat
            if (Array.isArray(si)) return {nama:"",rows:si.map(r=>({desc:r?.desc||"",amount:toRp(num(r?.amount)),date:r?.date||"",wallet:r?.wallet||""}))};
            return {nama:"",rows:[]};
          })(),
          log: Array.isArray(p?.log)?p.log:[],
        })));
      } else {
        try {
          const old=await storage.get("bt_periods_v2");
          if (old) {
            const ps=JSON.parse(old.value);
            setPeriods((Array.isArray(ps)?ps:[]).map(p=>({
              periode:p?.periode||"Periode",
              income:{main:toRp(num(p?.income?.main)),mainWallet:p?.income?.mainWallet||"",lastPeriod:toRp(num(p?.income?.lastPeriod)),mainLocked:p?.income?.mainLocked||false},
              kategoriList:(Array.isArray(p?.kategoriList)&&p.kategoriList.length)?injectHints(p.kategoriList.map(k=>({id:k?.id||uid(),nama:k?.nama||"",hint:k?.hint||"",rows:(k?.rows||[]).map(r=>({...r,desc:r?.desc||"",budget:toRp(num(r?.budget)),actual:toRp(num(r?.actual)),notes:(r?.notes||[]).map(n=>({...n,id:n.id||uid(),paid:n.paid||false})),isPiutangRow:r?.isPiutangRow||false,isPiutang:r?.isPiutang||false,budgetManual:r?.budgetManual||false}))}))):defaultKats(),
              sideIncome: (() => { const si=p?.sideIncome; if(!si) return {nama:"",rows:[{desc:"",amount:"",date:"",wallet:""}]}; if(si.rows) return {nama:si.nama||"",rows:si.rows.map(r=>({desc:r?.desc||"",amount:toRp(num(r?.amount)),date:r?.date||"",wallet:r?.wallet||""}))}; if(Array.isArray(si)) return {nama:"",rows:si.map(r=>({desc:r?.desc||"",amount:toRp(num(r?.amount)),date:r?.date||"",wallet:r?.wallet||""}))}; return {nama:"",rows:[]}; })(),
              log:Array.isArray(p?.log)?p.log:[],
            })));
          }
        } catch(_) {}
      }
      if (r2) {
        const ws=JSON.parse(r2.value);
        const parsed=(Array.isArray(ws)?ws:[]).map(w=>({name:w?.name||"Wallet",amount:toRp(num(w?.amount)),isMain:w?.isMain||false}));
        const hasMain=parsed.some(w=>w.isMain);
        if (!hasMain && parsed.length>0) parsed[0].isMain=true;
        setWallets(parsed.length>0 ? parsed : defaultWallets());
      } else {
        setWallets(defaultWallets());
      }
      if (r3) setCurIdx(Number(r3.value)||0);
      if (r4) {
        const gs=JSON.parse(r4.value);
        setGoals(ensureDanaDarurat((Array.isArray(gs)?gs:[]).map(g=>({...g,id:g?.id||uid(),nama:g?.nama||"",target:toRp(num(g?.target)),saved:toRp(num(g?.saved)),log:Array.isArray(g?.log)?g.log:[],walletPenyimpanan:g?.walletPenyimpanan||"",walletLocked:g?.walletLocked===true||g?.walletLocked==="true"}))));
      } else {
        setGoals(ensureDanaDarurat([]));
      }
    } catch(_) {}
    setLoaded(true);
  })();},[]);

  // -- STORAGE SAVE --
  useEffect(()=>{
    if (!loaded) return;
    const t = setTimeout(async()=>{
      setSaving(true);
      try {
        await storage.set("bt_v2",JSON.stringify(periods));
        await storage.set("bt_wallets",JSON.stringify(wallets));
        await storage.set("bt_cur",String(curIdx));
        await storage.set("bt_goals",JSON.stringify(goals));
        const now = new Date();
        await storage.set("bt_last_updated", now.toISOString());
        setLastUpdated(now);
      } catch(_) {}
      setTimeout(()=>setSaving(false),600);
    },600);
    return ()=>clearTimeout(t);
  },[periods,wallets,curIdx,goals,loaded]);

  // -- CALCS --
  const kats = budget?.kategoriList||[];
  const totalKat = useMemo(()=>{let b=0,a=0;kats.forEach(k=>k.rows.forEach(r=>{if(!r.isPiutangRow&&!r.isPiutang){b+=num(r.budget);a+=num(r.actual);}}));return{b,a};},[kats]);
  const totalSide = useMemo(()=>budget?(budget.sideIncome.rows||[]).reduce((s,r)=>s+num(r.amount),0):0,[budget]);

  // Total Income untuk Overview - murni pemasukan, tidak terpengaruh transfer
  const totalInOverview = budget ? num(budget.income.main)+num(budget.income.lastPeriod)+totalSide : 0;

  // Side income yang masuk ke Wallet Saldo saja (untuk Effective)
  const mainWalletName = wallets.find(w=>w.isMain)?.name||"";
  const sideIncomeMain = budget ? (budget.sideIncome.rows||[]).reduce((s,r)=>r.wallet===mainWalletName?s+num(r.amount):s, 0) : 0;
  // Piutang net: masuk dikurangi yang di-void — exclude side income
  const piutangMain = budget ? (budget.log||[]).reduce((s,l)=>{
    if (l.type==="in" && l.isPiutang===true && !l.isSideIncome && l.wallet===mainWalletName) return s+num(l.amount);
    if (l.type==="void" && l.isPiutang===true && l.wallet===mainWalletName) return s-num(l.amount);
    return s;
  }, 0) : 0;
  // Savings deposit dari Wallet Saldo kurangi Effective, withdraw ke Wallet Saldo tambah
  const savingsNet = budget ? (budget.log||[]).reduce((s,l)=>{
    if (l.type==="savings_deposit" && l.from===mainWalletName) return s-num(l.amount);
    if (l.type==="savings_withdraw" && l.to===mainWalletName) return s+num(l.amount);
    return s;
  },0) : 0;
  const totalInEffective = budget ? num(budget.income.main)+num(budget.income.lastPeriod)+sideIncomeMain+piutangMain+savingsNet : 0;

  // Total Effective untuk Budget - totalIn + transfer masuk - transfer keluar dari Wallet Saldo
  const transferNet = useMemo(()=>{
    if (!budget) return 0;
    return (budget.log||[]).reduce((s,l)=>{
      if (l.type!=="transfer") return s;
      const toMain = wallets.find(w=>w.isMain)?.name;
      if (l.to===toMain) return s+num(l.amount);   // transfer masuk ke Wallet Saldo
      if (l.from===toMain) return s-num(l.amount); // transfer keluar dari Wallet Saldo
      return s;
    },0);
  },[budget,wallets]);
  const totalInBudget = totalInEffective;

  const saldoUtama = useMemo(()=>{ const w=wallets.find(x=>x.isMain); return w?num(w.amount):0; },[wallets]);
  const sisaSaldo = saldoUtama;
  const danaBebas = totalInBudget - totalKat.b;
  const saldo = totalInOverview - totalKat.a; // untuk modal periode baru (sisa saldo lama)
  const walletTotal = useMemo(()=>wallets.reduce((s,w)=>s+num(w.amount),0),[wallets]);
  const mainWallet = wallets.find(w=>w.isMain) || wallets[0];
  const regularWallets = wallets.filter(w=>!w.isMain);
  const allWalletNames = wallets.map(w=>w.name);
  const goalsTotal = useMemo(()=>goals.reduce((s,g)=>s+num(g.saved),0),[goals]);
  const usedPct = totalInBudget>0 ? Math.min((totalKat.a/totalInBudget)*100,100) : 0;
  const trend = useMemo(()=>periods.map(p=>{let a=0;(p.kategoriList||[]).forEach(k=>k.rows.forEach(r=>{a+=num(r.actual);}));const si=p.sideIncome;const sideTotal=Array.isArray(si)?si.reduce((s,r)=>s+num(r.amount),0):(si?.rows||[]).reduce((s,r)=>s+num(r.amount),0);const inc=num(p.income.main)+num(p.income.lastPeriod)+sideTotal;return{label:p.periode.split(" ")[0].slice(0,3),in:inc,out:a};}),[periods]);
  const allLog = useMemo(()=>budget?(budget.log||[]):[],[budget]);
  const filteredLog = useMemo(()=>{
    let logs = allLog;
    if (search) { const q=search.toLowerCase(); logs=logs.filter(l=>(l.desc||"").toLowerCase().includes(q)||(l.rowDesc||"").toLowerCase().includes(q)||(l.note||"").toLowerCase().includes(q)||(l.katNama||"").toLowerCase().includes(q)); }
    if (dateFrom) logs=logs.filter(l=>l.date&&l.date>=dateFrom);
    if (dateTo) logs=logs.filter(l=>l.date&&l.date<=dateTo);
    if (logFilter!=="all") {
      if (logFilter==="savings_deposit") logs=logs.filter(l=>l.type==="savings_deposit"||l.type==="savings_withdraw"||l.type==="savings_withdraw_urgent");
      else if (logFilter==="in") logs=logs.filter(l=>l.type==="in"||l.type==="side_income");
      else if (logFilter==="out") {
        // Exclude entri out yang sudah di-void (ada pasangan void dengan desc+rowDesc yang sama)
        const voidSet = new Set(allLog.filter(l=>l.type==="void").map(l=>`${l.rowDesc}|${l.desc}|${l.ts}`));
        // Untuk setiap out, cek apakah ada void yang lebih baru dengan desc yang sama
        const voidsByDesc = {};
        allLog.filter(l=>l.type==="void").forEach(l=>{
          const key = `${l.rowDesc}|${l.desc}`;
          if (!voidsByDesc[key]) voidsByDesc[key]=[];
          voidsByDesc[key].push(l.ts);
        });
        logs = logs.filter(l=>{
          if (l.type!=="out") return false;
          const key = `${l.rowDesc}|${l.desc}`;
          const voids = voidsByDesc[key]||[];
          // Entry out valid kalau tidak ada void yang lebih baru setelahnya
          const hasLaterVoid = voids.some(vt=>vt>l.ts);
          return !hasLaterVoid;
        });
      }
      else logs=logs.filter(l=>l.type===logFilter);
    }
    return logs;
  },[allLog,search,dateFrom,dateTo,logFilter]);

  // -- KAT UPDATERS --
  const updKatNama = (id,v) => setBudget(b=>({...b,kategoriList:b.kategoriList.map(k=>k.id===id?{...k,nama:v}:k)}));
  const addKat = () => setBudget(b=>({...b,kategoriList:[...b.kategoriList,{id:uid(),nama:"",hint:"",rows:emptyRows()}]}));
  const [confirmDelKat, setConfirmDelKat] = useState(null); // kat object
  const delKat = (id) => {
    const kat = budget?.kategoriList?.find(k=>k.id===id);
    const allPaid = (kat?.rows||[]).flatMap(r=>(r.notes||[]).filter(n=>n.paid&&n.wallet&&num(n.amount)>0).map(n=>({...n,rowDesc:r.desc,isPiutang:r.isPiutangRow||r.isPiutang||false})));
    if (allPaid.length>0) {
      setConfirmDelKat(kat);
    } else {
      execDelKat(id);
    }
  };
  const execDelKat = (id) => {
    setBudget(b=>{
      const kat = b.kategoriList.find(k=>k.id===id);
      const walletDelta = {};
      const newLog = [];
      (kat?.rows||[]).forEach(r=>{
        const isPiutang = r.isPiutangRow||r.isPiutang||false;
        (r.notes||[]).filter(n=>n.paid&&n.wallet).forEach(n=>{
          walletDelta[n.wallet]=(walletDelta[n.wallet]||0)+(isPiutang?+num(n.amount):-num(n.amount));
          newLog.push({type:"void",katNama:kat?.nama||"",rowDesc:r.desc||"",desc:n.desc||r.desc||"",amount:n.amount,wallet:n.wallet,date:n.date,note:"",ts:Date.now()});
        });
      });
      if (Object.keys(walletDelta).length) {
        setWallets(w=>w.map(r=>walletDelta[r.name]!==undefined?{...r,amount:toRp(num(r.amount)-walletDelta[r.name])}:r));
      }
      return {...b,
        kategoriList:b.kategoriList.filter(k=>k.id!==id),
        log:[...newLog,...(b.log||[])]
      };
    });
    setConfirmDelKat(null);
  };
  const updRow = (katId,idx,field,val) => setBudget(b=>({...b,kategoriList:b.kategoriList.map(k=>k.id===katId?{...k,rows:k.rows.map((r,i)=>i===idx?{...r,[field]:val,...(field==="budget"?{budgetManual:!!val}:{})}:r)}:k)}));
  const [addDebtModal, setAddDebtModal] = useState(null); // katId
  const addRow = (katId, isPiutangRow=false) => setBudget(b=>({...b,kategoriList:b.kategoriList.map(k=>k.id===katId?{...k,rows:[...k.rows,{desc:"",budget:"",actual:"",notes:[],isPiutangRow:isPiutangRow}]}:k)}));
  const [confirmDelGoal, setConfirmDelGoal] = useState(null);
  const [confirmDelSide, setConfirmDelSide] = useState(null);
  const [confirmDelRow, setConfirmDelRow] = useState(null);
  const [confirmReset, setConfirmReset] = useState(false); // {katId, idx}
  const delRow = (katId,idx) => {
    const row = budget?.kategoriList?.find(k=>k.id===katId)?.rows?.[idx];
    const paidNotes = (row?.notes||[]).filter(n=>n.paid&&n.wallet&&num(n.amount)>0);
    if (paidNotes.length>0) {
      setConfirmDelRow({katId,idx});
    } else {
      execDelRow(katId,idx);
    }
  };
  const execDelRow = (katId,idx) => {
    setBudget(b=>{
      const kat = b.kategoriList.find(k=>k.id===katId);
      const row = kat?.rows?.[idx];
      const isPiutang = row?.isPiutangRow||row?.isPiutang||false;
      const paidNotes = (row?.notes||[]).filter(n=>n.paid&&n.wallet);
      const walletDelta = {};
      const newLog = [];
      paidNotes.forEach(n=>{
        // Refund: utang = tambah wallet (delta negatif), piutang = kurangi wallet (delta positif)
        walletDelta[n.wallet]=(walletDelta[n.wallet]||0)+(isPiutang?+num(n.amount):-num(n.amount));
        newLog.push({type:"void",katNama:kat?.nama||"",rowDesc:row?.desc||"",desc:n.desc||row?.desc||"",amount:n.amount,wallet:n.wallet,date:n.date,note:"",ts:Date.now()});
      });
      if (Object.keys(walletDelta).length) {
        setWallets(w=>w.map(r=>walletDelta[r.name]!==undefined?{...r,amount:toRp(num(r.amount)-walletDelta[r.name])}:r));
      }
      return {...b,
        kategoriList:b.kategoriList.map(k=>k.id===katId?{...k,rows:k.rows.filter((_,i)=>i!==idx)}:k),
        log:[...newLog,...(b.log||[])]
      };
    });
    setConfirmDelRow(null);
  };

  // Total budget dari notes: cicilan cuma ngitung bulan yang beneran udah
  // dikonfirmasi (paid/lockedPaid) -- bukan yang cuma "active" (dipilih tapi
  // belum diceklis) ataupun total/amount mentahnya. Biar input Total pas bikin
  // cicilan baru gak langsung "ngacak" budget header sebelum ada yang dicentang.
  const notesBudgetTotal = (notesArr) => (notesArr||[]).reduce((s,n)=>{
    // Cicilan: budget cuma dari bulan yang "disentuh" periode ini -- dipilih (active)
    // atau udah dibayar (paid, BUKAN locked). Bulan abu-abu & locked (carry-over dari
    // periode lalu) otomatis kelewat, keduanya punya active:false & paid:false.
    if (n.isCicilan) return s + (n.cicilan||[]).filter(c=>c.active||c.paid).reduce((ss,c)=>ss+num(c.amount),0);
    return s + num(n.amount);
  },0);
  const isDebtKategori = (k) => !!k && (k.hint==="debt" || (k.nama||"").toLowerCase().includes("debt") || (k.nama||"").toLowerCase().includes("hutang") || (k.nama||"").toLowerCase().includes("utang"));

  // -- NOTES CHANGE (dipakai popup Simpan/Pakai Aktual & toggle centang) --
  const applyNotesChange = (katId,idx,newNotesRaw) => {
    const clean = newNotesRaw.map(({_isNew,...rest})=>rest);
    setBudget(b=>{
      const kat=b.kategoriList.find(k=>k.id===katId);
      const row=kat?.rows?.[idx];
      const isPiutang = row?.isPiutang||row?.isPiutangRow||false;
      const oldById={}; (row?.notes||[]).forEach(n=>{ if(n.id) oldById[n.id]=n; });
      const walletDelta={};
      const newLog=[];
      // Cek item yang dihapus (ada di old tapi tidak di clean)
      Object.values(oldById).forEach(old=>{
        const stillExists = clean.some(n=>n.id===old.id);
        if (!stillExists && old.paid && old.wallet) {
          // Piutang: refund = kurangi wallet. Utang: refund = tambah wallet
          walletDelta[old.wallet]=(walletDelta[old.wallet]||0)+(isPiutang?-num(old.amount):+num(old.amount));
        }
      });
      clean.forEach(n=>{
        const old = n.id ? oldById[n.id] : null;
        const wasPaid = old?.paid || false;
        const wasPartial = old?.partial || false;
        // Centang cicilan: paid atau partial baru (belum ada sebelumnya)
        const nowActive = n.paid || n.partial;
        const wasActive = wasPaid || wasPartial;
        if (nowActive && !wasActive) {
          // Sedot wallet: hitung dari bulan yang baru jadi paid (active → paid)
          if (n.isCicilan && n.wallet) {
            // Hitung amount dari bulan yang baru paid (sebelumnya active)
            const newlyPaid = (n.cicilan||[]).filter((c,ci)=>{
              const oldC = old?.cicilan?.[ci];
              return c.paid && !c.lockedPaid && !(oldC?.paid);
            });
            const paidAmount = newlyPaid.reduce((s,c)=>s+num(c.amount),0);
            if (paidAmount>0) {
              walletDelta[n.wallet]=(walletDelta[n.wallet]||0)+(isPiutang?-paidAmount:+paidAmount);
              newLog.push({type:"out",katNama:kat?.nama||"",rowDesc:row?.desc||"",desc:n.desc||row?.desc||"",amount:toRp(paidAmount),wallet:n.wallet,date:n.date,note:"",ts:Date.now()});
            }
          } else if (!n.isCicilan) {
            if (n.wallet) walletDelta[n.wallet]=(walletDelta[n.wallet]||0)+(isPiutang?-num(n.amount):+num(n.amount));
            if (isPiutang) newLog.push({type:"in",isPiutang:true,katNama:"PIUTANG DITERIMA",rowDesc:row?.desc||"",desc:n.desc||row?.desc||"",amount:n.amount,wallet:n.wallet,date:n.date,note:"",ts:Date.now()});
            else newLog.push({type:"out",katNama:kat?.nama||"",rowDesc:row?.desc||"",desc:n.desc||row?.desc||"",amount:n.amount,wallet:n.wallet,date:n.date,note:"",ts:Date.now()});
          }
        } else if (!nowActive && wasActive) {
          // Uncheck: kembalikan wallet
          if (n.isCicilan && n.wallet) {
            // Hitung amount dari bulan yang baru un-paid (paid → active)
            const newlyUnpaid = (n.cicilan||[]).filter((c,ci)=>{
              const oldC = old?.cicilan?.[ci];
              return c.active && !c.lockedPaid && oldC?.paid;
            });
            const unpaidAmount = newlyUnpaid.reduce((s,c)=>s+num(c.amount),0);
            if (unpaidAmount>0) {
              walletDelta[n.wallet]=(walletDelta[n.wallet]||0)+(isPiutang?+unpaidAmount:-unpaidAmount);
              newLog.push({type:"void",katNama:kat?.nama||"",rowDesc:row?.desc||"",desc:n.desc||row?.desc||"",amount:toRp(unpaidAmount),wallet:n.wallet,date:n.date,note:"",ts:Date.now()});
            }
          } else if (!n.isCicilan) {
            if (n.wallet) walletDelta[n.wallet]=(walletDelta[n.wallet]||0)+(isPiutang?+num(n.amount):-num(n.amount));
            newLog.push({type:"void",isPiutang:isPiutang||false,katNama:isPiutang?"PIUTANG DITERIMA":kat?.nama||"",rowDesc:row?.desc||"",desc:n.desc||row?.desc||"",amount:n.amount,wallet:n.wallet,date:n.date,note:"",ts:Date.now()});
          }
        } else if (n.paid && wasPaid && !n.isCicilan && old && num(n.amount)!==num(old.amount) && n.wallet) {
          const diff = num(n.amount) - num(old.amount);
          walletDelta[n.wallet]=(walletDelta[n.wallet]||0)+(isPiutang?-diff:+diff);
        }
      });
      if (Object.keys(walletDelta).length) {
        setWallets(w=>w.map(r=>walletDelta[r.name]!==undefined?{...r,amount:toRp(num(r.amount)-walletDelta[r.name])}:r));
      }
      const totalBudget = notesBudgetTotal(clean);
      const totalActual = clean.reduce((s,n)=>{
        if (n.isCicilan) return s + (n.cicilan||[]).filter(c=>c.paid&&!c.lockedPaid).reduce((ss,c)=>ss+num(c.amount),0);
        return s + (n.paid?num(n.amount):0);
      },0);
      const isDebtKat = isDebtKategori(kat);
      return {...b,kategoriList:b.kategoriList.map(k=>k.id===katId?{...k,rows:k.rows.map((r,i)=>{
        if (i!==idx) return r;
        // Piutang: budget selalu kosong. Row debt (utang/piutang): budget selalu ikut
        // total rincian yang beneran udah diceklis, gak boleh "dibekukan" manual --
        // biar gak ada yang bisa nge-freeze budget di angka Total padahal belum lunas.
        // Kategori non-debt: tetap hormati budgetManual seperti biasa.
        const newBudget = r.isPiutangRow ? "" : (isDebtKat ? toRp(totalBudget) : (r.budgetManual ? r.budget : toRp(totalBudget)));
        return {...r,notes:clean,budget:newBudget,actual:toRp(totalActual)};
      })}:k),log:[...newLog,...(b.log||[])]};
    });
  };

  const saveNotes = (katId,idx,notes,isPiutang) => {
    setBudget(b=>({...b,kategoriList:b.kategoriList.map(k=>{
      if (k.id!==katId) return k;
      const isDebtKat = isDebtKategori(k);
      return {...k,rows:k.rows.map((r,i)=>{
        if (i!==idx) return r;
        const totalBudget = notesBudgetTotal(notes);
        const totalActual = notes.reduce((s,n)=>{
          if (n.isCicilan) return s + (n.cicilan||[]).filter(c=>c.paid&&!c.lockedPaid).reduce((ss,c)=>ss+num(c.amount),0);
          return s + (n.paid?num(n.amount):0);
        },0);
        // Piutang: budget selalu kosong. Row debt: budget selalu ikut total rincian
        // yang beneran udah diceklis (bukan yang cuma dipilih/active atau Total mentah),
        // gak boleh dibekukan manual. Kategori non-debt: tetap hormati budgetManual.
        const newBudget = r.isPiutangRow ? "" : (isDebtKat ? toRp(totalBudget) : (r.budgetManual ? r.budget : toRp(totalBudget)));
        return {...r,notes,budget:newBudget,actual:toRp(totalActual),isPiutang:isPiutang??r.isPiutang};
      })};
    })}));
  };
  const adjustBudget = (katId,idx,newBudget) => {
    setBudget(b=>({...b,kategoriList:b.kategoriList.map(k=>k.id===katId?{...k,rows:k.rows.map((r,i)=>i===idx?{...r,budget:toRp(newBudget),budgetManual:true}:r)}:k)}));
  };

  // -- SIDE INCOME --
  // -- SIDE INCOME --
  const updSideNama = v => setBudget(b=>({...b,sideIncome:{...b.sideIncome,nama:v}}));
  const updSide = (idx,f,v) => setBudget(b=>{
    const old=b.sideIncome.rows[idx];
    if (f==="amount" && old?.wallet) {
      const diff=num(v)-num(old.amount);
      if (diff!==0) setWallets(w=>w.map(r=>r.name===old.wallet?{...r,amount:toRp(num(r.amount)+diff)}:r));
    }
    if (f==="wallet" && num(old?.amount)>0) {
      setWallets(w=>w.map(r=>{
        if(r.name===old.wallet) return{...r,amount:toRp(num(r.amount)-num(old.amount))};
        if(r.name===v) return{...r,amount:toRp(num(r.amount)+num(old.amount))};
        return r;
      }));
    }
    const rows=[...b.sideIncome.rows]; rows[idx]={...rows[idx],[f]:v};
    return{...b,sideIncome:{...b.sideIncome,rows}};
  });
  const addSide = () => setBudget(b=>({...b,sideIncome:{...b.sideIncome,rows:[...b.sideIncome.rows,{desc:"",amount:"",date:"",wallet:""}]}}));
  const delSide = idx => {
    const r=budget?.sideIncome?.rows?.[idx];
    if (num(r?.amount)>0) {
      setConfirmDelSide(idx);
    } else {
      setBudget(b=>({...b,sideIncome:{...b.sideIncome,rows:b.sideIncome.rows.filter((_,i)=>i!==idx)}}));
    }
  };
  const execDelSide = idx => {
    const r=budget?.sideIncome?.rows?.[idx];
    if(r?.wallet && num(r.amount)>0) setWallets(w=>w.map(x=>x.name===r.wallet?{...x,amount:toRp(num(x.amount)-num(r.amount))}:x));
    setBudget(b=>({...b,sideIncome:{...b.sideIncome,rows:b.sideIncome.rows.filter((_,i)=>i!==idx)},
      log:[{type:"void",katNama:"Income Add-Ons",desc:r?.desc||"",amount:r?.amount||"",wallet:r?.wallet||"",date:r?.date||"",note:"",ts:Date.now()},...(b.log||[])]}));
    setConfirmDelSide(null);
  };

  // -- WALLET --
  const updW = (idx,f,v) => setWallets(w=>w.map((r,i)=>i===idx?{...r,[f]:v}:r));

  const setMainIncome = val => setBudget(b=>{
    const oldVal = num(b.income.main);
    const newVal = num(val);
    const diff = newVal - oldVal;
    if (diff!==0 && mainWallet) {
      setWallets(w=>w.map(r=>r.isMain?{...r,amount:toRp(num(r.amount)+diff)}:r));
    }
    return {...b,income:{...b.income,main:val}};
  });
  const formatLastUpdated = (d) => {
    if (!d) return null;
    const now = new Date();
    const sameDay = d.getDate()===now.getDate()&&d.getMonth()===now.getMonth()&&d.getFullYear()===now.getFullYear();
    if (sameDay) return d.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"});
    const dd = String(d.getDate()).padStart(2,"0");
    const mm = String(d.getMonth()+1).padStart(2,"0");
    const yyyy = d.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  };
  const lockMainIncome = () => setBudget(b=>({...b,income:{...b.income,mainLocked:true}}));
  const setTabSafe = (t) => {
    if (t==="overview") setHeaderMetric("totalincome");
    else if (headerMetric==="totalincome") setHeaderMetric("saldo");
    setTab(t);
    // Reset scroll ke atas tiap ganti tab -- tanpa ini posisi scroll dari tab
    // sebelumnya kebawa ke tab baru (misal abis scroll bawah di Budget, pindah
    // ke Wallet langsung muncul dalam kondisi ke-scroll juga).
    scrollRef.current?.scrollTo({top:0});
    window.scrollTo({top:0});
  };
  // -- QUICK INPUT HANDLERS --
  const handleOut = ({katId,desc,amount,wallet,date,note,payNoteId}) => {
    setBudget(b=>{
      const kat=b.kategoriList.find(k=>k.id===katId);
      let logDesc = desc;
      const kl=b.kategoriList.map(k=>{
        if (k.id!==katId) return k;
        const rows=[...k.rows]; const idx=rows.findIndex(r=>r.desc===desc);
        if (idx>=0) {
          const row=rows[idx];
          let newNotes;
          if (payNoteId) {
            const target = row.notes.find(n=>n.id===payNoteId);
            logDesc = target?.desc || desc;
            newNotes = row.notes.map(n=>n.id===payNoteId?{...n,amount,wallet,date:date||n.date,paid:true}:n);
          } else {
            const noteDesc = note||desc;
            logDesc = note ? note : desc;
            newNotes = [...row.notes,{id:uid(),desc:noteDesc,amount,date,wallet,paid:true}];
          }
          const totalActual = newNotes.reduce((s,n)=>s+(n.paid?num(n.amount):0),0);
          rows[idx]={...row,notes:newNotes,actual:toRp(totalActual)};
        } else {
          rows.push({desc,budget:"",actual:toRp(num(amount)),notes:[{id:uid(),desc:note||desc,amount,date,wallet,paid:true}]});
        }
        return {...k,rows};
      });
      return {...b,kategoriList:kl,log:[{type:"out",katNama:kat?.nama||"",rowDesc:desc,desc:logDesc,amount,wallet,date,note,ts:Date.now()},...(b.log||[])]};
    });
    setWallets(w=>w.map(r=>r.name===wallet?{...r,amount:toRp(num(r.amount)-num(amount))}:r));
  };
  const handleInMain = ({amount,wallet,date}) => {
    setWallets(w=>w.map(r=>r.name===wallet?{...r,amount:toRp(num(r.amount)+num(amount))}:r));
    setBudget(b=>({...b,log:[{type:"in",desc:"Pemasukan Utama",amount,wallet,date,note:"",ts:Date.now()},...(b.log||[])]}));
  };
  const handleIn = ({desc,amount,wallet,date,note}) => {
    setBudget(b=>{
      const rows = b.sideIncome.rows;
      const idx = rows.findIndex(r=>r.desc && r.desc.toLowerCase()===desc.toLowerCase());
      let newRows;
      if (idx>=0) {
        newRows = rows.map((r,i)=>i===idx?{...r,amount:toRp(num(r.amount)+num(amount)),date,wallet}:r);
      } else {
        newRows = [...rows, {desc,amount,date,wallet}];
      }
      // Update wallet atomik dalam callback yang sama
      setWallets(w=>w.map(r=>r.name===wallet?{...r,amount:toRp(num(r.amount)+num(amount))}:r));
      return {...b,sideIncome:{...b.sideIncome,rows:newRows},log:[{type:"side_income",isSideIncome:true,desc,amount,wallet,date,note,ts:Date.now()},...(b.log||[])]};
    });
  };
  const handleTransfer = ({from,to,amount,date,note}) => {
    setWallets(w=>w.map(r=>{if(r.name===from)return{...r,amount:toRp(num(r.amount)-num(amount))};if(r.name===to)return{...r,amount:toRp(num(r.amount)+num(amount))};return r;}));
    setBudget(b=>({...b,log:[{type:"transfer",from,to,amount,date,note,ts:Date.now()},...(b.log||[])]}));
  };

  // -- GOALS --
  const addGoal = () => setGoals(g=>[...g,{id:uid(),nama:"",target:"",saved:"",log:[],walletPenyimpanan:"",walletLocked:false}]);
  const execReset = () => {
    const namaPeriode = budget?.periode || new Date().toLocaleString("id-ID",{month:"long",year:"numeric"});
    const blank = blankPeriod(namaPeriode);
    setBudget(() => blank);
    setWallets(prev=>prev.map(w=>({...w,amount:"Rp0"})));
    setGoals([]);
    setConfirmReset(false);
  };

  const exportData = () => {
    const data = { periods, wallets, goals, curIdx, exportedAt: new Date().toISOString() };
    const blob = new Blob([JSON.stringify(data, null, 2)], {type:"application/json"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `kalibrasi-budget-${new Date().toLocaleDateString("id-ID").replace(/\//g,"-")}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const importData = e => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const data = JSON.parse(ev.target.result);
        if (data.periods) setPeriods(data.periods);
        if (data.wallets) setWallets(data.wallets);
        if (data.goals) setGoals(data.goals);
        if (typeof data.curIdx === "number") setCurIdx(data.curIdx);
        alert("Data berhasil diimport!");
      } catch {
        alert("File tidak valid. Pastikan file JSON dari Kalibrasi Budget.");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  const delGoal = (id, mode) => {
    if (id===DANA_DARURAT_ID) return; // permanen, gak bisa dihapus
    const goal = goals.find(g=>g.id===id);
    if (mode==='confirm' && num(goal?.saved)>0) {
      setConfirmDelGoal(goal);
    } else {
      setGoals(g=>g.filter(x=>x.id!==id));
      setConfirmDelGoal(null);
    }
  };
  const updGoal = (id,f,v) => {
    const updated = goals.map(x=>x.id===id?{...x,[f]:v}:x);
    setGoals(updated);
    // Force save langsung kalau yang diubah adalah walletLocked atau walletPenyimpanan
    if (f==="walletLocked"||f==="walletPenyimpanan") {
      try { storage.set("bt_goals", JSON.stringify(updated)); } catch(_) {}
    }
  };
  const goalDeposit = ({goalId,amount,wallet,date,note}) => {
    const goal = goals.find(g=>g.id===goalId);
    const walletPenyimpanan = goal?.walletPenyimpanan;
    setGoals(g=>g.map(x=>x.id===goalId?{...x,saved:toRp(num(x.saved)+num(amount)),log:[{type:"deposit",amount,wallet,date,note,ts:Date.now()},...(x.log||[])]}:x));
    // Kurangi wallet asal
    setWallets(w=>w.map(r=>r.name===wallet?{...r,amount:toRp(num(r.amount)-num(amount))}:r));
    // Tambah ke wallet penyimpanan kalau ada dan bukan wallet yang sama
    if (walletPenyimpanan && walletPenyimpanan!==wallet) {
      setWallets(w=>w.map(r=>r.name===walletPenyimpanan?{...r,amount:toRp(num(r.amount)+num(amount))}:r));
    }
    // Catat ke budget log - Effective turun kalau dari Wallet Saldo
    setBudget(b=>({...b,log:[{type:"savings_deposit",from:wallet,to:walletPenyimpanan||"savings",goalNama:goal?.nama||"",amount,date,note,ts:Date.now()},...(b.log||[])]}));
  };
  const goalWithdrawW = ({goalId,amount,wallet,date,note}) => {
    const goal = goals.find(g=>g.id===goalId);
    const walletPenyimpanan = goal?.walletPenyimpanan;
    setGoals(g=>g.map(x=>x.id===goalId?{...x,saved:toRp(Math.max(0,num(x.saved)-num(amount))),log:[{type:"withdraw_wallet",amount,wallet,date,note,ts:Date.now()},...(x.log||[])]}:x));
    // Tambah ke wallet tujuan
    setWallets(w=>w.map(r=>r.name===wallet?{...r,amount:toRp(num(r.amount)+num(amount))}:r));
    // Kurangi wallet penyimpanan kalau ada dan bukan wallet yang sama
    if (walletPenyimpanan && walletPenyimpanan!==wallet) {
      setWallets(w=>w.map(r=>r.name===walletPenyimpanan?{...r,amount:toRp(num(r.amount)-num(amount))}:r));
    }
    // Catat ke budget log
    setBudget(b=>({...b,log:[{type:"savings_withdraw",from:walletPenyimpanan||"savings",to:wallet,goalNama:goal?.nama||"",amount,date,note,ts:Date.now()},...(b.log||[])]}));
  };
  const goalWithdrawU = ({goalId,amount,date,note}) => {
    const goal = goals.find(g=>g.id===goalId);
    const walletPenyimpanan = goal?.walletPenyimpanan;
    // Kurangi saldo wallet penyimpanan
    if (walletPenyimpanan) setWallets(w=>w.map(r=>r.name===walletPenyimpanan?{...r,amount:toRp(num(r.amount)-num(amount))}:r));
    // Masuk Activity Log
    setBudget(b=>({...b,log:[{type:"savings_withdraw_urgent",from:walletPenyimpanan||"savings",goalNama:goal?.nama||"",amount,date,note,ts:Date.now()},...(b.log||[])]}));
    setGoals(g=>g.map(x=>x.id===goalId?{...x,saved:toRp(Math.max(0,num(x.saved)-num(amount))),log:[{type:"withdraw_urgent",amount,date,note,ts:Date.now()},...(x.log||[])]}:x));
  };

  // -- PERIOD --
  const startNew = () => {
    if (!newName.trim()) return;
    const np = blankPeriod(newName.trim());
    np.income.lastPeriod = saldoUtama!==0 ? toRp(saldoUtama) : "";
    if (budget) {
      // Copy struktur kategori
      np.kategoriList = budget.kategoriList.map(k=>{
        const isDebt = k.hint==="debt" || k.nama.toLowerCase().includes("debt") || k.nama.toLowerCase().includes("hutang")||k.nama.toLowerCase().includes("utang");
        const newRows = k.rows.map(r=>{
          // Carry over kalau isDebt ATAU ada cicilan yang beneran masih ada sisa
          // (sisa dihitung dari Total dikurangi yang udah settled -- BUKAN dari slot bulan yang belum dicentang,
          // soalnya semua slot bisa aja udah dicentang tapi total-nya masih kurang, lihat fix togglePaid).
          const hasCicilan = (r.notes||[]).some(n=>{
            if (!n.isCicilan) return false;
            const paidSoFar = (n.cicilan||[]).filter(c=>c.paid||c.lockedPaid).reduce((s,c)=>s+num(c.amount),0);
            return num(n.totalUtang) - paidSoFar > 0;
          });
          if (isDebt || hasCicilan) {
            const unpaidNotes = (r.notes||[]).filter(n=>(!n.paid || n.partial) && (isDebt || n.isCicilan)).map(n=>{
              if (!n.isCicilan) return {...n,_isNew:false};
              // Carry over: bulan yang paid → lockedPaid (hijau locked), active/abu → reset
              const newCicilan = (n.cicilan||[]).map(c=>{
                if (c.paid) return {...c,paid:false,active:false,lockedPaid:true};
                return {...c,active:false};
              });
              const totalPaidSoFar = newCicilan.filter(c=>c.lockedPaid).reduce((s,c)=>s+num(c.amount),0);
              const sisaAmount = Math.max(0, num(n.totalUtang) - totalPaidSoFar);
              if (sisaAmount<=0) return null;
              return {...n,_isNew:false,cicilan:newCicilan,paid:false,partial:false,
                amount:"",totalUtang:n.totalUtang,tenor:n.tenor,isCicilan:true,_sisaAmount:sisaAmount};
            }).filter(Boolean);
            const unpaidTotal = unpaidNotes.reduce((s,n)=>{
              if (n.isCicilan) return s + (n._sisaAmount||0);
              return s + num(n.amount);
            }, 0);
            // Bersihkan _sisaAmount sebelum simpan
            const cleanNotes = unpaidNotes.map(({_sisaAmount,...rest})=>rest);
            const isPiutangRow = r.isPiutangRow||r.isPiutang||false;
            return {
              desc:r.desc, budget:"", actual:"",
              notes:cleanNotes, budgetManual:false,
              isPiutangRow, isPiutang:isPiutangRow
            };
          }
          return {desc:r.desc,budget:"",actual:"",notes:[],budgetManual:false};
        });
        return {id:uid(),nama:k.nama,hint:k.hint||"",rows:newRows};
      });
      np.sideIncome = {
        nama: budget.sideIncome.nama||"",
        rows: (budget.sideIncome.rows||[]).map(r=>({desc:r.desc,amount:"",date:"",wallet:""}))
      };
      // Flag kalau ada debt carry over
      const hasDebtCarryOver = budget.kategoriList.some(k=>{
        const isDebt = k.hint==="debt" || k.nama.toLowerCase().includes("debt") || k.nama.toLowerCase().includes("hutang")||k.nama.toLowerCase().includes("utang");
        return isDebt && k.rows.some(r=>(r.notes||[]).some(n=>!n.paid));
      });
      if (hasDebtCarryOver) np._hasDebtCarryOver = true;
    }
    // Wallet Saldo diisi dari sisa Wallet Saldo periode lalu
    setWallets(w=>w.map(x=>x.isMain?{...x,amount:saldoUtama>0?toRp(saldoUtama):""}:x));
    setPeriods(ps=>[...ps,np]); setCurIdx(periods.length); setShowNew(false); setNewName("");
  };
  const delPeriod = idx => {
    setPeriods(ps=>ps.filter((_,i)=>i!==idx));
    setCurIdx(ci=>idx<ci?ci-1:idx===ci?Math.max(0,idx-1):ci);
    setConfirmDel(null);
  };

  // -- EXPORT --
  const exportExcel = async () => {
    setExporting(true); setExportErr("");
    try {
      const p=periods[Number(exportIdx)]; if (!p){setExportErr("Periode tidak ditemukan.");setExporting(false);return;}
      let tb=0,ta=0; (p.kategoriList||[]).forEach(k=>k.rows.forEach(r=>{tb+=num(r.budget);ta+=num(r.actual);}));
      const ts=(budget?.sideIncome?.rows||[]).reduce((s,r)=>s+num(r.amount),0);
      const ti=num(p.income.main)+num(p.income.lastPeriod)+ts;
      const wb=XLSX.utils.book_new();
      const ws2=XLSX.utils.aoa_to_sheet([["Kategori","Deskripsi","Budget","Aktual"],...(p.kategoriList||[]).flatMap(k=>k.rows.filter(r=>r.desc||num(r.budget)||num(r.actual)).map(r=>[k.nama,r.desc,num(r.budget),num(r.actual)]))]);
      ws2["!cols"]=[{wch:20},{wch:30},{wch:14},{wch:14}]; XLSX.utils.book_append_sheet(wb,ws2,"Kategori");
      const ws3=XLSX.utils.aoa_to_sheet([["Tipe","Deskripsi","Jumlah","Wallet","Tanggal","Catatan"],...(p.log||[]).map(l=>[l.type==="out"?"Pengeluaran":l.type==="in"?"Pemasukan":"Transfer",l.type==="transfer"?`${l.from}->${l.to}`:(l.desc||""),num(l.amount),l.wallet||"",l.date?isoToDisplay(l.date):"",l.note||""])]);
      ws3["!cols"]=[{wch:14},{wch:26},{wch:14},{wch:16},{wch:14},{wch:24}]; XLSX.utils.book_append_sheet(wb,ws3,"Log");
      XLSX.writeFile(wb,`Rekap_${p.periode.replace(/\s+/g,"_")}.xlsx`);
      setShowExport(false);
    } catch(e){setExportErr(e?.message||"Gagal.");}
    setExporting(false);
  };

  const exportImage = () => {
    setExporting(true); setExportErr("");
    try {
      const p=periods[Number(exportIdx)]; if(!p){setExportErr("Periode tidak ditemukan.");setExporting(false);return;}
      let tb=0,ta=0; (p.kategoriList||[]).forEach(k=>k.rows.forEach(r=>{tb+=num(r.budget);ta+=num(r.actual);}));
      const ts=(p.sideIncome?.rows||[]).reduce((s,r)=>s+num(r.amount),0);
      const ti=num(p.income.main)+num(p.income.lastPeriod)+ts;
      const kb=(p.kategoriList||[]).map(k=>({nama:k.nama||"...",a:k.rows.reduce((s,r)=>s+num(r.actual),0),b:k.rows.reduce((s,r)=>s+num(r.budget),0),items:k.rows.filter(r=>r.desc).map(r=>({desc:r.desc,a:num(r.actual),b:num(r.budget)}))})).filter(k=>k.nama!=="..."||k.items.length>0);
      const W=720, katH=26, itmH=20, hdrH=200, ftrH=60;
      const totalItems=kb.reduce((s,k)=>s+k.items.length,0);
      const H=hdrH+50+kb.length*katH+totalItems*itmH+kb.length*8+ftrH;
      const canvas=document.createElement("canvas"); canvas.width=W; canvas.height=H;
      const ctx=canvas.getContext("2d");
      const fp=n=>"Rp"+Math.round(n).toLocaleString("id-ID");
      ctx.fillStyle="#FAFAFA"; ctx.fillRect(0,0,W,H);
      ctx.fillStyle=C.red; ctx.fillRect(0,0,W,hdrH);
      ctx.fillStyle="#fff"; ctx.font="13px sans-serif"; ctx.fillText(`Rekap Periode ${p.periode}`,28,34);
      const cw4=(W-56)/4;
      // Baris 1: Total Masuk, Effective, Total Keluar, Sisa Saldo
      [[ti,"Total Masuk"],[totalInEffective,"Effective"],[ta,"Total Keluar"],[saldoUtama,"Sisa Saldo"]].forEach(([v,l],i)=>{
        const x=28+i*cw4;
        ctx.fillStyle="rgba(255,255,255,0.55)"; ctx.font="10px sans-serif"; ctx.fillText(l.toUpperCase(),x,72);
        ctx.fillStyle="#fff"; ctx.font="bold 15px sans-serif"; ctx.fillText(fp(v),x,92);
      });
      // Baris 2: Budget, Dana Bebas, Wallet, Savings
      [[tb,"Budget"],[totalInEffective-tb,"Dana Bebas"],[walletTotal,"Wallet"],[goalsTotal,"Savings"]].forEach(([v,l],i)=>{
        const x=28+i*cw4;
        ctx.fillStyle="rgba(255,255,255,0.55)"; ctx.font="10px sans-serif"; ctx.fillText(l.toUpperCase(),x,124);
        ctx.fillStyle="#fff"; ctx.font="bold 15px sans-serif"; ctx.fillText(fp(v),x,144);
      });
      let y=hdrH+26;
      ctx.fillStyle=C.text; ctx.font="bold 14px sans-serif"; ctx.fillText("Breakdown per Kategori",28,y); y+=16;
      ctx.strokeStyle="#E5E5E5"; ctx.beginPath(); ctx.moveTo(28,y); ctx.lineTo(W-28,y); ctx.stroke(); y+=16;
      kb.forEach(k=>{
        const over=k.a>k.b;
        ctx.fillStyle=C.text; ctx.font="bold 13px sans-serif"; ctx.fillText(k.nama,28,y);
        ctx.fillStyle=over?C.danger:C.text; const t=`${fp(k.a)} / ${fp(k.b)}`; const tw=ctx.measureText(t).width; ctx.fillText(t,W-28-tw,y); y+=katH;
        k.items.forEach(it=>{
          ctx.fillStyle="#8A8A8A"; ctx.font="12px sans-serif"; ctx.fillText(it.desc||"...",44,y);
          ctx.fillStyle=it.a>it.b&&it.b>0?C.danger:"#595959";
          const aktual = it.a>0 ? fp(it.a) : "-";
          const budget = it.b>0 ? fp(it.b) : "-";
          const it2 = it.b>0 ? `${aktual} / ${budget}` : (it.a>0 ? fp(it.a) : "");
          const itw=ctx.measureText(it2).width; ctx.fillText(it2,W-28-itw,y); y+=itmH;
        });
        y+=8;
      });
      ctx.strokeStyle="#E5E5E5"; ctx.beginPath(); ctx.moveTo(28,y+8); ctx.lineTo(W-28,y+8); ctx.stroke();
      ctx.fillStyle="#8A8A8A"; ctx.font="11px sans-serif"; ctx.fillText(`Dibuat ${new Date().toLocaleDateString("id-ID")}    .    Kalibrasi Ulang Cara Kamu Pegang Uang`,28,y+28);
      const a=document.createElement("a"); a.download=`Rekap_${p.periode.replace(/\s+/g,"_")}.png`; a.href=canvas.toDataURL("image/png"); document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setShowExport(false);
    } catch(e){setExportErr(e?.message||"Gagal membuat gambar.");}
    setExporting(false);
  };

  // -- DERIVED --
  const openNote = (notesOpen&&budget) ? kats.find(k=>k.id===notesOpen.katId)?.rows?.[notesOpen.idx] : null;
  const NAV = [["overview","Overview"],["budget","Kalibrasi Budget"],["wallet","Wallet"],["savings","Savings"],["log","Activity"]];

  // -- SPLASH -- logo M20S doang, clean white (gak pakai gradient), muncul ~2
  // detik pas app pertama dibuka.
  if (showSplash) return (
    <div style={{minHeight:"100vh",background:"#fff",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontFamily:"system-ui,sans-serif"}}>
      <Logo height={13}/>
    </div>
  );

  // -- LOADING --
  if (!loaded) return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:C.bg,fontSize:14,color:C.muted,fontFamily:"sans-serif"}}>
      Memuat data...
    </div>
  );

  // -- WELCOME --
  if (periods.length===0) return (
    <div style={{minHeight:"100vh",background:C.bg,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:32,fontFamily:"system-ui,sans-serif",position:"relative",overflow:"hidden"}}><div style={{position:"absolute",bottom:-160,left:"50%",transform:"translateX(-50%)",width:800,height:400,background:"radial-gradient(ellipse at center, rgba(196,30,30,0.4) 0%, transparent 70%)",pointerEvents:"none",zIndex:0}}/>
      {showNew && <Modal>
        <div style={{fontWeight:700,fontSize:15,marginBottom:6}}>Periode Baru</div>
        <div style={{fontSize:12,color:C.muted,marginBottom:14}}>Beri nama periode pertamamu.</div>
        <Lbl>Nama Periode</Lbl>
        <TxtInput value={newName} onChange={setNewName} placeholder="contoh: Agustus 2026"/>
        <div style={{display:"flex",gap:8,marginTop:16,justifyContent:"flex-end"}}>
          <Btn color={C.muted} outline onClick={()=>setShowNew(false)}>Batal</Btn>
          <Btn color={THEMES.red.red} onClick={()=>{if(!newName.trim())return;setPeriods([blankPeriod(newName.trim())]);setCurIdx(0);setShowNew(false);setNewName("");}}>Mulai</Btn>
        </div>
      </Modal>}
      <div style={{textAlign:"center",maxWidth:300,position:"relative",zIndex:1,transform:"translateY(-36px)"}}>
        <div style={{marginBottom:6,display:"flex",justifyContent:"center"}}><Logo height={13}/></div>
        <div style={{fontFamily:"Cormorant Garamond, serif",fontWeight:700,fontStyle:"italic",fontSize:26,color:C.text,lineHeight:1.1,marginBottom:8}}>KALIBRASI BUDGET</div>
        <div style={{fontSize:13,color:C.muted,marginBottom:28,lineHeight:1.8,whiteSpace:"nowrap"}}>
          Langkah pertama menuju keuangan yang lebih sadar<br/>dimulai dari sini.
        </div>
        <Btn onClick={()=>setShowNew(true)} color={THEMES.red.red} style={{fontSize:13,padding:"10px 24px"}}>+ Mulai Budgeting</Btn>
      </div>
    </div>
  );

  // ---------------------------------------------
  // MAIN RENDER
  // ---------------------------------------------
  return (
    <div style={{fontFamily:"system-ui,sans-serif",background:"#F2F2F7",height:"100vh",display:"flex",alignItems:"center",justifyContent:"center",color:C.text}}>
    <div style={{
      width:isMobile?"100%":"60%",
      height:isMobile?"100vh":"calc(100vh - 24px)",
      display:"flex",
      flexDirection:isMobile?"column":"row",
      background:isMobile?"transparent":C.card,
      borderRadius:isMobile?0:4,
      overflow:isMobile?"hidden":"hidden",
      boxShadow:isMobile?"none":"0 8px 40px rgba(0,0,0,0.10)",
      position:"relative"
    }}>
      <style>{`
        .m2os-date-input::-webkit-calendar-picker-indicator {
          position: absolute; inset: 0; width: 100%; height: 100%;
          margin: 0; padding: 0; opacity: 0; cursor: pointer;
        }
        .m2os-date-input::-webkit-datetime-edit { color: transparent; }
        .m2os-period-item:hover .m2os-period-del { opacity: 1 !important; width: auto !important; padding-left: 4px !important; }
        .m2os-note-row:hover .m2os-note-del { opacity: 1 !important; }
        @media (hover: none) { .m2os-note-del { opacity: 0.25 !important; } }
        .m2os-period-scroll { overflow-x: auto; }
        .m2os-period-scroll::-webkit-scrollbar { display: none; }
        .m2os-period-scroll { -ms-overflow-style: none; scrollbar-width: none; }
        @import url("https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@1,700&display=swap"); button:focus { outline: none; }
        body::-webkit-scrollbar { display: none; }
        body { -ms-overflow-style: none; scrollbar-width: none; }
        *::-webkit-scrollbar { display: none; }
        * { -ms-overflow-style: none; scrollbar-width: none; }
        input:focus { outline: none !important; box-shadow: none !important; -webkit-appearance: none; background-color: white !important; }
        .m2os-input:focus { border-color: #C41E1E !important; }
        .m2os-datalist-input::-webkit-calendar-picker-indicator { display: none !important; opacity: 0; }
        .m2os-datalist-input::-webkit-list-button { display: none !important; }
        input::selection { background: rgba(196,30,30,0.15); color: #1A1A1A; }
        input::placeholder { color: #BFBFBF; }
        .m2os-popup-scroll::-webkit-scrollbar { display: none; }
        .m2os-popup-scroll { -ms-overflow-style: none; scrollbar-width: none; }
        input:focus { color: #1A1A1A; }
      `}</style>

      {/* -- POPUPS -- */}
      {/* -- QUICK INPUT POPUP (modal tengah) -- */}
      {quickAction && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={()=>setQuickAction("")}>
          <div className="m2os-popup-scroll" style={{width:"92vw",maxWidth:420,maxHeight:"85vh",overflowY:"auto",overflowX:"visible"}} onClick={e=>e.stopPropagation()}>
            <QuickInput wallets={wallets} kategoriList={kats}
              onOut={v=>{handleOut(v);setQuickAction("");}}
              onIn={v=>{handleIn(v);setQuickAction("");}}
              onTransfer={v=>{handleTransfer(v);setQuickAction("");}}
              isMobile={true} defaultAction={quickAction}
              onClose={()=>setQuickAction("")}/>
          </div>
        </div>
      )}

      {/* -- MODAL HAPUS GOAL SAVINGS -- */}
      {confirmDelGoal && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={()=>setConfirmDelGoal(null)}>
          <div style={{background:C.card,borderRadius:14,padding:24,width:300,maxWidth:"90vw",boxShadow:"0 16px 40px rgba(0,0,0,0.2)"}} onClick={e=>e.stopPropagation()}>
            <div style={{fontWeight:700,fontSize:14,marginBottom:4}}>Hapus goal "{confirmDelGoal.nama}"?</div>
            <div style={{fontSize:11,color:C.muted,marginBottom:12}}>
              Goal ini punya tabungan <strong>{fmt(num(confirmDelGoal.saved))}</strong> yang tersimpan di wallet penyimpanan. Menghapus goal tidak otomatis mengembalikan uang ke Wallet Saldo.
            </div>
            <div style={{display:"flex",gap:8}}>
              <button onClick={()=>setConfirmDelGoal(null)} style={{flex:1,background:C.soft,border:"none",borderRadius:8,padding:"8px 0",fontSize:12,fontWeight:600,color:C.muted,cursor:"pointer",fontFamily:"inherit"}}>Batal</button>
              <button onClick={()=>delGoal(confirmDelGoal.id)} style={{flex:1,background:C.danger,border:"none",borderRadius:8,padding:"8px 0",fontSize:12,fontWeight:600,color:"#fff",cursor:"pointer",fontFamily:"inherit"}}>Ya, Hapus</button>
            </div>
          </div>
        </div>
      )}

      {/* -- MODAL HAPUS SIDE INCOME -- */}
      {confirmDelSide !== null && (() => {
        const r = budget?.sideIncome?.rows?.[confirmDelSide];
        return (
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={()=>setConfirmDelSide(null)}>
            <div style={{background:C.card,borderRadius:14,padding:24,width:300,maxWidth:"90vw",boxShadow:"0 16px 40px rgba(0,0,0,0.2)"}} onClick={e=>e.stopPropagation()}>
              <div style={{fontWeight:700,fontSize:14,marginBottom:4}}>Hapus "{r?.desc||"pemasukan ini"}"?</div>
              <div style={{fontSize:11,color:C.muted,marginBottom:12}}>
                Pemasukan <strong>{fmt(num(r?.amount))}</strong> dari {r?.wallet||"wallet"} akan dihapus.{' '}
                {isWalletSaldo
                  ? <strong style={{color:"#E57373"}}>Saldo Wallet Utama akan berkurang.</strong>
                  : <strong style={{color:"#E57373"}}>Saldo {r?.wallet} akan berkurang.</strong>
                }
                {isWalletSaldo && <span style={{color:"#E57373"}}> Effective juga akan turun.</span>}
              </div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>setConfirmDelSide(null)} style={{flex:1,background:C.soft,border:"none",borderRadius:8,padding:"8px 0",fontSize:12,fontWeight:600,color:C.muted,cursor:"pointer",fontFamily:"inherit"}}>Batal</button>
                <button onClick={()=>execDelSide(confirmDelSide)} style={{flex:1,background:C.danger,border:"none",borderRadius:8,padding:"8px 0",fontSize:12,fontWeight:600,color:"#fff",cursor:"pointer",fontFamily:"inherit"}}>Ya, Hapus</button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* -- MODAL HAPUS ROW -- */}
      {confirmDelRow && (() => {
        const row = budget?.kategoriList?.find(k=>k.id===confirmDelRow.katId)?.rows?.[confirmDelRow.idx];
        const paidNotes = (row?.notes||[]).filter(n=>n.paid&&n.wallet);
        return (
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={()=>setConfirmDelRow(null)}>
            <div style={{background:C.card,borderRadius:14,padding:24,width:300,maxWidth:"90vw",boxShadow:"0 16px 40px rgba(0,0,0,0.2)"}} onClick={e=>e.stopPropagation()}>
              <div style={{fontWeight:700,fontSize:14,marginBottom:4}}>Hapus "{row?.desc||"(tanpa nama)"}"?</div>
              <div style={{fontSize:11,color:C.muted,marginBottom:12}}>Ini akan mempengaruhi budget, saldo, dan effective.</div>
              {paidNotes.length>0 && (
                <div style={{background:C.soft,borderRadius:8,padding:"8px 12px",marginBottom:12,fontSize:11}}>
                  <div style={{fontWeight:600,color:C.danger,marginBottom:4}}>{paidNotes.length} transaksi akan dibatalkan:</div>
                  {paidNotes.map((n,i)=>(
                    <div key={i} style={{display:"flex",justifyContent:"space-between",color:C.muted,marginBottom:2}}>
                      <span>{n.desc||row?.desc||"(tanpa nama)"}</span>
                      <span>{fmt(num(n.amount))}</span>
                    </div>
                  ))}
                </div>
              )}
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>setConfirmDelRow(null)} style={{flex:1,background:C.soft,border:"none",borderRadius:8,padding:"8px 0",fontSize:12,fontWeight:600,color:C.muted,cursor:"pointer",fontFamily:"inherit"}}>Batal</button>
                <button onClick={()=>execDelRow(confirmDelRow.katId,confirmDelRow.idx)} style={{flex:1,background:C.danger,border:"none",borderRadius:8,padding:"8px 0",fontSize:12,fontWeight:600,color:"#fff",cursor:"pointer",fontFamily:"inherit"}}>Ya, Hapus</button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* -- MODAL HAPUS KATEGORI -- */}
      {confirmDelKat && (() => {
        const allPaid = (confirmDelKat?.rows||[]).flatMap(r=>(r.notes||[]).filter(n=>n.paid&&n.wallet).map(n=>({...n,rowDesc:r.desc})));
        return (
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={()=>setConfirmDelKat(null)}>
            <div style={{background:C.card,borderRadius:14,padding:24,width:320,maxWidth:"90vw",maxHeight:"80vh",overflowY:"auto",boxShadow:"0 16px 40px rgba(0,0,0,0.2)"}} onClick={e=>e.stopPropagation()}>
              <div style={{fontWeight:700,fontSize:14,marginBottom:4}}>Hapus kategori "{confirmDelKat?.nama||"(tanpa nama)"}"?</div>
              <div style={{fontSize:11,color:C.muted,marginBottom:12}}>Semua baris dan transaksi di kategori ini akan dihapus.</div>
              {allPaid.length>0 && (
                <div style={{background:C.soft,borderRadius:8,padding:"8px 12px",marginBottom:12,fontSize:11}}>
                  <div style={{fontWeight:600,color:C.danger,marginBottom:6}}>{allPaid.length} transaksi akan dibatalkan:</div>
                  {allPaid.map((n,i)=>(
                    <div key={i} style={{display:"flex",justifyContent:"space-between",color:C.muted,marginBottom:3}}>
                      <div>
                        <span style={{fontWeight:500}}>{n.rowDesc||"-"}</span>
                        {n.desc&&n.desc!==n.rowDesc&&<span style={{color:C.faint}}>    .    {n.desc}</span>}
                      </div>
                      <span>{fmt(num(n.amount))}</span>
                    </div>
                  ))}
                </div>
              )}
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>setConfirmDelKat(null)} style={{flex:1,background:C.soft,border:"none",borderRadius:8,padding:"8px 0",fontSize:12,fontWeight:600,color:C.muted,cursor:"pointer",fontFamily:"inherit"}}>Batal</button>
                <button onClick={()=>execDelKat(confirmDelKat.id)} style={{flex:1,background:C.danger,border:"none",borderRadius:8,padding:"8px 0",fontSize:12,fontWeight:600,color:"#fff",cursor:"pointer",fontFamily:"inherit"}}>Ya, Hapus</button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* -- MODAL UTANG/PIUTANG -- */}
      {addDebtModal && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={()=>setAddDebtModal(null)}>
          <div style={{background:C.card,borderRadius:14,padding:24,width:280,maxWidth:"90vw",boxShadow:"0 16px 40px rgba(0,0,0,0.2)"}} onClick={e=>e.stopPropagation()}>
            <div style={{fontWeight:700,fontSize:14,marginBottom:4}}>Tambah Baris</div>
            <div style={{fontSize:11,color:C.muted,marginBottom:16}}>Pilih jenis transaksi untuk baris ini:</div>
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              <button onClick={()=>{addRow(addDebtModal,false);setAddDebtModal(null);}}
                style={{background:C.soft,border:`1px solid ${C.border}`,borderRadius:10,padding:"10px 14px",cursor:"pointer",fontFamily:"inherit",textAlign:"left"}}>
                <div style={{fontWeight:600,fontSize:12,marginBottom:2}}>Utang</div>
                <div style={{fontSize:10,color:C.muted}}>Uang yang kamu pinjam dari orang lain - saldo keluar saat dilunasi</div>
              </button>
              <button onClick={()=>{addRow(addDebtModal,true);setAddDebtModal(null);}}
                style={{background:C.soft,border:`1px solid ${C.border}`,borderRadius:10,padding:"10px 14px",cursor:"pointer",fontFamily:"inherit",textAlign:"left"}}>
                <div style={{fontWeight:600,fontSize:12,marginBottom:2}}>Piutang</div>
                <div style={{fontSize:10,color:C.muted}}>Uang kamu yang dipinjam orang lain - saldo masuk saat dibayar balik</div>
              </button>
            </div>
            <button onClick={()=>setAddDebtModal(null)} style={{marginTop:12,width:"100%",background:"none",border:"none",color:C.muted,fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>Batal</button>
          </div>
        </div>
      )}

      {notesOpen&&openNote&&(
        <NotesPopup item={openNote} wallets={wallets} isMobile={isMobile}
          liveBudget={openNote.budget}
          liveActual={openNote.actual}
          isDebt={(() => { const k=kats.find(k=>k.id===notesOpen.katId); return k?.hint==="debt"||k?.nama?.toLowerCase().includes("debt")||k?.nama?.toLowerCase().includes("hutang"); })()}
          sisaSaldo={sisaSaldo}
          danaBebas={danaBebas}
          onClose={()=>setNotesOpen(null)}
          onSave={(n,ip)=>saveNotes(notesOpen.katId,notesOpen.idx,n,ip)}
          onTogglePaid={(n)=>applyNotesChange(notesOpen.katId,notesOpen.idx,n)}
          onAdjustBudget={(v)=>adjustBudget(notesOpen.katId,notesOpen.idx,v)}/>
      )}
      {confirmDel!==null&&<Modal>
        <div style={{fontWeight:700,fontSize:15,marginBottom:6}}>Hapus Periode?</div>
        <div style={{fontSize:12,color:C.muted,marginBottom:16}}>Periode <strong>{periods[confirmDel]?.periode}</strong> akan terhapus permanen.</div>
        <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
          <Btn color={C.muted} outline onClick={()=>setConfirmDel(null)}>Batal</Btn>
          <Btn color={C.danger} onClick={()=>delPeriod(confirmDel)}>Hapus</Btn>
        </div>
      </Modal>}
      {showNew&&<Modal>
        <div style={{fontWeight:700,fontSize:15,marginBottom:6}}>Periode Baru</div>
        <div style={{fontSize:12,color:C.muted,marginBottom:14}}>Saldo & wallet terbawa otomatis ke periode baru.</div>
        <Lbl>Nama Periode</Lbl>
        <TxtInput value={newName} onChange={setNewName} placeholder="contoh: Agustus 2026"/>
        {saldoUtama>0 && (
          <div style={{marginTop:10,padding:"8px 12px",background:C.soft,borderRadius:8,fontSize:11,color:C.muted}}>
            Sisa saldo periode ini <strong style={{color:C.green}}>{fmt(saldoUtama)}</strong> akan masuk sebagai "Sisa Periode Lalu".
          </div>
        )}
        {wallets.length>0 && (
          <div style={{marginTop:8,padding:"8px 12px",background:C.soft,borderRadius:8,fontSize:11,color:C.muted}}>
            <div style={{fontWeight:600,marginBottom:4}}>Saldo wallet terbawa:</div>
            {wallets.map((w,i)=>(
              <div key={i} style={{display:"flex",justifyContent:"space-between",marginTop:2}}>
                <span>{w.name}</span>
                <strong style={{color:C.text}}>{fmt(num(w.amount))}</strong>
              </div>
            ))}
            
          </div>
        )}
        {(() => {
          const debtItems = (budget?.kategoriList||[]).flatMap(k=>{
            const isDebt = k.hint==="debt"||k.nama.toLowerCase().includes("debt")||k.nama.toLowerCase().includes("hutang")||k.nama.toLowerCase().includes("utang");
            if (!isDebt) return [];
            return k.rows.flatMap(r=>{
              const isPiutang = r.isPiutangRow||r.isPiutang||false;
              return (r.notes||[]).filter(n=>!n.paid||n.partial).map(n=>{
                let sisaAmount;
                if (n.isCicilan) {
                  // Sisa = Total dikurangi yang udah settled (paid/locked) -- bukan cuma dari slot
                  // bulan yang belum dicentang, soalnya bisa aja semua slot udah dicentang tapi
                  // total-nya masih kurang (lihat fix togglePaid untuk status "partial" akibat shortfall).
                  const totalPaid = (n.cicilan||[]).filter(c=>c.paid||c.lockedPaid).reduce((s,c)=>s+num(c.amount),0);
                  sisaAmount = Math.max(0, num(n.totalUtang) - totalPaid);
                  return {katNama:k.nama||"DEBT",rowDesc:r.desc,desc:n.desc,amount:toRp(sisaAmount),totalUtang:n.totalUtang,paidSoFar:toRp(totalPaid),isPiutang,isCicilan:true};
                }
                return {katNama:k.nama||"DEBT",rowDesc:r.desc,desc:n.desc,amount:n.amount,isPiutang};
              });
            });
          });
          if (!debtItems.length) return null;
          const totalDebt = debtItems.reduce((s,n)=>s+num(n.amount),0);
          return (
            <div style={{marginTop:8,padding:"8px 12px",background:"#FFFBEB",borderRadius:8,border:"1px solid #FDE68A",fontSize:11}}>
              <div style={{fontWeight:600,marginBottom:4,color:C.ink}}>Terbawa ke periode baru:</div>
              {debtItems.slice(0,4).map((n,i)=>(
                <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:4,color:C.muted}}>
                  <div style={{display:"flex",flexDirection:"column",gap:1}}>
                    <div style={{display:"flex",alignItems:"center",gap:6}}>
                      <span>{n.desc||n.rowDesc||"(tanpa nama)"}</span>
                    </div>
                    <div style={{fontSize:9,color:C.faint}}>{n.isPiutang?"Piutang":"Utang"}</div>
                    {n.isCicilan && n.paidSoFar && num(n.paidSoFar)>0 && (
                      <span style={{fontSize:9,color:"#22C55E"}}>Terbayar: {fmt(num(n.paidSoFar))} dari {fmt(num(n.totalUtang))}</span>
                    )}
                  </div>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontWeight:700,color:n.isPiutang?"#059669":C.danger}}>{fmt(num(n.amount))}</div>
                    {n.isCicilan && <div style={{fontSize:9,color:C.faint}}>sisa</div>}
                  </div>
                </div>
              ))}
              {debtItems.length>4 && <div style={{color:C.faint,marginTop:2}}>+{debtItems.length-4} lainnya...</div>}
              <div style={{marginTop:6,color:C.ink,fontWeight:600}}>Total sisa: {fmt(totalDebt)}</div>
            </div>
          );
        })()}
        {budget?.kategoriList?.length>0 && (
          <div style={{marginTop:8,padding:"8px 12px",background:C.soft,borderRadius:8,fontSize:11,color:C.muted}}>
            <div style={{fontWeight:600,marginBottom:4}}>Struktur kategori di-copy:</div>
            {budget.kategoriList.map((k,i)=>(
              <div key={i} style={{marginTop:2,color:C.muted}}>
                   .    {k.nama||`(${k.hint||"tanpa nama"})`}
                <span style={{color:C.faint}}>  -  {k.rows.filter(r=>r.desc).length} baris</span>
              </div>
            ))}
            <div style={{marginTop:6,color:C.faint,fontSize:10}}>Angka budget & aktual dikosongkan, tinggal isi ulang.</div>
          </div>
        )}
        <div style={{display:"flex",gap:8,marginTop:16,justifyContent:"flex-end"}}>
          <Btn color={C.muted} outline onClick={()=>setShowNew(false)}>Batal</Btn>
          <Btn onClick={startNew}>Mulai Periode</Btn>
        </div>
      </Modal>}
      {showExport&&<Modal>
        <div style={{fontWeight:700,fontSize:15,marginBottom:6}}>Export Rekap</div>
        <div style={{fontSize:12,color:C.muted,marginBottom:14}}>Pilih periode yang mau didownload.</div>
        <Lbl>Periode</Lbl>
        <Sel value={String(exportIdx)} onChange={v=>setExportIdx(Number(v))} options={periods.map((p,i)=>[String(i),p.periode])}/>
        {exportErr&&<div style={{fontSize:11,color:C.danger,marginTop:8}}>{exportErr}</div>}
        <div style={{display:"flex",gap:8,marginTop:16,justifyContent:"flex-end",flexWrap:"wrap"}}>
          <Btn color={C.muted} outline onClick={()=>setShowExport(false)}>Batal</Btn>
          <Btn color={C.ink} onClick={exportImage}>{exporting?"...":"Gambar"}</Btn>
          <Btn onClick={exportExcel}>{exporting?"...":"Excel"}</Btn>
        </div>
      </Modal>}

      {/* -- SIDEBAR (desktop only) -- */}
      {!isMobile && (
        <div style={{width:208,background:C.card,borderRight:`1px solid ${C.border}`,padding:"20px 14px",display:"flex",flexDirection:"column",flexShrink:0}}>
          {NAV.map(([k,l])=>(
            <button key={k} onClick={()=>setTab(k)} style={{display:"flex",alignItems:"center",padding:"9px 12px",borderRadius:8,border:"none",background:"transparent",color:tab===k?C.red:C.ink,fontSize:13,fontWeight:tab===k?700:500,cursor:"pointer",fontFamily:"inherit",marginBottom:3,textAlign:"left"}}>
              {l}
            </button>
          ))}
          <div style={{marginTop:22,padding:"0 6px"}}><Lbl>Periode</Lbl></div>
          <div style={{display:"flex",flexDirection:"column",gap:4,marginBottom:8}}>
            {periods.map((p,i)=>(
              <div key={i} style={{display:"flex",alignItems:"center"}} className="m2os-note-row">
                <button onClick={()=>setCurIdx(i)} style={{flex:1,textAlign:"left",background:"transparent",border:"none",padding:"6px 10px",fontSize:12,fontWeight:i===curIdx?700:500,color:i===curIdx?C.red:C.muted,cursor:"pointer",fontFamily:"inherit"}}>{p.periode}</button>
                <button onClick={()=>setConfirmDel(i)} className="m2os-note-del" style={{background:"none",border:"none",padding:"0 8px",cursor:"pointer",opacity:0,transition:"opacity 0.15s",display:"flex",alignItems:"center"}}><svg width="10" height="13" viewBox="0 0 10 13" fill="none" xmlns="http://www.w3.org/2000/svg">
  <polygon points="1,2 9,2 8,12 2,12" stroke="#E5E7EB" strokeWidth="1.2" fill="none" strokeLinejoin="round"/>
  <line x1="0" y1="2" x2="10" y2="2" stroke="#E5E7EB" strokeWidth="1.2" strokeLinecap="round"/>
  <line x1="3" y1="4.5" x2="3" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
  <line x1="5" y1="4.5" x2="5" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
  <line x1="7" y1="4.5" x2="7" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
</svg></button>
              </div>
            ))}
          </div>
          <button onClick={()=>setShowNew(true)} style={{background:"transparent",border:"none",padding:"6px 10px",fontSize:11,color:C.faint,cursor:"pointer",fontFamily:"inherit"}}>+ Periode Baru</button>
          <div style={{marginTop:"auto",paddingTop:16,fontSize:10,color:C.faint}}>{saving?"Menyimpan...":"Tersimpan"}</div>
        </div>
      )}

      {/* -- KONTEN KANAN (mobile full, desktop flex-1) -- */}
      <div style={{flex:1,display:"flex",flexDirection:"column",minWidth:0,overflow:"hidden",position:"relative"}}>
        {/* -- HEADER MERAH (semua device) -- */}
        <div style={{background:theme==='white'?'#FFFFFF':`linear-gradient(180deg, ${T.dark} 0%, ${T.red} 100%)`,paddingBottom:isMobile?0:16,paddingTop:"calc(env(safe-area-inset-top) + 2px)",flexShrink:0,position:"sticky",top:0,zIndex:40}}>
          {/* Angka besar + pill selector */}
          <div style={{textAlign:"center",padding:"12px 20px 0",paddingBottom:(scrolled||(tab==="wallet"||tab==="savings"||tab==="log"))?12:0}}>
            {/* Dulu titik 3 tema ada di sini (ketutupan notch/dynamic island di iPhone).
                Sekarang dipindah ke row judul tab (antara nama tab & tombol Export).
                Div ini dibiarkan kosong supaya jarak/spacing di atas "3 selector" tidak berubah. */}
            <div style={{marginBottom:12}}/>
            {/* 3 selector */}
            <div style={{display:"flex",justifyContent:"center",alignItems:"center",gap:20,marginBottom:12}}>
              {[
                ...(tab==="overview"?[["totalincome","Total Income"]]:[]),
                ["saldo","Sisa Saldo"],
                ["effective","Effective"],
                ["wallet","Wallet"],
                ["savings","Savings"],
              ].map(([k,l])=>(
                <button key={k} onClick={()=>setHeaderMetric(k)}
                  style={{background:"none",border:"none",padding:"0",
                  fontSize:11,fontWeight:headerMetric===k?700:400,
                  color:headerMetric===k?headerText:headerTextMuted,
                  cursor:"pointer",fontFamily:"inherit"}}>
                  {l}
                </button>
              ))}
            </div>
            {/* Angka besar */}
            <div style={{fontSize:42,fontWeight:700,color:headerText,letterSpacing:-1.5,lineHeight:1,marginTop:12,marginBottom:17}}>
              {headerMetric==="saldo"?fmt(sisaSaldo):headerMetric==="effective"?fmt(totalInBudget):headerMetric==="wallet"?fmt(walletTotal):headerMetric==="totalincome"?fmt(totalInOverview):fmt(goalsTotal)}
            </div>

            {/* Budget    .    Aktual    .    Dana Bebas row */}
            <div style={{display:"flex",justifyContent:"center",gap:20,marginBottom:5,marginTop:5}}>
              {[["Budget",fmt(totalKat.b),"rgba(255,255,255,0.5)"],["Aktual",fmt(totalKat.a),totalKat.a>totalKat.b?"#FCA5A5":"rgba(255,255,255,0.5)"],["Dana Bebas",fmt(danaBebas),danaBebas<0?"#FCA5A5":"rgba(255,255,255,0.5)"]].map(([l,v,col])=>(
                <div key={l} style={{textAlign:"center"}}>
                  <div style={{fontSize:8,color:theme==="white"?"rgba(0,0,0,0.3)":"rgba(255,255,255,0.3)",fontWeight:600,letterSpacing:0.5,marginBottom:2}}>{l.toUpperCase()}</div>
                  <div style={{fontSize:12,fontWeight:600,color:theme==="white"&&col==="rgba(255,255,255,0.5)"?"rgba(0,0,0,0.6)":col}}>{v}</div>
                </div>
              ))}
            </div>
          </div>

          {/* 3 Quick Actions */}
          <div style={{
            maxHeight:qaExpanded?93:0,
            overflow:"hidden",
            transition:"max-height 0.22s ease"}}
            onTransitionEnd={(e)=>{ if(e.propertyName==="max-height" && !qaHide) setQaFadeIn(true); }}>
            <div>
              <div style={{display:"flex",justifyContent:"center",gap:32,padding:"14px 20px 16px",
                pointerEvents:qaHide?"none":"auto",
                opacity:qaFadeIn?1:0,
                transition:"opacity 0.18s ease"}}
                onTransitionEnd={(e)=>{ if(e.propertyName==="opacity" && qaHide) setQaExpanded(false); }}>
            {[["- Keluar","out"],["+ Masuk","in"],["Transfer","transfer"]].map(([l,k])=>(
              <button key={k} onClick={()=>setQuickAction(quickAction===k?"":k)} style={{background:"none",border:"none",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:6}}>
                <div style={{width:44,height:44,borderRadius:22,background:theme==="white"?(quickAction===k?"rgba(0,0,0,0.12)":"rgba(0,0,0,0.06)"):(quickAction===k?"rgba(255,255,255,0.35)":"rgba(255,255,255,0.15)"),display:"flex",alignItems:"center",justifyContent:"center",transition:"background 0.2s"}}>
                  <span style={{fontSize:16,color:theme==="white"?"#1A1A1A":"#fff",display:"flex",alignItems:"center",justifyContent:"center",lineHeight:1}}>{k==="out"?"-":k==="in"?"+":"⇌"}</span>
                </div>
                <span style={{fontSize:10,color:theme==="white"?(quickAction===k?"#1A1A1A":"rgba(0,0,0,0.5)"):(quickAction===k?"#fff":"rgba(255,255,255,0.7)"),fontWeight:quickAction===k?700:600,whiteSpace:"nowrap"}}>{l}</span>
              </button>
            ))}
            </div>
            </div>
          </div>

          {/* Period selector - mobile only */}
          {isMobile && (
          <div style={{background:theme==="white"?"rgba(0,0,0,0.04)":"rgba(0,0,0,0.15)",padding:"8px 12px",display:"flex",alignItems:"center"}}>
            <div className="m2os-period-scroll" style={{display:"flex",alignItems:"center",gap:0}}>
              {periods.map((p,i)=>(
                <div key={i} style={{display:"flex",alignItems:"center",flexShrink:0}}>
                  {i>0 && <span style={{color:theme==="white"?"rgba(0,0,0,0.2)":"rgba(255,255,255,0.25)",margin:"0 6px",fontSize:10}}>·</span>}
                  <div style={{display:"flex",alignItems:"center",flexShrink:0}} className="m2os-period-item">
                  <button onClick={()=>setCurIdx(i)} style={{background:"transparent",border:"none",padding:"2px 0",fontSize:11,fontWeight:i===curIdx?700:400,color:theme==="white"?(i===curIdx?"rgba(0,0,0,0.65)":"rgba(0,0,0,0.4)"):(i===curIdx?"#fff":"rgba(255,255,255,0.35)"),cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap"}}>{p.periode}</button>
                  <button onClick={()=>setConfirmDel(i)} className="m2os-period-del" style={{background:"none",border:"none",padding:"0",fontSize:10,color:theme==="white"?"rgba(0,0,0,0.3)":"rgba(255,255,255,0.3)",cursor:"pointer",lineHeight:1,opacity:0,width:0,overflow:"hidden",transition:"opacity 0.15s, width 0.15s, padding 0.15s"}}>×</button>
                  </div>
                </div>
              ))}
              <span style={{color:theme==="white"?"rgba(0,0,0,0.2)":"rgba(255,255,255,0.25)",margin:"0 6px",fontSize:10}}>·</span>
              <button onClick={()=>setShowNew(true)} style={{background:"transparent",border:"none",padding:"2px 0",fontSize:11,color:theme==="white"?"rgba(0,0,0,0.4)":"rgba(255,255,255,0.35)",cursor:"pointer",fontFamily:"inherit",flexShrink:0,whiteSpace:"nowrap"}}>+ Periode</button>
            </div>
          </div>
          )}
        </div>



      {/* -- MAIN CONTENT -- */}
      <div ref={scrollRef} onScroll={handleScroll} className="m2os-scroll-area" style={{flex:1,padding:"16px 14px 120px",minWidth:0,overflowY:"auto",minHeight:0}}>
        {/* Sentinel untuk detect scroll */}
        <div ref={sentinelRef} style={{height:1,pointerEvents:"none"}}/>

        {/* Tab title */}
        <div style={{marginBottom:14,display:"grid",gridTemplateColumns:"1fr auto 1fr",alignItems:"center",gap:8}}>
          <div>
            <div style={{fontWeight:600,fontSize:17}}>{NAV.find(n=>n[0]===tab)?.[1]}</div>
            <div style={{fontSize:11,color:C.muted,marginTop:2}}>Periode {budget.periode}</div>
          </div>

          {/* Titik 3 ganti tema - cuma tampil di tab Overview.
              Popup "Pilih Tema": dropdown nempel persis di bawah tombol titik 3 (bukan
              center ke content area lagi) - wrapper-nya position:relative jadi anchor,
              dropdown-nya position:absolute top:100% + sedikit gap, center horizontal
              relatif ke tombol. Sama treatment-nya di desktop & mobile, jadi gak perlu lagi
              itung-itungan tinggi header/sidebar. */}
          {tab==="overview" && (
          <div style={{justifySelf:"center",position:"relative"}}>
            <button onClick={()=>setShowTheme(v=>!v)}
              style={{background:"none",border:"none",cursor:"pointer",padding:"2px 8px",
              color:C.muted,fontSize:16,lineHeight:1,letterSpacing:2,display:"inline-block"}}>
              ···
            </button>
            {showTheme && (<>
              <div style={{position:"fixed",inset:0,zIndex:99}} onClick={()=>setShowTheme(false)}/>
              <div style={{
                position:"absolute",
                top:"calc(100% + 6px)",
                left:"50%",transform:"translateX(-50%)",background:"#fff",borderRadius:12,
                boxShadow:"0 8px 32px rgba(0,0,0,0.2)",zIndex:100,minWidth:180,overflow:"hidden"}}>
                <div style={{padding:"12px 16px",fontSize:12,fontWeight:700,color:"#8A8A8A",borderBottom:"1px solid #F0F0F0"}}>Pilih Tema</div>
                {[["red","Original Red"],["black","Deep Black"],["white","Clean White"],["pink","Pink Rose"],["ocean","Ocean Blue"]].map(([t,l])=>(
                  <button key={t} onClick={()=>{saveTheme(t);setShowTheme(false);}}
                    style={{display:"flex",alignItems:"center",gap:12,width:"100%",padding:"12px 16px",
                    background:theme===t?"#F5F5F5":"#fff",border:"none",cursor:"pointer",fontFamily:"inherit",
                    fontSize:13,color:"#1A1A1A",textAlign:"left"}}>
                    <span style={{width:12,height:12,borderRadius:"50%",boxSizing:"border-box",background:t==="white"?"#FFFFFF":THEMES[t].red,flexShrink:0,display:"inline-block",border:t==="white"?"1.5px solid #D0D0D0":"none"}}/>
                    {l}
                    {theme===t && <span style={{marginLeft:"auto",color:THEMES[t].red,fontWeight:700}}>✓</span>}
                  </button>
                ))}
              </div>
            </>)}
          </div>
          )}

          <div style={{justifySelf:"end"}}>
            {tab==="overview"&&<Btn outline color={C.red} onClick={()=>{setExportIdx(curIdx);setShowExport(true);}} style={{padding:"5px 10px",fontSize:11}}>Export</Btn>}
          </div>
        </div>

        {/* -- OVERVIEW -- */}
        {/* Banner debt carry over */}
        {budget?._hasDebtCarryOver && (
          <div style={{background:"#FFFBEB",border:"1px solid #FDE68A",borderRadius:10,padding:"10px 16px",marginBottom:14,display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,flexWrap:"wrap"}}>
            <span style={{fontSize:12,color:C.ink}}>Ada hutang yang terbawa dari periode lalu - cek kategori <strong>DEBT</strong>.</span>
            <button onClick={()=>{setBudget(b=>({...b,_hasDebtCarryOver:false}));setTab("budget");}} style={{background:"none",border:"none",fontSize:11,color:C.muted,cursor:"pointer",textDecoration:"underline",fontFamily:"inherit",flexShrink:0}}>Lihat</button>
          </div>
        )}

        {tab==="overview" && (
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1.3fr 1fr",gap:14}}>
            <div>
              {/* Compact financial rows */}
              <Card style={{padding:"6px 0",marginBottom:14}}>
                {[
                  ["Sisa Saldo", fmt(sisaSaldo), sisaSaldo>=0?C.green:C.danger],
                  ["Total Spending", fmt(totalKat.a), C.danger],
                  ["Budget", fmt(totalKat.b), C.text],
                  ["Dana Bebas", fmt(danaBebas), danaBebas>=0?C.text:C.danger],
                ].map(([l,v,col],i,arr)=>(
                  <div key={l} style={{display:"flex",justifyContent:"space-between",alignItems:"center",
                    padding:"10px 16px",
                    borderBottom:i<arr.length-1?`1px solid ${C.light}`:"none"}}>
                    <span style={{fontSize:11,color:C.muted,fontWeight:400}}>{l}</span>
                    <span style={{fontSize:12,fontWeight:500,color:col}}>{v}</span>
                  </div>
                ))}
              </Card>

              {/* Spending limit */}
              <Card style={{padding:16,marginBottom:14}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                  <span style={{fontSize:11,color:C.muted}}>Spending limit</span>
                  <span style={{fontSize:11,fontWeight:600,color:usedPct>85?C.danger:C.muted}}>{fmt(totalKat.a)} / {fmt(totalInBudget)}</span>
                </div>
                <div style={{background:C.soft,borderRadius:6,height:4,overflow:"hidden"}}>
                  <div style={{width:`${usedPct}%`,background:usedPct>85?C.danger:C.red,height:"100%",borderRadius:6,transition:"width 0.4s"}}/>
                </div>
              </Card>
              <Card style={{padding:18,marginBottom:14}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
                  <Lbl>Profit and Loss</Lbl>
                  <div style={{display:"flex",gap:12,fontSize:10}}>
                    <span style={{color:C.green}}>* Pemasukan</span>
                    <span style={{color:C.danger}}>* Pengeluaran</span>
                  </div>
                </div>
                {trend.length>0 ? <BarChart data={trend}/> : <div style={{textAlign:"center",color:C.muted,fontSize:12,padding:40}}>Grafik muncul setelah ada lebih dari satu periode.</div>}
              </Card>
              <SpendingInsights kats={kats} totalIn={totalInBudget} totalKat={totalKat} saldo={sisaSaldo} uangDingin={danaBebas} periods={periods} curIdx={curIdx} goals={goals}/>
              <BudgetBreakdown kats={kats} totalInBudget={totalInBudget}/>
            </div>
            <div>
              <Card style={{padding:18,marginBottom:14}}>
                <Lbl>Wallets    .    {wallets.length}</Lbl>
                {wallets.map((w,i)=>(
                  <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"9px 0",borderBottom:i<wallets.length-1?`1px solid ${C.light}`:"none"}}>
                    <span style={{fontSize:12,fontWeight:600}}>{w.name}</span>
                    <span style={{fontSize:12,fontWeight:700}}>{fmt(num(w.amount))}</span>
                  </div>
                ))}
              </Card>
              <Card style={{padding:18,marginBottom:14}}>
                <Lbl>Spending by Category</Lbl>
                {(() => {
                  const catRamp = themeRamp(T.red, Math.max(kats.length,1));
                  return (
                    <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:14,marginTop:10}}>
                      <Donut size={132} data={kats.map((k,i)=>({label:k.nama,value:k.rows.reduce((s,r)=>s+num(r.actual),0),color:catRamp[i]}))}/>
                      <div style={{width:"100%"}}>
                        {kats.map((k,i)=>{const v=k.rows.reduce((s,r)=>s+num(r.actual),0);return v>0?(
                          <div key={k.id} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 0",fontSize:11,borderTop:i>0?`1px solid ${C.light}`:"none"}}>
                            <div style={{width:8,height:8,borderRadius:2,background:catRamp[i],flexShrink:0}}/>
                            <span style={{flex:1,color:C.muted,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{k.nama||"(tanpa nama)"}</span>
                            <span style={{fontWeight:700}}>{fmt(v)}</span>
                          </div>
                        ):null;})}
                      </div>
                    </div>
                  );
                })()}
              </Card>
              <Card>
                <div style={{padding:"16px 18px 10px",display:"flex",justifyContent:"space-between"}}>
                  <Lbl>Recent Activities</Lbl>
                  <span onClick={()=>setTab("log")} style={{fontSize:11,color:C.muted,fontWeight:600,cursor:"pointer",textDecoration:"underline"}}>Lihat semua</span>
                </div>
                {allLog.length===0 && <div style={{padding:"10px 18px 18px",fontSize:12,color:C.muted}}>Belum ada transaksi. Catat pengeluaran pertama via Quick Input di tab Budget.</div>}
                {allLog.slice(0,5).map((l,i)=>(
                  <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 18px",borderTop:`1px solid ${C.light}`,gap:8}}>
                    <span style={{fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",flex:1}}>{l.type==="transfer"?`${l.from}->${l.to}`:(l.desc||"(tanpa nama)")}</span>
                    <span style={{fontSize:12,fontWeight:700,color:l.type==="in"?C.green:l.type==="out"?C.danger:C.text,flexShrink:0}}>{l.type==="in"?"+":l.type==="out"?"-":""}{fmt(num(l.amount))}</span>
                  </div>
                ))}
                <div style={{height:6}}/>
              </Card>
              {/* Import / Export / Reset */}
              <div style={{marginTop:14,display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                <button onClick={exportData}
                  style={{background:"none",border:`1px solid ${C.border}`,borderRadius:4,padding:"9px 0",fontSize:11,color:C.muted,cursor:"pointer",fontFamily:"inherit",fontWeight:600,letterSpacing:0.2}}>
                  Export Data
                </button>
                <label style={{background:"none",border:`1px solid ${C.border}`,borderRadius:4,padding:"9px 0",fontSize:11,color:C.muted,cursor:"pointer",fontFamily:"inherit",fontWeight:600,letterSpacing:0.2,textAlign:"center",display:"block"}}>
                  Import Data
                  <input type="file" accept=".json" onChange={importData} style={{display:"none"}}/>
                </label>
              </div>
              <button onClick={()=>setConfirmReset(true)}
                style={{width:"100%",marginTop:8,background:"none",border:`1px solid ${C.border}`,borderRadius:4,padding:"9px 0",fontSize:11,color:C.muted,cursor:"pointer",fontFamily:"inherit",fontWeight:600,letterSpacing:0.2}}>
                Reset Data
              </button>
            </div>
          </div>
        )}

        {/* -- BUDGET -- */}
        {tab==="budget" && (<>

          {/* Main Income */}
          <Card style={{marginBottom:12}}>
            <div style={{padding:"14px 18px"}}>
              <div style={{display:"flex",justifyContent:"flex-end",marginBottom:8}}>
                <button onClick={()=>setBudget(b=>({...b,income:{...b.income,mainLocked:!b.income.mainLocked}}))}
                  style={{background:"none",border:"none",cursor:"pointer",fontSize:16,padding:"0 2px"}}>
                  {budget.income.mainLocked?"🔒":"🔓"}
                </button>
              </div>
              {budget.income.mainLocked ? (
                <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"repeat(2,1fr)",gap:10}}>
                  <div>
                    <Lbl>Pemasukan Utama</Lbl>
                    <div style={{background:"transparent",border:"none",borderBottom:`1px solid ${C.light}`,padding:"6px 2px",fontSize:13,color:C.muted,textAlign:"right"}}>
                      {budget.income.main||"Rp0"}
                    </div>
                  </div>
                  <div>
                    <Lbl>Sisa Periode Lalu</Lbl>
                    <div style={{background:"transparent",border:"none",borderBottom:`1px solid ${C.light}`,padding:"6px 2px",fontSize:13,color:C.muted,textAlign:"right"}}>
                      {budget.income.lastPeriod||"Rp0"}
                    </div>
                  </div>
                </div>
              ) : (
                <>
                  <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"repeat(2,1fr)",gap:10,marginBottom:12}}>
                    <div><Lbl>Pemasukan Utama</Lbl><RpInput value={budget.income.main} onChange={setMainIncome}/></div>
                    <div><Lbl>Sisa Periode Lalu</Lbl><RpInput value={budget.income.lastPeriod} onChange={v=>setBudget(b=>({...b,income:{...b.income,lastPeriod:v}}))}/></div>
                  </div>
                  <div style={{fontSize:11,color:C.muted}}>Masuk otomatis ke <strong>{mainWallet?.name||"Saldo Utama"}</strong>.</div>
                </>
              )}
            </div>
          </Card>

          {/* Kategori */}
          {kats.map(kat => {
            const totB=kat.rows.reduce((s,r)=>(!r.isPiutangRow&&!r.isPiutang)?s+num(r.budget):s,0);
            const totA=kat.rows.reduce((s,r)=>(!r.isPiutangRow&&!r.isPiutang)?s+num(r.actual):s,0);
            const over=totA>totB; const pct=totB>0?(totA/totB)*100:0;
            const notesBtnStyle = r => {
              const has=r.notes?.length>0;
              const allPaid=has&&r.notes.every(n=>n.paid);
              const color = has ? (allPaid ? "#22C55E" : "#EF4444") : "#D1D5DB";
              const glow = has ? (allPaid
                ? "0 0 6px 2px rgba(34,197,94,0.5)"
                : "0 0 6px 2px rgba(239,68,68,0.5)")
                : "none";
              return {
                background:"none", border:"none", padding:"4px 6px",
                cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center",
                lineHeight:1,
                // dot styles passed separately
                _color: color, _glow: glow,
              };
            };
            const NotesDot = ({r}) => {
              const has=r.notes?.length>0;
              const allPaid=has&&r.notes.every(n=>n.paid);
              const color = has ? (allPaid ? "#22C55E" : "#EF4444") : "#D1D5DB";
              const glow = has ? (allPaid
                ? "0 0 5px 2px rgba(34,197,94,0.55)"
                : "0 0 5px 2px rgba(239,68,68,0.55)")
                : "none";
              return (
                <span style={{
                  display:"inline-block",
                  width:15, height:15, borderRadius:"50%",
                  background:color,
                  boxShadow:glow,
                  flexShrink:0,
                }}/>
              );
            };
            return (
              <Card key={kat.id} style={{marginBottom:10}}>
                <div style={{padding:"12px 18px 0"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8,gap:8}}>
                    <div style={{flex:1,minWidth:0}}>
                      {kat.hint && <div style={{fontSize:9,color:C.faint,fontWeight:700,letterSpacing:0.5,marginBottom:2}}>{kat.hint.toUpperCase()}</div>}
                      <TxtInput value={kat.nama} onChange={v=>updKatNama(kat.id,v)} placeholder="..."
                        style={{fontWeight:700,fontSize:12,border:"none",background:"transparent",padding:"2px 0"}}/>
                      <div style={{fontSize:11,color:over?C.danger:C.muted}}>{fmt(totA)} / {fmt(totB)}{totB>0&&totalInBudget>0?<span style={{color:"#C4C4C4",marginLeft:4}}>. {Math.round(totB/totalInBudget*100)}% effective</span>:null}</div>
                    </div>
                    <Tag color={over?C.danger:C.muted}>{pct.toFixed(0)}%</Tag>
                    <button onClick={()=>delKat(kat.id)} style={{background:"none",border:"none",color:C.faint,cursor:"pointer",fontSize:13,flexShrink:0,padding:0,lineHeight:1,alignSelf:"flex-start",marginTop:0}}>x</button>
                  </div>
                  <div style={{background:C.soft,borderRadius:6,height:4,overflow:"hidden",marginTop:8,marginBottom:over?6:10}}>
                    <div style={{width:`${Math.min(pct,100)}%`,background:over?C.danger:C.red,height:"100%",borderRadius:6}}/>
                  </div>
                  {over && (
                    <div style={{fontSize:10,color:C.danger,background:"#FEF2F2",border:"1px solid #FECACA",borderRadius:4,padding:"4px 10px",marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                      <span>{totA>totB?`Aktual melebihi budget ${fmt(totA-totB)}`:`Rencana melebihi budget ${fmt(totB>0?pct/100*totB-totB:0)}`}</span>
                    </div>
                  )}
                </div>

                {/* Desktop table */}
                {!isMobile && (
                  <div style={{overflowX:"auto"}}>
                    <div style={{display:"grid",gridTemplateColumns:"1.6fr 110px 110px 48px",gap:8,background:"#fff",padding:"5px 18px",fontSize:9,color:C.muted,fontWeight:700,minWidth:400,borderBottom:"1px solid #F0F0F0"}}>
                      <span>DESKRIPSI</span><span style={{textAlign:"center"}}>BUDGET</span><span style={{textAlign:"center"}}>AKTUAL</span><span style={{textAlign:"center"}}>NOTES</span>
                    </div>
                    {kat.rows.map((r,i)=>(
                      <div key={i} style={{display:"grid",gridTemplateColumns:"1.6fr 110px 110px 48px",gap:8,padding:"5px 18px",borderTop:`1px solid ${C.light}`,alignItems:"center",background:"#fff",minWidth:400}} className="m2os-note-row">
                        <div style={{display:"flex",alignItems:"center",gap:6,minWidth:0}}>
                          <TxtInput value={r.desc} onChange={v=>updRow(kat.id,i,"desc",v)} placeholder="..."
                            style={{border:"none",background:"transparent",padding:"4px 2px",fontSize:12}}/>
                          
                        </div>
                        {(r.isPiutangRow||r.isPiutang)
                          ? <div style={{fontSize:11,color:C.faint,textAlign:"center"}}>-</div>
                          : <RpInput value={r.budget} onChange={v=>updRow(kat.id,i,"budget",v)} style={{background:"transparent",border:"none",fontSize:12,textAlign:"center",padding:"4px 2px",color:(!r.budget||r.budget==="Rp0"||r.budget==="")?"#BFBFBF":"#1A1A1A"}}/>
                        }
                        <RpInput value={r.actual} onChange={v=>updRow(kat.id,i,"actual",v)}
                          style={{background:"transparent",border:"none",fontSize:12,textAlign:"center",padding:"4px 2px",color:num(r.actual)>num(r.budget)?C.danger:(!r.actual||r.actual==="Rp0"?"#BFBFBF":"#1A1A1A"),fontWeight:num(r.actual)>num(r.budget)?700:400}}/>
                        <div style={{display:"flex",justifyContent:"flex-end",alignItems:"center",gap:20,paddingLeft:4}}>
                          <button onClick={()=>setNotesOpen({katId:kat.id,idx:i})} style={{background:"none",border:"none",padding:0,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",marginLeft:12}}><NotesDot r={r}/></button>
                          <button onClick={()=>delRow(kat.id,i)} className="m2os-note-del" style={{background:"none",border:"none",cursor:"pointer",opacity:0,transition:"opacity 0.15s",padding:0,display:"flex",alignItems:"center"}}><svg width="10" height="13" viewBox="0 0 10 13" fill="none" xmlns="http://www.w3.org/2000/svg">
  <polygon points="1,2 9,2 8,12 2,12" stroke="#E5E7EB" strokeWidth="1.2" fill="none" strokeLinejoin="round"/>
  <line x1="0" y1="2" x2="10" y2="2" stroke="#E5E7EB" strokeWidth="1.2" strokeLinecap="round"/>
  <line x1="3" y1="4.5" x2="3" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
  <line x1="5" y1="4.5" x2="5" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
  <line x1="7" y1="4.5" x2="7" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
</svg></button>
                          
                        </div>
                      </div>
                    ))}
                  </div>
                )}

                {/* Mobile rows */}
                {isMobile && (<>
                  <div style={{display:"flex",alignItems:"center",gap:8,padding:"4px 16px"}}>
                    <div style={{flex:1.35,fontSize:9,color:"#999",fontWeight:700}}>DESKRIPSI</div>
                    <div style={{flex:1.325,fontSize:9,color:"#999",fontWeight:700,textAlign:"center"}}>BUDGET</div>
                    <div style={{flex:1.325,fontSize:9,color:"#999",fontWeight:700,textAlign:"center"}}>AKTUAL</div>
                    <div style={{width:40}}/>
                  </div>
                  {kat.rows.map((r,i)=>(
                    <div key={i} style={{padding:"8px 16px",borderTop:`1px solid ${C.light}`}} className="m2os-note-row">
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <div style={{display:"flex",alignItems:"center",gap:4,flex:1.35,minWidth:0}}>
                          <TxtInput value={r.desc} onChange={v=>updRow(kat.id,i,"desc",v)} placeholder="..." style={{fontSize:12,flex:1,border:"none",background:"transparent",padding:"0"}}/>

                        </div>
                        {(r.isPiutangRow||r.isPiutang)
                          ? <div style={{flex:1.325,fontSize:11,color:C.faint,textAlign:"center"}}>-</div>
                          : <RpInput value={r.budget} onChange={v=>updRow(kat.id,i,"budget",v)} style={{flex:1.325,fontSize:12,border:"none",background:"transparent",textAlign:"center",padding:"6px 2px",color:(!r.budget||r.budget==="Rp0"||r.budget==="")?"#BFBFBF":"#1A1A1A"}}/>
                        }
                        <RpInput value={r.actual} onChange={v=>updRow(kat.id,i,"actual",v)} style={{flex:1.325,fontSize:12,border:"none",background:"transparent",textAlign:"center",padding:"6px 2px",color:num(r.actual)>num(r.budget)?C.danger:(!r.actual||r.actual==="Rp0"?"#BFBFBF":"#1A1A1A"),fontWeight:num(r.actual)>num(r.budget)?700:400}}/>
                        <div style={{display:"flex",gap:16,flexShrink:0,alignItems:"center"}}>
                          <button onClick={()=>setNotesOpen({katId:kat.id,idx:i})} style={{background:"none",border:"none",padding:0,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",marginLeft:8}}><NotesDot r={r}/></button>
                          <button onClick={()=>delRow(kat.id,i)} className="m2os-note-del" style={{background:"none",border:"none",cursor:"pointer",opacity:0,transition:"opacity 0.15s",padding:0,display:"flex",alignItems:"center"}}><svg width="10" height="13" viewBox="0 0 10 13" fill="none" xmlns="http://www.w3.org/2000/svg">
  <polygon points="1,2 9,2 8,12 2,12" stroke="#E5E7EB" strokeWidth="1.2" fill="none" strokeLinejoin="round"/>
  <line x1="0" y1="2" x2="10" y2="2" stroke="#E5E7EB" strokeWidth="1.2" strokeLinecap="round"/>
  <line x1="3" y1="4.5" x2="3" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
  <line x1="5" y1="4.5" x2="5" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
  <line x1="7" y1="4.5" x2="7" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
</svg></button>

                        </div>
                      </div>
                    </div>
                  ))}
                </>)}

                <div style={{padding:"8px 18px",borderTop:`1px solid ${C.light}`,display:"flex",alignItems:"center"}}>
                  <button onClick={()=>{const isD=kat.hint==="debt"||kat.nama.toLowerCase().includes("debt")||kat.nama.toLowerCase().includes("hutang");isD?setAddDebtModal(kat.id):addRow(kat.id);}} style={{background:"none",border:"none",padding:"4px 0",fontSize:11,fontWeight:400,fontFamily:"inherit",color:C.faint,cursor:"pointer"}}>+ tambah baris</button>
                </div>
              </Card>
            );
          })}

          <div style={{marginBottom:14}}>
            <button onClick={addKat} style={{width:"100%",background:"none",border:`1px solid ${C.border}`,borderRadius:4,padding:"9px 0",fontSize:11,color:C.muted,cursor:"pointer",fontFamily:"inherit",fontWeight:600,letterSpacing:0.2}}>+ Tambah Kategori</button>
          </div>

          {/* ETC - Side Income */}
          <Card style={{marginBottom:14}}>
            <div style={{padding:"12px 18px 0"}}>
              <div style={{fontSize:9,color:C.faint,fontWeight:700,letterSpacing:0.5,marginBottom:4}}>INCOME ADD-ONS</div>
              <TxtInput value={budget.sideIncome.nama} onChange={updSideNama} placeholder="..."
                style={{fontWeight:700,fontSize:12,border:"none",background:"transparent",padding:"2px 0 8px 0"}}/>
              <div style={{fontSize:10,color:C.green,marginBottom:10}}>{fmt(totalSide)} masuk</div>
            </div>
            {!isMobile && (
              <div style={{display:"grid",gridTemplateColumns:"1.4fr 110px 40px 120px 24px",gap:8,background:"#fff",padding:"5px 18px",fontSize:9,color:C.muted,fontWeight:700,minWidth:500,borderBottom:"1px solid #F0F0F0"}}>
                <span>DESKRIPSI</span><span style={{textAlign:"right"}}>JUMLAH</span><span></span><span>WALLET</span><span/>
              </div>
            )}
            {(budget.sideIncome.rows||[]).map((r,i) => isMobile ? (
              <div key={i} style={{padding:"8px 10px",borderTop:`1px solid ${C.border}`,background:C.card}} className="m2os-note-row">
                <div style={{display:"grid",gridTemplateColumns:"1fr 110px 18px",gap:8,alignItems:"center",marginBottom:6}}>
                  <TxtInput value={r.desc||""} onChange={v=>updSide(i,"desc",v)} placeholder="..." style={{fontSize:12}}/>
                  <RpInput value={r.amount} onChange={v=>updSide(i,"amount",v)} style={{color:C.green,fontWeight:400}}/>
                  <button onClick={()=>delSide(i)} className="m2os-note-del" style={{background:"none",border:"none",cursor:"pointer",padding:0,opacity:0.25,display:"flex",alignItems:"center",justifySelf:"center"}}><svg width="10" height="13" viewBox="0 0 10 13" fill="none" xmlns="http://www.w3.org/2000/svg"><polygon points="1,2 9,2 8,12 2,12" stroke="#E5E7EB" strokeWidth="1.2" fill="none" strokeLinejoin="round"/><line x1="0" y1="2" x2="10" y2="2" stroke="#E5E7EB" strokeWidth="1.2" strokeLinecap="round"/><line x1="3" y1="4.5" x2="3" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/><line x1="5" y1="4.5" x2="5" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/><line x1="7" y1="4.5" x2="7" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/></svg></button>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 110px 18px",gap:8,alignItems:"center"}}>
                  <Sel value={r.wallet||""} onChange={v=>updSide(i,"wallet",v)} options={[[""," Pilih Wallet"],...wallets.map(w=>w.name)]}/>
                  <DateBtn value={r.date} onChange={v=>updSide(i,"date",v)}/>
                  <span/>
                </div>
              </div>
            ) : (
              <div key={i} style={{display:"grid",gridTemplateColumns:"1.4fr 110px 40px 120px 24px",gap:8,padding:"5px 18px",borderTop:`1px solid ${C.light}`,alignItems:"center",background:"#fff",minWidth:500}} className="m2os-note-row">
                <TxtInput value={r.desc} onChange={v=>updSide(i,"desc",v)} placeholder="..." style={{border:"none",background:"transparent",padding:"4px 2px",fontSize:12}}/>
                <RpInput value={r.amount} onChange={v=>updSide(i,"amount",v)} style={{background:"transparent",border:"none",fontSize:12,color:C.green,fontWeight:700}}/>
                <DateBtn value={r.date} onChange={v=>updSide(i,"date",v)}/>
                <Sel value={r.wallet||""} onChange={v=>updSide(i,"wallet",v)} options={[[""," - "],...wallets.map(w=>w.name)]} style={{fontSize:11,padding:"4px 6px",border:"none",background:"transparent"}}/>
                <button onClick={()=>delSide(i)} className="m2os-note-del" style={{background:"none",border:"none",cursor:"pointer",padding:0,opacity:0,transition:"opacity 0.15s",display:"flex",alignItems:"center"}}><svg width="10" height="13" viewBox="0 0 10 13" fill="none" xmlns="http://www.w3.org/2000/svg">
  <polygon points="1,2 9,2 8,12 2,12" stroke="#E5E7EB" strokeWidth="1.2" fill="none" strokeLinejoin="round"/>
  <line x1="0" y1="2" x2="10" y2="2" stroke="#E5E7EB" strokeWidth="1.2" strokeLinecap="round"/>
  <line x1="3" y1="4.5" x2="3" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
  <line x1="5" y1="4.5" x2="5" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
  <line x1="7" y1="4.5" x2="7" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/>
</svg></button>
              </div>
            ))}
            <div style={{padding:"8px 18px",borderTop:`1px solid ${C.light}`}}>
              <button onClick={addSide} style={{background:"none",border:"none",padding:"4px 0",fontSize:11,fontWeight:400,fontFamily:"inherit",color:C.faint,cursor:"pointer"}}>+ tambah baris</button>
            </div>
          </Card>
        </>)}

        {/* -- WALLET -- */}
        {tab==="wallet" && (
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:14}}>
            <div>
              {/* -- WALLET SALDO (main) -- */}
              <Card style={{marginBottom:14}}>
                <div style={{padding:"12px 16px 4px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <Lbl>Wallet Saldo</Lbl>
                  <span style={{fontSize:9,color:C.faint,fontWeight:600}}>Linked ke Pemasukan Utama</span>
                </div>
                {wallets.filter(w=>w.isMain).map((w,i)=>(
                  <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 16px 12px"}}>
                    <TxtInput value={w.name} onChange={v=>setWallets(ws=>ws.map(x=>x.isMain?{...x,name:v}:x))} style={{fontWeight:700,fontSize:13,flex:1}}/>
                    <div style={{textAlign:"right",flexShrink:0}}>
                      <div style={{fontSize:16,fontWeight:700,color:C.text}}>{fmt(num(w.amount))}</div>
                      <div style={{fontSize:9,color:C.muted}}>saldo otomatis</div>
                    </div>
                  </div>
                ))}
              </Card>

              {/* -- WALLET BIASA -- */}
              <Card>
                <div style={{padding:"12px 16px 4px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <Lbl>Wallet</Lbl>
                </div>
                {!isMobile && regularWallets.length>0 && (
                  <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 32px",gap:8,background:C.soft,padding:"6px 8px 6px 16px",fontSize:9,color:C.muted,fontWeight:700,borderBottom:`1px solid ${C.border}`}}>
                    <span>NAMA</span><span style={{textAlign:"right"}}>SALDO</span><span/>
                  </div>
                )}
                {regularWallets.length===0 && (
                  <div style={{padding:"8px 16px 12px",fontSize:12,color:C.muted}}>Belum ada wallet operasional. Tambah wallet harian seperti BCA, GoPay, atau Cash.</div>
                )}
                {wallets.map((w,i)=>w.isMain?null: isMobile ? (
                  <div key={i} style={{display:"flex",alignItems:"center",gap:11,padding:"10px 11px 10px 14px",borderBottom:`1px solid ${C.light}`,background:"#fff"}} className="m2os-note-row">
                    <TxtInput value={w.name} onChange={v=>updW(i,"name",v)} style={{fontWeight:600,fontSize:13,flex:1}}/>
                    <RpInput value={w.amount} onChange={v=>updW(i,"amount",v)} style={{fontWeight:700,fontSize:13,flex:1}}/>
                    <button onClick={()=>setWallets(ws=>ws.filter((_,j)=>j!==i))} className="m2os-note-del" style={{background:"none",border:"none",cursor:"pointer",padding:0,opacity:0.25,display:"flex",alignItems:"center",justifySelf:"center"}}><svg width="10" height="13" viewBox="0 0 10 13" fill="none" xmlns="http://www.w3.org/2000/svg"><polygon points="1,2 9,2 8,12 2,12" stroke="#E5E7EB" strokeWidth="1.2" fill="none" strokeLinejoin="round"/><line x1="0" y1="2" x2="10" y2="2" stroke="#E5E7EB" strokeWidth="1.2" strokeLinecap="round"/><line x1="3" y1="4.5" x2="3" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/><line x1="5" y1="4.5" x2="5" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/><line x1="7" y1="4.5" x2="7" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/></svg></button>
                  </div>
                ) : (
                  <div key={i} style={{display:"grid",gridTemplateColumns:"2fr 1fr 32px",gap:8,padding:"8px 8px 8px 16px",borderBottom:`1px solid ${C.light}`,alignItems:"center",background:"#fff"}} className="m2os-note-row">
                    <TxtInput value={w.name} onChange={v=>updW(i,"name",v)} style={{border:"none",background:"transparent",fontWeight:600,fontSize:13,padding:"2px"}}/>
                    <RpInput value={w.amount} onChange={v=>updW(i,"amount",v)} style={{fontWeight:700,background:"transparent",border:"none",fontSize:13}}/>
                    <button onClick={()=>setWallets(ws=>ws.filter((_,j)=>j!==i))} className="m2os-note-del" style={{background:"none",border:"none",cursor:"pointer",padding:0,opacity:0.25,display:"flex",alignItems:"center",justifySelf:"center"}}><svg width="10" height="13" viewBox="0 0 10 13" fill="none" xmlns="http://www.w3.org/2000/svg"><polygon points="1,2 9,2 8,12 2,12" stroke="#E5E7EB" strokeWidth="1.2" fill="none" strokeLinejoin="round"/><line x1="0" y1="2" x2="10" y2="2" stroke="#E5E7EB" strokeWidth="1.2" strokeLinecap="round"/><line x1="3" y1="4.5" x2="3" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/><line x1="5" y1="4.5" x2="5" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/><line x1="7" y1="4.5" x2="7" y2="10" stroke="#E5E7EB" strokeWidth="1.0" strokeLinecap="round"/></svg></button>
                  </div>
                ))}
                <div style={{padding:"8px 16px"}}>
                  <button onClick={()=>setWallets(w=>[...w,{name:"Wallet Baru",amount:"",isMain:false}])} style={{background:"none",border:"none",padding:"4px 0",fontSize:11,fontWeight:400,fontFamily:"inherit",color:C.faint,cursor:"pointer"}}>+ tambah wallet</button>
                </div>
              </Card>
            </div>

            <Card>
              <div style={{padding:16}}>
                <Lbl>Distribusi Wallet</Lbl>
                <div style={{fontSize:11,fontWeight:700,marginBottom:10,color:C.text}}>Total: {fmt(walletTotal)}</div>
                {walletTotal>0 && (
                  <div style={{display:"flex",height:8,borderRadius:4,overflow:"hidden",gap:1,marginBottom:14}}>
                    {wallets.filter(w=>num(w.amount)>0).map((w,i)=><div key={i} style={{flex:num(w.amount),background:w.isMain?C.red:GRAYS[(i)%GRAYS.length]}}/>)}
                  </div>
                )}
                {wallets.map((w,i)=>num(w.amount)>0&&(
                  <div key={i} style={{display:"flex",alignItems:"center",gap:6,marginBottom:6,fontSize:12}}>
                    <div style={{width:8,height:8,borderRadius:2,background:w.isMain?C.red:GRAYS[i%GRAYS.length],flexShrink:0}}/>
                    <span style={{flex:1,color:C.muted}}>{w.name}{w.isMain?" *":""}</span>
                    <span style={{fontWeight:700}}>{walletTotal>0?((num(w.amount)/walletTotal)*100).toFixed(1):0}%</span>
                  </div>
                ))}
              </div>
            </Card>
          </div>
        )}

        {/* -- SAVINGS -- */}
        {tab==="savings" && (
          <SavingsTab goals={goals} wallets={wallets} isMobile={isMobile}
            onAdd={addGoal} onDel={delGoal} onUpd={updGoal}
            onDeposit={goalDeposit} onWithdrawW={goalWithdrawW} onWithdrawU={goalWithdrawU}/>
        )}

        {/* -- LOG -- */}
        {tab==="log" && (() => {
          const totalOut = filteredLog.filter(l=>l.type==="out").reduce((s,l)=>s+num(l.amount),0);
          const totalIn = filteredLog.filter(l=>l.type==="in"||l.type==="side_income").reduce((s,l)=>s+num(l.amount),0);
          const byKat = filteredLog.filter(l=>l.type==="out"&&l.katNama).reduce((acc,l)=>{acc[l.katNama]=(acc[l.katNama]||0)+num(l.amount);return acc;},{});
          return (
          <div>
            {/* Search */}
            <div style={{marginBottom:8,position:"relative"}}>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Cari transaksi, kategori..."
                style={{width:"100%",border:`1px solid ${C.border}`,borderRadius:10,padding:"9px 14px",fontSize:13,fontFamily:"inherit",background:C.card,boxSizing:"border-box",outline:"none"}}/>
              {search&&<button onClick={()=>setSearch("")} style={{position:"absolute",right:12,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:C.muted,cursor:"pointer",fontSize:15}}>x</button>}
            </div>

            {/* Filter tipe */}
            <div style={{display:"flex",gap:6,marginBottom:8,flexWrap:"wrap"}}>
              {[["all","Semua"],["out","Pengeluaran"],["in","Pemasukan"],["transfer","Transfer"],["void","Dibatalkan"],["savings_deposit","Savings"]].map(([v,l])=>(
                <button key={v} onClick={()=>setLogFilter(v)}
                  style={{background:"none",border:`1px solid ${logFilter===v?C.red:C.border}`,
                  borderRadius:20,padding:"4px 12px",fontSize:11,fontWeight:logFilter===v?600:400,
                  color:logFilter===v?C.red:C.muted,cursor:"pointer",fontFamily:"inherit"}}>
                  {l}
                </button>
              ))}
            </div>

            {/* Date range */}
            <div style={{display:"flex",gap:8,marginBottom:12,alignItems:"center"}}>
              <DateBtn value={dateFrom} onChange={setDateFrom} style={{flex:1}}/>
              <span style={{fontSize:12,color:C.muted,flexShrink:0}}>s/d</span>
              <DateBtn value={dateTo} onChange={setDateTo} style={{flex:1}}/>
              {(dateFrom||dateTo)&&<button onClick={()=>{setDateFrom("");setDateTo("");}} style={{background:"none",border:"none",color:C.muted,cursor:"pointer",fontSize:11,flexShrink:0}}>Reset</button>}
            </div>

            {/* Summary */}
            {(dateFrom||dateTo||search) && filteredLog.length>0 && (
              <div style={{display:"flex",gap:8,marginBottom:12}}>
                <Card style={{flex:1,padding:"10px 14px"}}>
                  <Lbl>Pengeluaran</Lbl>
                  <div style={{fontSize:15,fontWeight:700,color:C.danger}}>{fmt(totalOut)}</div>
                </Card>
                <Card style={{flex:1,padding:"10px 14px"}}>
                  <Lbl>Pemasukan</Lbl>
                  <div style={{fontSize:15,fontWeight:700,color:C.green}}>{fmt(totalIn)}</div>
                </Card>
              </div>
            )}

            {/* Breakdown per kategori */}
            {(dateFrom||dateTo) && Object.keys(byKat).length>0 && (
              <Card style={{marginBottom:12,padding:"12px 16px"}}>
                <Lbl style={{marginBottom:8}}>Breakdown Pengeluaran</Lbl>
                {Object.entries(byKat).sort((a,b)=>b[1]-a[1]).map(([k,v])=>(
                  <div key={k} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
                    <span style={{fontSize:12,color:C.text}}>{k}</span>
                    <span style={{fontSize:12,fontWeight:700,color:C.danger}}>{fmt(v)}</span>
                  </div>
                ))}
              </Card>
            )}
            <Card>
              {filteredLog.length===0&&<div style={{padding:28,textAlign:"center",color:C.muted,fontSize:13}}>{search?"Tidak ada transaksi yang cocok dengan pencarian.":"Belum ada transaksi bulan ini. Mulai catat via Quick Input di tab Budget."}</div>}
              {filteredLog.map((l,i)=>(
                <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 18px",borderBottom:`1px solid ${C.light}`,gap:8}}>
                  <div style={{minWidth:0}}>
                    <div style={{fontWeight:600,fontSize:13}}>
                      {l.type==="transfer"?`${l.from} -> ${l.to}`:
                       l.type==="savings_deposit"?`Setor    .    ${l.goalNama||"Savings"}`:
                       l.type==="savings_withdraw"?`Tarik    .    ${l.goalNama||"Savings"}`:l.type==="savings_withdraw_urgent"?`Ambil Urgent    .    ${l.goalNama||"Savings"}`:
                       (l.rowDesc||l.desc||"(tanpa nama)")}
                    </div>
                    <div style={{fontSize:10,color:C.muted}}>
                      {(l.type==="out"||l.type==="void") && [l.katNama, l.desc&&l.desc!==l.rowDesc?l.desc:null, l.wallet].filter(Boolean).join("    .    ")}
                      {(l.type==="in"||l.type==="side_income") && `+ ${l.wallet||""}`}
                      {(l.type==="in"||l.type==="side_income") && l.note && ` · ${l.note}`}
                      {l.type==="transfer" && [l.note, l.date?isoToDisplay(l.date):""].filter(Boolean).join("    .    ")}
                      {l.type==="savings_deposit" && [l.from+" -> "+(l.to||"savings"), l.date?isoToDisplay(l.date):""].filter(Boolean).join("    .    ")}
                      {(l.type==="savings_withdraw"||l.type==="savings_withdraw_urgent") && [(l.from||"savings")+(l.type==="savings_withdraw_urgent"?" (urgent)":l.to?" -> "+l.to:""), l.date?isoToDisplay(l.date):""].filter(Boolean).join("    .    ")}
                      {(l.type==="out"||l.type==="void"||l.type==="in"||l.type==="side_income") && ("  "+(l.date?isoToDisplay(l.date):""))}
                    </div>
                  </div>
                  <div style={{textAlign:"right",flexShrink:0}}>
                    <div style={{fontWeight:700,fontSize:13,color:
                      l.type==="in"||l.type==="side_income"||l.type==="savings_withdraw"?C.green:l.type==="savings_withdraw_urgent"?C.danger:
                      l.type==="void"?(l.isPiutang?C.danger:C.green):
                      l.type==="out"||l.type==="savings_deposit"?C.danger:C.text}}>
                      {l.type==="in"||l.type==="side_income"||l.type==="savings_withdraw"?"+":
                       l.type==="void"?(l.isPiutang?"-":"+"):
                       l.type==="savings_withdraw_urgent"?"-":
                       l.type==="out"||l.type==="savings_deposit"?"-":""}{fmt(num(l.amount))}
                    </div>
                    {l.type==="void" && <div style={{fontSize:9,fontWeight:600,color:C.muted,marginTop:2}}>dibatalkan</div>}
                  </div>
                </div>
              ))}
            </Card>
          </div>
          );
        })()}
      </div>

      {/* -- SCROLL TO TOP -- */}
      {scrolled && (
        <button onClick={()=>{ scrollRef.current?.scrollTo({top:0,behavior:'smooth'}); window.scrollTo({top:0,behavior:'smooth'}); }}
          style={{position:isMobile?"fixed":"absolute",
          top:isMobile?(scrolled?"calc(env(safe-area-inset-top) + 230px)":"calc(env(safe-area-inset-top) + 280px)"):"240px",
          left:"50%",
          transform:"translateX(-50%)",zIndex:50,transition:"top 0.25s ease",
          width:36,height:36,borderRadius:18,
          background:"rgba(255,255,255,0.9)",backdropFilter:"blur(8px)",
          border:`1px solid ${C.border}`,color:C.red,fontSize:16,cursor:"pointer",
          display:"flex",alignItems:"center",justifyContent:"center",
          boxShadow:"0 2px 12px rgba(0,0,0,0.15)",lineHeight:1}}>
          ↑
        </button>
      )}

      {/* -- BOTTOM NAV (mobile) -- */}
      {isMobile && (
        <div style={{position:"fixed",bottom:"max(24px, env(safe-area-inset-bottom))",left:"50%",transform:"translateX(-50%)",zIndex:50,background:theme==='white'?'rgba(255,255,255,0.95)':`linear-gradient(135deg, ${T.dark} 0%, ${T.red} 100%)`,borderRadius:40,padding:"10px 16px",display:"flex",justifyContent:"space-around",alignItems:"center",gap:8,boxShadow:theme==='white'?'0 4px 20px rgba(0,0,0,0.12), 0 0 0 1px rgba(0,0,0,0.08)':`0 8px 32px ${T.dark}55`,minWidth:"72vw"}}>
          {[
            ["overview", <svg width="20" height="20" viewBox="0 0 20 20" fill="none"><rect x="2" y="2" width="7" height="7" rx="1.5" fill="currentColor"/><rect x="11" y="2" width="7" height="7" rx="1.5" fill="currentColor"/><rect x="2" y="11" width="7" height="7" rx="1.5" fill="currentColor"/><rect x="11" y="11" width="7" height="7" rx="1.5" fill="currentColor"/></svg>],
            ["budget", <svg width="20" height="20" viewBox="0 0 20 20" fill="none"><rect x="4" y="2" width="12" height="16" rx="2" stroke="currentColor" strokeWidth="2"/><path d="M7 8h6M7 12h6" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/></svg>],
            ["wallet", <svg width="20" height="20" viewBox="0 0 20 20" fill="none"><rect x="2" y="4" width="16" height="12" rx="2" stroke="currentColor" strokeWidth="2"/><line x1="3.5" y1="9" x2="16.5" y2="9" stroke="currentColor" strokeWidth="1.5"/></svg>],
            ["savings", <svg width="20" height="20" viewBox="0 0 20 20" fill="none"><circle cx="10" cy="10" r="7" stroke="currentColor" strokeWidth="2"/><circle cx="10" cy="10" r="3" fill="currentColor"/></svg>],
            ["log", <svg width="20" height="20" viewBox="0 0 20 20" fill="none"><path d="M4 5h12M4 10h8M4 15h10" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/></svg>],
          ].map(([k,icon])=>(
            <button key={k} onClick={()=>setTabSafe(k)} style={{background:"none",border:"none",cursor:"pointer",padding:"6px 10px",color:theme==='white'?(tab===k?"#1A1A1A":"rgba(0,0,0,0.25)"):(tab===k?"#fff":"rgba(255,255,255,0.35)"),display:"flex",alignItems:"center",justifyContent:"center",transition:"color 0.2s",outline:"none"}}>
              {icon}
            </button>
          ))}
        </div>
      )}
      {/* -- MODAL RESET DATA -- */}
      {confirmReset && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={()=>setConfirmReset(false)}>
          <div style={{background:C.card,borderRadius:4,padding:24,width:320,maxWidth:"92vw",boxShadow:"0 16px 40px rgba(0,0,0,0.2)"}} onClick={e=>e.stopPropagation()}>
            <div style={{fontWeight:700,fontSize:14,marginBottom:10,color:C.ink}}>Reset Data</div>
            <div style={{fontSize:12,color:C.muted,lineHeight:1.6,marginBottom:20}}>Reset data artinya reset semua data yang ada di periode ini tanpa menghilangkan dokumentasi di periode sebelumnya.</div>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>setConfirmReset(false)} style={{flex:1,background:C.soft,border:"none",borderRadius:8,padding:"8px 0",fontSize:12,fontWeight:600,color:C.muted,cursor:"pointer",fontFamily:"inherit"}}>Batal</button>
              <button onClick={execReset} style={{flex:1,background:C.danger,border:"none",borderRadius:8,padding:"8px 0",fontSize:12,fontWeight:600,color:"#fff",cursor:"pointer",fontFamily:"inherit"}}>Reset</button>
            </div>
          </div>
        </div>
      )}
      </div>
    </div>
    </div>
  );
}
