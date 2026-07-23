import { useState, useEffect, useReducer, useMemo, useRef } from "react";
import { BarChart, Bar, LineChart, Line, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from "recharts";
import * as XLSX from "xlsx";
import { Plus, Trash2, Edit2, ChevronDown, ChevronUp, Home, BarChart2, User, Brain, Sun, Moon, Download, FileText, AlertTriangle, DollarSign, TrendingUp, TrendingDown, Check, Building2, Target, Sparkles, Wallet, RefreshCw, Info, Calendar, Clock, HelpCircle, ChevronLeft, ChevronRight, CheckCircle2, Zap, Upload, RotateCcw, ArrowRight, List, BookOpen, Star } from "lucide-react";

/* ══ CONSTANTES ══════════════════════════════════════════ */
const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
const DIAS = ["Lun","Mar","Mié","Jue","Vie","Sáb","Dom"];
const PAY_METHODS = ["Transferencia","Efectivo","Tarjeta débito","Tarjeta crédito","Otro"];
const NOW = new Date(), CY = NOW.getFullYear(), CM = NOW.getMonth();
const YEARS = [CY-2, CY-1, CY, CY+1];
const COLORS = ["#7C3AED","#10B981","#F59E0B","#EF4444","#06B6D4","#F97316","#8B5CF6","#84CC16","#EC4899","#14B8A6","#0EA5E9","#A855F7"];
const CAT_PRESETS = [["Vivienda","🏠"],["Alimentación","🛒"],["Transporte","🚌"],["Salud","💊"],["Educación","📚"],["Entretenimiento","🎭"],["Ropa y Calzado","👔"],["Tecnología","💻"],["Comunicaciones","📱"],["Seguros","🛡️"],["Mascotas","🐾"],["Ahorro","🏦"],["Otros","📦"]];
const INC_TYPES = ["Sueldo","Beca","Aporte familiar","Freelance","Renta","Pensión","Bonificación","Otro"];
const FREQS = [{v:"mensual",l:"Mensual"},{v:"quincenal",l:"Quincenal"},{v:"semanal",l:"Semanal"},{v:"anual",l:"Anual"}];
const MF = {mensual:1,quincenal:2,semanal:4.33,anual:1/12};
const SK = "gestor_v3";
const AK_KEY = "gestor_v3_ak";

/* ══ UTILIDADES ══════════════════════════════════════════ */
let _c=0;
const uid=()=>`${Date.now().toString(36)}_${++_c}`;
const fCLP=n=>new Intl.NumberFormat("es-CL",{style:"currency",currency:"CLP",maximumFractionDigits:0}).format(Number(String(n||0).replace(/\D/g,""))||0);
const pCLP=s=>parseInt(String(s||"").replace(/\D/g,""))||0;
const toM=i=>Math.round((i.amount||0)*(MF[i.frequency]||1));
const sumInc=p=>(p?.incomes||[]).reduce((s,i)=>s+toM(i),0);
const todayStr=()=>new Date().toISOString().split("T")[0];
const fmtDate=d=>{if(!d)return"";try{return new Date(d+"T12:00").toLocaleDateString("es-CL",{day:"2-digit",month:"short"});}catch{return d;}};
const fmtGreeting=()=>{const h=NOW.getHours();return h<12?"Buenos días":h<19?"Buenas tardes":"Buenas noches";};

// Plantilla helpers — el corazón del nuevo modelo
const tmplTotal=addr=>(addr?.template?.categories||[]).reduce((s,c)=>(c.items||[]).reduce((ss,it)=>ss+Math.round((it.amount||0)*(MF[it.frequency||"mensual"]||1)),s),0);
const allTmplItems=addr=>(addr?.template?.categories||[]).flatMap(c=>(c.items||[]).map(it=>{const mAmt=Math.round((it.amount||0)*(MF[it.frequency||"mensual"]||1));return{...it,catName:c.name,catIcon:c.icon,catColor:c.color,catId:c.id,monthlyAmt:mAmt};}));
const itemAccumulated=(addr,y,m,itemId)=>mPays(addr,y,m).filter(p=>p.templateItemId===itemId).reduce((s,p)=>s+(p.amount||0),0);
const mPays=(addr,y,m)=>(addr?.months||[]).find(x=>x.year===y&&x.month===m)?.payments||[];
const totalPaid=(addr,y,m)=>mPays(addr,y,m).reduce((s,p)=>s+(p.amount||0),0);

// Ítems pendientes de la plantilla (no pagados este mes)
const pendingTmpl=(addr,y,m)=>allTmplItems(addr).map(it=>({...it,accumulated:itemAccumulated(addr,y,m,it.id)})).filter(it=>it.accumulated<it.monthlyAmt);
// Ítems de plantilla ya pagados este mes
const paidTmpl=(addr,y,m)=>allTmplItems(addr).map(it=>{const acc=itemAccumulated(addr,y,m,it.id);const pays=mPays(addr,y,m).filter(p=>p.templateItemId===it.id);return acc>0?{...it,accumulated:acc,pays,isPaid:acc>=it.monthlyAmt}:null;}).filter(Boolean);
// Pagos extra (no de plantilla)
const extraPays=(addr,y,m)=>mPays(addr,y,m).filter(p=>p.isExtra);

// Sugerencias: extras que se repiten 2+ veces y no están en la plantilla
const smartSuggestions=addr=>{
  const tmplNames=new Set(allTmplItems(addr).map(i=>i.name.trim().toLowerCase()));
  const counts={};
  (addr?.months||[]).flatMap(m=>(m.payments||[]).filter(p=>p.isExtra)).forEach(p=>{
    const k=p.name.trim().toLowerCase();
    if(tmplNames.has(k))return;
    counts[k]=counts[k]||{name:p.name,count:0,lastAmt:0};
    counts[k].count++;counts[k].lastAmt=p.amount;
  });
  return Object.values(counts).filter(x=>x.count>=2).slice(0,3);
};

// Gastos frecuentes conocidos — ordenados por popularidad de uso (más usado primero)
const knownExtrasSorted=addr=>[...(addr?.knownExtras||[])].sort((a,b)=>(b.useCount||0)-(a.useCount||0)||new Date(b.lastDate||0)-new Date(a.lastDate||0));
const matchKnownExtra=(addr,name)=>{
  const k=(name||"").trim().toLowerCase();
  if(!k)return null;
  return (addr?.knownExtras||[]).find(x=>x.name.trim().toLowerCase()===k)||null;
};
const searchKnownExtras=(addr,query)=>{
  const q=(query||"").trim().toLowerCase();
  const all=knownExtrasSorted(addr);
  if(!q)return all.slice(0,6);
  return all.filter(k=>k.name.toLowerCase().includes(q)).slice(0,6);
};
// Cuántas veces se ha usado realmente una categoría en pagos históricos (para ordenar el gestor por frecuencia real)
const catUsageCount=(addr,catId)=>(addr?.months||[]).reduce((s,m)=>s+(m.payments||[]).filter(p=>p.catId===catId).length,0);

/* ══ STORAGE (localStorage para web real) ════════════════ */
const ldata=async()=>{try{const r=localStorage.getItem(SK);return r?JSON.parse(r):null;}catch{return null;}};
const sdata=async d=>{try{localStorage.setItem(SK,JSON.stringify(d));}catch{}};
const getAK=()=>{try{return localStorage.getItem(AK_KEY)||"";}catch{return"";}};
const saveAK=k=>{try{localStorage.setItem(AK_KEY,k);}catch{}};

/* ══ ESTADO — modelo nuevo limpio ════════════════════════
   address.template.categories[].items[] = gastos fijos recurrentes
   address.months[].payments[]           = lo que realmente pagaste
   Sin "categorías en meses" — eso era el problema de la v2
══════════════════════════════════════════════════════════ */
const INIT={
  addresses:[],selAddr:null,selYear:CY,selMonth:CM,
  profile:{name:"",incomes:[],availability:{hoursPerWeek:0,canWork:true,studying:false,entrepreneur:false,notes:""},goals:[],paymentMethods:[]},
  settings:{dark:false,alertPct:25,linkProf:false,tutDone:false,iaGuideDismissed:false},
  aiHistory:[]
};
const upA=(addrs,id,fn)=>addrs.map(a=>a.id===id?fn(a):a);
const gOrM=(addr,y,m)=>{
  const ex=(addr.months||[]).find(x=>x.year===y&&x.month===m);
  if(ex)return{months:addr.months,mid:ex.id};
  const nm={id:uid(),year:y,month:m,payments:[]};
  return{months:[...(addr.months||[]),nm],mid:nm.id};
};
const mapM=(addr,y,m,fn)=>({...addr,months:(addr.months||[]).map(x=>x.year===y&&x.month===m?fn(x):x)});

function red(s,a){
  switch(a.t){
    case"LOAD":return{...INIT,...a.d};
    case"DARK":return{...s,settings:{...s.settings,dark:!s.settings.dark}};
    case"SS":return{...s,settings:{...s.settings,[a.k]:a.v}};
    case"YM":return{...s,...(a.year!==undefined&&{selYear:a.year}),...(a.month!==undefined&&{selMonth:a.month})};
    case"SA":return{...s,selAddr:a.id};
    // Direcciones
    case"AA":{const n={id:uid(),name:a.name,address:a.addr||"",template:{categories:[]},months:[],knownExtras:[]};return{...s,addresses:[...s.addresses,n],selAddr:n.id};}
    case"UA":return{...s,addresses:upA(s.addresses,a.id,x=>({...x,name:a.name,address:a.addr||""}))};
    case"DA":return{...s,addresses:s.addresses.filter(x=>x.id!==a.id),selAddr:s.selAddr===a.id?null:s.selAddr};
    // Plantilla — categorías
    case"ATC":return{...s,addresses:upA(s.addresses,a.aid,ad=>({...ad,template:{...ad.template,categories:[...(ad.template?.categories||[]),{id:a.id||uid(),name:a.name,icon:a.icon,color:a.color,items:[]}]}}))};
    case"UTC":return{...s,addresses:upA(s.addresses,a.aid,ad=>{
      const cats=(ad.template?.categories||[]).map(c=>c.id===a.cid?{...c,name:a.name,icon:a.icon,color:a.color}:c);
      const ke=(ad.knownExtras||[]).map(k=>k.catId===a.cid?{...k,catName:a.name}:k);
      const months=(ad.months||[]).map(m=>({...m,payments:(m.payments||[]).map(p=>p.catId===a.cid?{...p,catName:a.name}:p)}));
      return{...ad,template:{...ad.template,categories:cats},knownExtras:ke,months};
    })};
    case"DTC":return{...s,addresses:upA(s.addresses,a.aid,ad=>({...ad,template:{...ad.template,categories:(ad.template?.categories||[]).filter(c=>c.id!==a.cid)},knownExtras:(ad.knownExtras||[]).map(k=>k.catId===a.cid?{...k,catId:null}:k)}))};
    case"MERGE_CAT":return{...s,addresses:upA(s.addresses,a.aid,ad=>{
      const cats=ad.template?.categories||[];
      const src=cats.find(c=>c.id===a.srcId),tgt=cats.find(c=>c.id===a.tgtId);
      if(!src||!tgt||src.id===tgt.id)return ad;
      const newCats=cats.filter(c=>c.id!==src.id).map(c=>c.id===tgt.id?{...c,items:[...(c.items||[]),...(src.items||[])]}:c);
      const newKE=(ad.knownExtras||[]).map(k=>k.catId===src.id?{...k,catId:tgt.id,catName:tgt.name}:k);
      const newMonths=(ad.months||[]).map(m=>({...m,payments:(m.payments||[]).map(p=>(p.catId===src.id||(!p.catId&&p.catName===src.name))?{...p,catId:tgt.id,catName:tgt.name}:p)}));
      return{...ad,template:{...ad.template,categories:newCats},knownExtras:newKE,months:newMonths};
    })};
    // Plantilla — ítems
    case"ATI":return{...s,addresses:upA(s.addresses,a.aid,ad=>({...ad,template:{...ad.template,categories:(ad.template?.categories||[]).map(c=>c.id===a.cid?{...c,items:[...(c.items||[]),{id:uid(),name:a.name,amount:a.amount||0,isVariable:a.isVariable||false,frequency:a.frequency||"mensual"}]}:c)}}))};
    case"UTI":return{...s,addresses:upA(s.addresses,a.aid,ad=>({...ad,template:{...ad.template,categories:(ad.template?.categories||[]).map(c=>c.id===a.cid?{...c,items:(c.items||[]).map(it=>it.id===a.iid?{...it,name:a.name,amount:a.amount||0,isVariable:a.isVariable||false,frequency:a.frequency||"mensual"}:it)}:c)}}))};
    case"DTI":return{...s,addresses:upA(s.addresses,a.aid,ad=>({...ad,template:{...ad.template,categories:(ad.template?.categories||[]).map(c=>c.id===a.cid?{...c,items:(c.items||[]).filter(it=>it.id!==a.iid)}:c)}}))};
    // Pagos del mes
    case"ADD_PAY":return{...s,addresses:upA(s.addresses,a.aid,ad=>{
      const{months,mid}=gOrM(ad,s.selYear,s.selMonth);
      return{...ad,months:months.map(m=>m.id===mid?{...m,payments:[...(m.payments||[]),{id:uid(),...a.d}]}:m)};
    })};
    case"DEL_PAY":return{...s,addresses:upA(s.addresses,a.aid,ad=>mapM(ad,s.selYear,s.selMonth,m=>({...m,payments:(m.payments||[]).filter(p=>p.id!==a.pid)})))};
    case"UPD_PAY":return{...s,addresses:upA(s.addresses,a.aid,ad=>({...ad,months:(ad.months||[]).map(m=>({...m,payments:(m.payments||[]).map(p=>p.id===a.pid?{...p,...a.d}:p)}))}))};
    // Gastos frecuentes conocidos (autocompletado inteligente)
    case"AKE":return{...s,addresses:upA(s.addresses,a.aid,ad=>({...ad,knownExtras:[...(ad.knownExtras||[]),{useCount:1,lastAmount:0,lastDate:"",...a.d}]}))};
    case"BKE":return{...s,addresses:upA(s.addresses,a.aid,ad=>({...ad,knownExtras:(ad.knownExtras||[]).map(k=>k.id===a.id?{...k,useCount:(k.useCount||0)+1,lastAmount:a.amount,lastDate:a.date}:k)}))};
    case"DKE":return{...s,addresses:upA(s.addresses,a.aid,ad=>({...ad,knownExtras:(ad.knownExtras||[]).filter(k=>k.id!==a.id)}))};
    // Perfil
    case"SP":return{...s,profile:{...s.profile,...a.d}};
    case"AINC":return{...s,profile:{...s.profile,incomes:[...s.profile.incomes,{id:uid(),...a.d}]}};
    case"UINC":return{...s,profile:{...s.profile,incomes:s.profile.incomes.map(i=>i.id===a.id?{...i,...a.d}:i)}};
    case"DINC":return{...s,profile:{...s.profile,incomes:s.profile.incomes.filter(i=>i.id!==a.id)}};
    case"AG":return{...s,profile:{...s.profile,goals:[...s.profile.goals,{id:uid(),...a.d}]}};
    case"DG":return{...s,profile:{...s.profile,goals:s.profile.goals.filter(g=>g.id!==a.id)}};
    case"SAV":return{...s,profile:{...s.profile,availability:{...s.profile.availability,...a.d}}};
    case"APM":return{...s,profile:{...s.profile,paymentMethods:[...(s.profile.paymentMethods||[]),{id:uid(),name:a.name,isDefault:(s.profile.paymentMethods||[]).length===0}]}};
    case"DPM":return{...s,profile:{...s.profile,paymentMethods:(s.profile.paymentMethods||[]).filter(m=>m.id!==a.id)}};
    case"SETDEFPM":return{...s,profile:{...s.profile,paymentMethods:(s.profile.paymentMethods||[]).map(m=>({...m,isDefault:m.id===a.id}))}};
    // Tutorial
    case"TUT_DONE":return{...s,settings:{...s.settings,tutDone:true}};
    case"TUT_SHOW":return{...s,settings:{...s.settings,tutDone:false}};
    // IA
    case"ADAI":return{...s,aiHistory:[{id:uid(),date:new Date().toLocaleDateString("es-CL"),...a.d},...s.aiHistory].slice(0,20)};
    case"CLAI":return{...s,aiHistory:[]};
    default:return s;
  }
}

/* ══ TEMA ════════════════════════════════════════════════ */
const mkT=dark=>({
  bg:dark?"#0D1117":"#EEF2F7",card:dark?"#161B22":"#FFFFFF",card2:dark?"#21262D":"#F8FAFC",card3:dark?"#2D333B":"#EEF2FF",
  border:dark?"#30363D":"#E2E8F0",text:dark?"#E6EDF3":"#1E293B",muted:dark?"#7D8590":"#64748B",dim:dark?"#484F58":"#94A3B8",
  pri:"#7C3AED",priL:dark?"#A78BFA":"#6D28D9",ok:"#10B981",warn:"#F59E0B",err:"#EF4444",info:"#06B6D4",
  shadow:dark?"0 4px 24px rgba(0,0,0,.55)":"0 4px 24px rgba(0,0,0,.08)",sm:dark?"0 2px 8px rgba(0,0,0,.4)":"0 2px 8px rgba(0,0,0,.06)",
});

/* ══ COMPONENTES COMPARTIDOS ═════════════════════════════ */
function Btn({onClick,children,v="pri",sz="md",dis,icon,full,sx={}}){
  const VS={pri:{bg:"#7C3AED",cl:"#fff",b:"none"},ok:{bg:"#10B981",cl:"#fff",b:"none"},err:{bg:"#EF4444",cl:"#fff",b:"none"},warn:{bg:"#F59E0B",cl:"#1e293b",b:"none"},ghost:{bg:"transparent",cl:"inherit",b:"1px solid #94A3B8"},out:{bg:"transparent",cl:"#7C3AED",b:"1px solid #7C3AED"}};
  const SS={sm:{p:"0.3rem 0.65rem",fs:"0.78rem",r:"0.4rem"},md:{p:"0.5rem 1rem",fs:"0.85rem",r:"0.5rem"},lg:{p:"0.75rem 1.5rem",fs:"0.97rem",r:"0.6rem"}};
  const st=VS[v]||VS.pri,s=SS[sz]||SS.md;
  return <button onClick={onClick} disabled={dis} style={{display:"flex",alignItems:"center",gap:"0.4rem",cursor:dis?"not-allowed":"pointer",opacity:dis?0.45:1,fontWeight:600,width:full?"100%":"auto",justifyContent:"center",transition:"opacity .15s",background:st.bg,color:st.cl,border:st.b,padding:s.p,fontSize:s.fs,borderRadius:s.r,...sx}}>{icon}{children}</button>;
}
function CCard({ch,t,sx={},onClick}){return <div onClick={onClick} style={{background:t.card,border:`1px solid ${t.border}`,borderRadius:"0.75rem",padding:"1.25rem",boxShadow:t.sm,...sx,cursor:onClick?"pointer":"default"}}>{ch}</div>;}
function Modal({title,onClose,ch,t,w="490px"}){
  return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.65)",zIndex:500,display:"flex",alignItems:"center",justifyContent:"center",padding:"1rem"}}>
    <div style={{background:t.card,border:`1px solid ${t.border}`,borderRadius:"0.875rem",padding:"1.5rem",width:"100%",maxWidth:w,maxHeight:"90vh",overflowY:"auto",boxShadow:t.shadow}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"1.25rem"}}>
        <h3 style={{margin:0,color:t.text,fontSize:"1.05rem",fontWeight:700}}>{title}</h3>
        <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",color:t.muted,fontSize:"1.6rem",lineHeight:1,padding:"0 0.3rem"}}>×</button>
      </div>{ch}
    </div>
  </div>;
}
function Cfm({msg,onOk,onCancel,t}){
  return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.65)",zIndex:600,display:"flex",alignItems:"center",justifyContent:"center"}}>
    <div style={{background:t.card,border:`1px solid ${t.border}`,borderRadius:"0.75rem",padding:"1.75rem",maxWidth:"360px",width:"90%",textAlign:"center",boxShadow:t.shadow}}>
      <AlertTriangle size={32} color={t.warn} style={{marginBottom:"0.75rem"}}/>
      <p style={{color:t.text,margin:"0 0 1.5rem",lineHeight:1.6}}>{msg}</p>
      <div style={{display:"flex",gap:"0.75rem",justifyContent:"center"}}><Btn onClick={onCancel} v="ghost">Cancelar</Btn><Btn onClick={onOk} v="err">Confirmar</Btn></div>
    </div>
  </div>;
}
function MoneyInput({val,onChange,ph="$0",t,sx={},af}){
  const[disp,setDisp]=useState(val>0?fCLP(val):"");const ref=useRef();
  useEffect(()=>{setDisp(val>0?fCLP(val):"");},[val]);
  useEffect(()=>{if(af&&ref.current)ref.current.focus();},[af]);
  return <input ref={ref} type="text" value={disp}
    onChange={e=>{const n=pCLP(e.target.value);setDisp(fCLP(n));onChange(n);}}
    onFocus={e=>setTimeout(()=>e.target.select(),10)}
    onBlur={()=>setDisp(val>0?fCLP(val):"")}
    placeholder={ph} style={{width:"100%",padding:"0.5rem 0.75rem",background:t.card2,border:`1px solid ${t.border}`,borderRadius:"0.5rem",color:t.text,fontSize:"0.9rem",boxSizing:"border-box",outline:"none",...sx}}/>;
}
function TI({val,onChange,ph,t,type="text",sx={}}){return <input type={type} value={val} onChange={e=>onChange(e.target.value)} placeholder={ph} style={{width:"100%",padding:"0.5rem 0.75rem",background:t.card2,border:`1px solid ${t.border}`,borderRadius:"0.5rem",color:t.text,fontSize:"0.875rem",boxSizing:"border-box",outline:"none",...sx}}/>;}
function Sel({val,onChange,opts,t}){return <select value={val} onChange={e=>onChange(e.target.value)} style={{width:"100%",padding:"0.5rem 0.75rem",background:t.card2,border:`1px solid ${t.border}`,borderRadius:"0.5rem",color:t.text,fontSize:"0.875rem",outline:"none"}}>{opts.map(o=><option key={o.v??o} value={o.v??o}>{o.l??o}</option>)}</select>;}
function TA({val,onChange,ph,rows=3,t}){return <textarea value={val} onChange={e=>onChange(e.target.value)} placeholder={ph} rows={rows} style={{width:"100%",padding:"0.5rem 0.75rem",background:t.card2,border:`1px solid ${t.border}`,borderRadius:"0.5rem",color:t.text,fontSize:"0.875rem",boxSizing:"border-box",resize:"vertical",outline:"none"}}/>;}
function Fld({label,ch,t}){return <div style={{marginBottom:"0.85rem"}}>{label&&<label style={{display:"block",color:t.muted,fontSize:"0.73rem",fontWeight:700,marginBottom:"0.3rem",textTransform:"uppercase",letterSpacing:"0.04em"}}>{label}</label>}{ch}</div>;}
function Badge({ch,color="#7C3AED"}){return <span style={{display:"inline-flex",alignItems:"center",padding:"0.15rem 0.55rem",borderRadius:"9999px",fontSize:"0.71rem",fontWeight:700,background:`${color}22`,color}}>{ch}</span>;}
function StCard({label,val,sub,color,icon,t}){
  return <div style={{background:t.card,border:`1px solid ${t.border}`,borderRadius:"0.75rem",padding:"1rem",boxShadow:t.sm,flex:1,minWidth:"130px"}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
      <div style={{flex:1,minWidth:0}}>
        <p style={{margin:0,color:t.muted,fontSize:"0.68rem",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.05em"}}>{label}</p>
        <p style={{margin:"0.3rem 0 0",color:color||t.text,fontSize:"1.2rem",fontWeight:700,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{val}</p>
        {sub&&<p style={{margin:"0.15rem 0 0",color:t.dim,fontSize:"0.71rem"}}>{sub}</p>}
      </div>
      {icon&&<div style={{background:`${color||t.pri}18`,borderRadius:"0.5rem",padding:"0.45rem",marginLeft:"0.75rem",display:"flex",flexShrink:0}}>{icon}</div>}
    </div>
  </div>;
}
function SH({title,sub,action,t}){
  return <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"1.1rem"}}>
    <div><h2 style={{margin:0,color:t.text,fontSize:"1.3rem",fontWeight:800}}>{title}</h2>{sub&&<p style={{margin:"0.2rem 0 0",color:t.muted,fontSize:"0.82rem"}}>{sub}</p>}</div>
    {action}
  </div>;
}
function Empty({icon,title,sub,action,t}){
  return <div style={{textAlign:"center",padding:"2.5rem 1rem"}}>
    <div style={{color:t.dim,marginBottom:"0.75rem"}}>{icon}</div>
    <h3 style={{color:t.text,margin:"0 0 0.5rem",fontSize:"1rem"}}>{title}</h3>
    {sub&&<p style={{color:t.muted,margin:"0 0 1.25rem",fontSize:"0.875rem",lineHeight:1.5}}>{sub}</p>}
    {action}
  </div>;
}
function ProgressBar({value,max,t}){
  const pct=max>0?Math.min(100,Math.round(value/max*100)):0;
  const color=pct>90?t.err:pct>70?t.warn:t.ok;
  return <div>
    <div style={{height:"8px",background:t.border,borderRadius:"4px",overflow:"hidden"}}>
      <div style={{height:"100%",width:`${pct}%`,background:`linear-gradient(90deg,${color},${color}cc)`,borderRadius:"4px",transition:"width .5s ease"}}/>
    </div>
    <div style={{display:"flex",justifyContent:"space-between",marginTop:"0.3rem"}}>
      <span style={{color:t.muted,fontSize:"0.72rem"}}>Pagado: {fCLP(value)}</span>
      <span style={{color:t.muted,fontSize:"0.72rem"}}>{pct}% de {fCLP(max)}</span>
    </div>
  </div>;
}

/* Modal de pago rápido — usado en Inicio e Historial */
/* Autocompletado de gastos frecuentes: recuerda lugares/ítems usados antes,
   ordenados por frecuencia de uso (los usados una sola vez se hunden solos) */
function ExtraNameField({value,onChange,onPick,addr,t}){
  const[open,setOpen]=useState(false);
  const[hi,setHi]=useState(0);
  const sugg=searchKnownExtras(addr,value);
  const showDropdown=open&&sugg.length>0;
  const pick=k=>{onPick(k);setOpen(false);};
  return <div style={{position:"relative"}}>
    <input value={value}
      onChange={e=>{onChange(e.target.value);setOpen(true);setHi(0);}}
      onFocus={()=>setOpen(true)}
      onBlur={()=>setTimeout(()=>setOpen(false),150)}
      onKeyDown={e=>{
        if(!showDropdown)return;
        if(e.key==="ArrowDown"){e.preventDefault();setHi(p=>Math.min(p+1,sugg.length-1));}
        else if(e.key==="ArrowUp"){e.preventDefault();setHi(p=>Math.max(p-1,0));}
        else if(e.key==="Tab"||e.key==="Enter"){if(sugg[hi]){e.preventDefault();pick(sugg[hi]);}}
        else if(e.key==="Escape")setOpen(false);
      }}
      placeholder="Luz, Almacén, Netflix, Uber..."
      style={{width:"100%",padding:"0.5rem 0.75rem",background:t.card2,border:`1px solid ${t.border}`,borderRadius:"0.5rem",color:t.text,fontSize:"0.875rem",boxSizing:"border-box",outline:"none"}}/>
    {showDropdown&&<div style={{position:"absolute",top:"calc(100% + 4px)",left:0,right:0,background:t.card,border:`1px solid ${t.border}`,borderRadius:"0.5rem",boxShadow:t.shadow,zIndex:50,maxHeight:"220px",overflowY:"auto"}}>
      {sugg.map((k,i)=><div key={k.id} onMouseDown={()=>pick(k)} style={{display:"flex",alignItems:"center",gap:"0.5rem",padding:"0.5rem 0.7rem",background:i===hi?t.pri+"15":"transparent",cursor:"pointer",borderBottom:i<sugg.length-1?`1px solid ${t.border}`:"none"}}>
        <span style={{flex:1,color:t.text,fontSize:"0.83rem",fontWeight:i===hi?700:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{k.name}</span>
        {k.catName&&<Badge ch={k.catName} color={t.pri}/>}
        <span style={{color:t.dim,fontSize:"0.68rem",whiteSpace:"nowrap"}}>{k.useCount||1}×</span>
      </div>)}
      <div style={{padding:"0.3rem 0.7rem",color:t.dim,fontSize:"0.66rem",borderTop:`1px solid ${t.border}`}}>↑↓ navegar · Tab/Enter elegir</div>
    </div>}
  </div>;
}
function InlineCatPicker({cats,onPick,t}){
  const[adding,setAdding]=useState(false);const[nn,setNn]=useState("");
  return <div>
    <div style={{display:"flex",flexWrap:"wrap",gap:"0.4rem"}}>
      {cats.map(c=><button key={c.id} onClick={()=>onPick(c)} style={{display:"flex",alignItems:"center",gap:"0.3rem",padding:"0.3rem 0.65rem",background:t.card2,border:`1px solid ${t.border}`,borderRadius:"9999px",cursor:"pointer",color:t.text,fontSize:"0.78rem"}}>
        <span>{c.icon||"📦"}</span>{c.name}
      </button>)}
      {!adding&&<button onClick={()=>setAdding(true)} style={{display:"flex",alignItems:"center",gap:"0.3rem",padding:"0.3rem 0.65rem",background:"transparent",border:`1px dashed ${t.pri}`,borderRadius:"9999px",cursor:"pointer",color:t.pri,fontSize:"0.78rem",fontWeight:600}}>
        <Plus size={11}/> Nueva categoría
      </button>}
    </div>
    {adding&&<div style={{display:"flex",gap:"0.4rem",marginTop:"0.5rem"}}>
      <TI val={nn} onChange={setNn} ph="Nombre de la categoría..." t={t}/>
      <Btn onClick={()=>{if(nn.trim()){onPick({__new:true,name:nn.trim()});setAdding(false);setNn("");}}} dis={!nn.trim()} sz="sm" icon={<Check size={12}/>}>Crear</Btn>
    </div>}
  </div>;
}
/* Selector de forma de pago: prioriza las cuentas/métodos frecuentes que el
   usuario definió en su Perfil (con la principal pre-seleccionada), y deja
   los genéricos (Efectivo, etc.) siempre disponibles debajo */
function methodOptionsFor(profile){
  const custom=[...(profile?.paymentMethods||[])].sort((a,b)=>(b.isDefault?1:0)-(a.isDefault?1:0));
  return{custom,all:[...new Set([...custom.map(m=>m.name),...PAY_METHODS])]};
}
function MethodChips({value,onChange,profile,t}){
  const{custom,all}=methodOptionsFor(profile);
  return <div>
    {custom.length>0&&<div style={{display:"flex",flexWrap:"wrap",gap:"0.4rem",marginBottom:"0.5rem"}}>
      {custom.map(m=><button key={m.id} onClick={()=>onChange(m.name)} style={{display:"flex",alignItems:"center",gap:"0.3rem",padding:"0.35rem 0.7rem",background:value===m.name?t.pri:t.card2,color:value===m.name?"#fff":t.text,border:`1px solid ${value===m.name?t.pri:t.border}`,borderRadius:"9999px",cursor:"pointer",fontSize:"0.79rem",fontWeight:value===m.name?700:500}}>
        {m.isDefault&&<Star size={11} fill={value===m.name?"#fff":t.warn} color={value===m.name?"#fff":t.warn}/>}{m.name}
      </button>)}
    </div>}
    <Sel val={value} onChange={onChange} opts={all} t={t}/>
  </div>;
}
function PayModal({state,t,onClose,onSave,addr,d,profile}){
  const[name,setName]=useState(state.itemName||"");
  const[amount,setAmount]=useState(state.remaining!==undefined?state.remaining:(state.estimated||0));
  const[date,setDate]=useState(state.date||todayStr());
  const[note,setNote]=useState("");const[method,setMethod]=useState(()=>profile?.paymentMethods?.find(m=>m.isDefault)?.name||"Transferencia");
  const[matched,setMatched]=useState(null);
  const[pickedNewCat,setPickedNewCat]=useState(null);
  const isExtra=state.isExtra!==false;
  const cats=addr?.template?.categories||[];
  const liveMatch=isExtra?matchKnownExtra(addr,name):null;
  const effectiveMatch=matched||liveMatch;
  const needsCategory=isExtra&&name.trim().length>0&&!effectiveMatch&&!pickedNewCat;

  const handlePick=k=>{setMatched(k);setName(k.name);setAmount(k.lastAmount||0);setPickedNewCat(null);};
  const handleCatPick=c=>{
    if(c.__new){
      const newCatId=uid();
      d({t:"ATC",aid:addr.id,id:newCatId,name:c.name,icon:"📦",color:COLORS[cats.length%COLORS.length]});
      setPickedNewCat({id:newCatId,name:c.name});
    }else setPickedNewCat({id:c.id,name:c.name});
  };
  const handleSave=()=>{
    if(!name||!amount)return;
    let catId=null,catName="",knownExtraId=null;
    if(isExtra){
      if(effectiveMatch){
        catId=effectiveMatch.catId;catName=effectiveMatch.catName;knownExtraId=effectiveMatch.id;
        d({t:"BKE",aid:addr.id,id:effectiveMatch.id,amount,date});
      }else if(pickedNewCat){
        catId=pickedNewCat.id;catName=pickedNewCat.name;
        const keId=uid();
        d({t:"AKE",aid:addr.id,d:{id:keId,name,catId,catName}});
        knownExtraId=keId;
      }
    }else{catId=state.catId||null;catName=state.catName||"";}
    onSave({name,amount,date,method,note,catId,catName,knownExtraId,templateItemId:state.itemId||null,isExtra});
  };

  return <Modal title={isExtra?"⚡ Registrar Pago":"✓ Marcar como Pagado"} onClose={onClose} t={t} ch={<>
    {!isExtra&&<div style={{padding:"0.6rem 0.75rem",background:t.ok+"12",borderRadius:"0.5rem",marginBottom:"0.85rem",fontSize:"0.82rem",color:t.muted}}>
      <div style={{display:"flex",gap:"0.5rem",alignItems:"center",marginBottom:state.accumulated>0?"0.4rem":"0"}}><CheckCircle2 size={14} color={t.ok}/><span>Ítem: <strong style={{color:t.text}}>{state.itemName}</strong> · Total mes: <strong style={{color:t.text}}>{fCLP(state.estimated)}</strong></span></div>
      {state.accumulated>0&&<div style={{display:"flex",gap:"0.75rem",fontSize:"0.79rem",paddingLeft:"0.25rem"}}><span>✅ Ya abonado: <strong style={{color:t.ok}}>{fCLP(state.accumulated)}</strong></span><span>⏳ Pendiente: <strong style={{color:t.warn}}>{fCLP(state.remaining)}</strong></span></div>}
    </div>}
    {isExtra?<Fld label="¿Qué pagaste?" ch={<ExtraNameField value={name} onChange={v=>{setName(v);setMatched(null);setPickedNewCat(null);}} onPick={handlePick} addr={addr} t={t}/>} t={t}/>
             :<Fld label="¿Qué pagaste?" ch={<TI val={name} onChange={setName} ph="Nombre del gasto" t={t}/>} t={t}/>}
    {isExtra&&effectiveMatch&&<div style={{display:"flex",alignItems:"center",gap:"0.5rem",padding:"0.5rem 0.7rem",background:t.ok+"12",border:`1px solid ${t.ok}35`,borderRadius:"0.5rem",marginBottom:"0.85rem",fontSize:"0.79rem",color:t.muted}}>
      <CheckCircle2 size={13} color={t.ok}/><span>Ya lo conocemos — va en <strong style={{color:t.text}}>{effectiveMatch.catName||"sin categoría"}</strong> · usado {effectiveMatch.useCount||1}× antes</span>
    </div>}
    {isExtra&&pickedNewCat&&<div style={{display:"flex",alignItems:"center",gap:"0.5rem",padding:"0.5rem 0.7rem",background:t.info+"12",border:`1px solid ${t.info}35`,borderRadius:"0.5rem",marginBottom:"0.85rem",fontSize:"0.79rem",color:t.muted}}>
      <Sparkles size={13} color={t.info}/><span>🆕 Nuevo — quedará guardado en <strong style={{color:t.text}}>{pickedNewCat.name}</strong> para la próxima vez</span>
    </div>}
    {needsCategory&&<div style={{marginBottom:"0.85rem"}}>
      <div style={{color:t.text,fontSize:"0.82rem",fontWeight:600,marginBottom:"0.4rem"}}>Es la primera vez — ¿en qué categoría va "{name}"?</div>
      <InlineCatPicker cats={cats} onPick={handleCatPick} t={t}/>
    </div>}
    <Fld label={isExtra?"Monto pagado":"Monto real (puede diferir del estimado)"} ch={<div>
      <MoneyInput val={amount} onChange={setAmount} t={t} af/>
      {isExtra&&effectiveMatch&&effectiveMatch.lastAmount>0&&amount!==effectiveMatch.lastAmount&&<button onClick={()=>setAmount(effectiveMatch.lastAmount)} style={{background:"none",border:"none",cursor:"pointer",color:t.pri,fontSize:"0.74rem",marginTop:"0.3rem",display:"flex",alignItems:"center",gap:"0.25rem"}}><Sparkles size={11}/> Usar último monto: {fCLP(effectiveMatch.lastAmount)}</button>}
    </div>} t={t}/>
    <Fld label="Fecha" ch={<TI val={date} onChange={setDate} type="date" t={t}/>} t={t}/>
    <Fld label="Forma de pago" ch={<MethodChips value={method} onChange={setMethod} profile={profile} t={t}/>} t={t}/>
    <Fld label="Nota (opcional)" ch={<TI val={note} onChange={setNote} ph="Info extra..." t={t}/>} t={t}/>
    <Btn onClick={handleSave} dis={!name||!amount} full icon={<Check size={14}/>}>{isExtra?"Registrar pago ⚡":"Confirmar pago ✓"}</Btn>
  </>}/>;
}

/* Editar un pago ya registrado — monto, fecha, forma de pago y nota,
   sin tener que borrarlo y crearlo de nuevo */
function EditPayModal({pay,t,onClose,onSave,profile}){
  const[name,setName]=useState(pay.name||"");
  const[amount,setAmount]=useState(pay.amount||0);
  const[date,setDate]=useState(pay.date||todayStr());
  const[method,setMethod]=useState(pay.method||"Transferencia");
  const[note,setNote]=useState(pay.note||"");
  return <Modal title="✏️ Editar Pago" onClose={onClose} t={t} ch={<>
    <Fld label="Nombre" ch={<TI val={name} onChange={setName} t={t}/>} t={t}/>
    <Fld label="Monto" ch={<MoneyInput val={amount} onChange={setAmount} t={t} af/>} t={t}/>
    <Fld label="Fecha" ch={<TI val={date} onChange={setDate} type="date" t={t}/>} t={t}/>
    <Fld label="Forma de pago" ch={<MethodChips value={method} onChange={setMethod} profile={profile} t={t}/>} t={t}/>
    <Fld label="Nota (opcional)" ch={<TI val={note} onChange={setNote} ph="Info extra..." t={t}/>} t={t}/>
    <Btn onClick={()=>{if(name&&amount)onSave({name,amount,date,method,note});}} dis={!name||!amount} full icon={<Check size={14}/>}>Guardar cambios</Btn>
  </>}/>;
}

/* ══ TUTORIAL ════════════════════════════════════════════ */
function Tutorial({onDone,t}){
  const[step,setStep]=useState(0);
  const STEPS=[
    {e:"👋",title:"¡Hola! Bienvenido a GestorGastos",desc:"Esta app está hecha para ser simple en el uso diario.\n\nEl objetivo: abres la app, registras lo que pagaste, la cierras.\n\nTe mostramos cómo funciona todo. Puedes volver aquí con el botón ❓ del menú.",hint:""},
    {e:"🏠",title:"Inicio — Tu centro de operaciones",desc:"Aquí está todo lo que necesitas en el día a día:\n\n• Cuánto llevas pagado este mes vs tu presupuesto\n• Tus gastos fijos pendientes de pagar (un tap para marcarlos)\n• Botón grande para registrar cualquier gasto en segundos\n• Chips de \"pago frecuente\" — repite un gasto que ya hiciste antes con un solo tap\n• Al escribir, la app recuerda lugares que ya usaste y sugiere su categoría automáticamente. Si es nuevo, te deja crear la categoría ahí mismo",hint:"💡 Esta es la pantalla que abrirás todos los días. El flujo es: abrir → pagar → cerrar."},
    {e:"📋",title:"Gastos Fijos — Configuras una vez",desc:"Aquí defines los gastos que tienes TODOS los meses:\n• Arriendo: $350.000 (fijo)\n• Luz: ~$40.000 (variable, cambia cada mes)\n• Agua, internet, suscripciones...\n\nLos configuras UNA VEZ y aparecen como pendientes automáticamente cada mes. Nunca más tienes que volver a ingresarlos.",hint:"💡 Si un monto varía (luz, agua), márcalo como 'variable'. El monto que pones es solo una estimación."},
    {e:"📅",title:"Historial — El mes en detalle",desc:"El calendario mensual muestra todo:\n• Los días con pagos tienen puntos de color\n• Verde = gasto fijo pagado · Naranja = gasto extra\n• Click en un día para ver o agregar pagos\n• Navega entre meses con las flechas ‹ ›",hint:""},
    {e:"📊",title:"Análisis — Entiende tus gastos",desc:"Para cuando quieres ver el panorama:\n• Gráfico de presupuesto vs lo que realmente pagaste\n• Distribución por categoría\n• Evolución mes a mes\n\nExporta a Excel o imprime como PDF.",hint:"💡 Úsalo una vez al mes. El resto del tiempo, usa solo la pantalla de Inicio."},
    {e:"👤",title:"Perfil — Tus datos e ingresos",desc:"Configura:\n• Tu nombre (aparece en el saludo)\n• Fuentes de ingreso (sueldo, beca, etc.)\n• Objetivos de ahorro\n• API key para la IA\n\n¡Importante! Aquí también puedes EXPORTAR tus datos como respaldo JSON e IMPORTARLOS si cambias de dispositivo.",hint:"💡 Exporta tus datos regularmente como respaldo. Los datos viven en tu navegador."},
    {e:"🧠",title:"Plan IA — Tu asesor financiero",desc:"Usando tu presupuesto y gastos reales, la IA genera:\n💰 Plan de ahorro personalizado\n🔍 Diagnóstico categoría por categoría\n🆘 Presupuesto de supervivencia\n🚀 Ideas de ingresos extra para Chile\n📊 Análisis del método 50/30/20",hint:"🔑 Necesitas API key gratuita de Anthropic (platform.claude.com). En la pestaña Plan IA hay una guía paso a paso."},
  ];
  const cur=STEPS[step],isLast=step===STEPS.length-1;
  return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.88)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:"1rem"}}>
    <div style={{background:t.card,border:`1px solid ${t.border}`,borderRadius:"1.1rem",padding:"2rem",maxWidth:"520px",width:"100%",boxShadow:"0 20px 60px rgba(0,0,0,.6)"}}>
      <div style={{display:"flex",justifyContent:"center",gap:"0.35rem",marginBottom:"1.5rem"}}>
        {STEPS.map((_,i)=><div key={i} style={{width:i===step?"22px":"7px",height:"7px",borderRadius:"4px",background:i===step?t.pri:t.border,transition:"all .3s"}}/>)}
      </div>
      <div style={{textAlign:"center",marginBottom:"1.5rem"}}>
        <div style={{fontSize:"2.8rem",marginBottom:"0.75rem",lineHeight:1}}>{cur.e}</div>
        <h2 style={{margin:"0 0 0.85rem",color:t.text,fontSize:"1.15rem",fontWeight:800}}>{cur.title}</h2>
        <p style={{margin:0,color:t.muted,fontSize:"0.87rem",lineHeight:1.75,whiteSpace:"pre-line",textAlign:"left"}}>{cur.desc}</p>
        {cur.hint&&<div style={{marginTop:"0.8rem",padding:"0.6rem 0.85rem",background:t.card2,borderRadius:"0.5rem",color:t.info,fontSize:"0.79rem",textAlign:"left",lineHeight:1.5}}>{cur.hint}</div>}
      </div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <button onClick={onDone} style={{background:"none",border:"none",cursor:"pointer",color:t.dim,fontSize:"0.8rem",padding:"0.3rem"}}>Saltar</button>
        <div style={{display:"flex",gap:"0.5rem"}}>
          {step>0&&<Btn onClick={()=>setStep(p=>p-1)} v="ghost" sz="sm" icon={<ChevronLeft size={13}/>}>Anterior</Btn>}
          <Btn onClick={()=>isLast?onDone():setStep(p=>p+1)} v={isLast?"ok":"pri"}>{isLast?"¡Empezar! 🚀":"Siguiente →"}</Btn>
        </div>
      </div>
      <div style={{textAlign:"center",marginTop:"0.75rem",color:t.dim,fontSize:"0.72rem"}}>Paso {step+1} de {STEPS.length}</div>
    </div>
  </div>;
}

/* ══ NAVBAR ══════════════════════════════════════════════ */
function Navbar({view,setView,s,d,t}){
  const addr=s.addresses.find(a=>a.id===s.selAddr);
  const tabs=[
    {id:"hoy",ic:<Home size={16}/>,lb:"Inicio"},
    {id:"plantilla",ic:<List size={16}/>,lb:"Gastos Fijos"},
    {id:"historial",ic:<Calendar size={16}/>,lb:"Historial"},
    {id:"analisis",ic:<BarChart2 size={16}/>,lb:"Análisis"},
    {id:"perfil",ic:<User size={16}/>,lb:"Perfil"},
    {id:"ia",ic:<Brain size={16}/>,lb:"IA"},
  ];
  return <nav style={{position:"sticky",top:0,zIndex:200,background:t.card,borderBottom:`1px solid ${t.border}`,boxShadow:t.sm}}>
    <div style={{maxWidth:"1000px",margin:"0 auto",padding:"0 0.75rem",display:"flex",alignItems:"center",gap:"0.15rem",height:"54px"}}>
      <div style={{display:"flex",alignItems:"center",gap:"0.55rem",marginRight:"0.65rem",paddingRight:"0.65rem",borderRight:`1px solid ${t.border}`,flexShrink:0}}>
        <div style={{background:"linear-gradient(135deg,#7C3AED,#A78BFA)",borderRadius:"0.45rem",width:"30px",height:"30px",display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontWeight:900,fontSize:"1rem"}}>$</div>
        <div>
          <div style={{color:t.text,fontWeight:800,fontSize:"0.87rem",lineHeight:1.1}}>GestorGastos</div>
          {addr&&<div style={{color:t.muted,fontSize:"0.65rem"}}>{addr.name}</div>}
        </div>
      </div>
      <div style={{display:"flex",gap:"0.05rem",flex:1,overflowX:"auto"}}>
        {tabs.map(tb=><button key={tb.id} onClick={()=>setView(tb.id)} style={{display:"flex",alignItems:"center",gap:"0.3rem",padding:"0.4rem 0.6rem",background:view===tb.id?`${t.pri}18`:"none",color:view===tb.id?t.pri:t.muted,border:"none",cursor:"pointer",borderRadius:"0.45rem",fontSize:"0.78rem",fontWeight:view===tb.id?700:400,whiteSpace:"nowrap",transition:"all .15s"}}>
          {tb.ic}<span style={{display:window.innerWidth<700?"none":"inline"}}>{tb.lb}</span>
        </button>)}
      </div>
      <div style={{display:"flex",gap:"0.15rem",flexShrink:0}}>
        <button onClick={()=>d({t:"DARK"})} style={{background:"none",border:"none",cursor:"pointer",color:t.muted,padding:"0.4rem",borderRadius:"0.4rem",display:"flex",alignItems:"center"}}>
          {s.settings.dark?<Sun size={17}/>:<Moon size={17}/>}
        </button>
        <button onClick={()=>d({t:"TUT_SHOW"})} title="Tutorial de uso" style={{background:"none",border:"none",cursor:"pointer",color:t.pri,padding:"0.4rem",borderRadius:"0.4rem",display:"flex",alignItems:"center"}}>
          <HelpCircle size={17}/>
        </button>
      </div>
    </div>
  </nav>;
}

/* ══ HOY — Pantalla principal (daily driver) ═════════════ */
function HoyView({s,d,t,setView}){
  const[payModal,setPayModal]=useState(null);
  const[cfm,setCfm]=useState(null);
  const[editPay,setEditPay]=useState(null);
  const addr=s.addresses.find(a=>a.id===s.selAddr);
  const budget=addr?tmplTotal(addr):0;
  const paid=addr?totalPaid(addr,s.selYear,s.selMonth):0;
  const inc=sumInc(s.profile);
  const pending=addr?pendingTmpl(addr,s.selYear,s.selMonth):[];
  const paidItems=addr?paidTmpl(addr,s.selYear,s.selMonth):[];
  const extras=addr?extraPays(addr,s.selYear,s.selMonth):[];
  const catNames=[...new Set((addr?.template?.categories||[]).map(c=>c.name))];
  const suggestions=addr?smartSuggestions(addr):[];

  return <div>
    {/* Saludo */}
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:"0.5rem",marginBottom:"1.25rem"}}>
      <div>
        <h2 style={{margin:0,color:t.text,fontSize:"1.4rem",fontWeight:800}}>
          {fmtGreeting()}{s.profile.name?`, ${s.profile.name}`:""}! 👋
        </h2>
        <p style={{margin:"0.2rem 0 0",color:t.muted,fontSize:"0.83rem"}}>{MESES[s.selMonth]} {s.selYear}{addr?` · ${addr.name}`:""}</p>
      </div>
      <div style={{display:"flex",gap:"0.4rem"}}>
        <select value={s.selMonth} onChange={e=>d({t:"YM",month:+e.target.value})} style={{padding:"0.3rem 0.5rem",background:t.card2,border:`1px solid ${t.border}`,borderRadius:"0.4rem",color:t.text,fontSize:"0.78rem"}}>
          {MESES.map((m,i)=><option key={i} value={i}>{m.substr(0,3)}</option>)}
        </select>
        <select value={s.selYear} onChange={e=>d({t:"YM",year:+e.target.value})} style={{padding:"0.3rem 0.5rem",background:t.card2,border:`1px solid ${t.border}`,borderRadius:"0.4rem",color:t.text,fontSize:"0.78rem"}}>
          {YEARS.map(y=><option key={y} value={y}>{y}</option>)}
        </select>
      </div>
    </div>

    {/* Onboarding si no hay dirección */}
    {!addr&&<div>
      <CCard ch={<div style={{textAlign:"center",padding:"1.5rem 0"}}>
        <div style={{fontSize:"3rem",marginBottom:"0.75rem"}}>💸</div>
        <h3 style={{color:t.text,margin:"0 0 0.75rem",fontSize:"1.1rem",fontWeight:800}}>Comencemos con 3 pasos</h3>
        <p style={{color:t.muted,margin:"0 0 1.5rem",fontSize:"0.875rem",lineHeight:1.6}}>GestorGastos te muestra cuánto gastas, cuánto tienes que pagar cada mes y dónde puedes ahorrar.</p>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"0.75rem",maxWidth:"380px",margin:"0 auto"}}>
          {[["1","Ingresa tu nombre e ingresos","👤",()=>setView("perfil")],["2","Configura tus gastos fijos","📋",()=>setView("plantilla")],["3","¡Ya está! Registra pagos","✅",null]].map(([n,desc,e,action])=><div key={n} onClick={action||undefined} style={{padding:"1rem",background:t.card2,borderRadius:"0.65rem",border:`1px solid ${t.border}`,textAlign:"left",cursor:action?"pointer":"default"}}>
            <div style={{fontSize:"1.2rem",marginBottom:"0.4rem"}}>{e}</div>
            <div style={{color:t.text,fontWeight:700,fontSize:"0.82rem",marginBottom:"0.15rem"}}>Paso {n}</div>
            <div style={{color:t.dim,fontSize:"0.75rem"}}>{desc}</div>
          </div>)}
        </div>
      </div>} t={t}/>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"0.75rem",marginTop:"1rem"}}>
        <Btn onClick={()=>setView("perfil")} v="out" icon={<User size={14}/>} full>Mi Perfil</Btn>
        <Btn onClick={()=>setView("plantilla")} icon={<List size={14}/>} full>Gastos Fijos</Btn>
      </div>
    </div>}

    {addr&&<>
      {/* Progreso del mes */}
      <CCard ch={<>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"0.75rem"}}>
          <div>
            <span style={{color:t.muted,fontSize:"0.73rem",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.04em"}}>Progreso del mes</span>
            {budget===0&&<p style={{margin:"0.25rem 0 0",color:t.warn,fontSize:"0.79rem"}}>Sin gastos fijos. <button onClick={()=>setView("plantilla")} style={{background:"none",border:"none",cursor:"pointer",color:t.pri,fontSize:"0.79rem",padding:0}}>Configurar →</button></p>}
          </div>
          <div style={{textAlign:"right"}}>
            <div style={{color:t.text,fontWeight:800,fontSize:"1.1rem"}}>{fCLP(paid)}</div>
            <div style={{color:t.dim,fontSize:"0.73rem"}}>de {fCLP(budget)} presupuestado</div>
          </div>
        </div>
        {budget>0&&<ProgressBar value={paid} max={budget} t={t}/>}
        {inc>0&&<div style={{display:"flex",justifyContent:"space-between",marginTop:"0.65rem",padding:"0.55rem 0.75rem",background:t.card2,borderRadius:"0.5rem"}}>
          <span style={{color:t.muted,fontSize:"0.8rem"}}>Ingresos mensuales: <strong style={{color:t.text}}>{fCLP(inc)}</strong></span>
          <span style={{color:paid>inc?t.err:t.ok,fontWeight:600,fontSize:"0.8rem"}}>{paid>inc?`⚠️ Sobre en ${fCLP(paid-inc)}`:`Queda: ${fCLP(inc-paid)}`}</span>
        </div>}
      </>} t={t} sx={{marginBottom:"1rem"}}/>

      {/* BOTÓN PRINCIPAL */}
      <button onClick={()=>setPayModal({date:todayStr(),isExtra:true})} style={{width:"100%",padding:"1rem",background:`linear-gradient(135deg,${t.pri},${t.priL})`,color:"#fff",border:"none",borderRadius:"0.75rem",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:"0.6rem",fontSize:"1.05rem",fontWeight:700,boxShadow:`0 4px 16px ${t.pri}44`,marginBottom:"1rem"}}>
        <Plus size={22}/> Registrar pago ahora
      </button>

      {/* Repetir pago frecuente — un tap, sin escribir nada */}
      {knownExtrasSorted(addr).length>0&&<div style={{marginBottom:"1rem"}}>
        <div style={{color:t.muted,fontSize:"0.73rem",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.04em",marginBottom:"0.5rem"}}>⚡ Repetir pago frecuente</div>
        <div style={{display:"flex",gap:"0.5rem",overflowX:"auto",paddingBottom:"0.2rem"}}>
          {knownExtrasSorted(addr).slice(0,6).map(k=><button key={k.id} onClick={()=>setPayModal({date:todayStr(),isExtra:true,itemName:k.name,estimated:k.lastAmount})} style={{display:"flex",flexDirection:"column",alignItems:"flex-start",gap:"0.15rem",padding:"0.5rem 0.75rem",background:t.card,border:`1px solid ${t.border}`,borderRadius:"0.6rem",cursor:"pointer",flexShrink:0,minWidth:"108px",boxShadow:t.sm}}>
            <span style={{color:t.text,fontSize:"0.79rem",fontWeight:700,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",maxWidth:"140px"}}>{k.name}</span>
            <span style={{color:t.dim,fontSize:"0.66rem"}}>{k.catName||"sin categoría"}</span>
            <span style={{color:t.pri,fontSize:"0.76rem",fontWeight:600}}>{fCLP(k.lastAmount)}</span>
          </button>)}
        </div>
      </div>}

      {/* Sugerencias inteligentes */}
      {suggestions.length>0&&<CCard ch={<>
        <div style={{display:"flex",alignItems:"center",gap:"0.5rem",marginBottom:"0.65rem"}}>
          <Sparkles size={15} color={t.warn}/>
          <span style={{color:t.text,fontWeight:700,fontSize:"0.88rem"}}>Sugerencias inteligentes</span>
        </div>
        {suggestions.map((sg,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:"0.75rem",padding:"0.5rem 0",borderBottom:i<suggestions.length-1?`1px solid ${t.border}`:"none"}}>
          <div style={{flex:1}}>
            <div style={{color:t.text,fontSize:"0.83rem",fontWeight:500}}>"{sg.name}" registrado {sg.count} veces como extra</div>
            <div style={{color:t.dim,fontSize:"0.73rem"}}>Últ. monto: {fCLP(sg.lastAmt)} — ¿Lo agregas a tus gastos fijos?</div>
          </div>
          <Btn onClick={()=>setView("plantilla")} v="ghost" sz="sm">Agregar →</Btn>
        </div>)}
      </>} t={t} sx={{marginBottom:"1rem"}}/>}

      {/* Pendientes */}
      {pending.length>0&&<CCard ch={<>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"0.75rem"}}>
          <h3 style={{margin:0,color:t.text,fontSize:"0.93rem",fontWeight:700}}>⏳ Pendiente de pagar ({pending.length})</h3>
          <Badge ch={fCLP(pending.reduce((s,i)=>s+(i.amount||0),0))} color={t.err}/>
        </div>
        {pending.map(it=><div key={it.id} style={{display:"flex",alignItems:"center",gap:"0.6rem",padding:"0.55rem 0.7rem",background:t.card2,borderRadius:"0.5rem",marginBottom:"0.4rem",border:`1px solid ${t.border}`}}>
          <span style={{fontSize:"0.95rem",lineHeight:1}}>{it.catIcon||"📦"}</span>
          <div style={{flex:1,minWidth:0}}>
            <div style={{color:t.text,fontSize:"0.84rem",fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{it.name}</div>
            <div style={{color:t.dim,fontSize:"0.7rem"}}>{it.catName}{it.isVariable?" · variable":""}</div>
          </div>
          <div style={{textAlign:"right",flexShrink:0}}>
            {it.accumulated>0&&<div style={{color:t.info,fontSize:"0.7rem",marginBottom:"0.15rem"}}>⟳ {fCLP(it.accumulated)}/{fCLP(it.monthlyAmt)}</div>}
            <div style={{color:t.muted,fontSize:"0.83rem",fontWeight:500,whiteSpace:"nowrap"}}>{it.isVariable?"~":""}{fCLP(it.monthlyAmt)}{it.frequency&&it.frequency!=="mensual"?` (${FREQS.find(f=>f.v===it.frequency)?.l})`:""}</div>
          </div>
          <Btn onClick={()=>setPayModal({date:todayStr(),isExtra:false,itemId:it.id,itemName:it.name,catId:it.catId,catName:it.catName,estimated:it.monthlyAmt,remaining:it.monthlyAmt-it.accumulated,accumulated:it.accumulated})} v={it.accumulated>0?"out":"pri"} sz="sm">{it.accumulated>0?"+ Abonar":"Pagar"}</Btn>
        </div>)}
      </>} t={t} sx={{marginBottom:"1rem"}}/>}

      {/* Pagado */}
      {(paidItems.length>0||extras.length>0)&&<CCard ch={<>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"0.75rem"}}>
          <h3 style={{margin:0,color:t.text,fontSize:"0.93rem",fontWeight:700}}>✅ Pagado este mes</h3>
          <Badge ch={fCLP(paid)} color={t.ok}/>
        </div>
        {paidItems.map(it=><div key={it.id} style={{padding:"0.45rem 0.7rem",background:t.ok+"10",borderRadius:"0.5rem",marginBottom:"0.3rem",border:`1px solid ${t.ok}30`}}>
          <div style={{display:"flex",alignItems:"center",gap:"0.6rem"}}>
            {it.isPaid?<CheckCircle2 size={15} color={t.ok}/>:<Clock size={15} color={t.warn}/>}
            <span style={{fontSize:"0.9rem",lineHeight:1}}>{it.catIcon||"📦"}</span>
            <div style={{flex:1}}>
              <div style={{color:t.text,fontSize:"0.83rem",fontWeight:500}}>{it.name}{it.pays.length>1?` · ${it.pays.length} abonos`:""}</div>
              <div style={{color:t.dim,fontSize:"0.7rem"}}>{fmtDate(it.pays[it.pays.length-1]?.date)}</div>
            </div>
            <div style={{textAlign:"right"}}>
              <div style={{color:it.isPaid?t.ok:t.warn,fontSize:"0.83rem",fontWeight:700}}>{fCLP(it.accumulated)}{!it.isPaid?` / ${fCLP(it.monthlyAmt)}`:""}</div>
              {!it.isPaid&&<div style={{color:t.dim,fontSize:"0.68rem"}}>parcial</div>}
            </div>
          </div>
          {it.pays.length>1&&<div style={{marginLeft:"1.9rem",marginTop:"0.35rem",display:"flex",flexDirection:"column",gap:"0.2rem"}}>
            {it.pays.map(p=><div key={p.id} style={{display:"flex",alignItems:"center",gap:"0.4rem",fontSize:"0.72rem",color:t.dim}}>
              <span style={{flex:1}}>{fmtDate(p.date)}{p.method?` · ${p.method}`:""}</span>
              <span style={{color:t.muted,fontWeight:600}}>{fCLP(p.amount)}</span>
              <button onClick={()=>setEditPay(p)} style={{background:"none",border:"none",cursor:"pointer",color:t.dim,padding:"0.1rem",display:"flex"}}><Edit2 size={10}/></button>
              <button onClick={()=>setCfm({msg:`¿Eliminar este abono de "${it.name}"?`,ok:()=>{d({t:"DEL_PAY",aid:addr.id,pid:p.id});setCfm(null);}})} style={{background:"none",border:"none",cursor:"pointer",color:t.dim,padding:"0.1rem",display:"flex"}}><Trash2 size={10}/></button>
            </div>)}
          </div>}
          {it.pays.length===1&&<div style={{textAlign:"right",marginTop:"0.15rem",display:"flex",justifyContent:"flex-end",gap:"0.6rem"}}>
            <button onClick={()=>setEditPay(it.pays[0])} style={{background:"none",border:"none",cursor:"pointer",color:t.dim,padding:"0.15rem",display:"inline-flex",alignItems:"center",gap:"0.2rem",fontSize:"0.7rem"}}><Edit2 size={11}/> Editar</button>
            <button onClick={()=>setCfm({msg:`¿Desmarcar "${it.name}" como pagado?`,ok:()=>{d({t:"DEL_PAY",aid:addr.id,pid:it.pays[0].id});setCfm(null);}})} style={{background:"none",border:"none",cursor:"pointer",color:t.dim,padding:"0.15rem",display:"inline-flex",alignItems:"center",gap:"0.2rem",fontSize:"0.7rem"}}><RotateCcw size={11}/> Desmarcar</button>
          </div>}
        </div>)}
        {extras.map(p=><div key={p.id} style={{display:"flex",alignItems:"center",gap:"0.6rem",padding:"0.45rem 0.7rem",background:t.warn+"10",borderRadius:"0.5rem",marginBottom:"0.3rem",border:`1px solid ${t.warn}30`}}>
          <Zap size={13} color={t.warn}/>
          <div style={{flex:1}}>
            <div style={{color:t.text,fontSize:"0.83rem",fontWeight:500}}>{p.name}</div>
            <div style={{color:t.dim,fontSize:"0.7rem"}}>{p.catName?`${p.catName} · `:""}{fmtDate(p.date)}</div>
          </div>
          <span style={{color:t.warn,fontSize:"0.83rem",fontWeight:700}}>{fCLP(p.amount)}</span>
          <button onClick={()=>setEditPay(p)} style={{background:"none",border:"none",cursor:"pointer",color:t.muted,padding:"0.2rem",display:"flex"}}><Edit2 size={12}/></button>
          <button onClick={()=>setCfm({msg:`¿Eliminar "${p.name}"?`,ok:()=>{d({t:"DEL_PAY",aid:addr.id,pid:p.id});setCfm(null);}})} style={{background:"none",border:"none",cursor:"pointer",color:t.err,padding:"0.2rem",display:"flex"}}><Trash2 size={12}/></button>
        </div>)}
      </>} t={t} sx={{marginBottom:"1rem"}}/>}

      {/* Accesos rápidos */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:"0.5rem"}}>
        {[{id:"plantilla",e:"📋",lb:"Gastos\nFijos"},{id:"historial",e:"📅",lb:"Historial"},{id:"analisis",e:"📊",lb:"Análisis"},{id:"perfil",e:"👤",lb:"Perfil"},{id:"ia",e:"🧠",lb:"Plan IA"}].map(q=><CCard key={q.id} ch={<><div style={{fontSize:"1.3rem",marginBottom:"0.3rem"}}>{q.e}</div><div style={{color:t.text,fontWeight:700,fontSize:"0.72rem",whiteSpace:"pre-line",lineHeight:1.3}}>{q.lb}</div></>} t={t} onClick={()=>setView(q.id)} sx={{cursor:"pointer",padding:"0.75rem",textAlign:"center"}}/>)}
      </div>
    </>}

    {payModal&&<PayModal state={payModal} t={t} addr={addr} d={d} profile={s.profile} onClose={()=>setPayModal(null)} onSave={dt=>{d({t:"ADD_PAY",aid:addr.id,d:dt});setPayModal(null);}}/>}
    {editPay&&<EditPayModal pay={editPay} t={t} profile={s.profile} onClose={()=>setEditPay(null)} onSave={dt=>{d({t:"UPD_PAY",aid:addr.id,pid:editPay.id,d:dt});setEditPay(null);}}/>}
    {cfm&&<Cfm msg={cfm.msg} onOk={cfm.ok} onCancel={()=>setCfm(null)} t={t}/>}
  </div>;
}

/* ══ PLANTILLA DE GASTOS FIJOS ═══════════════════════════ */
/* Gestor de categorías: ordenadas por frecuencia real de uso (las usadas una
   vez se hunden solas) y con opción de fusionar duplicados (ej. "Super" y "super") */
function CatManagerModal({addr,d,t,onClose}){
  const[merging,setMerging]=useState(null); // categoría origen que se quiere fusionar
  const[cfm,setCfm]=useState(null);
  const cats=addr?.template?.categories||[];
  const ranked=[...cats].map(c=>({...c,uses:catUsageCount(addr,c.id)+(addr?.knownExtras||[]).filter(k=>k.catId===c.id).length})).sort((a,b)=>b.uses-a.uses);

  if(cfm)return <Cfm msg={cfm.msg} onOk={cfm.ok} onCancel={()=>setCfm(null)} t={t}/>;

  return <div>
    <p style={{margin:"0 0 1rem",color:t.muted,fontSize:"0.82rem",lineHeight:1.55}}>Ordenadas por qué tanto las usas. Si tienes dos que en realidad son lo mismo (ej. "Super" y "super"), fusiónalas — todos sus gastos, ítems y montos históricos se combinan en una sola.</p>
    <div style={{display:"flex",flexDirection:"column",gap:"0.5rem"}}>
      {ranked.map(c=><div key={c.id} style={{border:`1px solid ${t.border}`,borderRadius:"0.6rem",padding:"0.65rem 0.8rem",background:t.card2}}>
        <div style={{display:"flex",alignItems:"center",gap:"0.6rem"}}>
          <span style={{fontSize:"1.1rem"}}>{c.icon||"📦"}</span>
          <div style={{flex:1}}>
            <div style={{color:t.text,fontWeight:700,fontSize:"0.87rem"}}>{c.name}</div>
            <div style={{color:t.dim,fontSize:"0.71rem"}}>{c.uses>0?`Usada ${c.uses}×`:"Sin uso todavía"} · {(c.items||[]).length} gasto(s) fijo(s)</div>
          </div>
          <Btn onClick={()=>setMerging(merging===c.id?null:c.id)} v="ghost" sz="sm">{merging===c.id?"Cancelar":"Fusionar en..."}</Btn>
        </div>
        {merging===c.id&&<div style={{marginTop:"0.6rem",paddingTop:"0.6rem",borderTop:`1px solid ${t.border}`}}>
          <div style={{color:t.muted,fontSize:"0.76rem",marginBottom:"0.4rem"}}>Elige la categoría destino — "{c.name}" desaparecerá y todo pasa a ella:</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:"0.4rem"}}>
            {ranked.filter(x=>x.id!==c.id).map(x=><button key={x.id} onClick={()=>setCfm({msg:`¿Fusionar "${c.name}" dentro de "${x.name}"? Esta acción no se puede deshacer.`,ok:()=>{d({t:"MERGE_CAT",aid:addr.id,srcId:c.id,tgtId:x.id});setMerging(null);setCfm(null);}})} style={{display:"flex",alignItems:"center",gap:"0.3rem",padding:"0.3rem 0.65rem",background:t.card,border:`1px solid ${t.border}`,borderRadius:"9999px",cursor:"pointer",color:t.text,fontSize:"0.78rem"}}>
              <span>{x.icon||"📦"}</span>{x.name}
            </button>)}
          </div>
        </div>}
      </div>)}
    </div>
    <Btn onClick={onClose} v="ghost" full sx={{marginTop:"1.1rem"}}>Cerrar</Btn>
  </div>;
}
function PlantillaView({s,d,t}){
  const[aM,setAM]=useState(null);const[cM,setCM]=useState(null);const[iM,setIM]=useState(null);const[cfm,setCfm]=useState(null);const[gM,setGM]=useState(false);
  const[expCat,setExpCat]=useState({});
  const addr=s.addresses.find(a=>a.id===s.selAddr);
  const cats=addr?.template?.categories||[];
  const total=addr?tmplTotal(addr):0;

  return <div>
    <SH title="Gastos Fijos Mensuales" sub="Configura una vez — aparecen automáticamente cada mes para siempre" t={t}
      action={<Btn onClick={()=>setAM("add")} icon={<Plus size={14}/>} sz="sm">+ Dirección</Btn>}/>
    <CCard ch={<div style={{display:"flex",gap:"0.75rem",alignItems:"center",flexWrap:"wrap"}}>
      <div style={{flex:1}}>
        <div style={{color:t.muted,fontSize:"0.73rem",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.04em",marginBottom:"0.3rem"}}>¿Qué son los gastos fijos?</div>
        <p style={{margin:0,color:t.text,fontSize:"0.84rem",lineHeight:1.5}}>Son los gastos que tienes <strong>todos los meses sin excepción</strong>: arriendo, luz, agua, internet, Netflix, etc. Los agregas aquí y la app los recuerda por ti, mes a mes, de forma automática.</p>
      </div>
      {total>0&&<div style={{textAlign:"center",padding:"0.75rem 1.25rem",background:t.pri+"15",borderRadius:"0.65rem",border:`1px solid ${t.pri}30`,flexShrink:0}}>
        <div style={{color:t.pri,fontWeight:800,fontSize:"1.3rem"}}>{fCLP(total)}</div>
        <div style={{color:t.dim,fontSize:"0.72rem"}}>total mensual estimado</div>
      </div>}
    </div>} t={t} sx={{marginBottom:"1rem"}}/>

    {/* Selector de dirección */}
    {s.addresses.length>0&&<CCard ch={<>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"0.65rem"}}>
        <span style={{color:t.muted,fontSize:"0.71rem",fontWeight:700,textTransform:"uppercase",letterSpacing:"0.04em"}}>Dirección activa</span>
        {addr&&<div style={{display:"flex",gap:"0.4rem"}}>
          <Btn onClick={()=>setAM({e:addr})} v="ghost" sz="sm" icon={<Edit2 size={12}/>}>Editar</Btn>
          <Btn onClick={()=>setCfm({msg:`¿Eliminar "${addr.name}" con TODOS sus datos? Esta acción es irreversible.`,ok:()=>{d({t:"DA",id:addr.id});setCfm(null);}})} v="ghost" sz="sm" sx={{color:t.err}} icon={<Trash2 size={12}/>}>Borrar</Btn>
        </div>}
      </div>
      <div style={{display:"flex",gap:"0.4rem",flexWrap:"wrap"}}>
        {s.addresses.map(a=><button key={a.id} onClick={()=>d({t:"SA",id:a.id})} style={{padding:"0.35rem 0.8rem",background:s.selAddr===a.id?t.pri:t.card2,color:s.selAddr===a.id?"#fff":t.text,border:`1px solid ${s.selAddr===a.id?t.pri:t.border}`,borderRadius:"0.45rem",cursor:"pointer",fontSize:"0.83rem",fontWeight:s.selAddr===a.id?700:400}}>
          📍 {a.name}
        </button>)}
      </div>
      {addr?.address&&<p style={{margin:"0.5rem 0 0",color:t.dim,fontSize:"0.77rem"}}>📌 {addr.address}</p>}
    </>} t={t} sx={{marginBottom:"1rem"}}/>}

    {!s.addresses.length&&<CCard ch={<Empty icon={<Building2 size={36}/>} title="Empieza creando tu dirección" sub="Agrega tu casa o departamento para organizar tus gastos fijos." action={<Btn onClick={()=>setAM("add")} icon={<Plus size={14}/>}>Agregar mi dirección</Btn>} t={t}/>} t={t} sx={{marginBottom:"1rem"}}/>}

    {/* Categorías */}
    {addr&&<>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"0.7rem"}}>
        <h3 style={{margin:0,color:t.text,fontSize:"0.97rem",fontWeight:700}}>Mis categorías recurrentes</h3>
        <div style={{display:"flex",gap:"0.4rem"}}>
          {cats.length>1&&<Btn onClick={()=>setGM(true)} v="ghost" icon={<List size={14}/>} sz="sm">Gestionar</Btn>}
          <Btn onClick={()=>setCM("add")} icon={<Plus size={14}/>} sz="sm">Categoría</Btn>
        </div>
      </div>
      {cats.length===0&&<CCard ch={<Empty icon={<List size={34}/>} title="Sin categorías aún" sub='Agrega categorías de gasto fijo como "Vivienda" o "Alimentación". Usa las plantillas para ir rápido.' action={<Btn onClick={()=>setCM("add")} icon={<Plus size={14}/>}>Agregar primera categoría</Btn>} t={t}/>} t={t}/>}
      {cats.map((cat,ci)=>{
        const catT=(cat.items||[]).reduce((s,it)=>s+(it.amount||0),0);const exp=expCat[cat.id];
        return <div key={cat.id} style={{marginBottom:"0.65rem",borderRadius:"0.75rem",overflow:"hidden",border:`1px solid ${t.border}`,boxShadow:t.sm}}>
          <div onClick={()=>setExpCat(p=>({...p,[cat.id]:!p[cat.id]}))} style={{display:"flex",alignItems:"center",gap:"0.7rem",padding:"0.85rem 1rem",background:t.card,cursor:"pointer",userSelect:"none"}}>
            <div style={{width:"8px",height:"8px",borderRadius:"50%",background:cat.color||COLORS[ci%COLORS.length],flexShrink:0}}/>
            <span style={{fontSize:"1.05rem",lineHeight:1}}>{cat.icon||"📦"}</span>
            <span style={{flex:1,color:t.text,fontWeight:700,fontSize:"0.92rem"}}>{cat.name}</span>
            <span style={{color:t.muted,fontSize:"0.88rem",fontWeight:600}}>{fCLP(catT)}/mes</span>
            <div style={{display:"flex",gap:"0.15rem"}} onClick={e=>e.stopPropagation()}>
              <button onClick={()=>setCM({e:cat})} style={{background:"none",border:"none",cursor:"pointer",color:t.muted,padding:"0.25rem",display:"flex"}}><Edit2 size={13}/></button>
              <button onClick={()=>setCfm({msg:`¿Eliminar categoría "${cat.name}" y todos sus ítems?`,ok:()=>{d({t:"DTC",aid:addr.id,cid:cat.id});setCfm(null);}})} style={{background:"none",border:"none",cursor:"pointer",color:t.err,padding:"0.25rem",display:"flex"}}><Trash2 size={13}/></button>
            </div>
            <span style={{color:t.dim}}>{exp?<ChevronUp size={15}/>:<ChevronDown size={15}/>}</span>
          </div>
          {exp&&<div style={{borderTop:`1px solid ${t.border}`,background:t.card2,padding:"0.65rem 0.85rem"}}>
            {(cat.items||[]).map(it=><div key={it.id} style={{display:"flex",alignItems:"center",gap:"0.5rem",padding:"0.45rem 0.65rem",background:t.card,borderRadius:"0.5rem",marginBottom:"0.35rem"}}>
              <div style={{flex:1}}>
                <div style={{color:t.text,fontSize:"0.84rem",fontWeight:500}}>{it.name}</div>
                {it.isVariable&&<div style={{color:t.info,fontSize:"0.7rem"}}>📊 Monto variable (estimado)</div>}
              {it.frequency&&it.frequency!=="mensual"&&<div style={{color:t.muted,fontSize:"0.69rem"}}>🔄 {FREQS.find(f=>f.v===it.frequency)?.l} → {new Intl.NumberFormat("es-CL",{style:"currency",currency:"CLP",maximumFractionDigits:0}).format(Math.round((it.amount||0)*(MF[it.frequency]||1)))}/mes</div>}
              </div>
              <div style={{textAlign:"right"}}><div style={{color:it.isVariable?t.info:t.muted,fontSize:"0.84rem",fontWeight:600}}>{it.isVariable?"~":""}{fCLP(it.amount)}</div>{it.frequency&&it.frequency!=="mensual"&&<div style={{color:t.dim,fontSize:"0.7rem"}}>{FREQS.find(f=>f.v===it.frequency)?.l}</div>}</div>
              <button onClick={()=>setIM({e:it,cid:cat.id})} style={{background:"none",border:"none",cursor:"pointer",color:t.muted,padding:"0.2rem",display:"flex"}}><Edit2 size={12}/></button>
              <button onClick={()=>setCfm({msg:`¿Quitar "${it.name}" de los gastos fijos?`,ok:()=>{d({t:"DTI",aid:addr.id,cid:cat.id,iid:it.id});setCfm(null);}})} style={{background:"none",border:"none",cursor:"pointer",color:t.err,padding:"0.2rem",display:"flex"}}><Trash2 size={12}/></button>
            </div>)}
            <Btn onClick={()=>setIM({cid:cat.id})} v="ghost" sz="sm" icon={<Plus size={12}/>} full sx={{marginTop:"0.25rem",borderColor:t.border}}>Agregar gasto a {cat.name}</Btn>
          </div>}
        </div>;
      })}
      {cats.length>0&&<div style={{padding:"0.65rem 0.85rem",background:t.card2,borderRadius:"0.5rem",border:`1px solid ${t.border}`,display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:"0.5rem"}}>
        <span style={{color:t.muted,fontSize:"0.84rem"}}>Total estimado mensual (para siempre)</span>
        <span style={{color:t.pri,fontWeight:800,fontSize:"1.05rem"}}>{fCLP(total)}/mes</span>
      </div>}
    </>}

    {aM&&<Modal title={aM==="add"?"Nueva Dirección":"Editar Dirección"} onClose={()=>setAM(null)} t={t} ch={<AddrF init={aM==="add"?{}:{name:aM.e.name,addr:aM.e.address}} onSave={dt=>{aM==="add"?d({t:"AA",...dt}):d({t:"UA",id:aM.e.id,...dt});setAM(null);}} t={t}/>}/>}
    {cM&&<Modal title={cM==="add"?"Nueva Categoría":"Editar Categoría"} onClose={()=>setCM(null)} t={t} ch={<CatF init={cM==="add"?{}:{name:cM.e.name,icon:cM.e.icon,color:cM.e.color}} onSave={dt=>{cM==="add"?d({t:"ATC",aid:addr.id,...dt}):d({t:"UTC",aid:addr.id,cid:cM.e.id,...dt});setCM(null);}} existing={cats.map(c=>c.name)} t={t}/>}/>}
    {iM&&<Modal title={iM.e?"Editar Gasto Fijo":"Nuevo Gasto Fijo"} onClose={()=>setIM(null)} t={t} ch={<TmplItemF init={iM.e?{name:iM.e.name,amount:iM.e.amount,isVariable:iM.e.isVariable}:{}} onSave={dt=>{iM.e?d({t:"UTI",aid:addr.id,cid:iM.cid,iid:iM.e.id,...dt}):d({t:"ATI",aid:addr.id,cid:iM.cid,...dt});setIM(null);}} t={t}/>}/>}
    {gM&&<Modal title="🗂️ Gestionar Categorías" onClose={()=>setGM(false)} t={t} w="540px" ch={<CatManagerModal addr={addr} d={d} t={t} onClose={()=>setGM(false)}/>}/>}
    {cfm&&<Cfm msg={cfm.msg} onOk={cfm.ok} onCancel={()=>setCfm(null)} t={t}/>}
  </div>;
}

function AddrF({init,onSave,t}){const[n,sN]=useState(init.name||""),[a,sA]=useState(init.addr||"");return <><Fld label="Nombre del lugar" ch={<TI val={n} onChange={sN} ph="Mi Casa en Ñuñoa" t={t}/>} t={t}/><Fld label="Dirección completa" ch={<TI val={a} onChange={sA} ph="Av. Irarrázaval 1234" t={t}/>} t={t}/><Btn onClick={()=>n&&onSave({name:n,addr:a})} dis={!n} full icon={<Check size={14}/>}>Guardar</Btn></>;}

function CatF({init,onSave,existing=[],t}){
  const[name,setName]=useState(init.name||"");const[icon,setIcon]=useState(init.icon||"📦");const[color,setColor]=useState(init.color||COLORS[0]);const[presets,setPresets]=useState(!init.name);
  return <>
    {presets&&<div style={{marginBottom:"1rem"}}>
      <p style={{color:t.muted,fontSize:"0.71rem",fontWeight:700,textTransform:"uppercase",marginBottom:"0.5rem"}}>Plantillas rápidas</p>
      <div style={{display:"flex",flexWrap:"wrap",gap:"0.35rem",marginBottom:"0.65rem"}}>
        {CAT_PRESETS.filter(([n])=>!existing.includes(n)).map(([n,ic],i)=><button key={n} onClick={()=>{setName(n);setIcon(ic);setColor(COLORS[i%COLORS.length]);setPresets(false);}} style={{padding:"0.3rem 0.65rem",background:t.card2,border:`1px solid ${t.border}`,borderRadius:"0.4rem",cursor:"pointer",color:t.text,fontSize:"0.79rem"}}>{ic} {n}</button>)}
      </div>
      <button onClick={()=>setPresets(false)} style={{background:"none",border:"none",cursor:"pointer",color:t.pri,fontSize:"0.79rem"}}>+ Crear personalizada</button>
    </div>}
    <Fld label="Nombre" ch={<TI val={name} onChange={setName} ph="Vivienda, Alimentación..." t={t}/>} t={t}/>
    <Fld label="Ícono" ch={<div style={{display:"flex",gap:"0.5rem",alignItems:"center"}}><span style={{fontSize:"1.4rem"}}>{icon}</span><TI val={icon} onChange={setIcon} ph="🏠" t={t} sx={{maxWidth:"80px"}}/></div>} t={t}/>
    <Fld label="Color" ch={<div style={{display:"flex",gap:"0.35rem",flexWrap:"wrap"}}>{COLORS.slice(0,12).map(c=><button key={c} onClick={()=>setColor(c)} style={{width:"26px",height:"26px",borderRadius:"50%",background:c,border:color===c?"3px solid #fff":"2px solid transparent",outline:color===c?`2px solid ${c}`:"none",cursor:"pointer"}}/>)}</div>} t={t}/>
    <Btn onClick={()=>name&&onSave({name,icon,color})} dis={!name} full icon={<Check size={14}/>} sx={{marginTop:"0.5rem"}}>{init.name?"Actualizar":"Crear Categoría"}</Btn>
  </>;
}

function TmplItemF({init,onSave,t}){
  const[name,setName]=useState(init.name||"");const[amount,setAmount]=useState(init.amount||0);const[isVariable,setIsVariable]=useState(init.isVariable||false);const[freq,setFreq]=useState(init.frequency||"mensual");
  const mEq=Math.round((amount||0)*(MF[freq]||1));
  return <>
    <Fld label="Nombre del gasto" ch={<TI val={name} onChange={setName} ph="Arriendo, Luz, Internet, Agua, Netflix..." t={t}/>} t={t}/>
    <Fld label="Frecuencia de pago" ch={<Sel val={freq} onChange={setFreq} opts={FREQS.map(f=>({v:f.v,l:f.l}))} t={t}/>} t={t}/>
    <Fld label={isVariable?"Monto estimado por pago":"Monto por pago"} ch={<MoneyInput val={amount} onChange={setAmount} t={t} af/>} t={t}/>
    {freq!=="mensual"&&amount>0&&<div style={{padding:"0.5rem 0.75rem",background:t.info+"12",borderRadius:"0.5rem",marginBottom:"0.85rem",fontSize:"0.8rem",color:t.muted}}>
      <span>Equivalente mensual: <strong style={{color:t.info}}>{new Intl.NumberFormat("es-CL",{style:"currency",currency:"CLP",maximumFractionDigits:0}).format(mEq)}</strong></span>
      <span style={{marginLeft:"0.75rem"}}>· Anual: <strong style={{color:t.info}}>{new Intl.NumberFormat("es-CL",{style:"currency",currency:"CLP",maximumFractionDigits:0}).format(mEq*12)}</strong></span>
    </div>}
    <label style={{display:"flex",alignItems:"flex-start",gap:"0.65rem",cursor:"pointer",padding:"0.65rem 0.75rem",background:t.card2,borderRadius:"0.5rem",marginBottom:"0.9rem"}}>
      <input type="checkbox" checked={isVariable} onChange={e=>setIsVariable(e.target.checked)} style={{width:"16px",height:"16px",accentColor:t.info,marginTop:"0.15rem"}}/>
      <div>
        <div style={{color:t.text,fontSize:"0.84rem",fontWeight:600}}>Monto variable</div>
        <div style={{color:t.dim,fontSize:"0.73rem"}}>Marca esto si el monto cambia cada mes (luz, agua, teléfono). El monto que pones aquí es solo una estimación que puedes ajustar cuando pagas.</div>
      </div>
    </label>
    <Btn onClick={()=>name&&onSave({name,amount,isVariable,frequency:freq})} dis={!name} full icon={<Check size={14}/>}>{init.name?"Actualizar":"Agregar a plantilla"}</Btn>
  </>;
}

/* ══ HISTORIAL — CALENDARIO MENSUAL ═════════════════════ */
function HistorialView({s,d,t}){
  const[selDay,setSelDay]=useState(null);const[payModal,setPayModal]=useState(null);const[cfm,setCfm]=useState(null);const[editPay,setEditPay]=useState(null);
  const addr=s.addresses.find(a=>a.id===s.selAddr);
  const pays=addr?mPays(addr,s.selYear,s.selMonth):[];
  const pending=addr?pendingTmpl(addr,s.selYear,s.selMonth):[];
  const paidItems=addr?paidTmpl(addr,s.selYear,s.selMonth):[];
  const extras=addr?extraPays(addr,s.selYear,s.selMonth):[];
  const catNames=[...new Set((addr?.template?.categories||[]).map(c=>c.name))];
  const budget=addr?tmplTotal(addr):0;const paid=addr?totalPaid(addr,s.selYear,s.selMonth):0;
  const daysInMonth=new Date(s.selYear,s.selMonth+1,0).getDate();
  const offset=(new Date(s.selYear,s.selMonth,1).getDay()+6)%7;
  const byDay={};pays.forEach(p=>{if(!p.date)return;const day=parseInt(p.date.split("-")[2]);if(!byDay[day])byDay[day]=[];byDay[day].push(p);});
  const selDayPays=selDay?(byDay[selDay]||[]):[];
  const goM=dir=>{let nm=s.selMonth+dir,ny=s.selYear;if(nm<0){nm=11;ny--;}if(nm>11){nm=0;ny++;}d({t:"YM",month:nm,year:ny});setSelDay(null);};
  const mkDate=day=>`${s.selYear}-${String(s.selMonth+1).padStart(2,"0")}-${String(day).padStart(2,"0")}`;

  if(!addr)return <CCard ch={<Empty icon={<Building2 size={36}/>} title="Sin dirección" sub="Configura una dirección en Gastos Fijos para ver el historial." t={t}/>} t={t}/>;

  return <div>
    <SH title="Historial de Pagos" sub={`${addr.name} · ${MESES[s.selMonth]} ${s.selYear}`} t={t}
      action={<Btn onClick={()=>setPayModal({date:todayStr(),isExtra:true})} icon={<Zap size={14}/>} sz="sm">⚡ Pago extra</Btn>}/>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:"0.75rem",marginBottom:"1rem"}}>
      <StCard label="Presupuestado" val={fCLP(budget)} color={t.pri} icon={<Target size={16} color={t.pri}/>} t={t}/>
      <StCard label="Pagado" val={fCLP(paid)} color={t.ok} icon={<CheckCircle2 size={16} color={t.ok}/>} sub={`${paidItems.length+extras.length} pagos`} t={t}/>
      <StCard label="Pendiente" val={fCLP(pending.reduce((s,i)=>s+Math.max(0,i.monthlyAmt-i.accumulated),0))} color={t.err} icon={<Clock size={16} color={t.err}/>} sub={`${pending.length} ítems`} t={t}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"1rem",marginBottom:"1rem",alignItems:"start"}}>
      {/* Calendario */}
      <CCard ch={<>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"0.75rem"}}>
          <button onClick={()=>goM(-1)} style={{background:"none",border:"none",cursor:"pointer",color:t.muted,padding:"0.3rem",display:"flex"}}><ChevronLeft size={17}/></button>
          <span style={{color:t.text,fontWeight:800,fontSize:"0.92rem"}}>{MESES[s.selMonth]} {s.selYear}</span>
          <button onClick={()=>goM(1)} style={{background:"none",border:"none",cursor:"pointer",color:t.muted,padding:"0.3rem",display:"flex"}}><ChevronRight size={17}/></button>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:"2px",marginBottom:"4px"}}>
          {DIAS.map(dn=><div key={dn} style={{textAlign:"center",color:t.dim,fontSize:"0.63rem",fontWeight:700,padding:"2px 0"}}>{dn}</div>)}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:"2px"}}>
          {Array(offset).fill(null).map((_,i)=><div key={"e"+i}/>)}
          {Array(daysInMonth).fill(null).map((_,i)=>{
            const day=i+1,dp=byDay[day]||[];
            const isToday=NOW.getDate()===day&&NOW.getMonth()===s.selMonth&&NOW.getFullYear()===s.selYear;
            const isSel=selDay===day;
            return <div key={day} onClick={()=>setSelDay(isSel?null:day)} style={{aspectRatio:"1",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",borderRadius:"0.4rem",cursor:"pointer",background:isSel?t.pri:isToday?t.pri+"28":dp.length?t.ok+"18":"none",border:isSel?`1px solid ${t.pri}`:isToday?`1px solid ${t.pri}60`:"1px solid transparent",transition:"all .15s"}}>
              <span style={{fontSize:"0.76rem",fontWeight:isToday||isSel?700:400,color:isSel?"#fff":isToday?t.pri:t.text,lineHeight:1}}>{day}</span>
              {dp.length>0&&<div style={{display:"flex",gap:"1px",marginTop:"2px"}}>{dp.slice(0,3).map((p,pi)=><div key={pi} style={{width:"4px",height:"4px",borderRadius:"50%",background:isSel?"rgba(255,255,255,.7)":p.isExtra?t.warn:t.ok}}/>)}</div>}
            </div>;
          })}
        </div>
        <div style={{display:"flex",gap:"1rem",marginTop:"0.75rem",fontSize:"0.67rem",color:t.dim,justifyContent:"center"}}>
          <span style={{display:"flex",alignItems:"center",gap:"0.3rem"}}><div style={{width:"6px",height:"6px",borderRadius:"50%",background:t.ok}}/> Fijo pagado</span>
          <span style={{display:"flex",alignItems:"center",gap:"0.3rem"}}><div style={{width:"6px",height:"6px",borderRadius:"50%",background:t.warn}}/> Extra</span>
        </div>
        {selDay&&<div style={{marginTop:"0.75rem",borderTop:`1px solid ${t.border}`,paddingTop:"0.75rem"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"0.5rem"}}>
            <span style={{color:t.text,fontWeight:700,fontSize:"0.85rem"}}>{selDay} de {MESES[s.selMonth]}</span>
            <Btn onClick={()=>setPayModal({date:mkDate(selDay),isExtra:true})} v="ghost" sz="sm" icon={<Plus size={12}/>}>Agregar</Btn>
          </div>
          {selDayPays.length===0&&<p style={{color:t.dim,fontSize:"0.79rem",margin:0}}>Sin pagos este día.</p>}
          {selDayPays.map(p=><div key={p.id} style={{display:"flex",gap:"0.5rem",padding:"0.35rem 0.6rem",background:t.card2,borderRadius:"0.4rem",marginBottom:"0.25rem",alignItems:"center"}}>
            <span style={{fontSize:"0.85rem"}}>{p.isExtra?"⚡":"✓"}</span>
            <div style={{flex:1,minWidth:0}}><div style={{color:t.text,fontSize:"0.79rem",fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</div>{p.method&&<div style={{color:t.dim,fontSize:"0.67rem"}}>{p.method}</div>}</div>
            <span style={{color:p.isExtra?t.warn:t.ok,fontSize:"0.79rem",fontWeight:700,whiteSpace:"nowrap"}}>{fCLP(p.amount)}</span>
            <button onClick={()=>setEditPay(p)} style={{background:"none",border:"none",cursor:"pointer",color:t.muted,padding:"0.15rem",display:"flex"}}><Edit2 size={11}/></button>
            <button onClick={()=>setCfm({msg:`¿Eliminar "${p.name}"?`,ok:()=>{d({t:"DEL_PAY",aid:addr.id,pid:p.id});setCfm(null);}})} style={{background:"none",border:"none",cursor:"pointer",color:t.err,padding:"0.15rem",display:"flex"}}><Trash2 size={11}/></button>
          </div>)}
        </div>}
      </>} t={t}/>
      {/* Panel ítems */}
      <div style={{display:"flex",flexDirection:"column",gap:"0.65rem"}}>
        {pending.length>0&&<CCard ch={<>
          <h4 style={{margin:"0 0 0.65rem",color:t.text,fontSize:"0.88rem",fontWeight:700}}>⏳ Pendiente ({pending.length})</h4>
          <div style={{maxHeight:"200px",overflowY:"auto",display:"flex",flexDirection:"column",gap:"0.35rem"}}>
            {pending.map(it=><div key={it.id} style={{display:"flex",alignItems:"center",gap:"0.45rem",padding:"0.45rem 0.6rem",background:t.card2,borderRadius:"0.45rem",border:`1px solid ${t.border}`}}>
              <span style={{fontSize:"0.9rem",lineHeight:1}}>{it.catIcon||"📦"}</span>
              <div style={{flex:1,minWidth:0}}>
                <div style={{color:t.text,fontSize:"0.81rem",fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{it.name}</div>
                <div style={{color:t.dim,fontSize:"0.68rem"}}>{it.catName}{it.isVariable?" · variable":""}{it.accumulated>0?` · abonado ${fCLP(it.accumulated)}`:""}</div>
              </div>
              <span style={{color:t.dim,fontSize:"0.78rem",whiteSpace:"nowrap"}}>{it.isVariable?"~":""}{fCLP(it.monthlyAmt)}</span>
              <Btn onClick={()=>setPayModal({date:todayStr(),isExtra:false,itemId:it.id,itemName:it.name,catId:it.catId,catName:it.catName,estimated:it.monthlyAmt,remaining:it.monthlyAmt-it.accumulated,accumulated:it.accumulated})} v={it.accumulated>0?"out":"pri"} sz="sm">{it.accumulated>0?"+ Abonar":"Pagar"}</Btn>
            </div>)}
          </div>
        </>} t={t}/>}
        {(paidItems.length>0||extras.length>0)&&<CCard ch={<>
          <h4 style={{margin:"0 0 0.65rem",color:t.text,fontSize:"0.88rem",fontWeight:700}}>✅ Pagado</h4>
          <div style={{maxHeight:"200px",overflowY:"auto",display:"flex",flexDirection:"column",gap:"0.3rem"}}>
            {paidItems.map(it=><div key={it.id} style={{padding:"0.4rem 0.6rem",background:t.ok+"10",borderRadius:"0.45rem",border:`1px solid ${t.ok}30`}}>
              <div style={{display:"flex",alignItems:"center",gap:"0.45rem"}}>
                {it.isPaid?<CheckCircle2 size={13} color={t.ok}/>:<Clock size={13} color={t.warn}/>}
                <div style={{flex:1,minWidth:0}}>
                  <div style={{color:t.text,fontSize:"0.8rem",fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{it.name}{it.pays.length>1?` (${it.pays.length})`:""}</div>
                  <div style={{color:t.dim,fontSize:"0.67rem"}}>{fmtDate(it.pays[it.pays.length-1]?.date)}</div>
                </div>
                <span style={{color:it.isPaid?t.ok:t.warn,fontSize:"0.79rem",fontWeight:700,whiteSpace:"nowrap"}}>{fCLP(it.accumulated)}{!it.isPaid?`/${fCLP(it.monthlyAmt)}`:""}</span>
                {it.pays.length===1&&<><button onClick={()=>setEditPay(it.pays[0])} style={{background:"none",border:"none",cursor:"pointer",color:t.dim,padding:"0.15rem",display:"flex"}}><Edit2 size={11}/></button><button onClick={()=>setCfm({msg:`¿Desmarcar "${it.name}"?`,ok:()=>{d({t:"DEL_PAY",aid:addr.id,pid:it.pays[0].id});setCfm(null);}})} style={{background:"none",border:"none",cursor:"pointer",color:t.dim,padding:"0.15rem",display:"flex"}}><RotateCcw size={11}/></button></>}
              </div>
              {it.pays.length>1&&<div style={{marginLeft:"1.6rem",marginTop:"0.25rem",display:"flex",flexDirection:"column",gap:"0.15rem"}}>
                {it.pays.map(p=><div key={p.id} style={{display:"flex",gap:"0.35rem",fontSize:"0.68rem",color:t.dim,alignItems:"center"}}>
                  <span style={{flex:1}}>{fmtDate(p.date)}</span><span>{fCLP(p.amount)}</span>
                  <button onClick={()=>setEditPay(p)} style={{background:"none",border:"none",cursor:"pointer",color:t.dim,padding:"0.1rem",display:"flex"}}><Edit2 size={9}/></button>
                  <button onClick={()=>setCfm({msg:`¿Eliminar este abono?`,ok:()=>{d({t:"DEL_PAY",aid:addr.id,pid:p.id});setCfm(null);}})} style={{background:"none",border:"none",cursor:"pointer",color:t.dim,padding:"0.1rem",display:"flex"}}><Trash2 size={9}/></button>
                </div>)}
              </div>}
            </div>)}
            {extras.map(p=><div key={p.id} style={{display:"flex",alignItems:"center",gap:"0.45rem",padding:"0.4rem 0.6rem",background:t.warn+"10",borderRadius:"0.45rem",border:`1px solid ${t.warn}30`}}>
              <Zap size={12} color={t.warn}/>
              <div style={{flex:1,minWidth:0}}><div style={{color:t.text,fontSize:"0.8rem",fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</div><div style={{color:t.dim,fontSize:"0.67rem"}}>{p.catName?`${p.catName} · `:""}{fmtDate(p.date)}</div></div>
              <span style={{color:t.warn,fontSize:"0.79rem",fontWeight:700,whiteSpace:"nowrap"}}>{fCLP(p.amount)}</span>
              <button onClick={()=>setEditPay(p)} style={{background:"none",border:"none",cursor:"pointer",color:t.muted,padding:"0.15rem",display:"flex"}}><Edit2 size={11}/></button>
              <button onClick={()=>setCfm({msg:`¿Eliminar "${p.name}"?`,ok:()=>{d({t:"DEL_PAY",aid:addr.id,pid:p.id});setCfm(null);}})} style={{background:"none",border:"none",cursor:"pointer",color:t.err,padding:"0.15rem",display:"flex"}}><Trash2 size={11}/></button>
            </div>)}
          </div>
        </>} t={t}/>}
        {pending.length===0&&paidItems.length===0&&extras.length===0&&<CCard ch={<Empty icon={<Calendar size={30}/>} title="Sin actividad este mes" sub='Marca gastos como pagados o agrega un pago extra.' t={t}/>} t={t}/>}
      </div>
    </div>
    {payModal&&<PayModal state={payModal} t={t} addr={addr} d={d} profile={s.profile} onClose={()=>setPayModal(null)} onSave={dt=>{d({t:"ADD_PAY",aid:addr.id,d:dt});setPayModal(null);}}/>}
    {editPay&&<EditPayModal pay={editPay} t={t} profile={s.profile} onClose={()=>setEditPay(null)} onSave={dt=>{d({t:"UPD_PAY",aid:addr.id,pid:editPay.id,d:dt});setEditPay(null);}}/>}
    {cfm&&<Cfm msg={cfm.msg} onOk={cfm.ok} onCancel={()=>setCfm(null)} t={t}/>}
  </div>;
}

/* ══ ANÁLISIS ════════════════════════════════════════════ */
function AnalisisView({s,d,t}){
  const addr=s.addresses.find(a=>a.id===s.selAddr);
  const budget=addr?tmplTotal(addr):0;
  const recentMonths=useMemo(()=>{if(!addr)return[];const months=[];for(let i=11;i>=0;i--){let m=CM-i,y=CY;if(m<0){m+=12;y--;}months.push({year:y,month:m});}return months;},[addr]);
  const[sel,setSel]=useState(()=>{
    const saved=s.settings.analisisSel;
    if(saved&&saved.length)return new Set(saved);
    return new Set(recentMonths.map(m=>`${m.year}-${m.month}`));
  });
  useEffect(()=>{d({t:"SS",k:"analisisSel",v:[...sel]});},[sel]);
  const[yearFilter,setYearFilter]=useState(null);
  const toggle=k=>setSel(p=>{const n=new Set(p);n.has(k)?n.delete(k):n.add(k);return n;});
  const selM=recentMonths.filter(m=>sel.has(`${m.year}-${m.month}`));
  const barData=selM.map(m=>{const p=addr?totalPaid(addr,m.year,m.month):0;const byCat={};(addr?mPays(addr,m.year,m.month):[]).forEach(pay=>{if(pay.catName)byCat[pay.catName]=(byCat[pay.catName]||0)+(pay.amount||0);});return{name:`${MESES[m.month].substr(0,3)} ${m.year}`,Presupuesto:budget,Pagado:p,...byCat};});
  const catAgg={};selM.forEach(m=>(addr?mPays(addr,m.year,m.month):[]).forEach(p=>{if(p.catName)catAgg[p.catName]=(catAgg[p.catName]||0)+(p.amount||0);}));
  const pieData=Object.entries(catAgg).map(([name,value])=>({name,value})).sort((a,b)=>b.value-a.value);
  const totalPaidAll=selM.reduce((s,m)=>s+(addr?totalPaid(addr,m.year,m.month):0),0);

  const exportExcel=()=>{
    if(!addr||!selM.length)return;
    const wb=XLSX.utils.book_new();
    const sumRows=selM.map(m=>({"Período":`${MESES[m.month]} ${m.year}`,"Presupuestado":budget,"Pagado":totalPaid(addr,m.year,m.month),"Diferencia":totalPaid(addr,m.year,m.month)-budget}));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(sumRows),"Resumen");
    // Hoja de gasto hormiga: agrupado por lugar/ítem con detalle de fechas
    const agg={};
    selM.forEach(m=>mPays(addr,m.year,m.month).filter(p=>p.isExtra).forEach(p=>{
      const k=(p.name||"").trim().toLowerCase();if(!k)return;
      agg[k]=agg[k]||{Lugar:p.name,Categoría:p.catName||"",Pagos:0,Total:0,detalle:[]};
      agg[k].Pagos++;agg[k].Total+=p.amount||0;agg[k].detalle.push(`${fmtDate(p.date)}: ${fCLP(p.amount)}`);
    }));
    const hormigaRows=Object.values(agg).filter(e=>e.Pagos>=2).sort((a,b)=>b.Total-a.Total).map(e=>({Lugar:e.Lugar,Categoría:e.Categoría,"N° Pagos":e.Pagos,"Total Gastado":e.Total,"Detalle por Fecha":e.detalle.join(" | ")}));
    if(hormigaRows.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(hormigaRows),"Gasto Hormiga");
    selM.forEach(m=>{const rows=(addr?mPays(addr,m.year,m.month):[]).map(p=>({Nombre:p.name,Monto:p.amount,Fecha:p.date,Categoría:p.catName||"",Tipo:p.isExtra?"Extra":"Fijo",Método:p.method||""}));if(rows.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),`${MESES[m.month].substr(0,3)} ${m.year}`.replace(/[:\\/\[\]*?]/g,""));});
    XLSX.writeFile(wb,`gastos_${addr.name||"reporte"}_${new Date().toISOString().split("T")[0]}.xlsx`);
  };

  const exportPDF=()=>{
    if(!addr||!selM.length)return;
    const f=n=>fCLP(n);
    const hAgg={};
    selM.forEach(m=>mPays(addr,m.year,m.month).filter(p=>p.isExtra).forEach(p=>{
      const k=(p.name||"").trim().toLowerCase();if(!k)return;
      hAgg[k]=hAgg[k]||{name:p.name,cat:p.catName||"",count:0,total:0,det:[]};
      hAgg[k].count++;hAgg[k].total+=p.amount||0;hAgg[k].det.push(`${fmtDate(p.date)}: ${f(p.amount)}`);
    }));
    const hormiga=Object.values(hAgg).filter(e=>e.count>=2).sort((a,b)=>b.total-a.total);
    const hormigaHTML=hormiga.length?`<h2>🐜 Gastos Frecuentes No Planificados (posible gasto hormiga)</h2><table><tr><th>Lugar/Ítem</th><th>Categoría</th><th>N° Pagos</th><th>Total</th><th>Detalle por fecha</th></tr>${hormiga.map(e=>`<tr><td>${e.name}</td><td>${e.cat||"—"}</td><td>${e.count}</td><td>${f(e.total)}</td><td style="font-size:10px">${e.det.join(", ")}</td></tr>`).join("")}</table>`:"";
    const html=`<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Reporte</title><style>body{font-family:system-ui,sans-serif;padding:2rem;color:#1e293b;font-size:13px}h1{color:#7C3AED;font-size:22px}h2{font-size:14px;color:#334155;margin:1.5rem 0 .5rem;border-bottom:2px solid #e2e8f0;padding-bottom:4px}table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:1rem}th{background:#7C3AED;color:#fff;padding:6px 10px;text-align:left}td{padding:5px 10px;border-bottom:1px solid #f1f5f9}tr:nth-child(even) td{background:#f8fafc}</style></head><body><h1>📊 Reporte — ${addr.name}</h1><p style="color:#64748b;font-size:11px">Generado el ${new Date().toLocaleDateString("es-CL")} · Presupuesto mensual: ${f(budget)}</p><h2>Resumen mensual</h2><table><tr><th>Mes</th><th>Presupuestado</th><th>Pagado</th><th>Diferencia</th></tr>${selM.map(m=>{const p=totalPaid(addr,m.year,m.month);return`<tr><td>${MESES[m.month]} ${m.year}</td><td>${f(budget)}</td><td>${f(p)}</td><td style="color:${p>budget?"#ef4444":"#10b981"}">${p>budget?"+":""}${f(p-budget)}</td></tr>`;}).join("")}<tr style="font-weight:700;background:#ede9fe"><td>TOTAL</td><td>${f(budget*selM.length)}</td><td>${f(totalPaidAll)}</td><td>${f(totalPaidAll-budget*selM.length)}</td></tr></table>${hormigaHTML}${selM.map(m=>{const pays=mPays(addr,m.year,m.month);if(!pays.length)return"";return`<h2>${MESES[m.month]} ${m.year} — ${f(totalPaid(addr,m.year,m.month))}</h2><table><tr><th>Nombre</th><th>Tipo</th><th>Categoría</th><th>Fecha</th><th>Monto</th></tr>${pays.map(p=>`<tr><td>${p.name}</td><td>${p.isExtra?"⚡":"✓"}</td><td>${p.catName||"—"}</td><td>${fmtDate(p.date)}</td><td>${f(p.amount)}</td></tr>`).join("")}</table>`;}).join("")}</body></html>`;
    const w=window.open("","_blank");if(w){w.document.write(html);w.document.close();setTimeout(()=>w.print(),500);}
  };

  if(!addr)return <CCard ch={<Empty icon={<Building2 size={36}/>} title="Sin dirección" sub="Configura una dirección en Gastos Fijos para ver el análisis." t={t}/>} t={t}/>;

  return <div>
    <SH title="Análisis de Gastos" sub={`${addr.name} · Presupuesto mensual: ${fCLP(budget)}`} t={t}
      action={<div style={{display:"flex",gap:"0.4rem"}}><Btn onClick={exportExcel} v="ok" sz="sm" icon={<Download size={13}/>}>Excel</Btn><Btn onClick={exportPDF} v="err" sz="sm" icon={<FileText size={13}/>}>PDF</Btn></div>}/>
    <CCard ch={<>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"0.65rem",flexWrap:"wrap",gap:"0.4rem"}}>
        <span style={{color:t.text,fontWeight:700,fontSize:"0.87rem"}}>Meses a analizar ({selM.length}/{recentMonths.length})</span>
        <div style={{display:"flex",gap:"0.35rem",flexWrap:"wrap"}}>
          <Btn onClick={()=>{setYearFilter(null);setSel(new Set(recentMonths.map(m=>`${m.year}-${m.month}`)));}} v="ghost" sz="sm">Todos</Btn>
          <Btn onClick={()=>{setYearFilter(CY);setSel(new Set(recentMonths.filter(m=>m.year===CY).map(m=>`${m.year}-${m.month}`)));}} v="ghost" sz="sm">Este año</Btn>
          <Btn onClick={()=>setSel(new Set([`${CY}-${CM}`]))} v="ghost" sz="sm">Este mes</Btn>
          <Btn onClick={()=>setSel(new Set(recentMonths.slice(-3).map(m=>`${m.year}-${m.month}`)))} v="ghost" sz="sm">Últ. 3</Btn>
          <Btn onClick={()=>setSel(new Set(recentMonths.slice(-6).map(m=>`${m.year}-${m.month}`)))} v="ghost" sz="sm">Últ. 6</Btn>
          <Btn onClick={()=>setSel(new Set(recentMonths.slice(-12).map(m=>`${m.year}-${m.month}`)))} v="ghost" sz="sm">Últ. 12</Btn>
        </div>
        <div style={{display:"flex",gap:"0.5rem",marginTop:"0.5rem",alignItems:"center"}}>
          <span style={{color:t.muted,fontSize:"0.75rem"}}>Filtrar por año:</span>
          <select value={yearFilter||""} onChange={e=>{const y=e.target.value?+e.target.value:null;setYearFilter(y);if(y)setSel(new Set(recentMonths.filter(m=>m.year===y).map(m=>`${m.year}-${m.month}`)));}} style={{padding:"0.25rem 0.5rem",background:t.card2,border:`1px solid ${t.border}`,borderRadius:"0.4rem",color:t.text,fontSize:"0.78rem"}}>
            <option value="">Todos</option>
            {[...new Set(recentMonths.map(m=>m.year))].sort((a,b)=>b-a).map(y=><option key={y} value={y}>{y}</option>)}
          </select>
        </div>
      </div>
      <div style={{display:"flex",flexWrap:"wrap",gap:"0.35rem"}}>
        {recentMonths.map(m=>{const k=`${m.year}-${m.month}`,isSel=sel.has(k),hp=(addr?mPays(addr,m.year,m.month):[]).length>0;return <button key={k} onClick={()=>toggle(k)} style={{padding:"0.25rem 0.6rem",borderRadius:"0.4rem",border:`1px solid ${isSel?t.pri:t.border}`,background:isSel?`${t.pri}18`:t.card2,color:isSel?t.pri:hp?t.text:t.dim,cursor:"pointer",fontSize:"0.75rem",fontWeight:isSel?700:hp?500:400}}>{MESES[m.month].substr(0,3)} {m.year}{hp?" ·":""}</button>;})}
      </div>
    </>} t={t} sx={{marginBottom:"1rem"}}/>
    {selM.length>0&&<>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(145px,1fr))",gap:"0.75rem",marginBottom:"1rem"}}>
        <StCard label="Presupuesto total" val={fCLP(budget*selM.length)} color={t.pri} icon={<Target size={16} color={t.pri}/>} t={t}/>
        <StCard label="Total pagado" val={fCLP(totalPaidAll)} color={t.ok} icon={<CheckCircle2 size={16} color={t.ok}/>} t={t}/>
        <StCard label="Diferencia" val={fCLP(Math.abs(totalPaidAll-budget*selM.length))} color={totalPaidAll>budget*selM.length?t.err:t.ok} sub={totalPaidAll>budget*selM.length?"sobre presupuesto":"bajo presupuesto"} t={t}/>
        <StCard label="Promedio mensual" val={fCLP(Math.round(totalPaidAll/selM.length))} color={t.warn} t={t}/>
      </div>
      <CCard ch={<><h4 style={{margin:"0 0 0.85rem",color:t.text,fontSize:"0.9rem",fontWeight:700}}>Presupuesto vs Pagado por mes</h4>
        <ResponsiveContainer width="100%" height={210}><BarChart data={barData} margin={{top:5,right:5,left:0,bottom:5}}><CartesianGrid strokeDasharray="3 3" stroke={t.border}/><XAxis dataKey="name" stroke={t.muted} tick={{fontSize:10,fill:t.muted}}/><YAxis stroke={t.muted} tick={{fontSize:10,fill:t.muted}} tickFormatter={v=>v>=1e6?`${(v/1e6).toFixed(1)}M`:v>=1e3?`${(v/1e3).toFixed(0)}K`:v}/><Tooltip cursor={{fill:t.pri,fillOpacity:0.08}} formatter={v=>fCLP(v)} contentStyle={{background:t.card,border:`1px solid ${t.border}`,borderRadius:"0.5rem",color:t.text,fontSize:"0.79rem"}}/><Legend wrapperStyle={{fontSize:"0.75rem"}}/><Bar dataKey="Presupuesto" fill={t.pri+"66"} radius={[3,3,0,0]}/><Bar dataKey="Pagado" fill={t.ok} radius={[3,3,0,0]}/></BarChart></ResponsiveContainer>
      </>} t={t} sx={{marginBottom:"1rem"}}/>
      {pieData.length>0&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"1rem",marginBottom:"1rem"}}>
        <CCard ch={<><h4 style={{margin:"0 0 0.85rem",color:t.text,fontSize:"0.9rem",fontWeight:700}}>Por categoría</h4>
          <ResponsiveContainer width="100%" height={185}><PieChart><Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={72} innerRadius={32}>{pieData.map((_,i)=><Cell key={i} fill={COLORS[i%COLORS.length]}/>)}</Pie><Tooltip formatter={v=>fCLP(v)} contentStyle={{background:t.card,border:`1px solid ${t.border}`,borderRadius:"0.5rem",color:t.text,fontSize:"0.79rem"}}/></PieChart></ResponsiveContainer>
        </>} t={t}/>
        <CCard ch={<><h4 style={{margin:"0 0 0.85rem",color:t.text,fontSize:"0.9rem",fontWeight:700}}>Detalle</h4>
          <div style={{maxHeight:"200px",overflowY:"auto"}}>{pieData.map((c,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"0.38rem 0",borderBottom:`1px solid ${t.border}`}}>
            <div style={{display:"flex",alignItems:"center",gap:"0.4rem"}}><div style={{width:"8px",height:"8px",borderRadius:"50%",background:COLORS[i%COLORS.length],flexShrink:0}}/><span style={{color:t.text,fontSize:"0.8rem"}}>{c.name}</span></div>
            <div style={{textAlign:"right"}}><div style={{color:t.text,fontSize:"0.8rem",fontWeight:600}}>{fCLP(c.value)}</div><div style={{color:t.dim,fontSize:"0.69rem"}}>{totalPaidAll>0?Math.round(c.value/totalPaidAll*100):0}%</div></div>
          </div>)}</div>
        </>} t={t}/>
      </div>}
      {selM.length>1&&<CCard ch={<><h4 style={{margin:"0 0 0.85rem",color:t.text,fontSize:"0.9rem",fontWeight:700}}>Evolución: Presupuesto vs Real</h4>
        <ResponsiveContainer width="100%" height={200}><LineChart data={barData} margin={{top:5,right:5,left:0,bottom:5}}><CartesianGrid strokeDasharray="3 3" stroke={t.border}/><XAxis dataKey="name" stroke={t.muted} tick={{fontSize:10,fill:t.muted}}/><YAxis stroke={t.muted} tick={{fontSize:10,fill:t.muted}} tickFormatter={v=>v>=1e3?`${(v/1e3).toFixed(0)}K`:v}/><Tooltip formatter={v=>fCLP(v)} contentStyle={{background:t.card,border:`1px solid ${t.border}`,borderRadius:"0.5rem",color:t.text,fontSize:"0.79rem"}}/><Legend wrapperStyle={{fontSize:"0.75rem"}}/><Line type="monotone" dataKey="Presupuesto" stroke={t.pri} strokeWidth={2} strokeDasharray="4 4" dot={false}/><Line type="monotone" dataKey="Pagado" stroke={t.ok} strokeWidth={2} dot={{r:3}} activeDot={{r:5}}/></LineChart></ResponsiveContainer>
      </>} t={t}/>}
    </>}
  </div>;
}

/* ══ PERFIL ══════════════════════════════════════════════ */
function PerfilView({s,d,t}){
  const[iM,setIM]=useState(null);const[gM,setGM]=useState(false);const[cfm,setCfm]=useState(null);const[pmM,setPmM]=useState(false);
  const[apiKey,setApiKey]=useState(getAK);const[showAK,setShowAK]=useState(false);
  const mInc=sumInc(s.profile);

  const exportData=()=>{try{const blob=new Blob([JSON.stringify(s,null,2)],{type:"application/json"});const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download=`gestor-gastos-backup-${new Date().toISOString().split("T")[0]}.json`;a.click();URL.revokeObjectURL(url);}catch(e){alert("Error al exportar: "+e.message);}};
  const importData=e=>{const file=e.target.files?.[0];if(!file)return;const reader=new FileReader();reader.onload=ev=>{try{const data=JSON.parse(ev.target.result);d({t:"LOAD",d:data});alert("✅ Datos importados correctamente.");}catch{alert("❌ Archivo inválido.");}};reader.readAsText(file);e.target.value="";};

  return <div>
    <SH title="Mi Perfil" sub="Ingresos, objetivos y configuración" t={t}/>
    <CCard ch={<>
      <h4 style={{margin:"0 0 0.75rem",color:t.text,fontSize:"0.9rem",fontWeight:700}}>👤 Información personal</h4>
      <Fld label="Tu nombre (aparece en el saludo)" ch={<TI val={s.profile.name} onChange={v=>d({t:"SP",d:{name:v}})} ph="Ej: Ricardo" t={t}/>} t={t}/>
      <div style={{display:"flex",alignItems:"center",gap:"0.6rem",padding:"0.6rem 0.75rem",background:t.card2,borderRadius:"0.5rem",fontSize:"0.84rem",color:t.text}}>
        <input type="checkbox" checked={s.settings.linkProf||false} onChange={e=>d({t:"SS",k:"linkProf",v:e.target.checked})} style={{width:"15px",height:"15px",accentColor:t.pri}}/>
        <label style={{cursor:"pointer",flex:1}}>Mostrar alerta cuando los gastos superan mis ingresos</label>
      </div>
    </>} t={t} sx={{marginBottom:"1rem"}}/>
    {/* API Key */}
    <CCard ch={<>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:(!apiKey||showAK)?"0.65rem":"0"}}>
        <div style={{display:"flex",alignItems:"center",gap:"0.5rem"}}>
          <span>🔑</span><h4 style={{margin:0,color:t.text,fontSize:"0.9rem",fontWeight:700}}>API Key (Plan IA)</h4>
          {apiKey&&<Badge ch="✓ Configurada" color={t.ok}/>}
        </div>
        <Btn onClick={()=>setShowAK(p=>!p)} v="ghost" sz="sm">{apiKey?(showAK?"Ocultar":"Cambiar"):"Agregar"}</Btn>
      </div>
      {(!apiKey||showAK)&&<>
        {!apiKey&&<p style={{color:t.muted,fontSize:"0.82rem",margin:"0 0 0.75rem",lineHeight:1.5}}>Para usar la IA necesitas una key gratuita de Anthropic. <a href="https://platform.claude.com" target="_blank" style={{color:t.pri}}>Obtener →</a> · Hay una guía paso a paso en la pestaña <strong style={{color:t.text}}>Plan IA</strong>.</p>}
        <Fld label="API Key (sk-ant-...)" ch={<TI val={apiKey} onChange={k=>{setApiKey(k);saveAK(k);}} ph="sk-ant-api03-..." t={t} type="password"/>} t={t}/>
        {apiKey&&<Btn onClick={()=>setShowAK(false)} v="ok" sz="sm" icon={<Check size={12}/>}>Guardar</Btn>}
      </>}
    </>} t={t} sx={{marginBottom:"1rem"}}/>
    {/* Ingresos */}
    <CCard ch={<>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"0.75rem"}}>
        <h4 style={{margin:0,color:t.text,fontSize:"0.9rem",fontWeight:700}}>💰 Fuentes de ingreso</h4>
        <Btn onClick={()=>setIM("add")} icon={<Plus size={13}/>} sz="sm">Agregar</Btn>
      </div>
      {s.profile.incomes.length===0&&<p style={{color:t.muted,fontSize:"0.84rem",textAlign:"center",padding:"0.75rem 0"}}>Agrega tu sueldo, beca, aporte familiar, etc.</p>}
      {s.profile.incomes.map(inc=><div key={inc.id} style={{display:"flex",alignItems:"center",gap:"0.65rem",padding:"0.55rem 0.75rem",background:t.card2,borderRadius:"0.5rem",marginBottom:"0.4rem"}}>
        <DollarSign size={14} color={t.ok}/>
        <div style={{flex:1}}><div style={{color:t.text,fontSize:"0.84rem",fontWeight:600}}>{inc.name}</div><div style={{color:t.muted,fontSize:"0.72rem"}}>{inc.type} · {FREQS.find(f=>f.v===inc.frequency)?.l}</div></div>
        <div style={{textAlign:"right"}}><div style={{color:t.ok,fontWeight:700,fontSize:"0.84rem"}}>{fCLP(inc.amount)}</div>{inc.frequency!=="mensual"&&<div style={{color:t.dim,fontSize:"0.69rem"}}>{fCLP(toM(inc))}/mes</div>}</div>
        <button onClick={()=>setIM({e:inc})} style={{background:"none",border:"none",cursor:"pointer",color:t.muted,padding:"0.2rem",display:"flex"}}><Edit2 size={13}/></button>
        <button onClick={()=>setCfm({msg:`¿Eliminar "${inc.name}"?`,ok:()=>{d({t:"DINC",id:inc.id});setCfm(null);}})} style={{background:"none",border:"none",cursor:"pointer",color:t.err,padding:"0.2rem",display:"flex"}}><Trash2 size={13}/></button>
      </div>)}
      {s.profile.incomes.length>0&&<div style={{display:"flex",justifyContent:"space-between",padding:"0.55rem 0.75rem",background:t.ok+"12",border:`1px solid ${t.ok}40`,borderRadius:"0.5rem",marginTop:"0.5rem"}}>
        <span style={{color:t.muted,fontSize:"0.84rem"}}>Total mensual estimado</span>
        <span style={{color:t.ok,fontWeight:800,fontSize:"0.95rem"}}>{fCLP(mInc)}</span>
      </div>}
    </>} t={t} sx={{marginBottom:"1rem"}}/>
    {/* Formas de pago frecuentes */}
    <CCard ch={<>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"0.75rem"}}>
        <h4 style={{margin:0,color:t.text,fontSize:"0.9rem",fontWeight:700}}>💳 Formas de pago frecuentes</h4>
        <Btn onClick={()=>setPmM(true)} icon={<Plus size={13}/>} sz="sm">Agregar</Btn>
      </div>
      {(!s.profile.paymentMethods||s.profile.paymentMethods.length===0)&&<p style={{color:t.muted,fontSize:"0.84rem",textAlign:"center",padding:"0.75rem 0"}}>Agrega tus cuentas o tarjetas habituales (ej. "Débito BancoEstado", "Mach") para elegirlas con un tap al registrar un pago.</p>}
      {(s.profile.paymentMethods||[]).map(m=><div key={m.id} style={{display:"flex",alignItems:"center",gap:"0.65rem",padding:"0.55rem 0.75rem",background:t.card2,borderRadius:"0.5rem",marginBottom:"0.4rem"}}>
        <Star size={14} color={m.isDefault?t.warn:t.dim} fill={m.isDefault?t.warn:"none"}/>
        <div style={{flex:1}}>
          <div style={{color:t.text,fontSize:"0.84rem",fontWeight:600}}>{m.name}</div>
          {m.isDefault&&<div style={{color:t.dim,fontSize:"0.7rem"}}>Principal — aparece preseleccionada</div>}
        </div>
        {!m.isDefault&&<Btn onClick={()=>d({t:"SETDEFPM",id:m.id})} v="ghost" sz="sm">Usar como principal</Btn>}
        <button onClick={()=>setCfm({msg:`¿Eliminar "${m.name}"?`,ok:()=>{d({t:"DPM",id:m.id});setCfm(null);}})} style={{background:"none",border:"none",cursor:"pointer",color:t.err,padding:"0.2rem",display:"flex"}}><Trash2 size={13}/></button>
      </div>)}
    </>} t={t} sx={{marginBottom:"1rem"}}/>
    {/* Disponibilidad */}
    <CCard ch={<>
      <h4 style={{margin:"0 0 0.75rem",color:t.text,fontSize:"0.9rem",fontWeight:700}}>⏰ Disponibilidad</h4>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"0.5rem",marginBottom:"0.75rem"}}>
        {[["canWork","Puedo trabajar"],["studying","Estoy estudiando"],["entrepreneur","Quiero emprender"]].map(([k,lb])=><label key={k} style={{display:"flex",alignItems:"center",gap:"0.5rem",cursor:"pointer",fontSize:"0.84rem",color:t.text,background:t.card2,padding:"0.5rem 0.65rem",borderRadius:"0.5rem"}}>
          <input type="checkbox" checked={!!s.profile.availability[k]} onChange={e=>d({t:"SAV",d:{[k]:e.target.checked}})} style={{width:"15px",height:"15px",accentColor:t.pri}}/>{lb}
        </label>)}
        <div style={{background:t.card2,padding:"0.5rem 0.65rem",borderRadius:"0.5rem"}}>
          <label style={{display:"block",color:t.muted,fontSize:"0.69rem",marginBottom:"0.25rem",fontWeight:700}}>HRS LIBRES/SEMANA</label>
          <input type="number" min="0" max="168" value={s.profile.availability.hoursPerWeek||0} onChange={e=>d({t:"SAV",d:{hoursPerWeek:+e.target.value}})} style={{width:"100%",padding:"0.25rem 0.4rem",background:t.card,border:`1px solid ${t.border}`,borderRadius:"0.35rem",color:t.text,fontSize:"0.85rem"}}/>
        </div>
      </div>
      <Fld label="Notas de tu situación" ch={<TA val={s.profile.availability.notes||""} onChange={v=>d({t:"SAV",d:{notes:v}})} ph="Busco empleo, estudiante de último año, tengo auto..." rows={2} t={t}/>} t={t}/>
    </>} t={t} sx={{marginBottom:"1rem"}}/>
    {/* Objetivos */}
    <CCard ch={<>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"0.75rem"}}>
        <h4 style={{margin:0,color:t.text,fontSize:"0.9rem",fontWeight:700}}>🎯 Objetivos financieros</h4>
        <Btn onClick={()=>setGM(true)} icon={<Plus size={13}/>} sz="sm">Agregar</Btn>
      </div>
      {s.profile.goals.length===0&&<p style={{color:t.muted,fontSize:"0.84rem",textAlign:"center",padding:"0.75rem 0"}}>Agrega metas: "Fondo emergencia", "GPU nueva", "Vacaciones"...</p>}
      {s.profile.goals.map(g=><div key={g.id} style={{display:"flex",alignItems:"center",gap:"0.65rem",padding:"0.55rem 0.75rem",background:t.card2,borderRadius:"0.5rem",marginBottom:"0.4rem"}}>
        <Target size={14} color={g.priority==="alta"?t.err:g.priority==="media"?t.warn:t.ok}/>
        <div style={{flex:1}}><div style={{color:t.text,fontSize:"0.84rem",fontWeight:600}}>{g.name}</div><div style={{color:t.muted,fontSize:"0.72rem"}}>{g.deadline?`Plazo: ${g.deadline} · `:""}Prioridad {g.priority||"media"}</div></div>
        <span style={{color:t.pri,fontWeight:700,fontSize:"0.84rem"}}>{fCLP(g.targetAmount)}</span>
        <button onClick={()=>setCfm({msg:`¿Eliminar objetivo "${g.name}"?`,ok:()=>{d({t:"DG",id:g.id});setCfm(null);}})} style={{background:"none",border:"none",cursor:"pointer",color:t.err,padding:"0.2rem",display:"flex"}}><Trash2 size={13}/></button>
      </div>)}
    </>} t={t} sx={{marginBottom:"1rem"}}/>
    {/* Export/Import — identidad de datos */}
    <CCard ch={<>
      <h4 style={{margin:"0 0 0.75rem",color:t.text,fontSize:"0.9rem",fontWeight:700}}>💾 Mis datos — Respaldo y restauración</h4>
      <div style={{display:"flex",gap:"0.6rem",padding:"0.7rem 0.85rem",background:t.err+"13",border:`1px solid ${t.err}45`,borderRadius:"0.55rem",marginBottom:"0.85rem"}}>
        <span style={{fontSize:"1.2rem",lineHeight:1.2,flexShrink:0}}>⚠️</span>
        <div>
          <div style={{color:t.err,fontWeight:700,fontSize:"0.84rem",marginBottom:"0.3rem"}}>¡Atención! Riesgo de pérdida de datos</div>
          <div style={{color:t.muted,fontSize:"0.79rem",lineHeight:1.55}}>
            Tus datos viven solo en <strong style={{color:t.text}}>este navegador</strong>. Estas acciones los borran permanentemente:<br/>
            • <strong style={{color:t.text}}>Borrar datos/caché del navegador</strong> (Configuración → Privacidad)<br/>
            • <strong style={{color:t.text}}>"Limpiar datos de sitios"</strong> en Chrome/Firefox<br/>
            • <strong style={{color:t.text}}>Modo incógnito/privado</strong> (se pierden al cerrar la ventana)<br/>
            • <strong style={{color:t.text}}>Reinstalar o cambiar de navegador</strong><br/><br/>
            👉 <strong style={{color:t.text}}>Exporta un respaldo regularmente</strong>, sobre todo antes de actualizar o limpiar el navegador.
          </div>
        </div>
      </div>
      <div style={{display:"flex",gap:"0.65rem",flexWrap:"wrap"}}>
        <Btn onClick={exportData} v="out" icon={<Download size={14}/>}>Exportar respaldo (JSON)</Btn>
        <label style={{display:"flex",alignItems:"center",gap:"0.4rem",padding:"0.5rem 1rem",background:"transparent",color:"inherit",border:"1px solid #94A3B8",borderRadius:"0.5rem",cursor:"pointer",fontWeight:600,fontSize:"0.85rem"}}>
          <Upload size={14}/> Importar respaldo
          <input type="file" accept=".json" onChange={importData} style={{display:"none"}}/>
        </label>
        <Btn onClick={()=>setCfm({msg:"¿Eliminar TODOS tus datos? Esta acción no se puede deshacer.",ok:()=>{d({t:"LOAD",d:INIT});setCfm(null);}})} v="ghost" sx={{color:t.err}}>🗑️ Borrar todo</Btn>
      </div>
    </>} t={t}/>
    {iM&&<Modal title={iM==="add"?"Nuevo Ingreso":"Editar Ingreso"} onClose={()=>setIM(null)} t={t} ch={<IncF init={iM==="add"?{}:{type:iM.e.type,name:iM.e.name,amount:iM.e.amount,frequency:iM.e.frequency}} onSave={dt=>{iM==="add"?d({t:"AINC",d:dt}):d({t:"UINC",id:iM.e.id,d:dt});setIM(null);}} t={t}/>}/>}
    {gM&&<Modal title="Nuevo Objetivo" onClose={()=>setGM(false)} t={t} ch={<GoalF onSave={dt=>{d({t:"AG",d:dt});setGM(false);}} t={t}/>}/>}
    {pmM&&<Modal title="Nueva Forma de Pago" onClose={()=>setPmM(false)} t={t} ch={<PayMethodF onSave={name=>{d({t:"APM",name});setPmM(false);}} t={t}/>}/>}
    {cfm&&<Cfm msg={cfm.msg} onOk={cfm.ok} onCancel={()=>setCfm(null)} t={t}/>}
  </div>;
}
function IncF({init,onSave,t}){
  const[type,setType]=useState(init.type||INC_TYPES[0]);const[name,setName]=useState(init.name||"");const[amount,setAmount]=useState(init.amount||0);const[freq,setFreq]=useState(init.frequency||"mensual");
  const eq=Math.round(amount*(MF[freq]||1));
  return <>
    <Fld label="Tipo" ch={<Sel val={type} onChange={setType} opts={INC_TYPES} t={t}/>} t={t}/>
    <Fld label="Descripción" ch={<TI val={name} onChange={setName} ph="Sueldo empresa X, Beca MINEDUC..." t={t}/>} t={t}/>
    <Fld label="Monto" ch={<MoneyInput val={amount} onChange={setAmount} t={t}/>} t={t}/>
    <Fld label="Frecuencia" ch={<Sel val={freq} onChange={setFreq} opts={FREQS.map(f=>({v:f.v,l:f.l}))} t={t}/>} t={t}/>
    {amount>0&&freq!=="mensual"&&<div style={{padding:"0.55rem 0.75rem",background:t.ok+"12",borderRadius:"0.5rem",marginBottom:"0.85rem",color:t.muted,fontSize:"0.8rem"}}>Equivalente mensual: <strong style={{color:t.ok}}>{fCLP(eq)}</strong></div>}
    <Btn onClick={()=>name&&amount&&onSave({type,name,amount,frequency:freq})} dis={!name||!amount} full icon={<Check size={14}/>}>{init.name?"Actualizar":"Agregar"}</Btn>
  </>;
}
function GoalF({onSave,t}){
  const[name,setName]=useState(""),[target,setTarget]=useState(0),[dl,setDl]=useState(""),[pri,setPri]=useState("media");
  const ps={alta:t.err,media:t.warn,baja:t.ok};
  return <>
    <Fld label="Objetivo" ch={<TI val={name} onChange={setName} ph="Fondo emergencia, GPU, Viaje..." t={t}/>} t={t}/>
    <Fld label="Monto objetivo" ch={<MoneyInput val={target} onChange={setTarget} t={t}/>} t={t}/>
    <Fld label="Fecha límite (opcional)" ch={<TI val={dl} onChange={setDl} type="month" t={t}/>} t={t}/>
    <Fld label="Prioridad" ch={<div style={{display:"flex",gap:"0.5rem"}}>{["alta","media","baja"].map(p=><button key={p} onClick={()=>setPri(p)} style={{flex:1,padding:"0.4rem",border:`2px solid ${pri===p?ps[p]:t.border}`,borderRadius:"0.5rem",background:pri===p?`${ps[p]}18`:"none",color:ps[p],cursor:"pointer",fontSize:"0.79rem",fontWeight:700,textTransform:"capitalize"}}>{p}</button>)}</div>} t={t}/>
    <Btn onClick={()=>name&&onSave({name,targetAmount:target,deadline:dl,priority:pri})} dis={!name} full icon={<Check size={14}/>}>Crear objetivo</Btn>
  </>;
}
function PayMethodF({onSave,t}){
  const[name,setName]=useState("");
  return <>
    <Fld label="Nombre de la cuenta o tarjeta" ch={<TI val={name} onChange={setName} ph="Débito BancoEstado, Mach, Efectivo..." t={t}/>} t={t}/>
    <p style={{color:t.dim,fontSize:"0.76rem",margin:"0 0 0.85rem",lineHeight:1.5}}>La primera que agregues queda marcada como principal. Puedes cambiarla después con "Usar como principal".</p>
    <Btn onClick={()=>name.trim()&&onSave(name.trim())} dis={!name.trim()} full icon={<Check size={14}/>}>Agregar</Btn>
  </>;
}

/* ══ IA ══════════════════════════════════════════════════ */
/* ══ GUÍA DE CONFIGURACIÓN API KEY ═══════════════════════
   Nota de honestidad: solo podemos detectar mecánicamente si YA hay
   una key guardada en este navegador (apiKey presente/ausente).
   No hay forma de saber si la persona ya tiene cuenta en Anthropic
   antes de que pegue su key — por eso el mensaje lo aclara y deja
   la puerta abierta a "ya tengo mi key" en todo momento.
══════════════════════════════════════════════════════════ */
function MockBrowser({t,url,ch}){
  return <div style={{border:`1px solid ${t.border}`,borderRadius:"0.6rem",overflow:"hidden",background:t.card2,margin:"0.6rem 0"}}>
    <div style={{display:"flex",alignItems:"center",gap:"0.4rem",padding:"0.4rem 0.6rem",background:t.card3,borderBottom:`1px solid ${t.border}`}}>
      <div style={{display:"flex",gap:"0.25rem"}}>{["#EF4444","#F59E0B","#10B981"].map(c=><div key={c} style={{width:"7px",height:"7px",borderRadius:"50%",background:c}}/>)}</div>
      <div style={{flex:1,background:t.card,borderRadius:"0.3rem",padding:"0.15rem 0.5rem",color:t.dim,fontSize:"0.68rem",textAlign:"center"}}>{url}</div>
    </div>
    <div style={{padding:"0.9rem"}}>{ch}</div>
  </div>;
}
const QUICK_STEPS=[
  {ic:"👤",lb:"Crear cuenta",sub:"platform.claude.com"},
  {ic:"🔑",lb:"Crear API Key",sub:"Settings → API Keys"},
  {ic:"📋",lb:"Copiar la key",sub:"empieza con sk-ant-"},
  {ic:"💳",lb:"Cargar crédito",sub:"desde $1.000 CLP"},
  {ic:"📌",lb:"Pegar aquí",sub:"¡y listo!"},
];
function DETAIL_STEPS(t){return[
  {e:"👋",time:"1 min",title:"¿Qué vamos a hacer?",desc:<>
    <p style={{margin:"0 0 0.6rem",lineHeight:1.65}}>Vamos a crear tu propia llave de acceso ("API key") para que la sección Plan IA pueda funcionar. Es gratuita de crear y tú controlas cuánto gastas — nunca se cobra nada sin que tú cargues crédito primero.</p>
    <p style={{margin:0,lineHeight:1.65}}>Son 5 pasos, toma unos 5 minutos la primera vez. Puedes omitir esto en cualquier momento con el botón de arriba.</p>
  </>},
  {e:"👤",time:"~2 min",title:"1. Crea tu cuenta",desc:<>
    <p style={{margin:"0 0 0.4rem",lineHeight:1.6}}>Ve a <strong style={{color:t.text}}>platform.claude.com</strong> y crea una cuenta gratis con tu correo o con Google. Esto es distinto de una cuenta de Claude.ai (el chat) — es el panel para desarrolladores, pero cualquier persona puede usarlo.</p>
    <MockBrowser t={t} url="platform.claude.com/signup" ch={<div style={{display:"flex",flexDirection:"column",gap:"0.4rem",alignItems:"center"}}>
      <div style={{fontSize:"0.78rem",color:t.muted,marginBottom:"0.2rem"}}>Crear cuenta</div>
      <div style={{width:"85%",padding:"0.35rem",background:t.card,border:`1px solid ${t.border}`,borderRadius:"0.35rem",fontSize:"0.72rem",color:t.dim,textAlign:"center"}}>🔵 Continuar con Google</div>
      <div style={{width:"85%",padding:"0.35rem",background:t.card,border:`1px solid ${t.border}`,borderRadius:"0.35rem",fontSize:"0.72rem",color:t.dim,textAlign:"center"}}>✉️ correo@ejemplo.com</div>
    </div>}/>
    <p style={{margin:0,color:t.dim,fontSize:"0.73rem",fontStyle:"italic"}}>Ilustración simplificada — la interfaz real puede verse distinta.</p>
  </>},
  {e:"✉️",time:"~1 min",title:"2. Verifica tu correo",desc:<>
    <p style={{margin:0,lineHeight:1.6}}>Te llegará un correo de confirmación. Ábrelo y haz click en el link de verificación. Con eso tu cuenta queda activa — todavía no has gastado ni cargado nada.</p>
  </>},
  {e:"🔑",time:"~1 min",title:"3. Crea tu API Key",desc:<>
    <p style={{margin:"0 0 0.4rem",lineHeight:1.6}}>Dentro del panel, ve a <strong style={{color:t.text}}>Settings → API Keys</strong> y haz click en <strong style={{color:t.text}}>"Create Key"</strong>. Puedes ponerle el nombre que quieras, por ejemplo "GestorGastos".</p>
    <MockBrowser t={t} url="platform.claude.com/settings/keys" ch={<div style={{display:"flex",flexDirection:"column",gap:"0.35rem"}}>
      <div style={{fontSize:"0.72rem",color:t.muted}}>API Keys</div>
      <div style={{display:"flex",justifyContent:"flex-end"}}><div style={{padding:"0.3rem 0.6rem",background:t.pri,color:"#fff",borderRadius:"0.3rem",fontSize:"0.7rem",fontWeight:700}}>+ Create Key</div></div>
    </div>}/>
  </>},
  {e:"📋",time:"~1 min",title:"4. Copia y guarda la key",desc:<>
    <p style={{margin:"0 0 0.5rem",lineHeight:1.6}}>Va a aparecer una key larga que empieza con <code style={{background:t.card2,padding:"0.1rem 0.35rem",borderRadius:"0.25rem",color:t.pri}}>sk-ant-...</code>. </p>
    <div style={{display:"flex",gap:"0.5rem",padding:"0.6rem 0.75rem",background:t.warn+"12",border:`1px solid ${t.warn}40`,borderRadius:"0.5rem"}}>
      <AlertTriangle size={14} color={t.warn} style={{flexShrink:0,marginTop:"0.1rem"}}/>
      <span style={{fontSize:"0.79rem",color:t.muted,lineHeight:1.5}}><strong style={{color:t.text}}>Solo se muestra una vez.</strong> Cópiala altiro. Trátala como una contraseña: no la compartas con nadie ni la publiques en redes o repositorios.</span>
    </div>
  </>},
  {e:"💳",time:"~2 min",title:"5. Cargar crédito (sin sorpresas)",desc:<>
    <p style={{margin:"0 0 0.6rem",lineHeight:1.6}}>El sistema funciona con <strong style={{color:t.text}}>saldo prepago</strong>: cargas un monto, y solo se descuenta de ahí. Algunas cuentas reciben un crédito de regalo (no está garantizado en todos los países) — si aparece, genial; si no, puedes cargar tú mismo desde muy poco.</p>
    <div style={{padding:"0.7rem 0.85rem",background:t.pri+"10",border:`1px solid ${t.pri}30`,borderRadius:"0.55rem",marginBottom:"0.6rem"}}>
      <div style={{color:t.pri,fontWeight:700,fontSize:"0.8rem",marginBottom:"0.4rem"}}>💰 ¿Cuánto cuesta realmente usar el Plan IA acá?</div>
      <div style={{color:t.muted,fontSize:"0.78rem",lineHeight:1.7}}>
        Esta app usa el modelo Sonnet 4.6 (~$3 USD por millón de palabras leídas, ~$15 USD por millón generadas).<br/>
        → Cada plan que generas cuesta aprox. <strong style={{color:t.text}}>$15-20 CLP</strong> (menos de un peso... bueno, menos de 20).<br/>
        → Con <strong style={{color:t.text}}>$1.000 CLP</strong> cargados alcanzas para <strong style={{color:t.text}}>+50 planes</strong>.<br/>
        <span style={{fontSize:"0.71rem",color:t.dim}}>Precios de referencia (jul. 2026) — pueden cambiar; el valor exacto siempre está en Billing → Cost de tu cuenta.</span>
      </div>
    </div>
    <div style={{color:t.text,fontWeight:700,fontSize:"0.82rem",marginBottom:"0.4rem"}}>Para no gastar de más:</div>
    <ul style={{margin:0,paddingLeft:"1.1rem",color:t.muted,fontSize:"0.79rem",lineHeight:1.85}}>
      <li><strong style={{color:t.text}}>No actives "recarga automática"</strong> — así jamás gastas más de lo que cargaste a propósito.</li>
      <li>Carga un monto chico para partir (con $5.000-$10.000 CLP tienes para meses de uso normal).</li>
      <li>Puedes poner además un tope mensual en Billing → Limits, como respaldo extra.</li>
      <li>Revisa tu gasto real cuando quieras en Billing → Cost.</li>
    </ul>
  </>},
  {e:"📌",time:"~1 min",title:"6. Pega tu key aquí y listo",desc:<>
    <p style={{margin:0,lineHeight:1.6}}>Vuelve a esta pestaña, pega tu key en el campo de abajo (o en tu Perfil) y guarda. Se queda guardada solo en este navegador — nunca se envía a ningún servidor nuestro.</p>
  </>},
];}
function QuickGuide({t,onDetail,onSkip,onNeverShow,apiKey,setApiKey}){
  const[kv,setKv]=useState("");
  return <CCard t={t} sx={{marginBottom:"1rem"}} ch={<>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"0.9rem"}}>
      <span style={{color:t.text,fontWeight:800,fontSize:"0.95rem"}}>🔑 Configura tu acceso a la IA</span>
      <button onClick={onSkip} style={{background:"none",border:"none",cursor:"pointer",color:t.dim,fontSize:"0.78rem",padding:"0.2rem 0.4rem"}}>Omitir ✕</button>
    </div>
    <p style={{margin:"0 0 0.9rem",color:t.muted,fontSize:"0.81rem",lineHeight:1.55}}>No podemos saber si ya tienes cuenta en Anthropic — pero si ya tienes tu key, pégala directo abajo. Si no, estos son los pasos:</p>
    <div style={{display:"flex",flexWrap:"wrap",alignItems:"center",gap:"0.2rem",marginBottom:"1rem",justifyContent:"center"}}>
      {QUICK_STEPS.flatMap((st,i)=>[
        <div key={`s${i}`} style={{display:"flex",flexDirection:"column",alignItems:"center",width:"84px",textAlign:"center"}}>
          <div style={{width:"38px",height:"38px",borderRadius:"50%",background:t.card2,border:`1px solid ${t.border}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"1.1rem",marginBottom:"0.3rem"}}>{st.ic}</div>
          <div style={{color:t.text,fontSize:"0.68rem",fontWeight:700,lineHeight:1.2}}>{st.lb}</div>
          <div style={{color:t.dim,fontSize:"0.62rem",marginTop:"0.1rem"}}>{st.sub}</div>
        </div>,
        i<QUICK_STEPS.length-1?<ChevronRight key={`a${i}`} size={14} color={t.dim} style={{flexShrink:0,marginTop:"-1.1rem"}}/>:null
      ])}
    </div>
    <div style={{display:"flex",gap:"0.5rem",marginBottom:"1rem",flexWrap:"wrap"}}>
      <Btn onClick={onDetail} v="out" sz="sm" icon={<BookOpen size={13}/>}>Ver guía detallada paso a paso</Btn>
      <a href="https://platform.claude.com" target="_blank" style={{fontSize:"0.79rem",color:t.pri,display:"flex",alignItems:"center",gap:"0.25rem"}}>Ir a platform.claude.com <ArrowRight size={12}/></a>
    </div>
    <div style={{borderTop:`1px solid ${t.border}`,paddingTop:"0.85rem"}}>
      <div style={{color:t.text,fontWeight:700,fontSize:"0.82rem",marginBottom:"0.5rem"}}>¿Ya tienes tu key? Pégala aquí:</div>
      <div style={{display:"flex",gap:"0.5rem"}}>
        <TI val={kv} onChange={setKv} ph="sk-ant-api03-..." t={t} type="password"/>
        <Btn onClick={()=>{if(kv){setApiKey(kv);saveAK(kv);}}} dis={!kv} icon={<Check size={13}/>}>Guardar</Btn>
      </div>
    </div>
    <button onClick={onNeverShow} style={{background:"none",border:"none",cursor:"pointer",color:t.dim,fontSize:"0.72rem",padding:"0.6rem 0 0",textDecoration:"underline"}}>No mostrar esta guía de nuevo</button>
  </>}/>;
}
function DetailGuide({t,onBack,onSkip,onNeverShow,apiKey,setApiKey}){
  const[step,setStep]=useState(0);
  const STEPS=DETAIL_STEPS(t);
  const cur=STEPS[step],isLast=step===STEPS.length-1;
  const[kv,setKv]=useState("");
  return <CCard t={t} sx={{marginBottom:"1rem"}} ch={<>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"0.75rem"}}>
      <button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",color:t.muted,fontSize:"0.78rem",display:"flex",alignItems:"center",gap:"0.25rem",padding:"0.2rem"}}><ChevronLeft size={13}/> Volver al resumen</button>
      <button onClick={onSkip} style={{background:"none",border:"none",cursor:"pointer",color:t.dim,fontSize:"0.78rem",padding:"0.2rem 0.4rem"}}>Omitir ✕</button>
    </div>
    <div style={{display:"flex",justifyContent:"center",gap:"0.3rem",marginBottom:"1rem"}}>
      {STEPS.map((_,i)=><div key={i} style={{width:i===step?"18px":"6px",height:"6px",borderRadius:"3px",background:i===step?t.pri:t.border,transition:"all .3s"}}/>)}
    </div>
    <div style={{textAlign:"center",marginBottom:"0.4rem"}}>
      <div style={{fontSize:"2.2rem",marginBottom:"0.4rem",lineHeight:1}}>{cur.e}</div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:"0.5rem",marginBottom:"0.7rem"}}>
        <h3 style={{margin:0,color:t.text,fontSize:"1.05rem",fontWeight:800}}>{cur.title}</h3>
        <Badge ch={`⏱ ${cur.time}`} color={t.info}/>
      </div>
    </div>
    <div style={{textAlign:"left",marginBottom:"1rem"}}>{cur.desc}</div>
    {isLast&&<div style={{borderTop:`1px solid ${t.border}`,paddingTop:"0.85rem",marginBottom:"0.75rem"}}>
      <div style={{display:"flex",gap:"0.5rem"}}>
        <TI val={kv} onChange={setKv} ph="sk-ant-api03-..." t={t} type="password"/>
        <Btn onClick={()=>{if(kv){setApiKey(kv);saveAK(kv);}}} dis={!kv} icon={<Check size={13}/>}>Guardar</Btn>
      </div>
    </div>}
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
      <Btn onClick={()=>setStep(p=>Math.max(0,p-1))} v="ghost" sz="sm" dis={step===0} icon={<ChevronLeft size={13}/>}>Anterior</Btn>
      <span style={{color:t.dim,fontSize:"0.72rem"}}>{step+1}/{STEPS.length}</span>
      {!isLast?<Btn onClick={()=>setStep(p=>p+1)} sz="sm">Siguiente →</Btn>:<Btn onClick={onBack} v="ok" sz="sm" icon={<Check size={13}/>}>Listo</Btn>}
    </div>
    <button onClick={onNeverShow} style={{background:"none",border:"none",cursor:"pointer",color:t.dim,fontSize:"0.72rem",padding:"0.75rem 0 0",textDecoration:"underline",display:"block",margin:"0.75rem auto 0"}}>No mostrar esta guía de nuevo</button>
  </>}/>;
}
function MinimalKeyPrompt({t,onShowGuide,apiKey,setApiKey}){
  const[kv,setKv]=useState("");
  return <CCard t={t} sx={{marginBottom:"1rem"}} ch={<div style={{display:"flex",gap:"0.75rem",alignItems:"flex-start"}}>
    <div style={{fontSize:"1.4rem"}}>🔑</div>
    <div style={{flex:1}}>
      <div style={{color:t.text,fontWeight:700,fontSize:"0.87rem",marginBottom:"0.35rem"}}>Necesitas una API key para usar la IA</div>
      <div style={{display:"flex",gap:"0.5rem",marginBottom:"0.5rem"}}>
        <TI val={kv} onChange={setKv} ph="Pega tu key: sk-ant-api03-..." t={t} type="password"/>
        <Btn onClick={()=>{if(kv){setApiKey(kv);saveAK(kv);}}} dis={!kv} sz="sm" icon={<Check size={13}/>}>Guardar</Btn>
      </div>
      <button onClick={onShowGuide} style={{background:"none",border:"none",cursor:"pointer",color:t.pri,fontSize:"0.78rem",padding:0}}>📖 ¿Cómo obtengo mi key? Ver guía paso a paso</button>
    </div>
  </div>}/>;
}

const PLANS=[
  {id:"ahorro",e:"💰",label:"Plan de Ahorro",desc:"Estrategia para maximizar ahorros"},
  {id:"diagnostico",e:"🔍",label:"Diagnóstico",desc:"Análisis detallado de mis gastos"},
  {id:"supervivencia",e:"🆘",label:"Supervivencia",desc:"Vivir con el mínimo posible"},
  {id:"ingresos",e:"🚀",label:"Generar ingresos",desc:"Ideas para Chile según mi disponibilidad"},
  {id:"gestion",e:"📊",label:"Presupuesto 50/30/20",desc:"Distribución óptima de ingresos"},
  {id:"libre",e:"❓",label:"Consulta libre",desc:"Pregunta lo que necesites"},
];
function IAView({s,d,t}){
  const[loading,setLoading]=useState(false);const[resp,setResp]=useState("");const[plan,setPlan]=useState("ahorro");const[cQ,setCQ]=useState("");
  const[useProf,setUseProf]=useState(true);const[showHist,setShowHist]=useState(false);const[apiKey,setApiKey]=useState(getAK);
  const[guideView,setGuideView]=useState(()=>(!getAK()&&!s.settings.iaGuideDismissed)?"quick":null);
  useEffect(()=>{if(apiKey)setGuideView(null);},[apiKey]);
  const addr=s.addresses.find(a=>a.id===s.selAddr);
  const budget=addr?tmplTotal(addr):0;
  const recMonths=useMemo(()=>{if(!addr)return[];const months=[];for(let i=2;i>=0;i--){let m=CM-i,y=CY;if(m<0){m+=12;y--;}months.push({year:y,month:m});}return months;},[addr]);
  const ctx=()=>{
    let c="";
    if(addr){
      c+=`\n\n### PLANTILLA DE GASTOS FIJOS:\nPresupuesto mensual total: ${fCLP(budget)}\n`;
      (addr.template?.categories||[]).forEach(cat=>{c+=`\nCategoría ${cat.name}:\n`;(cat.items||[]).forEach(it=>c+=`- ${it.name}: ${fCLP(it.amount)}${it.isVariable?" (variable)":""}\n`);});
      c+="\n### PAGOS REALES RECIENTES (con fecha):\n";
      recMonths.forEach(m=>{const pays=mPays(addr,m.year,m.month);if(pays.length){c+=`\n${MESES[m.month]} ${m.year} (Total: ${fCLP(totalPaid(addr,m.year,m.month))}, Presupuesto: ${fCLP(budget)}):\n`;pays.forEach(p=>c+=`- ${fmtDate(p.date)}: ${p.name} — ${fCLP(p.amount)} (${p.isExtra?"extra":"fijo"}${p.catName?`, ${p.catName}`:""})\n`);}});
      // Agregación de gasto hormiga: compras chicas y repetidas en el mismo lugar/ítem
      const extraAgg={};
      recMonths.forEach(m=>mPays(addr,m.year,m.month).filter(p=>p.isExtra).forEach(p=>{
        const k=(p.name||"").trim().toLowerCase();if(!k)return;
        extraAgg[k]=extraAgg[k]||{name:p.name,total:0,count:0,dates:[]};
        extraAgg[k].total+=p.amount||0;extraAgg[k].count++;extraAgg[k].dates.push(`${fmtDate(p.date)}: ${fCLP(p.amount)}`);
      }));
      const hormiga=Object.values(extraAgg).filter(e=>e.count>=2).sort((a,b)=>b.total-a.total);
      if(hormiga.length){
        c+="\n### GASTOS RECURRENTES NO PLANIFICADOS AGRUPADOS (posible gasto hormiga — analiza si conviene comprar todo junto en vez de de a poco):\n";
        hormiga.forEach(e=>c+=`- ${e.name}: ${e.count} pagos, total ${fCLP(e.total)} · detalle: ${e.dates.join(", ")}\n`);
      }
    }
    if(useProf&&s.profile){
      const p=s.profile;c+="\n### PERFIL:\n";
      if(p.name)c+=`Nombre: ${p.name}\n`;
      if(p.incomes?.length){c+=`Ingresos mensuales: ${fCLP(sumInc(p))}\n`;p.incomes.forEach(i=>c+=`- ${i.type} "${i.name}": ${fCLP(i.amount)} ${i.frequency}\n`);}
      const av=p.availability;
      c+=`Horas libres/semana: ${av.hoursPerWeek}\n`;
      if(av.canWork)c+="- Disponible para trabajar\n";if(av.studying)c+="- Estudiando\n";if(av.entrepreneur)c+="- Quiere emprender\n";
      if(av.notes)c+=`Situación: ${av.notes}\n`;
      if(p.goals?.length){c+="Objetivos:\n";p.goals.forEach(g=>c+=`- ${g.name}: ${fCLP(g.targetAmount)} (${g.priority}, plazo: ${g.deadline||"sin plazo"})\n`);}
    }
    return c;
  };
  const ASKS={ahorro:"Analiza mi presupuesto y gastos reales para crear un plan de ahorro. ¿Cuánto puedo ahorrar? ¿Dónde recortar? ¿Cómo estructurar ahorros en Chile (AFP cuenta 2, depósito a plazo, FFMM)? Dame montos y porcentajes concretos.",diagnostico:"Haz un diagnóstico completo. ¿Mis gastos fijos son razonables para Chile? ¿Dónde gasto de más vs lo normal? ¿Qué está bien distribuido? Sé específico por categoría.",supervivencia:"Necesito reducir gastos al mínimo. ¿Cuáles son absolutamente esenciales? ¿Qué eliminar primero? Dame un presupuesto de supervivencia realista para Chile.",ingresos:"Con mi perfil y disponibilidad de tiempo, ¿qué formas de generar ingresos extra son más realistas en Chile? Considera Mercado Libre, Workana, Yapo, servicios locales. Incluye inversión inicial y tiempo hasta primer ingreso.",gestion:"Aplica el método 50/30/20 a mi situación real. ¿Estoy bien distribuido? ¿Cómo debería quedar mi presupuesto óptimo? Dame montos concretos en CLP.",libre:cQ||"Dame consejos financieros para mi situación."};
  const callAI=async()=>{
    if(!apiKey){alert("Agrega tu API key de Anthropic en Perfil primero.");return;}
    if(plan==="libre"&&!cQ.trim())return;
    setLoading(true);setResp("");
    try{
      const r=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json","x-api-key":apiKey,"anthropic-version":"2023-06-01","anthropic-dangerous-direct-browser-access":"true"},body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:1000,system:"Eres un asesor financiero personal experto en la economía chilena 2025. Consejos prácticos, directos y accionables. Conoces AFP, Fonasa/Isapre, becas MINEDUC, precios en Chile. Usas método 50/30/20 y literatura económica real. Si ves la sección GASTOS RECURRENTES NO PLANIFICADOS AGRUPADOS, coméntala explícitamente: son compras chicas y repetidas (gasto hormiga) — analiza si comprar en volumen o cambiar de lugar convendría más. Respondes en español chileno con estructura clara y emojis. Eres honesto y realista. Finaliza siempre con sección 'Próximos pasos concretos' enumerados.",messages:[{role:"user",content:`${ASKS[plan]}${ctx()}`}]})});
      const data=await r.json();
      const text=data.content?.[0]?.text||"Sin respuesta.";
      setResp(text);d({t:"ADAI",d:{planType:plan,response:text}});
    }catch(e){setResp("❌ Error de conexión. Verifica tu API key e internet.");}
    setLoading(false);
  };
  const expPDF=()=>{if(!resp)return;const pi=PLANS.find(p=>p.id===plan);const html=`<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Plan IA</title><style>body{font-family:system-ui,sans-serif;padding:2rem;max-width:750px;margin:0 auto;line-height:1.75;color:#1e293b}h1{color:#7C3AED}pre{white-space:pre-wrap;font-family:inherit;font-size:13px}.meta{color:#64748b;font-size:11px;margin-bottom:2rem}</style></head><body><h1>${pi?.e} ${pi?.label||"Plan IA"}</h1><div class="meta">Generado el ${new Date().toLocaleDateString("es-CL")} con GestorGastos</div><pre>${resp}</pre></body></html>`;const w=window.open("","_blank");if(w){w.document.write(html);w.document.close();setTimeout(()=>w.print(),400);}};

  return <div>
    <SH title="Planificación IA" sub="Asesor financiero para la realidad chilena" t={t}/>
    {!apiKey&&guideView==="quick"&&<QuickGuide t={t} apiKey={apiKey} setApiKey={setApiKey}
      onDetail={()=>setGuideView("detail")}
      onSkip={()=>setGuideView(null)}
      onNeverShow={()=>{d({t:"SS",k:"iaGuideDismissed",v:true});setGuideView(null);}}/>}
    {!apiKey&&guideView==="detail"&&<DetailGuide t={t} apiKey={apiKey} setApiKey={setApiKey}
      onBack={()=>setGuideView("quick")}
      onSkip={()=>setGuideView(null)}
      onNeverShow={()=>{d({t:"SS",k:"iaGuideDismissed",v:true});setGuideView(null);}}/>}
    {!apiKey&&guideView===null&&<MinimalKeyPrompt t={t} apiKey={apiKey} setApiKey={setApiKey} onShowGuide={()=>setGuideView("quick")}/>}
    {addr&&budget>0&&<div style={{padding:"0.55rem 0.75rem",background:t.pri+"12",borderRadius:"0.5rem",marginBottom:"1rem",fontSize:"0.82rem",color:t.muted,border:`1px solid ${t.pri}30`}}>
      Analizando: <strong style={{color:t.text}}>{addr.name}</strong> · Presupuesto: <strong style={{color:t.pri}}>{fCLP(budget)}/mes</strong>
    </div>}
    <CCard ch={<div style={{display:"flex",alignItems:"center",gap:"0.5rem"}}>
      <input type="checkbox" checked={useProf} onChange={e=>setUseProf(e.target.checked)} style={{accentColor:t.pri}}/>
      <label style={{fontSize:"0.83rem",color:t.text,cursor:"pointer"}}>Incluir mi perfil de ingresos y disponibilidad en el análisis</label>
    </div>} t={t} sx={{marginBottom:"1rem"}}/>
    <CCard ch={<>
      <h4 style={{margin:"0 0 0.65rem",color:t.text,fontSize:"0.9rem",fontWeight:700}}>Tipo de análisis</h4>
      <div style={{display:"flex",flexDirection:"column",gap:"0.35rem"}}>
        {PLANS.map(p=><button key={p.id} onClick={()=>setPlan(p.id)} style={{display:"flex",alignItems:"center",gap:"0.75rem",padding:"0.6rem 0.85rem",background:plan===p.id?`${t.pri}15`:t.card2,border:`2px solid ${plan===p.id?t.pri:t.border}`,borderRadius:"0.5rem",cursor:"pointer",textAlign:"left",transition:"all .15s"}}>
          <span style={{fontSize:"1.05rem",lineHeight:1}}>{p.e}</span>
          <div><div style={{fontWeight:plan===p.id?700:500,fontSize:"0.84rem",color:plan===p.id?t.pri:t.text}}>{p.label}</div><div style={{color:t.muted,fontSize:"0.72rem"}}>{p.desc}</div></div>
        </button>)}
      </div>
      {plan==="libre"&&<div style={{marginTop:"0.75rem"}}><TA val={cQ} onChange={setCQ} ph="¿Cómo ahorro para una GPU en 6 meses? ¿Qué emprendimiento me conviene con 10 horas semanales libres?" rows={3} t={t}/></div>}
    </>} t={t} sx={{marginBottom:"1rem"}}/>
    <Btn onClick={callAI} dis={loading||!apiKey||(plan==="libre"&&!cQ.trim())} full icon={loading?<RefreshCw size={15} style={{animation:"spin 1s linear infinite"}}/>:<Sparkles size={15}/>} sx={{marginBottom:"1rem",padding:"0.7rem",fontSize:"0.9rem"}}>
      {loading?"Generando análisis...":"Generar plan con IA"}
    </Btn>
    {resp&&<CCard ch={<>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"0.75rem"}}>
        <h4 style={{margin:0,color:t.text,fontSize:"0.9rem",fontWeight:700}}>🧠 {PLANS.find(p=>p.id===plan)?.label}</h4>
        <Btn onClick={expPDF} v="ghost" sz="sm" icon={<FileText size={12}/>}>PDF</Btn>
      </div>
      <div style={{color:t.text,fontSize:"0.86rem",lineHeight:1.8,whiteSpace:"pre-wrap"}}>{resp}</div>
    </>} t={t} sx={{marginBottom:"1rem"}}/>}
    {s.aiHistory.length>0&&<CCard ch={<>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"0.65rem"}}>
        <h4 style={{margin:0,color:t.text,fontSize:"0.9rem",fontWeight:700}}>📜 Historial ({s.aiHistory.length})</h4>
        <div style={{display:"flex",gap:"0.35rem"}}>
          <Btn onClick={()=>setShowHist(p=>!p)} v="ghost" sz="sm">{showHist?"Ocultar":"Mostrar"}</Btn>
          <Btn onClick={()=>d({t:"CLAI"})} v="ghost" sz="sm" sx={{color:t.err}}>Limpiar</Btn>
        </div>
      </div>
      {showHist&&s.aiHistory.map(h=><div key={h.id} style={{padding:"0.65rem 0.75rem",background:t.card2,borderRadius:"0.5rem",marginBottom:"0.4rem"}}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:"0.35rem"}}>
          <Badge ch={PLANS.find(p=>p.id===h.planType)?.label||h.planType} color={t.pri}/>
          <span style={{color:t.dim,fontSize:"0.69rem"}}>{h.date}</span>
        </div>
        <div style={{color:t.muted,fontSize:"0.77rem",maxHeight:"55px",overflow:"hidden",WebkitMaskImage:"linear-gradient(to bottom,black 40%,transparent)"}}>{h.response?.slice(0,200)}</div>
        <button onClick={()=>setResp(h.response)} style={{background:"none",border:"none",cursor:"pointer",color:t.pri,fontSize:"0.77rem",marginTop:"0.35rem"}}>Ver completo →</button>
      </div>)}
    </>} t={t}/>}
    <style>{"@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}"}</style>
  </div>;
}

/* ══ APP ROOT ════════════════════════════════════════════ */
export default function App(){
  const[s,d]=useReducer(red,INIT);
  const[view,setView]=useState("hoy");
  const[ready,setReady]=useState(false);
  const saveRef=useRef();
  useEffect(()=>{ldata().then(data=>{if(data)d({t:"LOAD",d:data});setReady(true);});},[]);
  useEffect(()=>{if(!ready)return;clearTimeout(saveRef.current);saveRef.current=setTimeout(()=>sdata(s),900);return()=>clearTimeout(saveRef.current);},[s,ready]);
  const t=useMemo(()=>mkT(s.settings.dark),[s.settings.dark]);
  if(!ready)return <div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",background:"#0D1117"}}>
    <div style={{textAlign:"center"}}><RefreshCw size={28} color="#7C3AED" style={{animation:"spin 1s linear infinite",marginBottom:"0.75rem"}}/><p style={{color:"#7D8590",fontSize:"0.9rem"}}>Cargando...</p></div>
    <style>{"@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}"}</style>
  </div>;
  const VIEWS={hoy:<HoyView s={s} d={d} t={t} setView={setView}/>,plantilla:<PlantillaView s={s} d={d} t={t}/>,historial:<HistorialView s={s} d={d} t={t}/>,analisis:<AnalisisView s={s} d={d} t={t}/>,perfil:<PerfilView s={s} d={d} t={t}/>,ia:<IAView s={s} d={d} t={t}/>};
  return <div style={{minHeight:"100vh",background:t.bg,color:t.text,transition:"background .2s,color .2s",fontFamily:"system-ui,-apple-system,sans-serif"}}>
    <Navbar view={view} setView={setView} s={s} d={d} t={t}/>
    <div style={{maxWidth:"920px",margin:"0 auto",padding:"1.5rem 1rem 4rem"}}>{VIEWS[view]||VIEWS.hoy}</div>
    {!s.settings.tutDone&&<Tutorial onDone={()=>d({t:"TUT_DONE"})} t={t}/>}
  </div>;
}
