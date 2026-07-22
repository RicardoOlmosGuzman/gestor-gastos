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
const tmplTotal=addr=>(addr?.template?.categories||[]).reduce((s,c)=>(c.items||[]).reduce((ss,it)=>ss+(it.amount||0),s),0);
const allTmplItems=addr=>(addr?.template?.categories||[]).flatMap(c=>(c.items||[]).map(it=>({...it,catName:c.name,catIcon:c.icon,catColor:c.color,catId:c.id})));
const mPays=(addr,y,m)=>(addr?.months||[]).find(x=>x.year===y&&x.month===m)?.payments||[];
const totalPaid=(addr,y,m)=>mPays(addr,y,m).reduce((s,p)=>s+(p.amount||0),0);

// Ítems pendientes de la plantilla (no pagados este mes)
const pendingTmpl=(addr,y,m)=>{
  const pays=mPays(addr,y,m);
  return allTmplItems(addr).filter(it=>!pays.find(p=>p.templateItemId===it.id));
};
// Ítems de plantilla ya pagados este mes
const paidTmpl=(addr,y,m)=>{
  const pays=mPays(addr,y,m);
  return allTmplItems(addr).map(it=>{const pay=pays.find(p=>p.templateItemId===it.id);return pay?{...it,pay}:null;}).filter(Boolean);
};
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
  profile:{name:"",incomes:[],availability:{hoursPerWeek:0,canWork:true,studying:false,entrepreneur:false,notes:""},goals:[]},
  settings:{dark:false,alertPct:25,linkProf:false,tutDone:false},
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
    case"AA":{const n={id:uid(),name:a.name,address:a.addr||"",template:{categories:[]},months:[]};return{...s,addresses:[...s.addresses,n],selAddr:n.id};}
    case"UA":return{...s,addresses:upA(s.addresses,a.id,x=>({...x,name:a.name,address:a.addr||""}))};
    case"DA":return{...s,addresses:s.addresses.filter(x=>x.id!==a.id),selAddr:s.selAddr===a.id?null:s.selAddr};
    // Plantilla — categorías
    case"ATC":return{...s,addresses:upA(s.addresses,a.aid,ad=>({...ad,template:{...ad.template,categories:[...(ad.template?.categories||[]),{id:uid(),name:a.name,icon:a.icon,color:a.color,items:[]}]}}))};
    case"UTC":return{...s,addresses:upA(s.addresses,a.aid,ad=>({...ad,template:{...ad.template,categories:(ad.template?.categories||[]).map(c=>c.id===a.cid?{...c,name:a.name,icon:a.icon,color:a.color}:c)}}))};
    case"DTC":return{...s,addresses:upA(s.addresses,a.aid,ad=>({...ad,template:{...ad.template,categories:(ad.template?.categories||[]).filter(c=>c.id!==a.cid)}}))};
    // Plantilla — ítems
    case"ATI":return{...s,addresses:upA(s.addresses,a.aid,ad=>({...ad,template:{...ad.template,categories:(ad.template?.categories||[]).map(c=>c.id===a.cid?{...c,items:[...(c.items||[]),{id:uid(),name:a.name,amount:a.amount||0,isVariable:a.isVariable||false}]}:c)}}))};
    case"UTI":return{...s,addresses:upA(s.addresses,a.aid,ad=>({...ad,template:{...ad.template,categories:(ad.template?.categories||[]).map(c=>c.id===a.cid?{...c,items:(c.items||[]).map(it=>it.id===a.iid?{...it,name:a.name,amount:a.amount||0,isVariable:a.isVariable||false}:it)}:c)}}))};
    case"DTI":return{...s,addresses:upA(s.addresses,a.aid,ad=>({...ad,template:{...ad.template,categories:(ad.template?.categories||[]).map(c=>c.id===a.cid?{...c,items:(c.items||[]).filter(it=>it.id!==a.iid)}:c)}}))};
    // Pagos del mes
    case"ADD_PAY":return{...s,addresses:upA(s.addresses,a.aid,ad=>{
      const{months,mid}=gOrM(ad,s.selYear,s.selMonth);
      return{...ad,months:months.map(m=>m.id===mid?{...m,payments:[...(m.payments||[]),{id:uid(),...a.d}]}:m)};
    })};
    case"DEL_PAY":return{...s,addresses:upA(s.addresses,a.aid,ad=>mapM(ad,s.selYear,s.selMonth,m=>({...m,payments:(m.payments||[]).filter(p=>p.id!==a.pid)})))};
    // Perfil
    case"SP":return{...s,profile:{...s.profile,...a.d}};
    case"AINC":return{...s,profile:{...s.profile,incomes:[...s.profile.incomes,{id:uid(),...a.d}]}};
    case"UINC":return{...s,profile:{...s.profile,incomes:s.profile.incomes.map(i=>i.id===a.id?{...i,...a.d}:i)}};
    case"DINC":return{...s,profile:{...s.profile,incomes:s.profile.incomes.filter(i=>i.id!==a.id)}};
    case"AG":return{...s,profile:{...s.profile,goals:[...s.profile.goals,{id:uid(),...a.d}]}};
    case"DG":return{...s,profile:{...s.profile,goals:s.profile.goals.filter(g=>g.id!==a.id)}};
    case"SAV":return{...s,profile:{...s.profile,availability:{...s.profile.availability,...a.d}}};
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
    onFocus={e=>{setDisp("");setTimeout(()=>e.target.select(),10);}}
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
function PayModal({state,t,onClose,onSave,catNames=[]}){
  const[name,setName]=useState(state.itemName||"");
  const[amount,setAmount]=useState(state.estimated||0);
  const[date,setDate]=useState(state.date||todayStr());
  const[note,setNote]=useState("");const[method,setMethod]=useState("Transferencia");
  const[catName,setCatName]=useState(state.catName||"");
  const isExtra=state.isExtra!==false;
  return <Modal title={isExtra?"⚡ Registrar Pago Extra":"✓ Marcar como Pagado"} onClose={onClose} t={t} ch={<>
    {!isExtra&&<div style={{display:"flex",gap:"0.5rem",padding:"0.6rem 0.75rem",background:t.ok+"12",borderRadius:"0.5rem",marginBottom:"0.85rem",fontSize:"0.82rem",color:t.muted,alignItems:"center"}}>
      <CheckCircle2 size={14} color={t.ok}/><span>Ítem: <strong style={{color:t.text}}>{state.itemName}</strong> · Estimado: <strong style={{color:t.text}}>{fCLP(state.estimated)}</strong></span>
    </div>}
    {isExtra&&<div style={{display:"flex",gap:"0.5rem",padding:"0.6rem 0.75rem",background:t.warn+"12",borderRadius:"0.5rem",marginBottom:"0.85rem",fontSize:"0.82rem",color:t.muted,alignItems:"center"}}>
      <Zap size={14} color={t.warn}/><span>Se suma al total del mes. Puedes asociarlo a una categoría para el análisis.</span>
    </div>}
    <Fld label="¿Qué pagaste?" ch={<TI val={name} onChange={setName} ph="Luz, Almacén, Netflix, Uber..." t={t}/>} t={t}/>
    <Fld label={isExtra?"Monto pagado":"Monto real (puede diferir del estimado)"} ch={<MoneyInput val={amount} onChange={setAmount} t={t} af/>} t={t}/>
    {isExtra&&catNames.length>0&&<Fld label="Categoría (opcional, para análisis)" ch={
      <select value={catName} onChange={e=>setCatName(e.target.value)} style={{width:"100%",padding:"0.5rem 0.75rem",background:t.card2,border:`1px solid ${t.border}`,borderRadius:"0.5rem",color:t.text,fontSize:"0.875rem",outline:"none"}}>
        <option value="">Sin categoría</option>
        {catNames.map(c=><option key={c} value={c}>{c}</option>)}
      </select>
    } t={t}/>}
    <Fld label="Fecha" ch={<TI val={date} onChange={setDate} type="date" t={t}/>} t={t}/>
    <Fld label="Forma de pago" ch={<Sel val={method} onChange={setMethod} opts={PAY_METHODS} t={t}/>} t={t}/>
    <Fld label="Nota (opcional)" ch={<TI val={note} onChange={setNote} ph="Info extra..." t={t}/>} t={t}/>
    <Btn onClick={()=>name&&amount&&onSave({name,amount,date,method,note,catName:catName||state.catName||"",templateItemId:state.itemId||null,isExtra:state.isExtra!==false})} dis={!name||!amount} full icon={<Check size={14}/>}>
      {isExtra?"Registrar pago ⚡":"Confirmar pago ✓"}
    </Btn>
  </>}/>;
}

/* ══ TUTORIAL ════════════════════════════════════════════ */
function Tutorial({onDone,t}){
  const[step,setStep]=useState(0);
  const STEPS=[
    {e:"👋",title:"¡Hola! Bienvenido a GestorGastos",desc:"Esta app está hecha para ser simple en el uso diario.\n\nEl objetivo: abres la app, registras lo que pagaste, la cierras.\n\nTe mostramos cómo funciona todo. Puedes volver aquí con el botón ❓ del menú.",hint:""},
    {e:"🏠",title:"Inicio — Tu centro de operaciones",desc:"Aquí está todo lo que necesitas en el día a día:\n\n• Cuánto llevas pagado este mes vs tu presupuesto\n• Tus gastos fijos pendientes de pagar (un tap para marcarlos)\n• Botón grande para registrar cualquier gasto en segundos\n• Sugerencias inteligentes basadas en tus hábitos",hint:"💡 Esta es la pantalla que abrirás todos los días. El flujo es: abrir → pagar → cerrar."},
    {e:"📋",title:"Gastos Fijos — Configuras una vez",desc:"Aquí defines los gastos que tienes TODOS los meses:\n• Arriendo: $350.000 (fijo)\n• Luz: ~$40.000 (variable, cambia cada mes)\n• Agua, internet, suscripciones...\n\nLos configuras UNA VEZ y aparecen como pendientes automáticamente cada mes. Nunca más tienes que volver a ingresarlos.",hint:"💡 Si un monto varía (luz, agua), márcalo como 'variable'. El monto que pones es solo una estimación."},
    {e:"📅",title:"Historial — El mes en detalle",desc:"El calendario mensual muestra todo:\n• Los días con pagos tienen puntos de color\n• Verde = gasto fijo pagado · Naranja = gasto extra\n• Click en un día para ver o agregar pagos\n• Navega entre meses con las flechas ‹ ›",hint:""},
    {e:"📊",title:"Análisis — Entiende tus gastos",desc:"Para cuando quieres ver el panorama:\n• Gráfico de presupuesto vs lo que realmente pagaste\n• Distribución por categoría\n• Evolución mes a mes\n\nExporta a Excel o imprime como PDF.",hint:"💡 Úsalo una vez al mes. El resto del tiempo, usa solo la pantalla de Inicio."},
    {e:"👤",title:"Perfil — Tus datos e ingresos",desc:"Configura:\n• Tu nombre (aparece en el saludo)\n• Fuentes de ingreso (sueldo, beca, etc.)\n• Objetivos de ahorro\n• API key para la IA\n\n¡Importante! Aquí también puedes EXPORTAR tus datos como respaldo JSON e IMPORTARLOS si cambias de dispositivo.",hint:"💡 Exporta tus datos regularmente como respaldo. Los datos viven en tu navegador."},
    {e:"🧠",title:"Plan IA — Tu asesor financiero",desc:"Usando tu presupuesto y gastos reales, la IA genera:\n💰 Plan de ahorro personalizado\n🔍 Diagnóstico categoría por categoría\n🆘 Presupuesto de supervivencia\n🚀 Ideas de ingresos extra para Chile\n📊 Análisis del método 50/30/20",hint:"🔑 Necesitas API key gratuita de Anthropic (console.anthropic.com). La agregas en Perfil."},
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
          <span style={{color:t.muted,fontSize:"0.83rem",fontWeight:500,whiteSpace:"nowrap"}}>{it.isVariable?"~":""}{fCLP(it.amount)}</span>
          <Btn onClick={()=>setPayModal({date:todayStr(),isExtra:false,itemId:it.id,itemName:it.name,catName:it.catName,estimated:it.amount})} v="pri" sz="sm">Pagar</Btn>
        </div>)}
      </>} t={t} sx={{marginBottom:"1rem"}}/>}

      {/* Pagado */}
      {(paidItems.length>0||extras.length>0)&&<CCard ch={<>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"0.75rem"}}>
          <h3 style={{margin:0,color:t.text,fontSize:"0.93rem",fontWeight:700}}>✅ Pagado este mes</h3>
          <Badge ch={fCLP(paid)} color={t.ok}/>
        </div>
        {paidItems.map(it=><div key={it.id} style={{display:"flex",alignItems:"center",gap:"0.6rem",padding:"0.45rem 0.7rem",background:t.ok+"10",borderRadius:"0.5rem",marginBottom:"0.3rem",border:`1px solid ${t.ok}30`}}>
          <CheckCircle2 size={15} color={t.ok}/>
          <span style={{fontSize:"0.9rem",lineHeight:1}}>{it.catIcon||"📦"}</span>
          <div style={{flex:1}}>
            <div style={{color:t.text,fontSize:"0.83rem",fontWeight:500}}>{it.name}</div>
            <div style={{color:t.dim,fontSize:"0.7rem"}}>{fmtDate(it.pay.date)}{it.pay.method?` · ${it.pay.method}`:""}</div>
          </div>
          <div style={{textAlign:"right"}}>
            <div style={{color:t.ok,fontSize:"0.83rem",fontWeight:700}}>{fCLP(it.pay.amount)}</div>
            {it.pay.amount!==it.amount&&<div style={{color:t.dim,fontSize:"0.68rem"}}>est. {fCLP(it.amount)}</div>}
          </div>
          <button onClick={()=>setCfm({msg:`¿Desmarcar "${it.name}" como pagado?`,ok:()=>{d({t:"DEL_PAY",aid:addr.id,pid:it.pay.id});setCfm(null);}})} style={{background:"none",border:"none",cursor:"pointer",color:t.dim,padding:"0.2rem",display:"flex"}}><RotateCcw size={12}/></button>
        </div>)}
        {extras.map(p=><div key={p.id} style={{display:"flex",alignItems:"center",gap:"0.6rem",padding:"0.45rem 0.7rem",background:t.warn+"10",borderRadius:"0.5rem",marginBottom:"0.3rem",border:`1px solid ${t.warn}30`}}>
          <Zap size={13} color={t.warn}/>
          <div style={{flex:1}}>
            <div style={{color:t.text,fontSize:"0.83rem",fontWeight:500}}>{p.name}</div>
            <div style={{color:t.dim,fontSize:"0.7rem"}}>{p.catName?`${p.catName} · `:""}{fmtDate(p.date)}</div>
          </div>
          <span style={{color:t.warn,fontSize:"0.83rem",fontWeight:700}}>{fCLP(p.amount)}</span>
          <button onClick={()=>setCfm({msg:`¿Eliminar "${p.name}"?`,ok:()=>{d({t:"DEL_PAY",aid:addr.id,pid:p.id});setCfm(null);}})} style={{background:"none",border:"none",cursor:"pointer",color:t.err,padding:"0.2rem",display:"flex"}}><Trash2 size={12}/></button>
        </div>)}
      </>} t={t} sx={{marginBottom:"1rem"}}/>}

      {/* Accesos rápidos */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:"0.5rem"}}>
        {[{id:"plantilla",e:"📋",lb:"Gastos\nFijos"},{id:"historial",e:"📅",lb:"Historial"},{id:"analisis",e:"📊",lb:"Análisis"},{id:"perfil",e:"👤",lb:"Perfil"},{id:"ia",e:"🧠",lb:"Plan IA"}].map(q=><CCard key={q.id} ch={<><div style={{fontSize:"1.3rem",marginBottom:"0.3rem"}}>{q.e}</div><div style={{color:t.text,fontWeight:700,fontSize:"0.72rem",whiteSpace:"pre-line",lineHeight:1.3}}>{q.lb}</div></>} t={t} onClick={()=>setView(q.id)} sx={{cursor:"pointer",padding:"0.75rem",textAlign:"center"}}/>)}
      </div>
    </>}

    {payModal&&<PayModal state={payModal} t={t} onClose={()=>setPayModal(null)} catNames={catNames} onSave={dt=>{d({t:"ADD_PAY",aid:addr.id,d:dt});setPayModal(null);}}/>}
    {cfm&&<Cfm msg={cfm.msg} onOk={cfm.ok} onCancel={()=>setCfm(null)} t={t}/>}
  </div>;
}

/* ══ PLANTILLA DE GASTOS FIJOS ═══════════════════════════ */
function PlantillaView({s,d,t}){
  const[aM,setAM]=useState(null);const[cM,setCM]=useState(null);const[iM,setIM]=useState(null);const[cfm,setCfm]=useState(null);
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
        <Btn onClick={()=>setCM("add")} icon={<Plus size={14}/>} sz="sm">Categoría</Btn>
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
              </div>
              <span style={{color:it.isVariable?t.info:t.muted,fontSize:"0.84rem",fontWeight:600}}>{it.isVariable?"~":""}{fCLP(it.amount)}</span>
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
  const[name,setName]=useState(init.name||"");const[amount,setAmount]=useState(init.amount||0);const[isVariable,setIsVariable]=useState(init.isVariable||false);
  return <>
    <Fld label="Nombre del gasto" ch={<TI val={name} onChange={setName} ph="Arriendo, Luz, Internet, Agua, Netflix..." t={t}/>} t={t}/>
    <Fld label={isVariable?"Monto estimado (referencia)":"Monto fijo mensual"} ch={<MoneyInput val={amount} onChange={setAmount} t={t} af/>} t={t}/>
    <label style={{display:"flex",alignItems:"flex-start",gap:"0.65rem",cursor:"pointer",padding:"0.65rem 0.75rem",background:t.card2,borderRadius:"0.5rem",marginBottom:"0.9rem"}}>
      <input type="checkbox" checked={isVariable} onChange={e=>setIsVariable(e.target.checked)} style={{width:"16px",height:"16px",accentColor:t.info,marginTop:"0.15rem"}}/>
      <div>
        <div style={{color:t.text,fontSize:"0.84rem",fontWeight:600}}>Monto variable</div>
        <div style={{color:t.dim,fontSize:"0.73rem"}}>Marca esto si el monto cambia cada mes (luz, agua, teléfono). El monto que pones aquí es solo una estimación que puedes ajustar cuando pagas.</div>
      </div>
    </label>
    <Btn onClick={()=>name&&onSave({name,amount,isVariable})} dis={!name} full icon={<Check size={14}/>}>{init.name?"Actualizar":"Agregar a plantilla"}</Btn>
  </>;
}

/* ══ HISTORIAL — CALENDARIO MENSUAL ═════════════════════ */
function HistorialView({s,d,t}){
  const[selDay,setSelDay]=useState(null);const[payModal,setPayModal]=useState(null);const[cfm,setCfm]=useState(null);
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
      <StCard label="Pendiente" val={fCLP(Math.max(0,budget-paidItems.reduce((s,i)=>s+(i.pay?.amount||0),0)))} color={t.err} icon={<Clock size={16} color={t.err}/>} sub={`${pending.length} ítems`} t={t}/>
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
              <div style={{flex:1,minWidth:0}}><div style={{color:t.text,fontSize:"0.81rem",fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{it.name}</div><div style={{color:t.dim,fontSize:"0.68rem"}}>{it.catName}{it.isVariable?" · variable":""}</div></div>
              <span style={{color:t.dim,fontSize:"0.78rem",whiteSpace:"nowrap"}}>{it.isVariable?"~":""}{fCLP(it.amount)}</span>
              <Btn onClick={()=>setPayModal({date:todayStr(),isExtra:false,itemId:it.id,itemName:it.name,catName:it.catName,estimated:it.amount})} v="pri" sz="sm">Pagar</Btn>
            </div>)}
          </div>
        </>} t={t}/>}
        {(paidItems.length>0||extras.length>0)&&<CCard ch={<>
          <h4 style={{margin:"0 0 0.65rem",color:t.text,fontSize:"0.88rem",fontWeight:700}}>✅ Pagado</h4>
          <div style={{maxHeight:"200px",overflowY:"auto",display:"flex",flexDirection:"column",gap:"0.3rem"}}>
            {paidItems.map(it=><div key={it.id} style={{display:"flex",alignItems:"center",gap:"0.45rem",padding:"0.4rem 0.6rem",background:t.ok+"10",borderRadius:"0.45rem",border:`1px solid ${t.ok}30`}}>
              <CheckCircle2 size={13} color={t.ok}/>
              <div style={{flex:1,minWidth:0}}><div style={{color:t.text,fontSize:"0.8rem",fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{it.name}</div><div style={{color:t.dim,fontSize:"0.67rem"}}>{fmtDate(it.pay.date)}</div></div>
              <span style={{color:t.ok,fontSize:"0.79rem",fontWeight:700,whiteSpace:"nowrap"}}>{fCLP(it.pay.amount)}</span>
              <button onClick={()=>setCfm({msg:`¿Desmarcar "${it.name}"?`,ok:()=>{d({t:"DEL_PAY",aid:addr.id,pid:it.pay.id});setCfm(null);}})} style={{background:"none",border:"none",cursor:"pointer",color:t.dim,padding:"0.15rem",display:"flex"}}><RotateCcw size={11}/></button>
            </div>)}
            {extras.map(p=><div key={p.id} style={{display:"flex",alignItems:"center",gap:"0.45rem",padding:"0.4rem 0.6rem",background:t.warn+"10",borderRadius:"0.45rem",border:`1px solid ${t.warn}30`}}>
              <Zap size={12} color={t.warn}/>
              <div style={{flex:1,minWidth:0}}><div style={{color:t.text,fontSize:"0.8rem",fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.name}</div><div style={{color:t.dim,fontSize:"0.67rem"}}>{p.catName?`${p.catName} · `:""}{fmtDate(p.date)}</div></div>
              <span style={{color:t.warn,fontSize:"0.79rem",fontWeight:700,whiteSpace:"nowrap"}}>{fCLP(p.amount)}</span>
              <button onClick={()=>setCfm({msg:`¿Eliminar "${p.name}"?`,ok:()=>{d({t:"DEL_PAY",aid:addr.id,pid:p.id});setCfm(null);}})} style={{background:"none",border:"none",cursor:"pointer",color:t.err,padding:"0.15rem",display:"flex"}}><Trash2 size={11}/></button>
            </div>)}
          </div>
        </>} t={t}/>}
        {pending.length===0&&paidItems.length===0&&extras.length===0&&<CCard ch={<Empty icon={<Calendar size={30}/>} title="Sin actividad este mes" sub='Marca gastos como pagados o agrega un pago extra.' t={t}/>} t={t}/>}
      </div>
    </div>
    {payModal&&<PayModal state={payModal} t={t} onClose={()=>setPayModal(null)} catNames={catNames} onSave={dt=>{d({t:"ADD_PAY",aid:addr.id,d:dt});setPayModal(null);}}/>}
    {cfm&&<Cfm msg={cfm.msg} onOk={cfm.ok} onCancel={()=>setCfm(null)} t={t}/>}
  </div>;
}

/* ══ ANÁLISIS ════════════════════════════════════════════ */
function AnalisisView({s,t}){
  const addr=s.addresses.find(a=>a.id===s.selAddr);
  const budget=addr?tmplTotal(addr):0;
  const recentMonths=useMemo(()=>{if(!addr)return[];const months=[];for(let i=11;i>=0;i--){let m=CM-i,y=CY;if(m<0){m+=12;y--;}months.push({year:y,month:m});}return months;},[addr]);
  const[sel,setSel]=useState(()=>new Set(recentMonths.map(m=>`${m.year}-${m.month}`)));
  const toggle=k=>setSel(p=>{const n=new Set(p);n.has(k)?n.delete(k):n.add(k);return n;});
  const selM=recentMonths.filter(m=>sel.has(`${m.year}-${m.month}`));
  const barData=selM.map(m=>{const p=addr?totalPaid(addr,m.year,m.month):0;const byCat={};(addr?mPays(addr,m.year,m.month):[]).forEach(pay=>{if(pay.catName)byCat[pay.catName]=(byCat[pay.catName]||0)+(pay.amount||0);});return{name:`${MESES[m.month].substr(0,3)} ${m.year}`,Presupuesto:budget,Pagado:p,...byCat};});
  const catAgg={};selM.forEach(m=>(addr?mPays(addr,m.year,m.month):[]).forEach(p=>{if(p.catName)catAgg[p.catName]=(catAgg[p.catName]||0)+(p.amount||0);}));
  const pieData=Object.entries(catAgg).map(([name,value])=>({name,value})).sort((a,b)=>b.value-a.value);
  const totalPaidAll=selM.reduce((s,m)=>s+(addr?totalPaid(addr,m.year,m.month):0),0);

  const exportExcel=()=>{if(!addr||!selM.length)return;const wb=XLSX.utils.book_new();const sumRows=selM.map(m=>({"Período":`${MESES[m.month]} ${m.year}`,"Presupuestado":budget,"Pagado":totalPaid(addr,m.year,m.month),"Diferencia":totalPaid(addr,m.year,m.month)-budget}));XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(sumRows),"Resumen");selM.forEach(m=>{const rows=(addr?mPays(addr,m.year,m.month):[]).map(p=>({Nombre:p.name,Monto:p.amount,Fecha:p.date,Categoría:p.catName||"",Tipo:p.isExtra?"Extra":"Fijo",Método:p.method||""}));if(rows.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),`${MESES[m.month].substr(0,3)} ${m.year}`.replace(/[:\\/\[\]*?]/g,""));});XLSX.writeFile(wb,`gastos_${addr.name||"reporte"}_${new Date().toISOString().split("T")[0]}.xlsx`);};

  const exportPDF=()=>{if(!addr||!selM.length)return;const f=n=>fCLP(n);const html=`<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Reporte</title><style>body{font-family:system-ui,sans-serif;padding:2rem;color:#1e293b;font-size:13px}h1{color:#7C3AED;font-size:22px}h2{font-size:14px;color:#334155;margin:1.5rem 0 .5rem;border-bottom:2px solid #e2e8f0;padding-bottom:4px}table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:1rem}th{background:#7C3AED;color:#fff;padding:6px 10px;text-align:left}td{padding:5px 10px;border-bottom:1px solid #f1f5f9}tr:nth-child(even) td{background:#f8fafc}</style></head><body><h1>📊 Reporte — ${addr.name}</h1><p style="color:#64748b;font-size:11px">Generado el ${new Date().toLocaleDateString("es-CL")} · Presupuesto mensual: ${f(budget)}</p><h2>Resumen mensual</h2><table><tr><th>Mes</th><th>Presupuestado</th><th>Pagado</th><th>Diferencia</th></tr>${selM.map(m=>{const p=totalPaid(addr,m.year,m.month);return`<tr><td>${MESES[m.month]} ${m.year}</td><td>${f(budget)}</td><td>${f(p)}</td><td style="color:${p>budget?"#ef4444":"#10b981"}">${p>budget?"+":""}${f(p-budget)}</td></tr>`;}).join("")}<tr style="font-weight:700;background:#ede9fe"><td>TOTAL</td><td>${f(budget*selM.length)}</td><td>${f(totalPaidAll)}</td><td>${f(totalPaidAll-budget*selM.length)}</td></tr></table>${selM.map(m=>{const pays=mPays(addr,m.year,m.month);if(!pays.length)return"";return`<h2>${MESES[m.month]} ${m.year} — ${f(totalPaid(addr,m.year,m.month))}</h2><table><tr><th>Nombre</th><th>Tipo</th><th>Categoría</th><th>Fecha</th><th>Monto</th></tr>${pays.map(p=>`<tr><td>${p.name}</td><td>${p.isExtra?"⚡":"✓"}</td><td>${p.catName||"—"}</td><td>${fmtDate(p.date)}</td><td>${f(p.amount)}</td></tr>`).join("")}</table>`;}).join("")}</body></html>`;const w=window.open("","_blank");if(w){w.document.write(html);w.document.close();setTimeout(()=>w.print(),500);}};

  if(!addr)return <CCard ch={<Empty icon={<Building2 size={36}/>} title="Sin dirección" sub="Configura una dirección en Gastos Fijos para ver el análisis." t={t}/>} t={t}/>;

  return <div>
    <SH title="Análisis de Gastos" sub={`${addr.name} · Presupuesto mensual: ${fCLP(budget)}`} t={t}
      action={<div style={{display:"flex",gap:"0.4rem"}}><Btn onClick={exportExcel} v="ok" sz="sm" icon={<Download size={13}/>}>Excel</Btn><Btn onClick={exportPDF} v="err" sz="sm" icon={<FileText size={13}/>}>PDF</Btn></div>}/>
    <CCard ch={<>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"0.65rem",flexWrap:"wrap",gap:"0.4rem"}}>
        <span style={{color:t.text,fontWeight:700,fontSize:"0.87rem"}}>Meses a analizar ({selM.length}/{recentMonths.length})</span>
        <div style={{display:"flex",gap:"0.35rem"}}>
          <Btn onClick={()=>setSel(new Set(recentMonths.map(m=>`${m.year}-${m.month}`)))} v="ghost" sz="sm">Todos</Btn>
          <Btn onClick={()=>setSel(new Set(recentMonths.slice(-3).map(m=>`${m.year}-${m.month}`)))} v="ghost" sz="sm">Últ. 3</Btn>
          <Btn onClick={()=>setSel(new Set(recentMonths.slice(-6).map(m=>`${m.year}-${m.month}`)))} v="ghost" sz="sm">Últ. 6</Btn>
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
        <ResponsiveContainer width="100%" height={210}><BarChart data={barData} margin={{top:5,right:5,left:0,bottom:5}}><CartesianGrid strokeDasharray="3 3" stroke={t.border}/><XAxis dataKey="name" stroke={t.muted} tick={{fontSize:10,fill:t.muted}}/><YAxis stroke={t.muted} tick={{fontSize:10,fill:t.muted}} tickFormatter={v=>v>=1e6?`${(v/1e6).toFixed(1)}M`:v>=1e3?`${(v/1e3).toFixed(0)}K`:v}/><Tooltip formatter={v=>fCLP(v)} contentStyle={{background:t.card,border:`1px solid ${t.border}`,borderRadius:"0.5rem",color:t.text,fontSize:"0.79rem"}}/><Legend wrapperStyle={{fontSize:"0.75rem"}}/><Bar dataKey="Presupuesto" fill={t.pri+"66"} radius={[3,3,0,0]}/><Bar dataKey="Pagado" fill={t.ok} radius={[3,3,0,0]}/></BarChart></ResponsiveContainer>
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
  const[iM,setIM]=useState(null);const[gM,setGM]=useState(false);const[cfm,setCfm]=useState(null);
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
        {!apiKey&&<p style={{color:t.muted,fontSize:"0.82rem",margin:"0 0 0.75rem",lineHeight:1.5}}>Para usar la IA necesitas una key gratuita de Anthropic. <a href="https://console.anthropic.com" target="_blank" style={{color:t.pri}}>Obtener →</a></p>}
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
      <h4 style={{margin:"0 0 0.5rem",color:t.text,fontSize:"0.9rem",fontWeight:700}}>💾 Mis datos — Respaldo y restauración</h4>
      <p style={{color:t.muted,fontSize:"0.82rem",margin:"0 0 0.85rem",lineHeight:1.5}}>Tus datos viven en <strong style={{color:t.text}}>este navegador</strong>. Expórtalos como respaldo y recupéralos si cambias de dispositivo o borras el caché del navegador.</p>
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

/* ══ IA ══════════════════════════════════════════════════ */
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
  const addr=s.addresses.find(a=>a.id===s.selAddr);
  const budget=addr?tmplTotal(addr):0;
  const recMonths=useMemo(()=>{if(!addr)return[];const months=[];for(let i=2;i>=0;i--){let m=CM-i,y=CY;if(m<0){m+=12;y--;}months.push({year:y,month:m});}return months;},[addr]);
  const ctx=()=>{
    let c="";
    if(addr){
      c+=`\n\n### PLANTILLA DE GASTOS FIJOS:\nPresupuesto mensual total: ${fCLP(budget)}\n`;
      (addr.template?.categories||[]).forEach(cat=>{c+=`\nCategoría ${cat.name}:\n`;(cat.items||[]).forEach(it=>c+=`- ${it.name}: ${fCLP(it.amount)}${it.isVariable?" (variable)":""}\n`);});
      c+="\n### PAGOS REALES RECIENTES:\n";
      recMonths.forEach(m=>{const pays=mPays(addr,m.year,m.month);if(pays.length){c+=`\n${MESES[m.month]} ${m.year} (Total: ${fCLP(totalPaid(addr,m.year,m.month))}, Presupuesto: ${fCLP(budget)}):\n`;pays.forEach(p=>c+=`- ${p.name}: ${fCLP(p.amount)} (${p.isExtra?"extra":"fijo"}${p.catName?`, ${p.catName}`:""})\n`);}});
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
      const r=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json","x-api-key":apiKey,"anthropic-version":"2023-06-01","anthropic-dangerous-direct-browser-access":"true"},body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:1000,system:"Eres un asesor financiero personal experto en la economía chilena 2025. Consejos prácticos, directos y accionables. Conoces AFP, Fonasa/Isapre, becas MINEDUC, precios en Chile. Usas método 50/30/20 y literatura económica real. Respondes en español chileno con estructura clara y emojis. Eres honesto y realista. Finaliza siempre con sección 'Próximos pasos concretos' enumerados.",messages:[{role:"user",content:`${ASKS[plan]}${ctx()}`}]})});
      const data=await r.json();
      const text=data.content?.[0]?.text||"Sin respuesta.";
      setResp(text);d({t:"ADAI",d:{planType:plan,response:text}});
    }catch(e){setResp("❌ Error de conexión. Verifica tu API key e internet.");}
    setLoading(false);
  };
  const expPDF=()=>{if(!resp)return;const pi=PLANS.find(p=>p.id===plan);const html=`<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Plan IA</title><style>body{font-family:system-ui,sans-serif;padding:2rem;max-width:750px;margin:0 auto;line-height:1.75;color:#1e293b}h1{color:#7C3AED}pre{white-space:pre-wrap;font-family:inherit;font-size:13px}.meta{color:#64748b;font-size:11px;margin-bottom:2rem}</style></head><body><h1>${pi?.e} ${pi?.label||"Plan IA"}</h1><div class="meta">Generado el ${new Date().toLocaleDateString("es-CL")} con GestorGastos</div><pre>${resp}</pre></body></html>`;const w=window.open("","_blank");if(w){w.document.write(html);w.document.close();setTimeout(()=>w.print(),400);}};

  return <div>
    <SH title="Planificación IA" sub="Asesor financiero para la realidad chilena" t={t}/>
    {!apiKey&&<CCard ch={<div style={{display:"flex",gap:"0.75rem",alignItems:"flex-start"}}>
      <div style={{fontSize:"1.5rem"}}>🔑</div>
      <div><div style={{color:t.text,fontWeight:700,fontSize:"0.9rem",marginBottom:"0.3rem"}}>Necesitas una API key de Anthropic</div>
      <p style={{color:t.muted,fontSize:"0.82rem",margin:"0 0 0.75rem",lineHeight:1.5}}>Es gratis para empezar. Obtén tu key en <a href="https://console.anthropic.com" target="_blank" style={{color:t.pri}}>console.anthropic.com</a> y agrégala en Perfil.</p></div>
    </div>} t={t} sx={{marginBottom:"1rem"}}/>}
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
  const VIEWS={hoy:<HoyView s={s} d={d} t={t} setView={setView}/>,plantilla:<PlantillaView s={s} d={d} t={t}/>,historial:<HistorialView s={s} d={d} t={t}/>,analisis:<AnalisisView s={s} t={t}/>,perfil:<PerfilView s={s} d={d} t={t}/>,ia:<IAView s={s} d={d} t={t}/>};
  return <div style={{minHeight:"100vh",background:t.bg,color:t.text,transition:"background .2s,color .2s",fontFamily:"system-ui,-apple-system,sans-serif"}}>
    <Navbar view={view} setView={setView} s={s} d={d} t={t}/>
    <div style={{maxWidth:"920px",margin:"0 auto",padding:"1.5rem 1rem 4rem"}}>{VIEWS[view]||VIEWS.hoy}</div>
    {!s.settings.tutDone&&<Tutorial onDone={()=>d({t:"TUT_DONE"})} t={t}/>}
  </div>;
}
