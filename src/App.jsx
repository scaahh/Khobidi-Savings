import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { LineChart, Line, BarChart, Bar, XAxis, YAxis, Tooltip,
  ResponsiveContainer, CartesianGrid, Legend, PieChart, Pie, Cell, AreaChart, Area } from "recharts";

/* ─────────────────────────────────── DESIGN TOKENS */
const G = {
  bg:       "#F7F8F5",
  surface:  "#FFFFFF",
  border:   "#E8EDE6",
  primary:  "#3D6B4F",
  primaryD: "#2A4E38",
  primaryL: "#EBF3EE",
  accent:   "#5A9B72",
  gold:     "#C8973A",
  red:      "#C94040",
  blue:     "#3A72A8",
  purple:   "#6B4F9B",
  text:     "#1A2318",
  textMid:  "#4A5E44",
  textSoft: "#8A9E85",
  shadow:   "0 1px 3px rgba(30,50,25,.08), 0 4px 16px rgba(30,50,25,.05)",
  shadowMd: "0 2px 8px rgba(30,50,25,.10), 0 8px 32px rgba(30,50,25,.08)",
};

/* ─────────────────────────────────── CONSTANTS */
const SEASON_START = 2026;
const LOAN_AMOUNT  = 3750;
const BASE_RATE    = 0.20;
const PENALTY_AMT  = 20;
const MEM_FEE      = 110;
const ADMIN_CREDS  = { "Angella":"1234", "Sandra":"5678", "Fatima":"9012" };
const ADMINS       = Object.keys(ADMIN_CREDS);
const STORE_KEY    = "khobidi_v3";

const TIERS = [
  { label:"Platinum", min:85, rate:0.12, color:"#7B5EA7", bg:"#F3EFF9", icon:"💎" },
  { label:"Gold",     min:70, rate:0.16, color:"#C8973A", bg:"#FDF6E9", icon:"🥇" },
  { label:"Silver",   min:50, rate:0.20, color:"#5A7A8A", bg:"#EEF4F7", icon:"🥈" },
  { label:"Bronze",   min:0,  rate:0.25, color:"#9B6B3A", bg:"#FAF2EA", icon:"🥉" },
];

function getTier(score){ return TIERS.find(t=>score>=t.min)||TIERS[TIERS.length-1]; }

function creditScore(m){
  const weeks=Object.values(m.weeklyContributions);
  const paid=weeks.filter(w=>w.paid).length;
  const total=Math.max(Object.keys(m.weeklyContributions).length,1);
  const payRate=paid/total;
  const penDed=Math.min(m.penalties.length*5,30);
  const loanBonus=Math.min(m.loans.filter(l=>!l.active).length*10,20);
  const feePt=m.membershipFeePaid?5:0;
  const overdueD=m.loans.some(l=>l.active&&loanDueDate(l)&&new Date(loanDueDate(l))<new Date())?20:0;
  return Math.max(0,Math.min(100,Math.round(60+(payRate*25)+loanBonus+feePt-penDed-overdueD)));
}

function getSeasonSundays(){
  const out=[]; let d=new Date(SEASON_START,0,1);
  while(d.getDay()!==0) d.setDate(d.getDate()+1); // find first Sunday
  for(let w=1;w<=52;w++){out.push({week:w,date:new Date(d)});d.setDate(d.getDate()+7);}
  return out;
}
const SUNDAYS=getSeasonSundays();

function loanDueDate(loan){
  if(!loan.issuedDate)return null;
  const d=new Date(loan.issuedDate);d.setMonth(d.getMonth()+3);return d;
}
function loanBalance(loan){
  const total=loan.amount*(1+(loan.rate||BASE_RATE));
  const paid=(loan.repayments||[]).reduce((s,r)=>s+r.amount,0);
  return Math.max(0,total-paid);
}

const INIT_MEMBERS=["Angella","Fatima","Sandra","Netty","Member E","Member F","Member G","Member H","Member I","Member J","Member K","Member L"]
  .map((name,i)=>({id:i+1,name,email:`${name.toLowerCase()}@group.com`,
    membershipFeePaid:false,refundablePaid:false,weeklyContributions:{},loans:[],penalties:[]}));

function fmt(n){return "€"+Number(n).toFixed(2);}
function fd(d){return d?new Date(d).toLocaleDateString("en-GB"):"—";}
function fds(d){return d?new Date(d).toLocaleDateString("en-GB",{day:"numeric",month:"short"}):"—";}
function load(){try{const r=localStorage.getItem(STORE_KEY);if(r)return JSON.parse(r);}catch{}return null;}
function save(d){try{localStorage.setItem(STORE_KEY,JSON.stringify(d));}catch{}}
function mkLog(admin,action,detail,memberId){return{id:Date.now()+Math.random(),ts:new Date().toISOString(),admin,action,detail,memberId};}

/* ─────────────────────────────────── CONFIRM DIALOG */
function ConfirmDialog({msg,sub,onConfirm,onCancel,confirmLabel="Confirm",danger=false}){
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(15,25,18,.55)",backdropFilter:"blur(6px)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:9000}}>
      <div style={{background:G.surface,borderRadius:20,padding:"32px 36px",maxWidth:380,width:"90%",boxShadow:"0 24px 80px rgba(15,25,18,.25)",border:`1px solid ${G.border}`,textAlign:"center"}}>
        <div style={{fontSize:40,marginBottom:12}}>{danger?"⚠️":"✅"}</div>
        <h3 style={{margin:"0 0 8px",fontSize:18,fontWeight:800,color:G.text,fontFamily:"'Georgia',serif"}}>{msg}</h3>
        {sub&&<p style={{margin:"0 0 24px",fontSize:13,color:G.textMid,lineHeight:1.6}}>{sub}</p>}
        <div style={{display:"flex",gap:10,justifyContent:"center"}}>
          <button onClick={onCancel} style={{padding:"10px 22px",borderRadius:10,border:`1.5px solid ${G.border}`,background:G.bg,color:G.textMid,cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:"inherit"}}>Cancel</button>
          <button onClick={onConfirm} style={{padding:"10px 22px",borderRadius:10,border:"none",background:danger?"#C94040":G.primary,color:"white",cursor:"pointer",fontSize:13,fontWeight:700,fontFamily:"inherit",boxShadow:`0 4px 14px ${danger?"rgba(201,64,64,.35)":"rgba(61,107,79,.35)"}`}}>{confirmLabel}</button>
        </div>
      </div>
    </div>
  );
}

/* ─────────────────────────────────── TOAST */
function Toast({msg,type}){
  return(
    <div style={{position:"fixed",top:20,right:20,background:type==="error"?"#C94040":type==="warn"?"#C8973A":G.primary,color:"white",padding:"12px 20px",borderRadius:12,fontWeight:600,zIndex:9999,fontSize:13,boxShadow:"0 8px 32px rgba(0,0,0,.2)",display:"flex",alignItems:"center",gap:8,fontFamily:"inherit",maxWidth:320}}>
      {type==="error"?"❌":type==="warn"?"⚠️":"✅"} {msg}
    </div>
  );
}

/* ─────────────────────────────────── NAV ITEM */
function NavItem({icon,label,active,onClick,badge}){
  return(
    <button onClick={onClick} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 16px",borderRadius:12,border:"none",background:active?G.primary:"transparent",color:active?"white":G.textMid,cursor:"pointer",width:"100%",textAlign:"left",fontSize:13,fontWeight:active?700:500,fontFamily:"inherit",transition:"all .15s",position:"relative"}}>
      <span style={{fontSize:17,lineHeight:1}}>{icon}</span>
      <span style={{flex:1}}>{label}</span>
      {badge>0&&<span style={{background:active?"rgba(255,255,255,.3)":"#C94040",color:"white",borderRadius:99,padding:"1px 7px",fontSize:10,fontWeight:700,minWidth:18,textAlign:"center"}}>{badge}</span>}
    </button>
  );
}

/* ─────────────────────────────────── STAT CARD */
function StatCard({icon,label,value,color,sub}){
  return(
    <div style={{background:G.surface,borderRadius:16,padding:"18px 20px",boxShadow:G.shadow,border:`1px solid ${G.border}`,display:"flex",flexDirection:"column",gap:4}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <span style={{fontSize:22}}>{icon}</span>
        <span style={{fontSize:11,color:G.textSoft,fontWeight:600,textTransform:"uppercase",letterSpacing:.8}}>{label}</span>
      </div>
      <div style={{fontSize:24,fontWeight:800,color:color||G.text,fontFamily:"'Georgia',serif",marginTop:4}}>{value}</div>
      {sub&&<div style={{fontSize:11,color:G.textSoft}}>{sub}</div>}
    </div>
  );
}

/* ─────────────────────────────────── SECTION HEADER */
function PageHeader({title,sub}){
  return(
    <div style={{marginBottom:24}}>
      <h2 style={{margin:0,fontSize:26,fontWeight:800,color:G.text,fontFamily:"'Georgia',serif",letterSpacing:-.3}}>{title}</h2>
      {sub&&<p style={{margin:"4px 0 0",color:G.textSoft,fontSize:13}}>{sub}</p>}
    </div>
  );
}

/* ─────────────────────────────────── PILL */
function Pill({label,color,bg}){
  return <span style={{display:"inline-flex",alignItems:"center",gap:4,padding:"3px 10px",borderRadius:99,fontSize:11,fontWeight:700,color:color||G.primary,background:bg||G.primaryL}}>{label}</span>;
}

/* ─────────────────────────────────── TABLE */
function Table({cols,rows}){
  return(
    <div style={{overflowX:"auto",borderRadius:16,border:`1px solid ${G.border}`,boxShadow:G.shadow}}>
      <table style={{width:"100%",borderCollapse:"collapse",background:G.surface,fontSize:13}}>
        <thead><tr style={{background:G.bg}}>
          {cols.map(c=><th key={c} style={{padding:"12px 16px",textAlign:"left",fontWeight:700,fontSize:11,color:G.textSoft,textTransform:"uppercase",letterSpacing:.7,borderBottom:`1px solid ${G.border}`,whiteSpace:"nowrap"}}>{c}</th>)}
        </tr></thead>
        <tbody>{rows.map((row,i)=><tr key={i} style={{borderBottom:i<rows.length-1?`1px solid ${G.border}`:"none"}} onMouseEnter={e=>e.currentTarget.style.background=G.bg} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
          {row.map((cell,j)=><td key={j} style={{padding:"13px 16px",verticalAlign:"middle"}}>{cell}</td>)}
        </tr>)}</tbody>
      </table>
    </div>
  );
}

/* ─────────────────────────────────── CARD */
function Card({children,style={}}){
  return <div style={{background:G.surface,borderRadius:16,padding:"20px",boxShadow:G.shadow,border:`1px solid ${G.border}`,...style}}>{children}</div>;
}

/* ─────────────────────────────────── CARD TITLE */
function CardTitle({children}){
  return <h3 style={{margin:"0 0 14px",fontSize:14,fontWeight:700,color:G.text,display:"flex",alignItems:"center",gap:8,paddingBottom:10,borderBottom:`1.5px solid ${G.border}`}}>{children}</h3>;
}

/* ─────────────────────────────────── SCORE BAR */
function ScoreBar({score,size="md"}){
  const tier=getTier(score);
  const h=size==="sm"?6:8;
  return(
    <div style={{display:"flex",alignItems:"center",gap:8}}>
      <div style={{flex:1,background:G.bg,borderRadius:99,height:h,minWidth:60,overflow:"hidden"}}>
        <div style={{width:score+"%",height:"100%",borderRadius:99,background:tier.color,transition:"width .5s"}}/>
      </div>
      <span style={{fontWeight:800,fontSize:size==="sm"?12:14,color:tier.color,minWidth:28}}>{score}</span>
    </div>
  );
}

/* ─────────────────────────────────── ROOT APP */
export default function App(){
  const saved=load();
  const [members,setMembers]=useState(saved?.members||INIT_MEMBERS);
  const [auditLog,setAuditLog]=useState(saved?.auditLog||[]);
  const [notes,setNotes]=useState(saved?.notes||[]);
  const [admin,setAdmin]=useState(null);
  const [adminSel,setAdminSel]=useState("");
  const [pinErr,setPinErr]=useState("");
  const [view,setView]=useState("dashboard");
  const [selMid,setSelMid]=useState(null);
  const [week,setWeek]=useState(1);
  const [toast,setToast]=useState(null);
  const [confirm,setConfirm]=useState(null);
  const [portalMid,setPortalMid]=useState(null);
  const [portalStep,setPortalStep]=useState(null); // null | "pin" | "pick"
  const [sidebarOpen,setSidebarOpen]=useState(false);
  const [isMobile,setIsMobile]=useState(typeof window!=="undefined"&&window.innerWidth<1024);

  useEffect(()=>{
    const onResize=()=>setIsMobile(window.innerWidth<1024);
    window.addEventListener("resize",onResize);
    return()=>window.removeEventListener("resize",onResize);
  },[]);
  const [portalOk,setPortalOk]=useState(false);
  const [portalErr,setPortalErr]=useState("");

  useEffect(()=>save({members,auditLog,notes}),[members,auditLog,notes]);

  const log=(action,detail,mid)=>setAuditLog(l=>[mkLog(admin,action,detail,mid),...l]);
  const toast2=(msg,type="success")=>{setToast({msg,type});setTimeout(()=>setToast(null),3200);};
  const upd=(id,fn)=>setMembers(ms=>ms.map(m=>m.id===id?fn({...m}):m));
  const ask=(cfg)=>setConfirm(cfg);
  const closeConfirm=()=>setConfirm(null);

  // derived
  const totalContribs=members.reduce((s,m)=>s+Object.values(m.weeklyContributions).reduce((a,c)=>a+(c.paid?c.amount:0),0),0);
  const totalPenalties=members.reduce((s,m)=>s+m.penalties.reduce((a,p)=>a+p.amount,0),0);
  const totalLoansOut=members.reduce((s,m)=>s+m.loans.reduce((a,l)=>a+(l.active?l.amount:0),0),0);
  const totalInterest=members.reduce((s,m)=>s+m.loans.reduce((a,l)=>a+l.amount*(l.rate||BASE_RATE),0),0);
  const totalRepaid=members.reduce((s,m)=>s+m.loans.reduce((a,l)=>a+(l.repayments||[]).reduce((b,r)=>b+r.amount,0),0),0);
  const totalCollected=totalContribs+totalPenalties+members.filter(m=>m.membershipFeePaid).length*MEM_FEE;
  const overdueMembers=members.filter(m=>m.loans.some(l=>l.active&&loanDueDate(l)&&new Date(loanDueDate(l))<new Date()));
  const pendingWeek=members.filter(m=>!m.weeklyContributions[week]?.paid).length;
  const idleCapital=Math.max(0,totalContribs-totalLoansOut);
  const selMember=members.find(m=>m.id===selMid);

  /* portal flow: pin → pick member → dashboard */
  if(!admin&&portalStep==="pin"){
    return(
      <div style={LS.wrap}>
        <div style={LS.card}>
          <button style={{...LS.back}} onClick={()=>{setPortalStep(null);setPortalErr("");}}>← Back</button>
          <div style={{width:74,height:74,borderRadius:18,background:`linear-gradient(135deg,${G.primary} 0%,${G.primaryD} 100%)`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:32,color:"white",margin:"0 auto 12px",boxShadow:`0 8px 24px rgba(61,107,79,.35)`}}>👤</div>
          <h1 style={LS.title}>Member Login</h1>
          <p style={LS.sub}>Read-only access to your dashboard</p>
          <label style={LS.label}>Member PIN</label>
          <input id="mpin" type="password" placeholder="0000" style={LS.input} onKeyDown={e=>{if(e.key==="Enter"){if(e.target.value==="0000"){setPortalStep("pick");setPortalErr("");}else setPortalErr("Incorrect PIN");}}} autoFocus/>
          {portalErr&&<p style={LS.err}>{portalErr}</p>}
          <button style={LS.btn} onClick={()=>{const v=document.getElementById("mpin").value;if(v==="0000"){setPortalStep("pick");setPortalErr("");}else setPortalErr("Incorrect PIN");}}>Continue →</button>
        </div>
      </div>
    );
  }

  if(!admin&&portalStep==="pick"){
    return(
      <div style={LS.wrap}>
        <div style={LS.card}>
          <button style={{...LS.back}} onClick={()=>{setPortalStep("pin");setPortalErr("");}}>← Back</button>
          <div style={{width:74,height:74,borderRadius:18,background:`linear-gradient(135deg,${G.primary} 0%,${G.primaryD} 100%)`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:32,color:"white",margin:"0 auto 12px",boxShadow:`0 8px 24px rgba(61,107,79,.35)`}}>👤</div>
          <h1 style={LS.title}>Who Are You?</h1>
          <p style={LS.sub}>Select your name to view your dashboard</p>
          <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:8,maxHeight:280,overflowY:"auto",marginTop:8}}>
            {members.map(m=><button key={m.id} style={{...LS.memberBtn,textAlign:"center",padding:"10px",fontSize:13,fontWeight:600,color:G.text}} onClick={()=>{setPortalMid(m.id);setPortalOk(true);setPortalStep(null);}}>{m.name}</button>)}
          </div>
        </div>
      </div>
    );
  }

  if(!admin&&portalMid&&portalOk){
    const pm=members.find(m=>m.id===portalMid);
    return <MemberPortal member={pm} members={members} onClose={()=>{setPortalMid(null);setPortalOk(false);setPortalErr("");setPortalStep(null);}} />;
  }

  /* admin login */
  if(!admin){
    const adminInitial=a=>a.charAt(0).toUpperCase();
    const adminColors={"Angella":"#7B5EA7","Sandra":"#3A72A8","Fatima":"#C8973A"};
    const selColor=adminSel?adminColors[adminSel]:G.primary;
    return(
      <div style={LS.wrap}>
        {/* Decorative top strip */}
        <div style={LS.topStrip}>
          <div style={LS.stripBlob1}/>
          <div style={LS.stripBlob2}/>
          <div style={LS.topBar}>
            <div style={LS.logoMini}>KSG</div>
            <div style={LS.topBarText}>
              <div style={LS.topBarTitle}>Khobidi Savings Group</div>
              <div style={LS.topBarSeason}>Season {SEASON_START} · Week 1</div>
            </div>
          </div>
        </div>

        {/* Floating card */}
        <div style={LS.floatCard}>
          <h2 style={LS.welcome}>Welcome back</h2>
          <p style={LS.welcomeSub}>Sign in to manage the group</p>

          {/* Admin avatars row */}
          <div style={LS.avatarRow}>
            {ADMINS.map(a=>{
              const isSel=adminSel===a;
              const color=adminColors[a]||G.primary;
              return(
                <button key={a} onClick={()=>{setAdminSel(a);setPinErr("");setTimeout(()=>document.getElementById("apin")?.focus(),50);}} style={LS.avatarCol}>
                  <div style={{...LS.avatarBig,background:isSel?color:`${color}1A`,color:isSel?"white":color,transform:isSel?"scale(1.08)":"scale(1)",boxShadow:isSel?`0 8px 24px ${color}55`:"none",border:isSel?`3px solid ${G.surface}`:`3px solid transparent`,outline:isSel?`2px solid ${color}`:"none"}}>{adminInitial(a)}</div>
                  <div style={{fontSize:13,fontWeight:isSel?700:600,color:isSel?G.text:G.textMid,marginTop:9}}>{a}</div>
                  {isSel&&<div style={{fontSize:9,color:color,marginTop:1,fontWeight:700,letterSpacing:1}}>● ACTIVE</div>}
                </button>
              );
            })}
          </div>

          {/* PIN section */}
          <div style={{...LS.pinSection,borderColor:adminSel?`${selColor}33`:G.border,background:adminSel?`${selColor}08`:G.bg}}>
            <div style={LS.pinLabel}>
              <span>🔒</span>
              <span>{adminSel?`PIN for ${adminSel}`:"Select an account above"}</span>
            </div>
            <input id="apin" type="password" placeholder="• • • •" style={{...LS.input,...(pinErr?{borderColor:G.red,background:"#FEF7F7"}:{})}}
              onKeyDown={e=>{if(e.key==="Enter"){if(!adminSel){setPinErr("Please select your account first");return;}if(e.target.value===ADMIN_CREDS[adminSel]){setAdmin(adminSel);setPinErr("");}else setPinErr("Incorrect PIN. Try again.");}}}/>
          </div>
          {pinErr&&<div style={LS.errBox}>⚠️ {pinErr}</div>}

          <button style={{...LS.btn,background:selColor,boxShadow:`0 8px 24px ${selColor}66`,...(adminSel?{}:{opacity:.5,cursor:"not-allowed",boxShadow:"none"})}}
            onClick={()=>{if(!adminSel){setPinErr("Please select your account first");return;}const v=document.getElementById("apin").value;if(v===ADMIN_CREDS[adminSel]){setAdmin(adminSel);setPinErr("");}else setPinErr("Incorrect PIN. Try again.");}}>
            Sign In →
          </button>

          <div style={LS.divider}><span style={LS.dividerLine}/><span style={LS.dividerText}>OR</span><span style={LS.dividerLine}/></div>

          <button style={LS.memberLoginBtn} onClick={()=>setPortalStep("pin")}>
            <span style={{width:32,height:32,borderRadius:"50%",background:`${G.primary}1A`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14}}>👤</span>
            <span style={{flex:1,textAlign:"left"}}><div style={{fontSize:13,fontWeight:700,color:G.text}}>Member Login</div><div style={{fontSize:10,color:G.textSoft,marginTop:1}}>Read-only access</div></span>
            <span style={{color:G.textSoft,fontSize:16}}>›</span>
          </button>
        </div>

        <p style={LS.helpText}>Trouble signing in? Contact another admin.</p>
      </div>
    );
  }

  const NAV=[
    {v:"dashboard",icon:"🏠",label:"Dashboard"},
    {v:"members",icon:"👥",label:"Members"},
    {v:"weekly",icon:"💵",label:"Payments"},
    {v:"loans",icon:"💳",label:"Loans"},
    {v:"penalties",icon:"⚠️",label:"Penalties"},
    {v:"calendar",icon:"📅",label:"Calendar"},
    {v:"analytics",icon:"📊",label:"Analytics"},
    {v:"profit",icon:"🎯",label:"Profit Optimizer"},
    {v:"credit",icon:"⭐",label:"Credit Scoring"},
    {v:"p2p",icon:"📈",label:"P2P Investment"},
    {v:"cashflow",icon:"〰️",label:"Cashflow"},
    {v:"audit",icon:"📋",label:"Audit Log"},
  ];

  return(
    <div style={{minHeight:"100vh",background:G.bg,fontFamily:"'DM Sans','Helvetica Neue',sans-serif",color:G.text,display:"flex",flexDirection:isMobile?"column":"row"}}>

      {/* MOBILE TOP BAR */}
      {isMobile&&(
        <header style={{position:"sticky",top:0,zIndex:50,background:G.surface,borderBottom:`1px solid ${G.border}`,padding:"10px 14px",display:"flex",alignItems:"center",gap:10,boxShadow:"0 2px 8px rgba(30,50,25,.04)"}}>
          <button aria-label="Menu" onClick={()=>setSidebarOpen(true)} style={{width:38,height:38,borderRadius:10,border:`1px solid ${G.border}`,background:G.bg,color:G.text,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,padding:0,fontFamily:"inherit"}}>☰</button>
          <div style={{width:32,height:32,borderRadius:9,background:G.primary,display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:800,color:"white",fontFamily:"'Georgia',serif",letterSpacing:.7}}>KSG</div>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontWeight:800,fontSize:13,color:G.text,letterSpacing:-.2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>Khobidi Savings Group</div>
            <div style={{fontSize:10,color:G.textSoft}}>{admin} · Week {week}</div>
          </div>
        </header>
      )}

      {/* MOBILE OVERLAY (when sidebar open) */}
      {isMobile&&sidebarOpen&&(
        <div onClick={()=>setSidebarOpen(false)} style={{position:"fixed",inset:0,background:"rgba(15,30,18,.55)",backdropFilter:"blur(4px)",zIndex:90,animation:"fadeIn .2s"}}/>
      )}

      {/* SIDEBAR */}
      {(!isMobile||sidebarOpen)&&<aside style={{
        width:isMobile?280:232,
        background:G.surface,
        borderRight:`1px solid ${G.border}`,
        display:"flex",flexDirection:"column",
        ...(isMobile?{
          position:"fixed",top:0,left:0,
          height:"100vh",zIndex:100,
          boxShadow:"4px 0 32px rgba(0,0,0,.25)",
          animation:"slideIn .22s ease",
        }:{
          position:"sticky",top:0,height:"100vh",flexShrink:0,
          boxShadow:"2px 0 12px rgba(30,50,25,.04)",
        }),
        overflowY:"auto",
      }}>
        <div style={{padding:"22px 20px 14px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:36,height:36,borderRadius:10,background:G.primary,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:800,color:"white",fontFamily:"'Georgia',serif",letterSpacing:.8}}>KSG</div>
            <div><div style={{fontWeight:800,fontSize:14,color:G.text,letterSpacing:-.3}}>Khobidi Savings</div><div style={{fontSize:10,color:G.textSoft}}>Season {SEASON_START}</div></div>
          </div>
          {isMobile&&<button onClick={()=>setSidebarOpen(false)} style={{width:30,height:30,borderRadius:8,border:"none",background:G.bg,color:G.textMid,cursor:"pointer",fontSize:16,fontFamily:"inherit"}}>✕</button>}
        </div>
        <div style={{margin:"0 12px 12px",padding:"10px 12px",background:G.primaryL,borderRadius:12,display:"flex",alignItems:"center",gap:8}}>
          <div style={{width:28,height:28,borderRadius:8,background:G.primary,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,color:"white"}}>👤</div>
          <div><div style={{fontSize:12,fontWeight:700,color:G.primary}}>{admin}</div><div style={{fontSize:10,color:G.accent}}>Administrator</div></div>
        </div>
        <nav style={{flex:1,padding:"4px 12px",display:"flex",flexDirection:"column",gap:2}}>
          {NAV.map(n=><NavItem key={n.v} icon={n.icon} label={n.label} active={view===n.v} onClick={()=>{setView(n.v);if(n.v!=="members")setSelMid(null);if(isMobile)setSidebarOpen(false);}} badge={n.v==="weekly"?pendingWeek:n.v==="loans"?overdueMembers.length:0}/>)}
        </nav>
        <div style={{padding:"12px 16px",borderTop:`1px solid ${G.border}`}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
            <span style={{fontSize:11,color:G.textSoft,fontWeight:600}}>CURRENT WEEK</span>
            <div style={{display:"flex",alignItems:"center",gap:6,marginLeft:"auto"}}>
              <button style={{width:24,height:24,borderRadius:6,border:`1px solid ${G.border}`,background:G.bg,cursor:"pointer",fontSize:14,color:G.textMid,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"inherit"}} onClick={()=>setWeek(w=>Math.max(1,w-1))}>−</button>
              <span style={{fontWeight:800,fontSize:16,color:G.primary,minWidth:24,textAlign:"center"}}>{week}</span>
              <button style={{width:24,height:24,borderRadius:6,border:`1px solid ${G.border}`,background:G.bg,cursor:"pointer",fontSize:14,color:G.textMid,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"inherit"}} onClick={()=>setWeek(w=>Math.min(52,w+1))}>+</button>
            </div>
          </div>
          <div style={{fontSize:10,color:G.textSoft,marginBottom:10}}>{fds(SUNDAYS[week-1]?.date)}</div>
          <button style={{width:"100%",padding:"9px",borderRadius:9,border:`1px solid ${G.border}`,background:G.bg,color:G.textMid,cursor:"pointer",fontSize:12,fontWeight:600,fontFamily:"inherit"}} onClick={()=>setAdmin(null)}>Sign Out</button>
        </div>
      </aside>}

      {/* MAIN */}
      <main style={{flex:1,overflowY:"auto",padding:isMobile?"18px 16px 100px":"32px 36px",minWidth:0,width:"100%"}}>
        {toast&&<Toast msg={toast.msg} type={toast.type}/>}
        {confirm&&<ConfirmDialog {...confirm} onCancel={closeConfirm}/>}

        {view==="dashboard"&&<DashView members={members} totalCollected={totalCollected} totalPenalties={totalPenalties} totalLoansOut={totalLoansOut} totalInterest={totalInterest} week={week} pendingWeek={pendingWeek} overdueMembers={overdueMembers} idleCapital={idleCapital} setView={setView}/>}
        {view==="members"&&!selMid&&<MembersView members={members} onSelect={id=>setSelMid(id)} upd={upd} toast2={toast2} log={log}/>}
        {view==="members"&&selMid&&selMember&&<MemberDetail member={selMember} upd={upd} toast2={toast2} log={log} week={week} ask={ask} closeConfirm={closeConfirm} onBack={()=>setSelMid(null)}/>}
        {view==="weekly"&&<WeeklyView members={members} week={week} upd={upd} toast2={toast2} log={log} ask={ask} closeConfirm={closeConfirm}/>}
        {view==="loans"&&<LoansView members={members} upd={upd} toast2={toast2} log={log} ask={ask} closeConfirm={closeConfirm}/>}
        {view==="penalties"&&<PenaltiesView members={members} upd={upd} toast2={toast2} log={log}/>}
        {view==="calendar"&&<CalendarView members={members} week={week} setWeek={setWeek} setView={setView}/>}
        {view==="analytics"&&<AnalyticsView members={members} week={week}/>}
        {view==="profit"&&<ProfitView members={members} totalLoansOut={totalLoansOut} totalInterest={totalInterest} idleCapital={idleCapital}/>}
        {view==="credit"&&<CreditView members={members}/>}
        {view==="p2p"&&<P2PView idleCapital={idleCapital}/>}
        {view==="cashflow"&&<CashflowView members={members} week={week}/>}
        {view==="audit"&&<AuditView auditLog={auditLog} members={members}/>}
      </main>

      {/* MOBILE BOTTOM NAV BAR */}
      {isMobile&&(
        <nav style={{position:"fixed",bottom:0,left:0,right:0,background:G.surface,borderTop:`1px solid ${G.border}`,boxShadow:"0 -4px 16px rgba(15,30,18,.06)",display:"flex",zIndex:80,paddingBottom:"env(safe-area-inset-bottom)"}}>
          {[
            {v:"dashboard",icon:"🏠",label:"Home"},
            {v:"members",icon:"👥",label:"Members"},
            {v:"weekly",icon:"💵",label:"Pay"},
            {v:"loans",icon:"💳",label:"Loans"},
            {v:"__more",icon:"☰",label:"More"},
          ].map(t=>{
            const isActive=t.v==="__more"?false:view===t.v;
            return(
              <button key={t.v} onClick={()=>{if(t.v==="__more"){setSidebarOpen(true);}else{setView(t.v);setSelMid(null);}}} style={{flex:1,background:"none",border:"none",padding:"10px 4px 8px",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:3,color:isActive?G.primary:G.textSoft,fontFamily:"inherit"}}>
                <span style={{fontSize:20,filter:isActive?"none":"grayscale(.5)",opacity:isActive?1:.7}}>{t.icon}</span>
                <span style={{fontSize:10,fontWeight:isActive?700:500}}>{t.label}</span>
              </button>
            );
          })}
        </nav>
      )}
    </div>
  );
}

/* ─── LOGIN STYLES */
const LS={
  /* phone-app shell */
  wrap:{minHeight:"100vh",maxWidth:440,margin:"0 auto",background:G.bg,display:"flex",flexDirection:"column",position:"relative",overflow:"hidden",fontFamily:"'DM Sans','Helvetica Neue',sans-serif",boxShadow:"0 0 60px rgba(0,0,0,.08)"},

  /* top strip (smaller header with logo + name in a row) */
  topStrip:{position:"relative",padding:"22px 22px 70px",background:`linear-gradient(135deg,${G.primaryD} 0%,${G.primary} 100%)`,color:"white",overflow:"hidden"},
  stripBlob1:{position:"absolute",top:-40,right:-30,width:140,height:140,borderRadius:"50%",background:"rgba(200,151,58,.22)",filter:"blur(24px)"},
  stripBlob2:{position:"absolute",bottom:0,left:-30,width:120,height:120,borderRadius:"50%",background:"rgba(123,94,167,.18)",filter:"blur(20px)"},
  topBar:{display:"flex",alignItems:"center",gap:12,position:"relative",zIndex:1},
  logoMini:{width:44,height:44,borderRadius:12,background:"rgba(255,255,255,.18)",backdropFilter:"blur(10px)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:800,color:"white",fontFamily:"'Georgia',serif",letterSpacing:1.2,border:"1.5px solid rgba(255,255,255,.28)",flexShrink:0},
  topBarText:{display:"flex",flexDirection:"column"},
  topBarTitle:{fontSize:15,fontWeight:800,fontFamily:"'Georgia',serif",letterSpacing:-.2,lineHeight:1.2},
  topBarSeason:{fontSize:10,opacity:.85,marginTop:2,letterSpacing:.5,fontWeight:600},

  /* floating card pulled up over the strip */
  floatCard:{background:G.surface,margin:"-48px 18px 0",borderRadius:24,padding:"24px 22px 22px",boxShadow:"0 20px 50px rgba(15,30,18,.12), 0 2px 6px rgba(15,30,18,.04)",border:`1px solid ${G.border}`,position:"relative",zIndex:2},
  welcome:{fontSize:22,fontWeight:800,fontFamily:"'Georgia',serif",margin:"0 0 4px",color:G.text,letterSpacing:-.3,textAlign:"center"},
  welcomeSub:{fontSize:12,color:G.textSoft,margin:"0 0 22px",lineHeight:1.5,textAlign:"center"},

  /* admin avatars in a row */
  avatarRow:{display:"flex",justifyContent:"space-around",gap:8,marginBottom:22,paddingBottom:6},
  avatarCol:{flex:1,background:"none",border:"none",cursor:"pointer",padding:"4px 2px",display:"flex",flexDirection:"column",alignItems:"center",fontFamily:"inherit"},
  avatarBig:{width:58,height:58,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,fontWeight:800,fontFamily:"'Georgia',serif",transition:"all .2s ease"},

  /* PIN section grouped in a tinted container */
  pinSection:{padding:"14px 14px 12px",borderRadius:14,border:`1.5px solid`,marginBottom:8,transition:"all .2s"},
  pinLabel:{display:"flex",alignItems:"center",gap:7,fontSize:11,fontWeight:700,color:G.textMid,marginBottom:8,letterSpacing:.3},
  input:{width:"100%",padding:"14px",border:`1.5px solid ${G.border}`,borderRadius:11,fontSize:20,boxSizing:"border-box",outline:"none",fontFamily:"inherit",color:G.text,background:G.surface,transition:"all .15s",letterSpacing:6,textAlign:"center",fontWeight:700},
  errBox:{display:"flex",alignItems:"center",gap:6,color:G.red,fontSize:12,marginTop:6,marginBottom:6,fontWeight:600,padding:"8px 12px",background:"#FEF2F2",borderRadius:8,border:`1px solid #FCA5A5`},

  /* CTA + secondary */
  btn:{width:"100%",padding:"14px",color:"white",border:"none",borderRadius:14,fontSize:15,fontWeight:700,cursor:"pointer",marginTop:14,fontFamily:"inherit",letterSpacing:.3,transition:"all .2s"},
  divider:{display:"flex",alignItems:"center",gap:10,margin:"18px 0 10px"},
  dividerLine:{flex:1,height:1,background:G.border},
  dividerText:{fontSize:10,fontWeight:700,color:G.textSoft,letterSpacing:1.5},
  memberLoginBtn:{width:"100%",padding:"12px 14px",background:G.bg,color:G.text,border:`1.5px solid ${G.border}`,borderRadius:14,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",gap:11,transition:"all .15s"},
  helpText:{fontSize:11,color:G.textSoft,textAlign:"center",margin:"22px 24px 24px"},

  /* legacy / shared (for portal screens still using these) */
  card:{background:G.surface,borderRadius:24,padding:"40px 32px",width:340,maxWidth:"92vw",boxShadow:"0 32px 80px rgba(15,30,18,.35)",position:"relative",zIndex:1,margin:"0 auto"},
  logo:{fontSize:44,textAlign:"center",marginBottom:8},
  title:{textAlign:"center",fontSize:24,fontWeight:800,margin:"0 0 4px",fontFamily:"'Georgia',serif",color:G.text},
  sub:{textAlign:"center",color:G.textSoft,marginBottom:22,fontSize:12},
  label:{display:"block",fontSize:10,fontWeight:700,color:G.textSoft,marginBottom:8,textTransform:"uppercase",letterSpacing:1.2},
  adminBtn:{padding:"10px 8px",border:`1.5px solid ${G.border}`,borderRadius:10,cursor:"pointer",background:G.bg,fontSize:12,fontWeight:600,fontFamily:"inherit",color:G.textMid,whiteSpace:"nowrap"},
  adminBtnA:{background:G.primary,color:"white",borderColor:G.primary},
  err:{color:G.red,fontSize:12,marginTop:2,fontWeight:500},
  back:{background:"none",border:"none",color:G.blue,cursor:"pointer",fontSize:12,marginBottom:12,padding:0,fontWeight:600,fontFamily:"inherit"},
  memberBtn:{background:G.bg,border:`1px solid ${G.border}`,borderRadius:8,padding:"7px 12px",cursor:"pointer",fontSize:12,textAlign:"left",fontFamily:"inherit",color:G.textMid,fontWeight:500},
};

/* ═══════════════════════════════════ DASHBOARD */
function DashView({members,totalCollected,totalPenalties,totalLoansOut,totalInterest,week,pendingWeek,overdueMembers,idleCapital,setView}){
  const activeLoans=members.reduce((s,m)=>s+m.loans.filter(l=>l.active).length,0);
  return(
    <div>
      <PageHeader title="Dashboard" sub={`Week ${week} · ${fds(SUNDAYS[week-1]?.date)}`}/>

      {overdueMembers.length>0&&(
        <div style={{background:"#FEF2F2",border:`1px solid #FCA5A5`,borderRadius:14,padding:"14px 18px",marginBottom:20,display:"flex",alignItems:"center",gap:12}}>
          <div style={{fontSize:22}}>🚨</div>
          <div style={{flex:1}}>
            <div style={{fontSize:13,fontWeight:700,color:"#991B1B"}}>{overdueMembers.length} overdue loan{overdueMembers.length>1?"s":""}</div>
            <div style={{fontSize:11,color:"#7F1D1D",marginTop:2}}>{overdueMembers.map(m=>m.name).join(", ")}</div>
          </div>
          <button style={{padding:"6px 14px",borderRadius:8,border:"none",background:"#991B1B",color:"white",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}} onClick={()=>setView("loans")}>Review →</button>
        </div>
      )}

      {idleCapital>500&&(
        <div style={{background:"#EEF6FA",border:`1px solid #B8D9EC`,borderRadius:14,padding:"14px 18px",marginBottom:20,display:"flex",alignItems:"center",gap:12}}>
          <div style={{fontSize:22}}>💡</div>
          <div style={{flex:1}}>
            <div style={{fontSize:13,fontWeight:700,color:"#1E3A5F"}}>{fmt(idleCapital)} sitting idle</div>
            <div style={{fontSize:11,color:"#1E40AF",marginTop:2}}>Deploy to P2P platforms for 11–15% returns</div>
          </div>
          <button style={{padding:"6px 14px",borderRadius:8,border:"none",background:G.blue,color:"white",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}} onClick={()=>setView("p2p")}>Explore →</button>
        </div>
      )}

      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(170px,1fr))",gap:14,marginBottom:24}}>
        <StatCard icon="💰" label="Collected" value={fmt(totalCollected)} color={G.primary}/>
        <StatCard icon="⏳" label={`Week ${week} Pending`} value={pendingWeek+" / "+members.length} color={pendingWeek>0?G.gold:G.primary}/>
        <StatCard icon="💳" label="Loans Out" value={fmt(totalLoansOut)} color={G.blue}/>
        <StatCard icon="📈" label="Interest Revenue" value={fmt(totalInterest)} color={G.purple}/>
        <StatCard icon="⚠️" label="Penalties" value={fmt(totalPenalties)} color={G.red}/>
        <StatCard icon="🏦" label="Active Loans" value={activeLoans} color={G.accent}/>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:16}}>
        <Card>
          <CardTitle>👥 Member Overview · Week {week}</CardTitle>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:10}}>
            {members.map(m=>{
              const paid=m.weeklyContributions[week]?.paid;
              const score=creditScore(m);const tier=getTier(score);
              return(
                <div key={m.id} style={{padding:"10px 12px",borderRadius:10,background:G.bg,borderLeft:`3px solid ${tier.color}`,display:"flex",flexDirection:"column",gap:3}}>
                  <div style={{fontWeight:700,fontSize:12,color:G.text}}>{m.name}</div>
                  <div style={{display:"flex",alignItems:"center",gap:6,fontSize:10}}>
                    <span style={{color:paid?G.primary:G.gold}}>{paid?"✓ Paid":"⏳ Pending"}</span>
                    <span style={{color:tier.color,fontWeight:600}}>{tier.icon} {score}</span>
                  </div>
                </div>
              );
            })}
          </div>
        </Card>

        <Card>
          <CardTitle>⚡ Quick Actions</CardTitle>
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {[["weekly","💵","Log Payments"],["loans","💳","Manage Loans"],["analytics","📊","View Analytics"],["profit","🎯","Profit Optimizer"],["p2p","📈","P2P Investments"]].map(([v,i,l])=>(
              <button key={v} onClick={()=>setView(v)} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",borderRadius:10,border:`1px solid ${G.border}`,background:G.bg,cursor:"pointer",fontSize:12,fontWeight:600,color:G.textMid,fontFamily:"inherit",textAlign:"left"}}>
                <span style={{fontSize:16}}>{i}</span><span style={{flex:1}}>{l}</span><span style={{color:G.textSoft}}>→</span>
              </button>
            ))}
          </div>
        </Card>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════ MEMBERS LIST */
function MembersView({members,onSelect,upd,toast2,log}){
  return(
    <div>
      <PageHeader title="Members" sub={`${members.length} active members`}/>
      <Table
        cols={["#","Name","Credit","Rate","Fee","Contributions","Loan Bal","Penalties",""]}
        rows={members.map(m=>{
          const score=creditScore(m);const tier=getTier(score);
          const contribs=Object.values(m.weeklyContributions).reduce((s,c)=>s+(c.paid?c.amount:0),0);
          const loanBal=m.loans.reduce((s,l)=>s+loanBalance(l),0);
          return[
            <span style={{color:G.textSoft,fontWeight:600,fontSize:12}}>#{m.id}</span>,
            <div><div style={{fontWeight:700,color:G.text}}>{m.name}</div><div style={{fontSize:10,color:G.textSoft,marginTop:2}}>{m.email}</div></div>,
            <ScoreBar score={score} size="sm"/>,
            <Pill label={Math.round(tier.rate*100)+"%"} color={tier.color} bg={tier.bg}/>,
            m.membershipFeePaid
              ?<Pill label="✓ Paid" color={G.primary} bg={G.primaryL}/>
              :<button style={smallBtnPrim} onClick={()=>{upd(m.id,mm=>({...mm,membershipFeePaid:true}));log("FEE","€110",m.id);toast2(`${m.name} fee logged`);}}>Log €110</button>,
            <span style={{fontWeight:600,color:G.text}}>{fmt(contribs)}</span>,
            <span style={{fontWeight:700,color:loanBal>0?G.red:G.primary}}>{fmt(loanBal)}</span>,
            <span style={{color:m.penalties.length>0?G.red:G.textSoft,fontWeight:600}}>{m.penalties.length}</span>,
            <button style={viewBtn} onClick={()=>onSelect(m.id)}>View →</button>
          ];
        })}
      />
    </div>
  );
}

const smallBtnPrim={padding:"5px 12px",borderRadius:7,border:"none",background:G.primary,color:"white",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"};
const smallBtn={padding:"5px 12px",borderRadius:7,border:`1px solid ${G.border}`,background:"white",color:G.textMid,fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit"};
const smallBtnDanger={padding:"5px 12px",borderRadius:7,border:`1px solid #FCA5A5`,background:"#FEF2F2",color:G.red,fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit"};
const viewBtn={padding:"5px 14px",borderRadius:7,border:"none",background:G.text,color:"white",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit"};

/* ═══════════════════════════════════ MEMBER DETAIL */
function MemberDetail({member:m,upd,toast2,log,week,ask,closeConfirm,onBack}){
  const [wkInput,setWkInput]=useState(week);
  const [repayAmt,setRepayAmt]=useState("");
  const [repayLoanIdx,setRepayLoanIdx]=useState(0);
  const score=creditScore(m);const tier=getTier(score);
  const totalContribs=Object.values(m.weeklyContributions).reduce((s,c)=>s+(c.paid?c.amount:0),0);
  const totalPens=m.penalties.reduce((s,p)=>s+p.amount,0);
  const totalLoanBal=m.loans.reduce((s,l)=>s+loanBalance(l),0);

  function doLogWeek(w,amt){
    ask({
      msg:"Confirm Payment",
      sub:`Log €${amt} from ${m.name} for Week ${w}?`,
      confirmLabel:`Yes, log €${amt}`,
      onConfirm:()=>{
        upd(m.id,mm=>({...mm,weeklyContributions:{...mm.weeklyContributions,[w]:{amount:amt,paid:true,date:new Date().toISOString()}}}));
        log("WEEKLY",`Wk${w}: ${fmt(amt)}`,m.id);
        toast2(`Week ${w} payment of ${fmt(amt)} logged`);
        closeConfirm();
      }
    });
  }
  function doIssueLoan(){
    if(m.loans.length>=2){toast2("Maximum 2 loans reached","error");return;}
    ask({
      msg:"Issue New Loan?",
      sub:`Issue ${fmt(LOAN_AMOUNT)} loan to ${m.name} at ${Math.round(tier.rate*100)}% interest (${tier.label} tier). Total repayable: ${fmt(LOAN_AMOUNT*(1+tier.rate))}.`,
      confirmLabel:"Issue Loan",
      onConfirm:()=>{
        const loan={id:Date.now(),amount:LOAN_AMOUNT,rate:tier.rate,total:LOAN_AMOUNT*(1+tier.rate),issuedDate:new Date().toISOString(),active:true,repayments:[]};
        upd(m.id,mm=>({...mm,loans:[...mm.loans,loan]}));
        log("LOAN_ISSUED",`${fmt(LOAN_AMOUNT)} @ ${Math.round(tier.rate*100)}%`,m.id);
        toast2(`Loan issued at ${Math.round(tier.rate*100)}%`);
        closeConfirm();
      }
    });
  }
  function doRepayLoan(){
    const amt=parseFloat(repayAmt);
    if(!amt||amt<=0){toast2("Enter valid amount","error");return;}
    ask({
      msg:"Confirm Repayment",
      sub:`Log ${fmt(amt)} repayment on ${m.name}'s Loan ${repayLoanIdx+1}?`,
      confirmLabel:"Log Repayment",
      onConfirm:()=>{
        upd(m.id,mm=>{const loans=mm.loans.map((l,i)=>{if(i!==repayLoanIdx)return l;const repayments=[...(l.repayments||[]),{amount:amt,date:new Date().toISOString()}];return{...l,repayments,active:loanBalance({...l,repayments})>0};});return{...mm,loans};});
        log("REPAYMENT",`${fmt(amt)} loan ${repayLoanIdx+1}`,m.id);
        toast2(`Repayment of ${fmt(amt)} logged`);
        setRepayAmt("");
        closeConfirm();
      }
    });
  }
  function doAddPenalty(reason){
    ask({
      msg:"Add Penalty?",
      sub:`Apply €${PENALTY_AMT} penalty to ${m.name}? Reason: ${reason}`,
      confirmLabel:"Apply Penalty",
      danger:true,
      onConfirm:()=>{
        upd(m.id,mm=>({...mm,penalties:[...mm.penalties,{amount:PENALTY_AMT,reason,date:new Date().toISOString()}]}));
        log("PENALTY",`€${PENALTY_AMT} — ${reason}`,m.id);
        toast2(`Penalty added to ${m.name}`,"warn");
        closeConfirm();
      }
    });
  }

  return(
    <div>
      <button style={{background:"none",border:"none",color:G.blue,cursor:"pointer",fontSize:13,marginBottom:14,padding:0,fontWeight:600,fontFamily:"inherit"}} onClick={onBack}>← Back to Members</button>

      <div style={{background:G.surface,borderRadius:18,padding:24,marginBottom:18,boxShadow:G.shadow,border:`1px solid ${G.border}`,borderLeft:`5px solid ${tier.color}`}}>
        <div style={{display:"flex",gap:18,alignItems:"center",flexWrap:"wrap"}}>
          <div style={{width:64,height:64,borderRadius:16,background:tier.bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:32}}>{tier.icon}</div>
          <div style={{flex:1,minWidth:200}}>
            <h2 style={{margin:0,fontSize:24,fontWeight:800,fontFamily:"'Georgia',serif",color:G.text}}>{m.name}</h2>
            <div style={{display:"flex",gap:10,marginTop:4,flexWrap:"wrap",alignItems:"center"}}>
              <Pill label={`${tier.icon} ${tier.label} — ${score}/100`} color={tier.color} bg={tier.bg}/>
              <span style={{fontSize:12,color:G.textSoft}}>Loan Rate: <strong style={{color:tier.color}}>{Math.round(tier.rate*100)}%</strong></span>
              <span style={{fontSize:12,color:G.textSoft}}>Member #{m.id}</span>
            </div>
          </div>
          <div style={{display:"flex",gap:14,flexWrap:"wrap"}}>
            {[["Contributions",fmt(totalContribs),G.primary],["Loan Balance",fmt(totalLoanBal),G.blue],["Penalties",fmt(totalPens),G.red]].map(([l,v,c])=>(
              <div key={l} style={{textAlign:"center"}}>
                <div style={{fontSize:18,fontWeight:800,color:c,fontFamily:"'Georgia',serif"}}>{v}</div>
                <div style={{fontSize:10,color:G.textSoft,textTransform:"uppercase",letterSpacing:.5}}>{l}</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:14}}>
        <Card>
          <CardTitle>💵 Log Weekly Payment</CardTitle>
          <label style={{fontSize:10,fontWeight:700,color:G.textSoft,textTransform:"uppercase",letterSpacing:.6,display:"block",marginBottom:5}}>Week #</label>
          <input type="number" min={1} max={52} value={wkInput} style={inputStyle} onChange={e=>setWkInput(+e.target.value)}/>
          <div style={{fontSize:10,color:G.textSoft,marginBottom:10}}>Due: {fds(SUNDAYS[wkInput-1]?.date)}</div>
          {m.weeklyContributions[wkInput]?.paid
            ?<div style={{background:G.primaryL,color:G.primary,padding:"8px 12px",borderRadius:8,fontSize:12,fontWeight:600}}>✓ Week {wkInput}: {fmt(m.weeklyContributions[wkInput].amount)}</div>
            :<div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
              {[60,120,180,240,300].map(a=>(
                <button key={a} style={{padding:"7px 12px",borderRadius:8,border:`1px solid ${G.primary}`,background:G.primaryL,color:G.primary,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}} onClick={()=>doLogWeek(wkInput,a)}>{fmt(a)}</button>
              ))}
            </div>}
          <button style={{...smallBtnDanger,marginTop:10,width:"100%",padding:"8px"}} onClick={()=>doAddPenalty(`Missed week ${wkInput}`)}>+ Late Penalty (€{PENALTY_AMT})</button>
        </Card>

        <Card>
          <CardTitle>💳 Loans · {Math.round(tier.rate*100)}% Rate</CardTitle>
          {m.loans.length===0&&<p style={{color:G.textSoft,fontSize:12,margin:"0 0 10px"}}>No loans issued.</p>}
          {m.loans.map((l,i)=>{
            const bal=loanBalance(l);const due=loanDueDate(l);
            const isOD=due&&new Date(due)<new Date()&&l.active;
            return(
              <div key={l.id} style={{padding:10,borderRadius:9,marginBottom:8,background:G.bg,borderLeft:`3px solid ${isOD?G.red:l.active?G.blue:G.primary}`}}>
                <div style={{fontWeight:700,fontSize:12,color:G.text}}>Loan {i+1}: {fmt(l.amount)} @ {Math.round((l.rate||0.20)*100)}%</div>
                <div style={{fontSize:10,color:G.textSoft,marginTop:2}}>Issued: {fd(l.issuedDate)} · Due: {fd(due)}</div>
                <div style={{color:isOD?G.red:l.active?G.gold:G.primary,fontWeight:700,fontSize:12,marginTop:4}}>{isOD?"🚨 ":""}Bal: {fmt(bal)} {!l.active&&"✓"}</div>
                {(l.repayments||[]).map((r,j)=><div key={j} style={{color:G.primary,fontSize:10,marginTop:2}}>↳ {fmt(r.amount)} on {fd(r.date)}</div>)}
              </div>
            );
          })}
          {m.loans.length<2&&<button style={{padding:"9px 14px",borderRadius:9,border:"none",background:G.primary,color:"white",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit",width:"100%",marginTop:4}} onClick={doIssueLoan}>+ Issue Loan ({Math.round(tier.rate*100)}%)</button>}
          {m.loans.some(l=>l.active)&&(
            <div style={{marginTop:10,padding:10,background:G.bg,borderRadius:9}}>
              <select style={{...inputStyle,marginBottom:6}} value={repayLoanIdx} onChange={e=>setRepayLoanIdx(+e.target.value)}>
                {m.loans.map((l,i)=>l.active&&<option key={i} value={i}>Loan {i+1} — bal {fmt(loanBalance(l))}</option>)}
              </select>
              <div style={{display:"flex",gap:6}}>
                <input type="number" placeholder="Repayment €" value={repayAmt} style={{...inputStyle,marginBottom:0,flex:1}} onChange={e=>setRepayAmt(e.target.value)}/>
                <button style={{padding:"8px 14px",borderRadius:8,border:"none",background:G.primary,color:"white",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}} onClick={doRepayLoan}>Log</button>
              </div>
            </div>
          )}
        </Card>

        <Card>
          <CardTitle>⚠️ Penalties</CardTitle>
          {m.penalties.length===0&&<p style={{color:G.textSoft,fontSize:12,margin:0}}>No penalties on record.</p>}
          {m.penalties.map((p,i)=>(
            <div key={i} style={{padding:"6px 0",fontSize:12,borderBottom:`1px solid ${G.border}`,display:"flex",justifyContent:"space-between",gap:8}}>
              <span>{p.reason}</span>
              <span style={{color:G.red,fontWeight:700,whiteSpace:"nowrap"}}>€{p.amount}</span>
            </div>
          ))}
          {m.penalties.length>0&&<div style={{marginTop:8,fontSize:12,fontWeight:700,color:G.text}}>Total: {fmt(totalPens)}</div>}
        </Card>

        <Card style={{gridColumn:"1 / -1"}}>
          <CardTitle>📅 Contribution History · Full Season</CardTitle>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(48px,1fr))",gap:4}}>
            {SUNDAYS.map(s=>{
              const c=m.weeklyContributions[s.week];const isPast=s.date<new Date();
              return(
                <div key={s.week} style={{padding:"6px 3px",borderRadius:6,textAlign:"center",background:c?.paid?G.primaryL:isPast?"#FEF2F2":G.bg,color:c?.paid?G.primary:isPast?G.red:G.textSoft,border:`1px solid ${c?.paid?G.primary:isPast?"#FCA5A5":G.border}`}}>
                  <div style={{fontWeight:700,fontSize:10}}>W{s.week}</div>
                  <div style={{fontSize:9}}>{c?.paid?fmt(c.amount):"—"}</div>
                </div>
              );
            })}
          </div>
        </Card>
      </div>
    </div>
  );
}

const inputStyle={width:"100%",padding:"9px 12px",border:`1.5px solid ${G.border}`,borderRadius:9,fontSize:13,boxSizing:"border-box",marginBottom:8,outline:"none",fontFamily:"inherit",color:G.text,background:G.bg};

/* ═══════════════════════════════════ WEEKLY PAYMENTS */
function WeeklyView({members,week,upd,toast2,log,ask,closeConfirm}){
  const due=SUNDAYS[week-1]?.date;
  const paidCount=members.filter(m=>m.weeklyContributions[week]?.paid).length;
  const weekTotal=members.reduce((s,m)=>s+(m.weeklyContributions[week]?.paid?m.weeklyContributions[week].amount:0),0);

  function doLog(mid,name){
    ask({
      msg:"Confirm Payment",
      sub:`Log €60 from ${name} for Week ${week}? This action will be recorded in the audit log.`,
      confirmLabel:"Yes, log €60",
      onConfirm:()=>{
        upd(mid,m=>({...m,weeklyContributions:{...m.weeklyContributions,[week]:{amount:60,paid:true,date:new Date().toISOString()}}}));
        log("WEEKLY",`Wk${week}: €60`,mid);
        toast2(`${name} — Week ${week} logged`);
        closeConfirm();
      }
    });
  }
  function doPen(mid,name){
    ask({
      msg:"Apply Late Penalty?",
      sub:`Apply €${PENALTY_AMT} penalty to ${name} for missed Week ${week} payment?`,
      confirmLabel:"Apply Penalty",
      danger:true,
      onConfirm:()=>{
        upd(mid,m=>({...m,penalties:[...m.penalties,{amount:PENALTY_AMT,reason:`Late/missed week ${week}`,date:new Date().toISOString()}]}));
        log("PENALTY",`€${PENALTY_AMT} missed wk${week}`,mid);
        toast2(`Penalty added to ${name}`,"warn");
        closeConfirm();
      }
    });
  }

  return(
    <div>
      <PageHeader title={`Week ${week} Payments`} sub={`Due: ${fd(due)} by 5:00 PM`}/>

      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:12,marginBottom:18}}>
        <StatCard icon="✅" label="Paid" value={`${paidCount}/${members.length}`} color={G.primary}/>
        <StatCard icon="💰" label="Collected" value={fmt(weekTotal)} color={G.primary}/>
        <StatCard icon="⏳" label="Pending" value={members.length-paidCount} color={G.gold}/>
        <StatCard icon="📅" label="Due Date" value={fds(due)} color={G.blue}/>
      </div>

      <Table
        cols={["Member","Tier","Status","Amount","Logged","Actions"]}
        rows={members.map(m=>{
          const c=m.weeklyContributions[week];
          const score=creditScore(m);const tier=getTier(score);
          return[
            <div style={{fontWeight:700,color:G.text}}>{m.name}</div>,
            <Pill label={`${tier.icon} ${score}`} color={tier.color} bg={tier.bg}/>,
            c?.paid?<Pill label="✓ Paid" color={G.primary} bg={G.primaryL}/>:<Pill label="⏳ Pending" color={G.gold} bg="#FDF6E9"/>,
            <span style={{fontWeight:600}}>{c?.paid?fmt(c.amount):"—"}</span>,
            <span style={{fontSize:11,color:G.textSoft}}>{c?.paid?fd(c.date):"—"}</span>,
            !c?.paid
              ?<div style={{display:"flex",gap:5}}>
                <button style={smallBtnPrim} onClick={()=>doLog(m.id,m.name)}>Log €60</button>
                <button style={smallBtnDanger} onClick={()=>doPen(m.id,m.name)}>Penalty</button>
              </div>
              :<span style={{color:G.textSoft,fontSize:11}}>✓ Complete</span>
          ];
        })}
      />
    </div>
  );
}

/* ═══════════════════════════════════ LOANS */
function LoansView({members,upd,toast2,log,ask,closeConfirm}){
  const [repayAmts,setRepayAmts]=useState({});

  function doIssue(mid,name){
    const m=members.find(m=>m.id===mid);
    if(m.loans.length>=2){toast2("Maximum 2 loans","error");return;}
    const score=creditScore(m);const tier=getTier(score);
    ask({
      msg:"Issue New Loan?",
      sub:`Issue ${fmt(LOAN_AMOUNT)} to ${name} at ${Math.round(tier.rate*100)}% (${tier.label}). Total repayable: ${fmt(LOAN_AMOUNT*(1+tier.rate))} within 3 months.`,
      confirmLabel:"Issue Loan",
      onConfirm:()=>{
        const loan={id:Date.now(),amount:LOAN_AMOUNT,rate:tier.rate,total:LOAN_AMOUNT*(1+tier.rate),issuedDate:new Date().toISOString(),active:true,repayments:[]};
        upd(mid,mm=>({...mm,loans:[...mm.loans,loan]}));
        log("LOAN_ISSUED",`${fmt(LOAN_AMOUNT)} @ ${Math.round(tier.rate*100)}% to ${name}`,mid);
        toast2(`Loan to ${name} @ ${Math.round(tier.rate*100)}%`);
        closeConfirm();
      }
    });
  }
  function doRepay(mid,loanIdx,name){
    const key=`${mid}-${loanIdx}`;const amt=parseFloat(repayAmts[key]||0);
    if(!amt){toast2("Enter amount","error");return;}
    ask({
      msg:"Confirm Repayment",
      sub:`Log ${fmt(amt)} repayment from ${name} on Loan ${loanIdx+1}?`,
      confirmLabel:"Log Repayment",
      onConfirm:()=>{
        upd(mid,mm=>{const loans=mm.loans.map((l,i)=>{if(i!==loanIdx)return l;const repayments=[...(l.repayments||[]),{amount:amt,date:new Date().toISOString()}];return{...l,repayments,active:loanBalance({...l,repayments})>0};});return{...mm,loans};});
        log("REPAYMENT",`${fmt(amt)} loan${loanIdx+1}`,mid);
        toast2(`${fmt(amt)} logged`);
        setRepayAmts(p=>({...p,[key]:""}));
        closeConfirm();
      }
    });
  }

  return(
    <div>
      <PageHeader title="Loans Management" sub="Dynamic rates based on credit tier · 3-month repayment window"/>

      <div style={{background:G.primaryL,border:`1px solid ${G.primary}`,borderRadius:12,padding:"12px 16px",marginBottom:18,fontSize:12,color:G.primary,display:"flex",gap:14,flexWrap:"wrap"}}>
        {TIERS.map(t=><span key={t.label} style={{display:"flex",alignItems:"center",gap:4}}>{t.icon} <strong>{t.label}</strong> {Math.round(t.rate*100)}%</span>)}
      </div>

      <div style={{display:"flex",flexDirection:"column",gap:12}}>
        {members.map(m=>{
          const score=creditScore(m);const tier=getTier(score);
          const overdue=m.loans.some(l=>l.active&&loanDueDate(l)&&new Date(loanDueDate(l))<new Date());
          return(
            <Card key={m.id} style={{borderLeft:`4px solid ${overdue?G.red:tier.color}`}}>
              <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:10,flexWrap:"wrap"}}>
                <strong style={{fontSize:15,color:G.text}}>{m.name}</strong>
                <Pill label={`${tier.icon} ${tier.label} ${Math.round(tier.rate*100)}%`} color={tier.color} bg={tier.bg}/>
                {overdue&&<Pill label="🚨 OVERDUE" color="white" bg={G.red}/>}
                <span style={{fontSize:11,color:G.textSoft,marginLeft:"auto"}}>{m.loans.length}/2 loans used</span>
                {m.loans.length<2&&<button style={smallBtnPrim} onClick={()=>doIssue(m.id,m.name)}>+ Issue Loan @ {Math.round(tier.rate*100)}%</button>}
              </div>
              {m.loans.length===0&&<p style={{color:G.textSoft,fontSize:12,margin:0}}>No loans issued. Eligible rate: <strong style={{color:tier.color}}>{Math.round(tier.rate*100)}%</strong></p>}
              {m.loans.map((l,i)=>{
                const bal=loanBalance(l);const due=loanDueDate(l);
                const isOD=due&&new Date(due)<new Date()&&l.active;
                const key=`${m.id}-${i}`;
                return(
                  <div key={l.id} style={{padding:12,marginTop:8,background:G.bg,borderRadius:10,border:`1px solid ${G.border}`}}>
                    <div style={{display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8,alignItems:"center"}}>
                      <span style={{fontWeight:700,fontSize:13}}>Loan {i+1}: {fmt(l.amount)} @ {Math.round((l.rate||0.20)*100)}% = {fmt(l.total)}</span>
                      <span style={{color:isOD?G.red:l.active?G.gold:G.primary,fontWeight:800,fontSize:13}}>{l.active?`Balance: ${fmt(bal)}`:"✓ CLEARED"}{isOD?" 🚨":""}</span>
                    </div>
                    <div style={{fontSize:11,color:G.textSoft,marginTop:3}}>Issued: {fd(l.issuedDate)} · Due: {fd(due)}</div>
                    {(l.repayments||[]).map((r,j)=><div key={j} style={{color:G.primary,fontSize:11,marginTop:3}}>↳ {fmt(r.amount)} on {fd(r.date)}</div>)}
                    {l.active&&<div style={{display:"flex",gap:7,marginTop:8}}>
                      <input type="number" placeholder="Repayment amount" value={repayAmts[key]||""} style={{...inputStyle,marginBottom:0,flex:1}} onChange={e=>setRepayAmts(p=>({...p,[key]:e.target.value}))}/>
                      <button style={smallBtnPrim} onClick={()=>doRepay(m.id,i,m.name)}>Log Repayment</button>
                    </div>}
                  </div>
                );
              })}
            </Card>
          );
        })}
      </div>
    </div>
  );
}

/* ═══════════════════════════════════ PENALTIES */
function PenaltiesView({members,upd,toast2,log}){
  function doPen(mid,name,reason){
    upd(mid,m=>({...m,penalties:[...m.penalties,{amount:PENALTY_AMT,reason,date:new Date().toISOString()}]}));
    log("PENALTY",`€${PENALTY_AMT} — ${reason}`,mid);
    toast2(`Penalty added to ${name}`,"warn");
  }
  const all=members.flatMap(m=>m.penalties.map(p=>({...p,memberName:m.name}))).sort((a,b)=>new Date(b.date)-new Date(a.date));
  const totalPens=members.reduce((s,m)=>s+m.penalties.reduce((a,p)=>a+p.amount,0),0);

  return(
    <div>
      <PageHeader title="Penalties" sub={`€${PENALTY_AMT} per missed/late payment · Total collected: ${fmt(totalPens)}`}/>

      <Card style={{marginBottom:16}}>
        <CardTitle>Apply Penalty</CardTitle>
        <div style={{display:"flex",flexDirection:"column",gap:7}}>
          {members.map(m=>(
            <div key={m.id} style={{display:"flex",alignItems:"center",gap:10,padding:"9px 14px",background:G.bg,borderRadius:10,flexWrap:"wrap"}}>
              <strong style={{minWidth:80,fontSize:13,color:G.text}}>{m.name}</strong>
              <button style={smallBtnDanger} onClick={()=>doPen(m.id,m.name,"Late contribution")}>+ Missed Payment</button>
              <button style={smallBtnDanger} onClick={()=>doPen(m.id,m.name,"Late loan repayment")}>+ Late Loan</button>
              <span style={{marginLeft:"auto",fontSize:11,color:G.textSoft}}>Total: <strong style={{color:m.penalties.length>0?G.red:G.text}}>€{m.penalties.reduce((s,p)=>s+p.amount,0)}</strong> ({m.penalties.length}×)</span>
            </div>
          ))}
        </div>
      </Card>

      {all.length>0&&<>
        <h3 style={{fontSize:14,fontWeight:700,marginBottom:10,color:G.text}}>Penalty Log ({all.length})</h3>
        <Table cols={["Date","Member","Reason","Amount"]} rows={all.map(p=>[fd(p.date),p.memberName,p.reason,<span style={{color:G.red,fontWeight:700}}>€{p.amount}</span>])}/>
      </>}
    </div>
  );
}

/* ═══════════════════════════════════ CALENDAR */
function CalendarView({members,week,setWeek,setView}){
  const [vm,setVm]=useState(()=>{const d=SUNDAYS[week-1]?.date||new Date();return{year:d.getFullYear(),month:d.getMonth()};});
  const [sel,setSel]=useState(null);
  const MONTHS=["January","February","March","April","May","June","July","August","September","October","November","December"];
  const {year,month}=vm;
  const firstDow=new Date(year,month,1).getDay();
  const days=new Date(year,month+1,0).getDate();
  const cells=[];for(let i=0;i<firstDow;i++)cells.push(null);for(let d=1;d<=days;d++)cells.push(new Date(year,month,d));
  function ev(date){if(!date)return{week:null,paid:[],unpaid:[],loanDue:[],pens:[]};
    const idx=SUNDAYS.findIndex(s=>s.date.toDateString()===date.toDateString());
    const w=idx>=0?idx+1:null;
    return{week:w,
      paid:w?members.filter(m=>m.weeklyContributions[w]?.paid):[],
      unpaid:w?members.filter(m=>!m.weeklyContributions[w]?.paid):[],
      loanDue:members.filter(m=>m.loans.some(l=>{const dd=loanDueDate(l);return dd&&dd.toDateString()===date.toDateString();})),
      pens:members.filter(m=>m.penalties.some(p=>new Date(p.date).toDateString()===date.toDateString())),
    };
  }
  const today=new Date();const sEv=sel?ev(sel):null;
  return(
    <div>
      <PageHeader title="Payment Calendar" sub="Visual schedule of payment dates and loan deadlines"/>
      <div style={{display:"flex",gap:20,flexWrap:"wrap",alignItems:"flex-start"}}>
        <Card style={{flex:"1 1 480px",padding:18}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
            <button style={{background:G.primary,color:"white",border:"none",borderRadius:8,width:34,height:34,cursor:"pointer",fontSize:16,fontWeight:700,fontFamily:"inherit"}} onClick={()=>setVm(m=>{const d=new Date(m.year,m.month-1,1);return{year:d.getFullYear(),month:d.getMonth()};})}>‹</button>
            <h3 style={{fontSize:18,fontWeight:800,margin:0,color:G.text,fontFamily:"'Georgia',serif"}}>{MONTHS[month]} {year}</h3>
            <button style={{background:G.primary,color:"white",border:"none",borderRadius:8,width:34,height:34,cursor:"pointer",fontSize:16,fontWeight:700,fontFamily:"inherit"}} onClick={()=>setVm(m=>{const d=new Date(m.year,m.month+1,1);return{year:d.getFullYear(),month:d.getMonth()};})}>›</button>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:4}}>
            {["Sun","Mon","Tue","Wed","Thu","Fri","Sat"].map(d=><div key={d} style={{textAlign:"center",fontSize:10,fontWeight:700,color:G.textSoft,padding:"6px 0",textTransform:"uppercase",letterSpacing:.6}}>{d}</div>)}
            {cells.map((date,i)=>{
              if(!date)return<div key={`e${i}`}/>;
              const e=ev(date);
              const isToday=date.toDateString()===today.toDateString();
              const isSel=sel&&date.toDateString()===sel.toDateString();
              const isSunday=e.week;
              return(
                <div key={i} onClick={()=>setSel(date)} style={{background:isSunday?"#FEF9C3":G.surface,borderRadius:8,padding:"6px 5px",minHeight:62,cursor:"pointer",border:isToday?`2px solid ${G.primary}`:isSel?`2px solid ${G.blue}`:`1px solid ${G.border}`,transition:"all .15s"}}>
                  <div style={{fontSize:12,fontWeight:isToday?800:isSunday?700:500,color:isToday?G.primary:G.text}}>{date.getDate()}</div>
                  {e.week&&<div style={{fontSize:9,color:G.textSoft,fontWeight:600}}>Wk{e.week}</div>}
                  <div style={{display:"flex",flexWrap:"wrap",gap:2,marginTop:2}}>
                    {e.paid.length>0&&<span style={{fontSize:9,fontWeight:700,padding:"1px 4px",borderRadius:6,color:"white",background:G.primary}}>{e.paid.length}✓</span>}
                    {e.unpaid.length>0&&e.week&&<span style={{fontSize:9,fontWeight:700,padding:"1px 4px",borderRadius:6,color:"white",background:G.red}}>{e.unpaid.length}✗</span>}
                    {e.loanDue.length>0&&<span style={{fontSize:9,padding:"1px 4px",borderRadius:6,color:"white",background:G.blue}}>💳</span>}
                    {e.pens.length>0&&<span style={{fontSize:9,padding:"1px 4px",borderRadius:6,color:"white",background:G.gold}}>⚠</span>}
                  </div>
                </div>
              );
            })}
          </div>
          <div style={{display:"flex",flexWrap:"wrap",gap:14,marginTop:14,paddingTop:12,borderTop:`1px solid ${G.border}`}}>
            {[[G.primary,"Paid"],[G.red,"Unpaid"],[G.blue,"Loan due"],[G.gold,"Penalty"],["#FEF9C3","Payment Sunday"]].map(([c,l])=>(
              <div key={l} style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:G.textMid}}>
                <div style={{width:10,height:10,borderRadius:3,background:c,border:c==="#FEF9C3"?`1px solid ${G.border}`:"none"}}/>{l}
              </div>
            ))}
          </div>
        </Card>
        <div style={{flex:"0 0 280px"}}>
          {sEv?(
            <Card>
              <h3 style={{margin:"0 0 10px",fontSize:14,fontWeight:700,color:G.text}}>{sel.toLocaleDateString("en-GB",{weekday:"long",day:"numeric",month:"long"})}</h3>
              {sEv.week&&<div style={{background:"#FEF9C3",border:`1px solid #FDE047`,borderRadius:8,padding:"6px 10px",fontSize:11,fontWeight:700,marginBottom:10,color:"#854D0E"}}>📅 Week {sEv.week} Payment Day</div>}
              {sEv.week&&<>
                <div style={{borderBottom:`1px solid ${G.border}`,paddingBottom:8,marginBottom:8}}>
                  <div style={{fontWeight:700,color:G.primary,fontSize:12,marginBottom:4}}>✓ Paid ({sEv.paid.length})</div>
                  {sEv.paid.length===0?<div style={{color:G.textSoft,fontSize:11}}>None yet</div>:sEv.paid.map(m=><div key={m.id} style={{fontSize:11,padding:"1px 0"}}>{m.name} — {fmt(m.weeklyContributions[sEv.week].amount)}</div>)}
                </div>
                <div style={{borderBottom:`1px solid ${G.border}`,paddingBottom:8,marginBottom:8}}>
                  <div style={{fontWeight:700,color:G.red,fontSize:12,marginBottom:4}}>✗ Unpaid ({sEv.unpaid.length})</div>
                  {sEv.unpaid.length===0?<div style={{color:G.primary,fontSize:11}}>All paid!</div>:sEv.unpaid.map(m=><div key={m.id} style={{fontSize:11,padding:"1px 0"}}>{m.name}</div>)}
                </div>
              </>}
              {sEv.loanDue.length>0&&<div>
                <div style={{fontWeight:700,color:G.blue,fontSize:12,marginBottom:4}}>💳 Loans Due ({sEv.loanDue.length})</div>
                {sEv.loanDue.map(m=><div key={m.id} style={{fontSize:11}}>{m.name}</div>)}
              </div>}
            </Card>
          ):<Card><div style={{color:G.textSoft,textAlign:"center",padding:"24px 0"}}><div style={{fontSize:28,marginBottom:6}}>📅</div><div style={{fontSize:12}}>Click a date to view details</div></div></Card>}
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════ ANALYTICS */
function AnalyticsView({members,week}){
  const [tab,setTab]=useState("weekly");
  const weeklyData=Array.from({length:Math.min(week,16)},(_,i)=>{
    const w=week-15+i;if(w<1)return null;
    const collected=members.reduce((s,m)=>{const c=m.weeklyContributions[w];return s+(c?.paid?c.amount:0);},0);
    return{name:`W${w}`,collected,expected:members.length*60,rate:Math.round((collected/(members.length*60))*100)};
  }).filter(Boolean);
  const memberData=members.map(m=>({
    name:m.name.substring(0,6),
    paid:Object.values(m.weeklyContributions).reduce((s,c)=>s+(c.paid?c.amount:0),0),
    score:creditScore(m),
  }));
  const loanPie=[
    {name:"Active",value:members.reduce((s,m)=>s+m.loans.filter(l=>l.active).length,0),color:G.blue},
    {name:"Cleared",value:members.reduce((s,m)=>s+m.loans.filter(l=>!l.active).length,0),color:G.primary},
    {name:"No Loan",value:members.filter(m=>m.loans.length===0).length,color:G.border},
  ].filter(d=>d.value>0);
  const revData=[
    {name:"Contributions",value:members.reduce((s,m)=>s+Object.values(m.weeklyContributions).reduce((a,c)=>a+(c.paid?c.amount:0),0),0),color:G.primary},
    {name:"Loan Interest",value:members.reduce((s,m)=>s+m.loans.reduce((a,l)=>a+l.amount*(l.rate||BASE_RATE),0),0),color:G.purple},
    {name:"Penalties",value:members.reduce((s,m)=>s+m.penalties.reduce((a,p)=>a+p.amount,0),0),color:G.red},
    {name:"Membership",value:members.filter(m=>m.membershipFeePaid).length*MEM_FEE,color:G.gold},
  ];
  return(
    <div>
      <PageHeader title="Analytics" sub="Visual insights into payment trends and member behavior"/>
      <div style={{display:"flex",gap:6,marginBottom:18,flexWrap:"wrap"}}>
        {[["weekly","📅 Payment Trends"],["members","👥 Contributions"],["loans","💳 Loans"],["revenue","💰 Revenue"]].map(([t,l])=>(
          <button key={t} onClick={()=>setTab(t)} style={{padding:"8px 16px",borderRadius:9,border:`1.5px solid ${tab===t?G.primary:G.border}`,background:tab===t?G.primary:G.surface,color:tab===t?"white":G.textMid,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>{l}</button>
        ))}
      </div>

      {tab==="weekly"&&<>
        <Card style={{marginBottom:14}}>
          <CardTitle>Collection Trend</CardTitle>
          <ResponsiveContainer width="100%" height={240}>
            <AreaChart data={weeklyData}>
              <defs><linearGradient id="cg" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={G.primary} stopOpacity={.3}/><stop offset="100%" stopColor={G.primary} stopOpacity={0}/></linearGradient></defs>
              <CartesianGrid strokeDasharray="3 3" stroke={G.border}/>
              <XAxis dataKey="name" tick={{fontSize:10,fill:G.textMid}}/><YAxis tick={{fontSize:10,fill:G.textMid}}/>
              <Tooltip contentStyle={{borderRadius:8,border:`1px solid ${G.border}`}}/>
              <Area type="monotone" dataKey="collected" stroke={G.primary} strokeWidth={2.5} fill="url(#cg)" name="Collected (€)"/>
              <Line type="monotone" dataKey="expected" stroke={G.textSoft} strokeWidth={1.5} strokeDasharray="4 3" dot={false} name="Expected (€)"/>
            </AreaChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <CardTitle>Collection Rate %</CardTitle>
          <ResponsiveContainer width="100%" height={180}>
            <BarChart data={weeklyData}>
              <CartesianGrid strokeDasharray="3 3" stroke={G.border}/>
              <XAxis dataKey="name" tick={{fontSize:10,fill:G.textMid}}/><YAxis domain={[0,100]} tick={{fontSize:10,fill:G.textMid}} unit="%"/>
              <Tooltip contentStyle={{borderRadius:8,border:`1px solid ${G.border}`}}/>
              <Bar dataKey="rate" radius={[6,6,0,0]}>{weeklyData.map((e,i)=><Cell key={i} fill={e.rate>=90?G.primary:e.rate>=70?G.gold:G.red}/>)}</Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </>}

      {tab==="members"&&<>
        <Card style={{marginBottom:14}}>
          <CardTitle>Total Contributions per Member</CardTitle>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={memberData}>
              <CartesianGrid strokeDasharray="3 3" stroke={G.border}/>
              <XAxis dataKey="name" tick={{fontSize:10,fill:G.textMid}}/><YAxis tick={{fontSize:10,fill:G.textMid}}/>
              <Tooltip contentStyle={{borderRadius:8,border:`1px solid ${G.border}`}} formatter={v=>"€"+v}/>
              <Bar dataKey="paid" fill={G.primary} radius={[6,6,0,0]} name="Paid (€)"/>
            </BarChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <CardTitle>Credit Score per Member</CardTitle>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={memberData}>
              <CartesianGrid strokeDasharray="3 3" stroke={G.border}/>
              <XAxis dataKey="name" tick={{fontSize:10,fill:G.textMid}}/><YAxis domain={[0,100]} tick={{fontSize:10,fill:G.textMid}}/>
              <Tooltip contentStyle={{borderRadius:8,border:`1px solid ${G.border}`}}/>
              <Bar dataKey="score" radius={[6,6,0,0]}>{memberData.map((e,i)=><Cell key={i} fill={getTier(e.score).color}/>)}</Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </>}

      {tab==="loans"&&<div style={{display:"flex",gap:14,flexWrap:"wrap"}}>
        <Card style={{flex:"1 1 280px"}}>
          <CardTitle>Loan Status</CardTitle>
          <ResponsiveContainer width="100%" height={240}>
            <PieChart>
              <Pie data={loanPie} cx="50%" cy="50%" outerRadius={80} dataKey="value" label={({name,value})=>`${name}: ${value}`}>{loanPie.map((e,i)=><Cell key={i} fill={e.color}/>)}</Pie>
              <Tooltip contentStyle={{borderRadius:8,border:`1px solid ${G.border}`}}/>
            </PieChart>
          </ResponsiveContainer>
        </Card>
        <Card style={{flex:"1 1 280px"}}>
          <CardTitle>Loan Balance per Member</CardTitle>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={members.map(m=>({name:m.name.substring(0,6),bal:+m.loans.reduce((s,l)=>s+loanBalance(l),0).toFixed(0)}))}>
              <CartesianGrid strokeDasharray="3 3" stroke={G.border}/>
              <XAxis dataKey="name" tick={{fontSize:10,fill:G.textMid}}/><YAxis tick={{fontSize:10,fill:G.textMid}}/>
              <Tooltip contentStyle={{borderRadius:8,border:`1px solid ${G.border}`}} formatter={v=>"€"+v}/>
              <Bar dataKey="bal" fill={G.blue} radius={[6,6,0,0]} name="Balance (€)"/>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>}

      {tab==="revenue"&&<div style={{display:"flex",gap:14,flexWrap:"wrap"}}>
        <Card style={{flex:"1 1 280px"}}>
          <CardTitle>Revenue Streams</CardTitle>
          <ResponsiveContainer width="100%" height={250}>
            <PieChart>
              <Pie data={revData} cx="50%" cy="50%" outerRadius={88} dataKey="value" label={({value})=>`€${value.toFixed(0)}`}>{revData.map((e,i)=><Cell key={i} fill={e.color}/>)}</Pie>
              <Tooltip contentStyle={{borderRadius:8,border:`1px solid ${G.border}`}} formatter={v=>"€"+Number(v).toFixed(2)}/>
              <Legend/>
            </PieChart>
          </ResponsiveContainer>
        </Card>
        <Card style={{flex:"1 1 280px"}}>
          <CardTitle>Revenue Breakdown</CardTitle>
          <ResponsiveContainer width="100%" height={250}>
            <BarChart data={revData}>
              <CartesianGrid strokeDasharray="3 3" stroke={G.border}/>
              <XAxis dataKey="name" tick={{fontSize:10,fill:G.textMid}}/><YAxis tick={{fontSize:10,fill:G.textMid}}/>
              <Tooltip contentStyle={{borderRadius:8,border:`1px solid ${G.border}`}} formatter={v=>"€"+Number(v).toFixed(2)}/>
              <Bar dataKey="value" radius={[6,6,0,0]}>{revData.map((e,i)=><Cell key={i} fill={e.color}/>)}</Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>}
    </div>
  );
}

/* ═══════════════════════════════════ PROFIT OPTIMIZER */
function ProfitView({members,totalLoansOut,totalInterest,idleCapital}){
  const scenarios=[
    {name:"Current",rev:24*LOAN_AMOUNT*0.20,color:G.textSoft},
    {name:"Dynamic Rates",rev:24*LOAN_AMOUNT*0.175+3*35,color:G.blue},
    {name:"+ Origination 2.5%",rev:24*LOAN_AMOUNT*0.175+3*35+24*LOAN_AMOUNT*0.025,color:G.purple},
    {name:"+ P2P Deployment",rev:24*LOAN_AMOUNT*0.175+3*35+24*LOAN_AMOUNT*0.025+650,color:G.accent},
    {name:"Full Optimization",rev:28*LOAN_AMOUNT*0.19+5*35+28*LOAN_AMOUNT*0.025+1200,color:G.primary},
  ];
  const strategies=[
    {icon:"⭐",color:G.purple,title:"Dynamic Credit-Based Rates",cur:"Flat 20%",prop:"12-25% by tier",impact:`Platinum 12%, Bronze 25%. Estimated +${fmt(24*LOAN_AMOUNT*0.03)} per season.`},
    {icon:"⚠️",color:G.red,title:"Escalating Penalty Structure",cur:"Flat €20",prop:"€20 / €35 / €50",impact:"Stronger deterrent. ~€420 additional penalty revenue."},
    {icon:"💳",color:G.gold,title:"Loan Origination Fee (2.5%)",cur:"No fee",prop:"€93.75 per loan upfront",impact:`24 loans × 2.5% = +${fmt(24*LOAN_AMOUNT*0.025)} non-refundable upfront.`},
    {icon:"📈",color:G.accent,title:"Premium 3rd Loan Tier",cur:"Max 2 loans",prop:"Optional €2k @ 28%",impact:`4 members × ${fmt(2000*0.28)} = ${fmt(4*2000*0.28)} additional interest.`},
    {icon:"💰",color:G.primary,title:"Ring-Fence Penalties as Profit",cur:"Pooled with contributions",prop:"50% to zero-penalty members",impact:"Turns penalties into loyalty bonus. Drives compliance."},
    {icon:"🌐",color:G.blue,title:"Deploy Idle to P2P (11-15%)",cur:fmt(idleCapital)+" idle",prop:"Mintos + Nectaro",impact:`If ${fmt(Math.min(idleCapital,5000))} deployed at 13%: +${fmt(Math.min(idleCapital,5000)*0.13)}/yr passive.`},
  ];
  const uplift=scenarios[scenarios.length-1].rev-scenarios[0].rev;
  return(
    <div>
      <PageHeader title="🎯 Profit Optimizer" sub="The group's purpose is profit generation — these strategies maximize returns"/>

      <Card style={{marginBottom:16,background:`linear-gradient(135deg,${G.primary} 0%,${G.primaryD} 100%)`,color:"white",border:"none"}}>
        <div style={{display:"flex",alignItems:"center",gap:16,flexWrap:"wrap"}}>
          <div style={{fontSize:36}}>🎯</div>
          <div style={{flex:1,minWidth:200}}>
            <div style={{fontSize:12,opacity:.9,textTransform:"uppercase",letterSpacing:.8,fontWeight:600}}>Maximum Potential Uplift</div>
            <div style={{fontSize:30,fontWeight:800,fontFamily:"'Georgia',serif",margin:"4px 0"}}>{fmt(uplift)}</div>
            <div style={{fontSize:12,opacity:.9}}>{Math.round((scenarios[scenarios.length-1].rev/scenarios[0].rev-1)*100)}% increase from current revenue, same {members.length} members</div>
          </div>
        </div>
      </Card>

      <Card style={{marginBottom:16}}>
        <CardTitle>📊 Revenue Scenario Projection</CardTitle>
        <ResponsiveContainer width="100%" height={260}>
          <BarChart data={scenarios} margin={{top:10,right:20,left:10,bottom:50}}>
            <CartesianGrid strokeDasharray="3 3" stroke={G.border}/>
            <XAxis dataKey="name" tick={{fontSize:10,fill:G.textMid}} angle={-15} textAnchor="end" height={60}/>
            <YAxis tick={{fontSize:10,fill:G.textMid}} tickFormatter={v=>"€"+v.toFixed(0)}/>
            <Tooltip contentStyle={{borderRadius:8,border:`1px solid ${G.border}`}} formatter={v=>"€"+Number(v).toFixed(2)}/>
            <Bar dataKey="rev" radius={[8,8,0,0]}>{scenarios.map((s,i)=><Cell key={i} fill={s.color}/>)}</Bar>
          </BarChart>
        </ResponsiveContainer>
      </Card>

      <h3 style={{fontSize:15,fontWeight:700,color:G.text,margin:"24px 0 14px"}}>💡 Growth Strategies</h3>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(320px,1fr))",gap:14}}>
        {strategies.map((s,i)=>(
          <Card key={i} style={{borderLeft:`4px solid ${s.color}`}}>
            <div style={{display:"flex",gap:10,alignItems:"flex-start",marginBottom:10}}>
              <div style={{width:38,height:38,borderRadius:10,background:G.bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>{s.icon}</div>
              <h4 style={{margin:0,fontSize:14,fontWeight:800,color:G.text,fontFamily:"'Georgia',serif"}}>{s.title}</h4>
            </div>
            <div style={{display:"flex",gap:6,marginBottom:8,flexWrap:"wrap"}}>
              <Pill label={`Now: ${s.cur}`} color={G.red} bg="#FEF2F2"/>
              <Pill label={`Proposed: ${s.prop}`} color={G.primary} bg={G.primaryL}/>
            </div>
            <p style={{margin:0,fontSize:12,color:G.textMid,lineHeight:1.6}}>{s.impact}</p>
          </Card>
        ))}
      </div>
    </div>
  );
}

/* ═══════════════════════════════════ CREDIT SCORING */
function CreditView({members}){
  const scored=[...members].map(m=>({...m,score:creditScore(m),tier:getTier(creditScore(m))})).sort((a,b)=>b.score-a.score);
  return(
    <div>
      <PageHeader title="Credit Scoring" sub="Dynamic interest rates based on member payment behavior"/>

      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:12,marginBottom:18}}>
        {TIERS.map(t=>{
          const count=scored.filter(m=>m.tier.label===t.label).length;
          return(
            <Card key={t.label} style={{borderTop:`4px solid ${t.color}`}}>
              <div style={{fontSize:28,marginBottom:4}}>{t.icon}</div>
              <div style={{fontSize:24,fontWeight:800,color:t.color,fontFamily:"'Georgia',serif"}}>{count}</div>
              <div style={{fontSize:13,fontWeight:700,color:G.text}}>{t.label}</div>
              <div style={{fontSize:11,color:G.textSoft,marginTop:2}}>{Math.round(t.rate*100)}% rate · score ≥{t.min}</div>
            </Card>
          );
        })}
      </div>

      <Table
        cols={["Rank","Member","Score","Tier","Loan Rate","Weeks Paid","Penalties","Loans Cleared"]}
        rows={scored.map((m,i)=>{
          const wp=Object.values(m.weeklyContributions).filter(c=>c.paid).length;
          const cleared=m.loans.filter(l=>!l.active).length;
          return[
            <span style={{fontWeight:800,color:i<3?G.gold:G.textSoft,fontSize:14}}>#{i+1}</span>,
            <div style={{fontWeight:700,color:G.text}}>{m.name}</div>,
            <ScoreBar score={m.score}/>,
            <Pill label={`${m.tier.icon} ${m.tier.label}`} color={m.tier.color} bg={m.tier.bg}/>,
            <span style={{fontWeight:700,color:m.tier.color,fontSize:14}}>{Math.round(m.tier.rate*100)}%</span>,
            <span>{wp}/52</span>,
            <span style={{color:m.penalties.length>0?G.red:G.textSoft}}>{m.penalties.length}</span>,
            <span style={{color:G.primary,fontWeight:600}}>{cleared}</span>
          ];
        })}
      />

      <Card style={{marginTop:18}}>
        <CardTitle>📐 Scoring Formula</CardTitle>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:10}}>
          {[
            ["Base","60 pts","Starting baseline"],
            ["Payment Rate","+25 max","Scaled to consistency"],
            ["Loans Cleared","+10 each","Up to +20"],
            ["Membership Fee","+5","Paid registration"],
            ["Each Penalty","-5","Up to -30"],
            ["Overdue Loan","-20","Past 3-month deadline"],
          ].map(([t,v,d])=>(
            <div key={t} style={{padding:10,borderRadius:9,background:G.bg}}>
              <div style={{fontWeight:700,fontSize:11,color:G.text}}>{t}</div>
              <div style={{color:G.primary,fontWeight:800,fontSize:15,margin:"2px 0"}}>{v}</div>
              <div style={{color:G.textSoft,fontSize:10}}>{d}</div>
            </div>
          ))}
        </div>
      </Card>
    </div>
  );
}

/* ═══════════════════════════════════ P2P INVESTMENT */
function P2PView({idleCapital}){
  const platforms=[
    {name:"Nectaro",ret:"14.9%",reg:"MiFID II",min:"€50",liq:"Secondary market",badge:"⭐ Top 2025",color:G.purple,rec:"Highest regulated return"},
    {name:"Mintos",ret:"11–12%",reg:"MiFID II",min:"€10",liq:"Secondary market",badge:"🏆 Largest in EU",color:G.blue,rec:"Most liquid, 700k+ investors"},
    {name:"PeerBerry",ret:"11.14%",reg:"Unregulated",min:"€10",liq:"None",badge:"✅ Zero capital loss",color:G.primary,rec:"Solid safety record"},
    {name:"Hive5",ret:"14–16%",reg:"Unregulated",min:"€10",liq:"None",badge:"🔥 Highest yield",color:G.gold,rec:"Best yields, short history"},
    {name:"Income Marketplace",ret:"11–12%",reg:"ECSPR",min:"€10",liq:"Secondary",badge:"🛡 Cashflow Buffer",color:G.accent,rec:"Strongest protection"},
    {name:"Debitum",ret:"11–12%",reg:"MiFID II",min:"€50",liq:"Secondary",badge:"🏢 SME loans",color:"#5A7A8A",rec:"Business loan diversification"},
  ];
  return(
    <div>
      <PageHeader title="📈 P2P Investment" sub="Deploy idle capital to regulated European P2P platforms for additional returns"/>

      <Card style={{marginBottom:16,background:`linear-gradient(135deg,${G.blue} 0%,#1E40AF 100%)`,color:"white",border:"none"}}>
        <div style={{display:"flex",alignItems:"center",gap:18,flexWrap:"wrap"}}>
          <div style={{fontSize:36}}>💤</div>
          <div style={{flex:1,minWidth:200}}>
            <div style={{fontSize:12,opacity:.9,textTransform:"uppercase",letterSpacing:.8,fontWeight:600}}>Estimated Idle Capital</div>
            <div style={{fontSize:30,fontWeight:800,fontFamily:"'Georgia',serif",margin:"4px 0"}}>{fmt(idleCapital)}</div>
            <div style={{fontSize:12,opacity:.9}}>At 13% avg yield: <strong>+{fmt(idleCapital*0.13)}/year</strong> passive income</div>
          </div>
        </div>
      </Card>

      <div style={{background:"#FEF3C7",border:`1px solid #FCD34D`,borderRadius:12,padding:"12px 16px",marginBottom:18,fontSize:12,color:"#92400E"}}>
        <strong>⚠️ Risk Notice:</strong> P2P platforms carry investment risk. Only deploy capital not needed within 30 days. Past returns don't guarantee future performance. All recommendations below are EU-based and have been operating for 3+ years.
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(310px,1fr))",gap:14,marginBottom:18}}>
        {platforms.map((p,i)=>(
          <Card key={i} style={{borderTop:`4px solid ${p.color}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
              <h3 style={{margin:0,fontWeight:800,fontSize:18,color:p.color,fontFamily:"'Georgia',serif"}}>{p.name}</h3>
              <Pill label={p.badge} color={G.textMid} bg={G.bg}/>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,fontSize:11,marginBottom:10}}>
              {[["Return",p.ret,G.primary],["Min Invest",p.min,G.blue],["Regulation",p.reg,G.purple],["Liquidity",p.liq,G.gold]].map(([l,v,c])=>(
                <div key={l} style={{background:G.bg,padding:"6px 9px",borderRadius:7}}>
                  <div style={{fontSize:9,color:G.textSoft,textTransform:"uppercase",letterSpacing:.6,fontWeight:600}}>{l}</div>
                  <div style={{fontWeight:700,color:c,fontSize:12}}>{v}</div>
                </div>
              ))}
            </div>
            <div style={{fontSize:11,color:p.color,fontWeight:600}}>★ {p.rec}</div>
          </Card>
        ))}
      </div>

      <Card>
        <CardTitle>📋 Recommended Deployment Strategy</CardTitle>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:10}}>
          {[
            ["1. Identify Idle Capital","Calculate weekly surplus between loan cycles. Typically €1k–€5k sits idle."],
            ["2. Open Group Account","Register on Mintos (most liquid) + Nectaro (highest yield). Standard KYC for groups."],
            ["3. AutoInvest 30-60 Day Terms","Short terms keep capital accessible. Target ≥11% interest rate filter."],
            ["4. Withdraw Monthly","Log as 'P2P Revenue'. Distribute as profit at season end."],
            ["5. Diversify 60/40","60% Mintos (regulated, liquid) + 40% Nectaro (higher yield). Never >70% of idle."],
          ].map(([t,d])=>(
            <div key={t} style={{background:G.bg,borderRadius:9,padding:11,borderLeft:`3px solid ${G.primary}`}}>
              <div style={{fontWeight:700,fontSize:12,marginBottom:4,color:G.text}}>{t}</div>
              <div style={{fontSize:11,color:G.textMid,lineHeight:1.6}}>{d}</div>
            </div>
          ))}
        </div>
      </Card>
    </div>
  );
}

/* ═══════════════════════════════════ CASHFLOW */
function CashflowView({members,week}){
  // Weekly cashflow data
  const cashflowData=Array.from({length:Math.min(week+4,52)},(_,i)=>{
    const w=i+1;
    const inflows=members.reduce((s,m)=>{
      const c=m.weeklyContributions[w];
      const cb=s+(c?.paid?c.amount:0);
      const lr=m.loans.reduce((a,l)=>a+(l.repayments||[]).filter(r=>{const rw=Math.floor((new Date(r.date)-SUNDAYS[0].date)/(7*24*60*60*1000))+1;return rw===w;}).reduce((b,r)=>b+r.amount,0),0);
      return cb+lr;
    },0);
    const outflows=members.reduce((s,m)=>{
      return s+m.loans.filter(l=>{const iw=Math.floor((new Date(l.issuedDate)-SUNDAYS[0].date)/(7*24*60*60*1000))+1;return iw===w;}).reduce((a,l)=>a+l.amount,0);
    },0);
    return{name:`W${w}`,inflows,outflows:-outflows,net:inflows-outflows};
  });

  let runningBalance=0;
  const balanceData=cashflowData.map(d=>{runningBalance+=d.net;return{...d,balance:runningBalance};});

  const totalIn=cashflowData.reduce((s,d)=>s+d.inflows,0);
  const totalOut=cashflowData.reduce((s,d)=>s+Math.abs(d.outflows),0);

  return(
    <div>
      <PageHeader title="〰️ Cashflow" sub="Weekly inflows, outflows, and running balance"/>

      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))",gap:12,marginBottom:18}}>
        <StatCard icon="📥" label="Total Inflows" value={fmt(totalIn)} color={G.primary}/>
        <StatCard icon="📤" label="Total Outflows" value={fmt(totalOut)} color={G.red}/>
        <StatCard icon="💰" label="Net Position" value={fmt(totalIn-totalOut)} color={totalIn-totalOut>=0?G.primary:G.red}/>
        <StatCard icon="📊" label="Current Balance" value={fmt(runningBalance)} color={G.blue}/>
      </div>

      <Card style={{marginBottom:16}}>
        <CardTitle>📈 Weekly Cashflow (Inflows vs Outflows)</CardTitle>
        <ResponsiveContainer width="100%" height={260}>
          <BarChart data={cashflowData}>
            <CartesianGrid strokeDasharray="3 3" stroke={G.border}/>
            <XAxis dataKey="name" tick={{fontSize:10,fill:G.textMid}}/>
            <YAxis tick={{fontSize:10,fill:G.textMid}} tickFormatter={v=>"€"+v}/>
            <Tooltip contentStyle={{borderRadius:8,border:`1px solid ${G.border}`}} formatter={v=>"€"+Number(Math.abs(v)).toFixed(2)}/>
            <Legend/>
            <Bar dataKey="inflows" stackId="cf" fill={G.primary} name="Inflows" radius={[6,6,0,0]}/>
            <Bar dataKey="outflows" stackId="cf" fill={G.red} name="Outflows" radius={[0,0,6,6]}/>
          </BarChart>
        </ResponsiveContainer>
      </Card>

      <Card style={{marginBottom:16}}>
        <CardTitle>💰 Running Balance Over Time</CardTitle>
        <ResponsiveContainer width="100%" height={220}>
          <AreaChart data={balanceData}>
            <defs><linearGradient id="bg" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={G.blue} stopOpacity={.4}/><stop offset="100%" stopColor={G.blue} stopOpacity={0}/></linearGradient></defs>
            <CartesianGrid strokeDasharray="3 3" stroke={G.border}/>
            <XAxis dataKey="name" tick={{fontSize:10,fill:G.textMid}}/>
            <YAxis tick={{fontSize:10,fill:G.textMid}} tickFormatter={v=>"€"+v}/>
            <Tooltip contentStyle={{borderRadius:8,border:`1px solid ${G.border}`}} formatter={v=>"€"+Number(v).toFixed(2)}/>
            <Area type="monotone" dataKey="balance" stroke={G.blue} strokeWidth={2.5} fill="url(#bg)" name="Balance (€)"/>
          </AreaChart>
        </ResponsiveContainer>
      </Card>

      <Card>
        <CardTitle>📋 Weekly Breakdown</CardTitle>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{borderBottom:`2px solid ${G.border}`}}>
              {["Week","Inflows","Outflows","Net","Balance"].map(h=><th key={h} style={{padding:"8px 12px",textAlign:"left",fontWeight:700,color:G.textSoft,fontSize:10,textTransform:"uppercase",letterSpacing:.6}}>{h}</th>)}
            </tr></thead>
            <tbody>{balanceData.filter(d=>d.inflows>0||Math.abs(d.outflows)>0).map((d,i)=>(
              <tr key={i} style={{borderBottom:`1px solid ${G.border}`}}>
                <td style={{padding:"8px 12px",fontWeight:600}}>{d.name}</td>
                <td style={{padding:"8px 12px",color:G.primary,fontWeight:600}}>{fmt(d.inflows)}</td>
                <td style={{padding:"8px 12px",color:G.red,fontWeight:600}}>{fmt(Math.abs(d.outflows))}</td>
                <td style={{padding:"8px 12px",fontWeight:700,color:d.net>=0?G.primary:G.red}}>{d.net>=0?"+":""}{fmt(d.net)}</td>
                <td style={{padding:"8px 12px",fontWeight:700,color:G.blue}}>{fmt(d.balance)}</td>
              </tr>
            ))}</tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

/* ═══════════════════════════════════ AUDIT LOG */
function AuditView({auditLog,members}){
  const [filter,setFilter]=useState("all");
  const [search,setSearch]=useState("");
  let filtered=auditLog;
  if(filter!=="all")filtered=filtered.filter(a=>a.action===filter);
  if(search)filtered=filtered.filter(a=>JSON.stringify(a).toLowerCase().includes(search.toLowerCase()));

  function exportCSV(){
    const headers=["Timestamp","Admin","Action","Detail","Member"];
    const rows=filtered.map(a=>[fd(a.ts)+" "+new Date(a.ts).toLocaleTimeString("en-GB"),a.admin,a.action,a.detail,a.memberId?members.find(m=>m.id===a.memberId)?.name||"":""]);
    const csv=[headers,...rows].map(r=>r.map(c=>`"${c}"`).join(",")).join("\n");
    const blob=new Blob([csv],{type:"text/csv"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");a.href=url;a.download=`AuditLog_${new Date().toISOString().split("T")[0]}.csv`;a.click();URL.revokeObjectURL(url);
  }
  function exportExcel(){
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(filtered.map(a=>({
      Timestamp:fd(a.ts)+" "+new Date(a.ts).toLocaleTimeString("en-GB"),
      Admin:a.admin,Action:a.action,Detail:a.detail,
      Member:a.memberId?members.find(m=>m.id===a.memberId)?.name||"":""
    }))),"Audit Log");
    XLSX.writeFile(wb,`AuditLog_${new Date().toISOString().split("T")[0]}.xlsx`);
  }

  const actions=["all",...Array.from(new Set(auditLog.map(a=>a.action)))];
  const colorFor=a=>a==="WEEKLY"?G.primary:a==="LOAN_ISSUED"?G.blue:a==="REPAYMENT"?G.purple:a==="PENALTY"?G.red:a==="FEE"?G.gold:G.textMid;

  return(
    <div>
      <PageHeader title="📋 Audit Log" sub={`${auditLog.length} actions logged · Full timestamped history`}/>

      <Card style={{marginBottom:14}}>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
          <input placeholder="🔍 Search audit log..." value={search} onChange={e=>setSearch(e.target.value)} style={{...inputStyle,marginBottom:0,flex:"1 1 200px"}}/>
          <select value={filter} onChange={e=>setFilter(e.target.value)} style={{...inputStyle,marginBottom:0,width:160}}>
            {actions.map(a=><option key={a} value={a}>{a==="all"?"All Actions":a}</option>)}
          </select>
          <button style={{padding:"9px 14px",borderRadius:9,border:"none",background:G.primary,color:"white",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}} onClick={exportExcel}>📊 Export Excel</button>
          <button style={{padding:"9px 14px",borderRadius:9,border:`1px solid ${G.border}`,background:G.bg,color:G.textMid,fontSize:12,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}} onClick={exportCSV}>📋 CSV</button>
        </div>
      </Card>

      <div style={{background:G.surface,borderRadius:14,border:`1px solid ${G.border}`,boxShadow:G.shadow,overflow:"hidden"}}>
        {filtered.length===0?<div style={{padding:40,textAlign:"center",color:G.textSoft}}><div style={{fontSize:32,marginBottom:6}}>📋</div>No audit entries match</div>:
        filtered.slice(0,200).map((a,i)=>(
          <div key={a.id} style={{display:"flex",gap:14,padding:"12px 18px",borderBottom:i<filtered.length-1&&i<199?`1px solid ${G.border}`:"none",alignItems:"center"}} onMouseEnter={e=>e.currentTarget.style.background=G.bg} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
            <div style={{width:36,height:36,borderRadius:10,background:colorFor(a.action)+"22",color:colorFor(a.action),display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,fontSize:14,fontWeight:700}}>{a.action.charAt(0)}</div>
            <div style={{flex:1,minWidth:0}}>
              <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap",marginBottom:2}}>
                <Pill label={a.action} color={colorFor(a.action)} bg={colorFor(a.action)+"15"}/>
                <span style={{fontSize:13,color:G.text,fontWeight:600}}>{a.detail}</span>
                {a.memberId&&<span style={{fontSize:11,color:G.textSoft}}>· {members.find(m=>m.id===a.memberId)?.name}</span>}
              </div>
              <div style={{fontSize:11,color:G.textSoft}}>by {a.admin} · {fd(a.ts)} {new Date(a.ts).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"})}</div>
            </div>
          </div>
        ))}
      </div>
      {filtered.length>200&&<p style={{fontSize:11,color:G.textSoft,marginTop:10,textAlign:"center"}}>Showing 200 of {filtered.length} entries · Export for full data</p>}
    </div>
  );
}

/* ═══════════════════════════════════ MEMBER PORTAL (READ-ONLY) */
function MemberPortal({member:m,members,onClose}){
  const score=creditScore(m);const tier=getTier(score);
  const totalContribs=Object.values(m.weeklyContributions).reduce((s,c)=>s+(c.paid?c.amount:0),0);
  const totalPens=m.penalties.reduce((s,p)=>s+p.amount,0);
  const loanBal=m.loans.reduce((s,l)=>s+loanBalance(l),0);
  const weeksPaid=Object.values(m.weeklyContributions).filter(c=>c.paid).length;
  const rank=[...members].map(mm=>({id:mm.id,s:creditScore(mm)})).sort((a,b)=>b.s-a.s).findIndex(mm=>mm.id===m.id)+1;
  const chartData=Array.from({length:12},(_,i)=>{const w=i+1;const c=m.weeklyContributions[w];return{name:`W${w}`,amount:c?.paid?c.amount:0};});
  return(
    <div style={{position:"fixed",inset:0,background:G.bg,overflowY:"auto",zIndex:200,padding:24,fontFamily:"'DM Sans','Helvetica Neue',sans-serif"}}>
      <div style={{maxWidth:780,margin:"0 auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:22}}>
          <div>
            <h1 style={{color:G.text,margin:0,fontFamily:"'Georgia',serif",fontSize:26,letterSpacing:-.3}}>{m.name}'s Dashboard</h1>
            <p style={{color:G.textSoft,margin:"3px 0 0",fontSize:13}}>Read-only · Khobidi Savings Group {SEASON_START}</p>
          </div>
          <button style={{background:G.text,color:"white",border:"none",borderRadius:10,padding:"9px 18px",cursor:"pointer",fontSize:12,fontWeight:600,fontFamily:"inherit"}} onClick={onClose}>✕ Close</button>
        </div>

        <Card style={{borderLeft:`5px solid ${tier.color}`,marginBottom:14}}>
          <div style={{display:"flex",gap:18,alignItems:"center",flexWrap:"wrap"}}>
            <div style={{width:74,height:74,borderRadius:18,background:tier.bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:38}}>{tier.icon}</div>
            <div style={{flex:1,minWidth:180}}>
              <div style={{color:G.textSoft,fontSize:11,textTransform:"uppercase",letterSpacing:1,fontWeight:600}}>Your Credit Rating</div>
              <div style={{color:tier.color,fontWeight:800,fontSize:28,fontFamily:"'Georgia',serif",margin:"2px 0"}}>{tier.label}</div>
              <div style={{fontSize:13,color:G.textMid}}>Score: <strong>{score}/100</strong> · Rank #{rank} of {members.length}</div>
            </div>
            <div style={{textAlign:"right"}}>
              <div style={{color:G.textSoft,fontSize:11,textTransform:"uppercase",letterSpacing:1,fontWeight:600}}>Loan Rate</div>
              <div style={{color:tier.color,fontWeight:800,fontSize:30,fontFamily:"'Georgia',serif"}}>{Math.round(tier.rate*100)}%</div>
              <div style={{color:G.textSoft,fontSize:11}}>on €{LOAN_AMOUNT.toLocaleString()}</div>
            </div>
          </div>
          <div style={{background:G.bg,borderRadius:10,height:10,marginTop:16,overflow:"hidden"}}>
            <div style={{width:score+"%",height:"100%",borderRadius:10,background:tier.color,transition:"width .5s"}}/>
          </div>
        </Card>

        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:12,marginBottom:14}}>
          <StatCard icon="💰" label="Contributions" value={fmt(totalContribs)} color={G.primary}/>
          <StatCard icon="📅" label="Weeks Paid" value={`${weeksPaid}/52`} color={G.blue}/>
          <StatCard icon="💳" label="Loan Balance" value={fmt(loanBal)} color={G.purple}/>
          <StatCard icon="⚠️" label="Penalties" value={fmt(totalPens)} color={G.red}/>
        </div>

        <Card style={{marginBottom:14}}>
          <CardTitle>📊 My Contributions (first 12 weeks)</CardTitle>
          <ResponsiveContainer width="100%" height={160}>
            <BarChart data={chartData}>
              <XAxis dataKey="name" tick={{fontSize:10,fill:G.textMid}}/>
              <YAxis tick={{fontSize:10,fill:G.textMid}}/>
              <Tooltip contentStyle={{borderRadius:8,border:`1px solid ${G.border}`}} formatter={v=>"€"+v}/>
              <Bar dataKey="amount" radius={[6,6,0,0]}>{chartData.map((e,i)=><Cell key={i} fill={e.amount>0?G.primary:G.border}/>)}</Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>

        {m.loans.length>0&&<Card style={{marginBottom:14}}>
          <CardTitle>💳 My Loans</CardTitle>
          {m.loans.map((l,i)=>{
            const bal=loanBalance(l);const pct=Math.round(((l.total-bal)/l.total)*100);
            return(
              <div key={l.id} style={{marginBottom:12}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:13,marginBottom:6}}>
                  <span style={{fontWeight:700}}>Loan {i+1}: {fmt(l.amount)} @ {Math.round((l.rate||0.20)*100)}%</span>
                  <Pill label={l.active?"Active":"Cleared ✓"} color={l.active?G.gold:G.primary} bg={l.active?"#FDF6E9":G.primaryL}/>
                </div>
                <div style={{background:G.bg,borderRadius:8,height:8,overflow:"hidden"}}>
                  <div style={{width:pct+"%",height:"100%",borderRadius:8,background:G.primary,transition:"width .5s"}}/>
                </div>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:G.textSoft,marginTop:3}}>
                  <span>Repaid: {fmt(l.total-bal)}</span><span>{pct}% cleared</span><span>Owed: {fmt(bal)}</span>
                </div>
              </div>
            );
          })}
        </Card>}

        <div style={{background:G.bg,borderRadius:10,padding:12,fontSize:11,color:G.textSoft,textAlign:"center"}}>
          🔒 This is a read-only view. Contact your admin for any changes or questions.
        </div>
      </div>
    </div>
  );
}