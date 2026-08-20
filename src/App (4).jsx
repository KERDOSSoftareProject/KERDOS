import { useState, useEffect, useMemo } from "react";
import { createClient } from "@supabase/supabase-js";
import { parseDocument } from "./ingestion.js";

const supabase = createClient(
  "https://antpbtorhqghrjqzftub.supabase.co",
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImFudHBidG9yaHFnaHJqcXpmdHViIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQyOTMxMjMsImV4cCI6MjA5OTg2OTEyM30.vt4tdD7IaUKalzmpQrl5vD_hBb1lCOdu7D_LooC8qOQ"
);

function loadFileLib() {
  return new Promise((resolve, reject) => {
    if (window.XLSX) { resolve(window.XLSX); return; }
    const script = document.createElement("script");
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    script.onload = () => resolve(window.XLSX);
    script.onerror = reject;
    document.head.appendChild(script);
  });
}

function loadScript(src, globalName) {
  return new Promise((resolve, reject) => {
    if (window[globalName]) { resolve(window[globalName]); return; }
    const script = document.createElement("script");
    script.src = src;
    script.onload = () => resolve(window[globalName]);
    script.onerror = reject;
    document.head.appendChild(script);
  });
}

async function extractPdfText(file) {
  const pdfjsLib = await loadScript(
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js", "pdfjsLib"
  );
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;

  let fullText = "";
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const lineGroups = [];
    for (const it of content.items) {
      const y = it.transform[5];
      const x = it.transform[4];
      let group = lineGroups.find(g => Math.abs(g.y - y) < 3);
      if (!group) { group = { y, words: [] }; lineGroups.push(group); }
      group.words.push({ x, str: it.str });
    }
    lineGroups.sort((a, b) => b.y - a.y);
    for (const group of lineGroups) {
      group.words.sort((a, b) => a.x - b.x);
      fullText += group.words.map(w => w.str).join(" ") + "\n";
    }
  }

  if (fullText.trim().length > 40) {
    return fullText;
  }

  await loadScript(
    "https://cdnjs.cloudflare.com/ajax/libs/tesseract.js/4.1.1/tesseract.min.js", "Tesseract"
  );
  let ocrText = "";
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale: 2 });
    const canvas = document.createElement("canvas");
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const ctx = canvas.getContext("2d");
    await page.render({ canvasContext: ctx, viewport }).promise;
    const { data: { text } } = await window.Tesseract.recognize(canvas, "eng");
    ocrText += text + "\n";
  }
  return ocrText;
}

async function fileToText(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
    const XLSX = await loadFileLib();
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_csv(sheet);
  }
  if (name.endsWith(".pdf")) {
    return await extractPdfText(file);
  }
  return await file.text();
}

// ── FILE STORAGE ────────────────────────────────────────────────────
// Uploads the original, untouched file to Supabase Storage so it can
// be reopened later exactly as it was received — separate from
// whatever data got extracted from it.
async function uploadOriginalFile(orgId, vendorId, file) {
  const path = `${orgId}/${vendorId}/${Date.now()}_${file.name}`;
  const { error } = await supabase.storage.from("documents").upload(path, file);
  if (error) return { path: null, name: null, error };
  return { path, name: file.name, error: null };
}

async function viewStoredFile(path) {
  const { data, error } = await supabase.storage.from("documents").createSignedUrl(path, 3600);
  if (error) { alert("Couldn't open file: " + error.message); return; }
  window.open(data.signedUrl, "_blank");
}

// A business's own logo, so their home screen feels like theirs, not a
// generic KERDOS screen. Stored in the same private bucket as documents,
// under a dedicated "logo" folder per organization.
async function uploadOrgLogo(orgId, file) {
  const path = `${orgId}/logo/${Date.now()}_${file.name}`;
  const { error } = await supabase.storage.from("documents").upload(path, file, { upsert: true });
  if (error) return { path: null, error };
  return { path, error: null };
}

async function getSignedUrl(path) {
  if (!path) return null;
  const { data, error } = await supabase.storage.from("documents").createSignedUrl(path, 3600);
  if (error) return null;
  return data.signedUrl;
}

// ── UTILITIES ────────────────────────────────────────────────────────
function r2(n) { return Math.round(n * 100) / 100; }

const UNIT_ALIASES = {
  lb:"LB",lbs:"LB",oz:"OZ",gal:"GAL",gallon:"GAL",gl:"GAL",ga:"GAL",
  qt:"QT",pt:"PT",liter:"L",l:"L",ct:"CT",ea:"EA",each:"EA",
  doz:"DOZ",dz:"DOZ",cn:"EA",bu:"BU",ft:"FT",
};
const TO_BASE = {
  OZ:["OZ",1],LB:["OZ",16],GAL:["FLOZ",128],QT:["FLOZ",32],
  PT:["FLOZ",16],L:["FLOZ",33.814],FLOZ:["FLOZ",1],
  EA:["EA",1],CT:["EA",1],DOZ:["EA",12],
};

function parsePackSize(pack) {
  if (!pack) return null;
  const clean = pack.trim().toUpperCase().replace(/LBA?V?/g,"LB").trim();
  let m = clean.match(/^(\d+(?:\.\d+)?)\s*\/\s*(\d+(?:\.\d+)?)\s*([A-Z#]+)/);
  if (m) {
    const cq=parseFloat(m[1]),uq=parseFloat(m[2]),u=UNIT_ALIASES[m[3].toLowerCase()]||m[3];
    return {caseQty:cq,unitQty:uq,unit:u,total:cq*uq,eachStr:`${uq} ${u}`,caseStr:`${cq}/${uq} ${u}`};
  }
  m = clean.match(/^(\d+(?:\.\d+)?)\s+([A-Z#]+)/);
  if (m) {
    const q=parseFloat(m[1]),u=UNIT_ALIASES[m[2].toLowerCase()]||m[2];
    return {caseQty:1,unitQty:q,unit:u,total:q,eachStr:`${q} ${u}`,caseStr:`${q} ${u}`};
  }
  m = clean.match(/^(\d+(?:\.\d+)?)([A-Z#]+)/);
  if (m) {
    const q=parseFloat(m[1]),u=UNIT_ALIASES[m[2].toLowerCase()]||m[2];
    return {caseQty:1,unitQty:q,unit:u,total:q,eachStr:`${q} ${u}`,caseStr:`${q} ${u}`};
  }
  return null;
}

function normalizedPrice(price, pack) {
  const p = parsePackSize(pack);
  if (!p) return null;
  const conv = TO_BASE[p.unit];
  if (!conv) return {price:r2(price/p.total), unit:p.unit};
  return {price:Math.round(price/p.total/conv[1]*1000000)/1000000, unit:conv[0]};
}

function eachPrice(casePrice, pack) {
  const p = parsePackSize(pack);
  if (!p || p.caseQty <= 1) return null;
  return {price:r2(casePrice/p.caseQty), size:p.eachStr};
}

function formatDateMDY(isoDate) {
  if (!isoDate) return "";
  const [y,m,d] = isoDate.split("-");
  if (!y||!m||!d) return isoDate;
  return `${m}/${d}/${y}`;
}

// ── SOLVER ───────────────────────────────────────────────────────────
function solve(cartItems, vendors) {
  if (!cartItems.length) return [];
  const reqMap = new Map(vendors.map(v=>[v.id,v]));
  let assignments = cartItems.map(item => {
    const opts = [...item.options].sort((a,b)=>a.price-b.price);
    const best = opts[0];
    return {...item, assignedVendorId:best.vendorId, assignedVendorName:best.vendorName,
      vendorItemId:best.vendorItemId, price:best.price, packSize:best.packSize,
      orderUnit:best.orderUnit, lineTotal:r2(best.price*item.quantity),
      cheapestPrice:best.price, premiumPaid:0};
  });

  for (let iter = 0; iter < vendors.length*4; iter++) {
    const totals = new Map();
    for (const a of assignments) {
      const t = totals.get(a.assignedVendorId)||{dollar:0,units:0};
      t.dollar=r2(t.dollar+a.lineTotal); t.units+=a.quantity;
      totals.set(a.assignedVendorId,t);
    }
    const short = vendors.filter(v=>{
      const t=totals.get(v.id);
      return t&&((v.delivery_minimum_dollar&&t.dollar<v.delivery_minimum_dollar)||
                 (v.delivery_minimum_units&&t.units<v.delivery_minimum_units));
    });
    if (!short.length) break;

    let fixed=false;
    for (const req of short) {
      const fills = assignments
        .filter(a=>a.assignedVendorId!==req.id)
        .map(a=>{
          const opt=a.options.find(o=>o.vendorId===req.id);
          if(!opt) return null;
          return {a,opt,premium:(opt.price-a.price)*a.quantity,free:opt.price<=a.price};
        }).filter(Boolean).sort((x,y)=>(x.free?0:1)-(y.free?0:1)||x.premium-y.premium);

      let pathA=[...assignments];
      let fd=assignments.filter(a=>a.assignedVendorId===req.id).reduce((s,a)=>s+a.lineTotal,0);
      let fu=assignments.filter(a=>a.assignedVendorId===req.id).reduce((s,a)=>s+a.quantity,0);
      for (const {a,opt} of fills) {
        const idx=pathA.findIndex(x=>x.catalogItemId===a.catalogItemId);
        if(idx===-1) continue;
        const lt=r2(opt.price*a.quantity);
        pathA[idx]={...a,assignedVendorId:req.id,assignedVendorName:req.name,
          vendorItemId:opt.vendorItemId,price:opt.price,packSize:opt.packSize,
          lineTotal:lt,premiumPaid:r2(Math.max(0,lt-a.cheapestPrice*a.quantity))};
        fd=r2(fd+opt.price*a.quantity); fu+=a.quantity;
        const met=(!req.delivery_minimum_dollar||fd>=req.delivery_minimum_dollar)&&
                  (!req.delivery_minimum_units||fu>=req.delivery_minimum_units);
        if(met) break;
      }
      const pathAItems=pathA.filter(a=>a.assignedVendorId===req.id);
      const pathADollar=pathAItems.reduce((s,a)=>s+a.lineTotal,0);
      const pathAUnits=pathAItems.reduce((s,a)=>s+a.quantity,0);
      const pathAMeets=(!req.delivery_minimum_dollar||pathADollar>=req.delivery_minimum_dollar)&&
                       (!req.delivery_minimum_units||pathAUnits>=req.delivery_minimum_units);
      const pathASpend=pathA.reduce((s,a)=>s+a.lineTotal,0);

      const pathB=assignments.map(a=>{
        if(a.assignedVendorId!==req.id) return a;
        const alt=[...a.options].filter(o=>o.vendorId!==req.id).sort((x,y)=>x.price-y.price)[0];
        if(!alt) return a;
        const lt=r2(alt.price*a.quantity);
        return {...a,assignedVendorId:alt.vendorId,assignedVendorName:alt.vendorName,
          vendorItemId:alt.vendorItemId,price:alt.price,packSize:alt.packSize,
          lineTotal:lt,premiumPaid:r2(Math.max(0,lt-a.cheapestPrice*a.quantity))};
      });
      const pathBSpend=pathB.reduce((s,a)=>s+a.lineTotal,0);

      const chosen=(pathAMeets&&pathASpend<=pathBSpend)?pathA:pathB;
      const newTotals=new Map();
      for(const a of chosen){const t=newTotals.get(a.assignedVendorId)||{dollar:0,units:0};t.dollar=r2(t.dollar+a.lineTotal);t.units+=a.quantity;newTotals.set(a.assignedVendorId,t);}
      const t=newTotals.get(req.id);
      if(!t||( (!req.delivery_minimum_dollar||t.dollar>=req.delivery_minimum_dollar)&&(!req.delivery_minimum_units||t.units>=req.delivery_minimum_units))){
        assignments=chosen; fixed=true; break;
      }
    }
    if(!fixed) break;
  }
  return assignments;
}

// ── KERDOS MARK ──────────────────────────────────────────────────────
// The real Greek 1-drachma owl coin, used across login, setup, loading,
// and header. Embedded directly so no separate image file is needed.
const KERDOS_COIN_DATA_URI = "data:image/png;base64,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";

function KerdosMark({ size = 40 }) {
  return (
    <img
      src={KERDOS_COIN_DATA_URI}
      alt="KERDOS"
      width={size}
      height={size}
      style={{ flexShrink: 0, display: "block", borderRadius: "50%" }}
    />
  );
}


// ── STYLES ───────────────────────────────────────────────────────────
const PALETTE = [
  {bg:"#E3F2FD",accent:"#1565C0",light:"#BBDEFB"},
  {bg:"#E8F5E9",accent:"#2E7D32",light:"#C8E6C9"},
  {bg:"#FFF3E0",accent:"#E65100",light:"#FFE0B2"},
  {bg:"#F3E5F5",accent:"#6A1B9A",light:"#E1BEE7"},
  {bg:"#FCE4EC",accent:"#880E4F",light:"#F8BBD0"},
  {bg:"#E0F2F1",accent:"#00695C",light:"#B2DFDB"},
];
const inp = {width:"100%",padding:"10px 12px",border:"1px solid #E0E0E0",borderRadius:8,fontSize:14,outline:"none",boxSizing:"border-box"};
const btn = (bg,color="white",extra={}) => ({padding:"10px 18px",borderRadius:8,border:"none",cursor:"pointer",fontWeight:700,fontSize:14,background:bg,color,...extra});

// ── LOGIN ─────────────────────────────────────────────────────────────
function Login() {
  const [email,setEmail]=useState("");
  const [password,setPassword]=useState("");
  const [loading,setLoading]=useState(false);
  const [error,setError]=useState("");
  const [mode,setMode]=useState("login");

  async function submit(e) {
    e.preventDefault(); setLoading(true); setError("");
    const {error:err} = mode==="login" ? await supabase.auth.signInWithPassword({email,password}) : await supabase.auth.signUp({email,password});
    if(err) setError(err.message);
    else if(mode==="signup") setError("Check your email to confirm, then sign in.");
    setLoading(false);
  }

  return (
    <div style={{minHeight:"100vh",background:"#003584",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:"white",borderRadius:12,padding:32,width:"100%",maxWidth:380,boxShadow:"0 4px 20px rgba(0,0,0,0.3)"}}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <div style={{display:"flex",justifyContent:"center",marginBottom:6}}><KerdosMark size={56}/></div>
          <div style={{fontWeight:900,fontSize:20,letterSpacing:"0.18em",color:"#003584",marginTop:4}}>KERDOS</div>
          <div style={{fontSize:12,color:"#888",marginTop:2}}>Procurement Management</div>
        </div>
        <form onSubmit={submit}>
          <div style={{marginBottom:12}}>
            <input style={inp} type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="Email" required />
          </div>
          <div style={{marginBottom:16}}>
            <input style={inp} type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="Password" required />
          </div>
          {error&&<div style={{background:"#FFF3E0",color:"#E65100",padding:"10px 12px",borderRadius:8,fontSize:13,marginBottom:12}}>{error}</div>}
          <button type="submit" disabled={loading} style={{...btn("#003584"),width:"100%"}}>
            {loading?"Please wait...":mode==="login"?"Sign In":"Create Account"}
          </button>
        </form>
        <div style={{textAlign:"center",marginTop:14,fontSize:13,color:"#888"}}>
          {mode==="login"
            ?<span>No account? <button onClick={()=>setMode("signup")} style={{background:"none",border:"none",color:"#003584",cursor:"pointer",fontWeight:700}}>Sign up</button></span>
            :<span>Have account? <button onClick={()=>setMode("login")} style={{background:"none",border:"none",color:"#003584",cursor:"pointer",fontWeight:700}}>Sign in</button></span>}
        </div>
      </div>
    </div>
  );
}

// ── SETUP WIZARD ─────────────────────────────────────────────────────
function Setup({user,onComplete}) {
  const [step,setStep]=useState(1);
  const [orgName,setOrgName]=useState("");
  const [industry,setIndustry]=useState("");
  const [vendors,setVendors]=useState([{name:"",minDollar:"",minUnits:""}]);
  const [loading,setLoading]=useState(false);
  const [error,setError]=useState("");

  async function create() {
    setLoading(true); setError("");
    try {
      const slug=orgName.toLowerCase().replace(/[^a-z0-9]/g,"-").replace(/-+/g,"-");
      const {data:org,error:e}=await supabase.from("organizations").insert({name:orgName,slug,industry}).select().single();
      if(e) throw e;
      await supabase.from("organization_members").insert({organization_id:org.id,user_id:user.id,role:"owner"});
      const vrows=vendors.filter(v=>v.name.trim()).map(v=>({
        organization_id:org.id, name:v.name.trim(),
        delivery_minimum_dollar:v.minDollar?parseFloat(v.minDollar):null,
        delivery_minimum_units:v.minUnits?parseInt(v.minUnits):null,
      }));
      if(vrows.length) await supabase.from("vendors").insert(vrows);
      onComplete(org);
    } catch(err){setError(err.message);}
    setLoading(false);
  }

  return (
    <div style={{minHeight:"100vh",background:"#F0F2F5",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:"white",borderRadius:12,padding:28,width:"100%",maxWidth:480,boxShadow:"0 2px 8px rgba(0,0,0,0.1)"}}>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
          <KerdosMark size={22}/>
          <span style={{fontWeight:900,fontSize:16,letterSpacing:"0.18em",color:"#003584"}}>KERDOS</span>
        </div>
        <h2 style={{margin:"0 0 4px",fontSize:18}}>Welcome — let's get set up</h2>
        <p style={{color:"#888",fontSize:13,margin:"0 0 20px"}}>Takes about 2 minutes.</p>

        {step===1&&<>
          <div style={{marginBottom:12}}>
            <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Organization name</div>
            <input style={inp} value={orgName} onChange={e=>setOrgName(e.target.value)} placeholder="e.g. Hornet's Nest Deli" />
          </div>
          <div style={{marginBottom:20}}>
            <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Industry (optional)</div>
            <input style={inp} value={industry} onChange={e=>setIndustry(e.target.value)} placeholder="e.g. Food Service, Hardware, Medical..." />
          </div>
          <button onClick={()=>setStep(2)} disabled={!orgName.trim()} style={{...btn("#003584"),width:"100%"}}>Next →</button>
        </>}

        {step===2&&<>
          <p style={{fontSize:13,color:"#555",margin:"0 0 14px"}}>Add your vendors — you can add more later.</p>
          {vendors.map((v,i)=>(
            <div key={i} style={{background:"#F8F9FA",borderRadius:8,padding:12,marginBottom:10}}>
              <div style={{marginBottom:8}}>
                <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Vendor name</div>
                <input style={inp} value={v.name} onChange={e=>{const vv=[...vendors];vv[i].name=e.target.value;setVendors(vv);}} placeholder="e.g. US Foods" />
              </div>
              <div style={{display:"flex",gap:8}}>
                <div style={{flex:1}}>
                  <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Min order ($)</div>
                  <input style={inp} value={v.minDollar} onChange={e=>{const vv=[...vendors];vv[i].minDollar=e.target.value;setVendors(vv);}} placeholder="500" type="number" />
                </div>
                <div style={{flex:1}}>
                  <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Min items</div>
                  <input style={inp} value={v.minUnits} onChange={e=>{const vv=[...vendors];vv[i].minUnits=e.target.value;setVendors(vv);}} placeholder="20" type="number" />
                </div>
              </div>
            </div>
          ))}
          <button onClick={()=>setVendors([...vendors,{name:"",minDollar:"",minUnits:""}])}
            style={{...btn("#F0F2F5","#555"),width:"100%",marginBottom:10}}>+ Add Vendor</button>
          {error&&<div style={{color:"#E65100",fontSize:13,marginBottom:10}}>{error}</div>}
          <div style={{display:"flex",gap:8}}>
            <button onClick={()=>setStep(1)} style={{...btn("#EEE","#555"),flex:1}}>← Back</button>
            <button onClick={create} disabled={loading} style={{...btn("#003584"),flex:2}}>
              {loading?"Creating...":"Get Started →"}
            </button>
          </div>
        </>}
      </div>
    </div>
  );
}

function OrgGate({user,onComplete}) {
  const [path,setPath]=useState(null); // null | "create" | "join"

  if(path==="create") return <Setup user={user} onComplete={onComplete} />;
  if(path==="join") return <JoinWithCode user={user} onComplete={onComplete} onBack={()=>setPath(null)} />;

  return (
    <div style={{minHeight:"100vh",background:"#F0F2F5",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:"white",borderRadius:12,padding:28,width:"100%",maxWidth:420,boxShadow:"0 2px 8px rgba(0,0,0,0.1)",textAlign:"center"}}>
        <div style={{display:"flex",justifyContent:"center",marginBottom:8}}><KerdosMark size={48}/></div>
        <h2 style={{margin:"0 0 4px",fontSize:18}}>Welcome to KERDOS</h2>
        <p style={{color:"#888",fontSize:13,margin:"0 0 24px"}}>Are you starting a new organization, or joining one your team already set up?</p>
        <button onClick={()=>setPath("create")} style={{...btn("#003584"),width:"100%",marginBottom:10}}>Create a new organization</button>
        <button onClick={()=>setPath("join")} style={{...btn("#F0F2F5","#555"),width:"100%"}}>Join with an invite code</button>
      </div>
    </div>
  );
}

function JoinWithCode({user,onComplete,onBack}) {
  const [code,setCode]=useState("");
  const [loading,setLoading]=useState(false);
  const [error,setError]=useState("");

  async function join() {
    setLoading(true); setError("");
    try {
      const cleanCode=code.trim().toUpperCase();
      const {data:invite,error:e}=await supabase.from("invite_codes").select("*").eq("code",cleanCode).is("used_by",null).maybeSingle();
      if(e) throw e;
      if(!invite){setError("That code wasn't found, or it's already been used.");setLoading(false);return;}

      await supabase.from("organization_members").insert({organization_id:invite.organization_id,user_id:user.id,role:invite.role});
      await supabase.from("invite_codes").update({used_by:user.id,used_at:new Date().toISOString()}).eq("id",invite.id);

      const {data:org}=await supabase.from("organizations").select("*").eq("id",invite.organization_id).single();
      onComplete(org);
    } catch(err){setError(err.message);}
    setLoading(false);
  }

  return (
    <div style={{minHeight:"100vh",background:"#F0F2F5",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:"white",borderRadius:12,padding:28,width:"100%",maxWidth:420,boxShadow:"0 2px 8px rgba(0,0,0,0.1)"}}>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:16}}>
          <KerdosMark size={22}/>
          <span style={{fontWeight:900,fontSize:16,letterSpacing:"0.18em",color:"#003584"}}>KERDOS</span>
        </div>
        <h2 style={{margin:"0 0 4px",fontSize:18}}>Join your team</h2>
        <p style={{color:"#888",fontSize:13,margin:"0 0 20px"}}>Enter the invite code your owner or manager shared with you.</p>
        <input style={{...inp,textAlign:"center",fontSize:20,letterSpacing:"0.1em",fontWeight:700,marginBottom:14}}
          value={code} onChange={e=>setCode(e.target.value)} placeholder="XXXX-XXXX" />
        {error&&<div style={{color:"#E65100",fontSize:13,marginBottom:14}}>{error}</div>}
        <div style={{display:"flex",gap:8}}>
          <button onClick={onBack} style={{...btn("#EEE","#555"),flex:1}}>← Back</button>
          <button onClick={join} disabled={loading||!code.trim()} style={{...btn("#003584"),flex:2}}>
            {loading?"Joining...":"Join team →"}
          </button>
        </div>
      </div>
    </div>
  );
}

function generateInviteCode() {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; // no ambiguous chars
  const part = () => Array.from({length:4}, () => chars[Math.floor(Math.random()*chars.length)]).join("");
  return `${part()}-${part()}`;
}

function TeamPanel({orgId,myRole,currentUserId}) {
  const [members,setMembers]=useState([]);
  const [codes,setCodes]=useState([]);
  const [loading,setLoading]=useState(true);
  const [showInvite,setShowInvite]=useState(false);
  const [inviteRole,setInviteRole]=useState("employee");
  const [newCode,setNewCode]=useState(null);
  const [error,setError]=useState("");

  async function load() {
    setLoading(true);
    const [mr,cr]=await Promise.all([
      supabase.from("organization_members").select("*").eq("organization_id",orgId),
      supabase.from("invite_codes").select("*").eq("organization_id",orgId).order("created_at",{ascending:false}),
    ]);
    setMembers(mr.data||[]);
    setCodes(cr.data||[]);
    setLoading(false);
  }

  useEffect(()=>{ load(); },[orgId]);

  async function createInvite() {
    setError("");
    const code=generateInviteCode();
    const {error:e}=await supabase.from("invite_codes").insert({organization_id:orgId,code,role:inviteRole,created_by:currentUserId});
    if(e){ setError(e.message); return; }
    setNewCode(code);
    load();
  }

  const roleCounts = {
    owner: members.filter(m=>m.role==="owner").length,
    manager: members.filter(m=>m.role==="manager").length,
    employee: members.filter(m=>m.role==="employee").length,
  };
  const roleBadge = (role) => {
    const colors = {owner:{bg:"#E3F2FD",fg:"#1565C0"},manager:{bg:"#E8F5E9",fg:"#2E7D32"},employee:{bg:"#FFF3E0",fg:"#E65100"}};
    const c = colors[role]||colors.employee;
    return <span style={{fontSize:11,fontWeight:700,padding:"3px 9px",borderRadius:10,background:c.bg,color:c.fg,textTransform:"capitalize"}}>{role}</span>;
  };

  if(loading) return <p style={{color:"#888"}}>Loading team...</p>;

  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
        <h3 style={{margin:0,fontSize:16}}>Team</h3>
        <button onClick={()=>{setShowInvite(true);setNewCode(null);setInviteRole("employee");}} style={{...btn("#003584","white",{fontSize:12,padding:"8px 14px"})}}>+ Invite</button>
      </div>

      <div style={{display:"flex",gap:8,marginBottom:16}}>
        <div style={{flex:1,background:"white",borderRadius:10,padding:14,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <div style={{fontWeight:900,fontSize:20,color:"#1565C0"}}>{roleCounts.owner}</div>
          <div style={{fontSize:11,color:"#888"}}>Owner{roleCounts.owner!==1?"s":""}</div>
        </div>
        <div style={{flex:1,background:"white",borderRadius:10,padding:14,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <div style={{fontWeight:900,fontSize:20,color:"#2E7D32"}}>{roleCounts.manager}</div>
          <div style={{fontSize:11,color:"#888"}}>Manager{roleCounts.manager!==1?"s":""}</div>
        </div>
        <div style={{flex:1,background:"white",borderRadius:10,padding:14,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <div style={{fontWeight:900,fontSize:20,color:"#E65100"}}>{roleCounts.employee}</div>
          <div style={{fontSize:11,color:"#888"}}>Employee{roleCounts.employee!==1?"s":""}</div>
        </div>
      </div>

      <div style={{fontSize:11,fontWeight:800,color:"#999",letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:8}}>Invite codes</div>
      {codes.length===0?(
        <div style={{background:"white",borderRadius:10,padding:20,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <p style={{color:"#888",fontSize:13,margin:0}}>No invite codes yet — create one to bring someone onto your team.</p>
        </div>
      ):codes.map(c=>(
        <div key={c.id} style={{background:"white",borderRadius:8,padding:"10px 14px",marginBottom:6,
          display:"flex",justifyContent:"space-between",alignItems:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <div>
            <div style={{fontFamily:"monospace",fontWeight:700,fontSize:14}}>{c.code}</div>
            <div style={{fontSize:11,color:"#888"}}>{c.used_by?"Used":"Not yet used"}</div>
          </div>
          {roleBadge(c.role)}
        </div>
      ))}

      {showInvite&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div style={{background:"white",borderRadius:12,padding:24,width:"100%",maxWidth:380}}>
            {!newCode?(<>
              <h3 style={{margin:"0 0 14px",fontSize:16}}>Invite someone</h3>
              <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:6}}>Their role</div>
              <div style={{display:"flex",gap:8,marginBottom:18}}>
                {myRole==="owner"&&(
                  <button onClick={()=>setInviteRole("manager")} style={{...btn(inviteRole==="manager"?"#2E7D32":"#EEE",inviteRole==="manager"?"white":"#555"),flex:1}}>Manager</button>
                )}
                <button onClick={()=>setInviteRole("employee")} style={{...btn(inviteRole==="employee"?"#E65100":"#EEE",inviteRole==="employee"?"white":"#555"),flex:1}}>Employee</button>
              </div>
              {error&&<div style={{color:"#E65100",fontSize:12,marginBottom:12}}>{error}</div>}
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>setShowInvite(false)} style={{...btn("#EEE","#555"),flex:1}}>Cancel</button>
                <button onClick={createInvite} style={{...btn("#003584"),flex:2}}>Generate code</button>
              </div>
            </>):(<>
              <h3 style={{margin:"0 0 6px",fontSize:16}}>Share this code</h3>
              <p style={{color:"#888",fontSize:12,margin:"0 0 16px"}}>They'll enter this after signing up to join as {inviteRole}.</p>
              <div style={{background:"#F0F2F5",borderRadius:8,padding:"16px",textAlign:"center",fontFamily:"monospace",fontWeight:900,fontSize:24,letterSpacing:"0.1em",marginBottom:16}}>{newCode}</div>
              <button onClick={()=>setShowInvite(false)} style={{...btn("#003584"),width:"100%"}}>Done</button>
            </>)}
          </div>
        </div>
      )}
    </div>
  );
}

// ── PASTE MODAL ───────────────────────────────────────────────────────
function PasteModal({vendors,orgId,onClose,onDone,initialVendorId,initialMode}) {
  const [vendorId,setVendorId]=useState(initialVendorId||vendors[0]?.id||"");
  const [mode,setMode]=useState(initialMode||"pricelist");
  const [text,setText]=useState("");
  const [parsed,setParsed]=useState([]);
  const [skipped,setSkipped]=useState([]);
  const [rawFile,setRawFile]=useState(null);
  const [step,setStep]=useState(1);
  const [loading,setLoading]=useState(false);
  const [result,setResult]=useState(null);
  const [dragOver,setDragOver]=useState(false);
  const [fileBusy,setFileBusy]=useState(false);

  async function handleDroppedFiles(files){
    const file=files[0];
    if(!file) return;
    setRawFile(file);
    setFileBusy(true);
    try{
      const content=await fileToText(file);
      setText(content);
    }catch(err){
      alert(err.message);
    }
    setFileBusy(false);
  }

  function doParse(){
    const result = parseDocument(text);
    setParsed(result.rows);
    setSkipped(result.skipped);
    setStep(2);
  }

  async function doSave(){
    setLoading(true);
    const vendor=vendors.find(v=>v.id===vendorId);
    let updated=0,created=0,invoiceTotal=0;
    let saveError=null;

    for(const row of parsed){
      if(mode==="pricelist"){
        const q=row.code
          ?supabase.from("vendor_items").select("id,price").eq("organization_id",orgId).eq("vendor_id",vendorId).eq("vendor_item_code",row.code).maybeSingle()
          :supabase.from("vendor_items").select("id,price").eq("organization_id",orgId).eq("vendor_id",vendorId).eq("description",row.description).maybeSingle();
        const {data:ex}=await q;
        if(ex){
          if(Math.abs((ex.price||0)-row.price)>0.001){
            await supabase.from("price_history").insert({vendor_item_id:ex.id,organization_id:orgId,price:row.price,source:"price_list"});
            await supabase.from("vendor_items").update({price:row.price,pack_size:row.packSize,last_updated:new Date().toISOString()}).eq("id",ex.id);
          }
          updated++;
        } else {
          const {data:ni}=await supabase.from("vendor_items").insert({organization_id:orgId,vendor_id:vendorId,vendor_item_code:row.code,description:row.description,pack_size:row.packSize,price:row.price}).select().single();
          if(ni) await supabase.from("price_history").insert({vendor_item_id:ni.id,organization_id:orgId,price:row.price,source:"price_list"});
          created++;
        }
      } else {
        invoiceTotal+=(row.amount!=null?row.amount:row.price);
      }
    }

    // Upload the original file, unaltered, so it can be reopened later
    // exactly as received — separate from whatever data was extracted.
    let filePath=null, fileName=null;
    if(mode==="invoice" && rawFile){
      const upload = await uploadOriginalFile(orgId, vendorId, rawFile);
      if(upload.error){
        saveError = "Data was extracted, but the original file couldn't be saved: " + upload.error.message;
      } else {
        filePath = upload.path;
        fileName = upload.name;
      }
    }

    if(mode==="invoice"){
      const {data:inv,error:invErr}=await supabase.from("invoices").insert({
        organization_id:orgId, vendor_id:vendorId, total_amount:r2(invoiceTotal),
        raw_text:text, invoice_date:new Date().toISOString().split("T")[0], status:"recorded",
        file_path:filePath, file_name:fileName,
      }).select().single();

      if(invErr){
        saveError = (saveError?saveError+" ":"") + "Couldn't save this invoice: " + invErr.message;
      } else if(inv){
        const {error:linesErr}=await supabase.from("invoice_lines").insert(parsed.map(row=>({
          invoice_id:inv.id, vendor_item_code:row.code, description:row.description,
          unit_price:row.price, line_total:(row.amount!=null?row.amount:row.price),
        })));
        if(linesErr){
          saveError = (saveError?saveError+" ":"") + "Invoice saved, but its line items didn't: " + linesErr.message;
        }
      }
    }

    setResult({mode,vendor:vendor?.name,updated,created,invoiceTotal:r2(invoiceTotal),count:parsed.length,error:saveError});
    setStep(3);setLoading(false);
  }

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:1000,display:"flex",alignItems:"flex-end",justifyContent:"center"}}>
      <div style={{background:"white",borderRadius:"16px 16px 0 0",padding:20,width:"100%",maxWidth:600,maxHeight:"90vh",overflowY:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <h3 style={{margin:0,fontSize:16}}>{mode==="pricelist"?"📋 Import Price List":"🧾 Record Invoice"}</h3>
          <button onClick={onClose} style={{background:"none",border:"none",fontSize:24,cursor:"pointer",color:"#888"}}>×</button>
        </div>

        {step===1&&<>
          <div style={{display:"flex",gap:8,marginBottom:14}}>
            <button onClick={()=>setMode("pricelist")} style={{...btn(mode==="pricelist"?"#003584":"#EEE",mode==="pricelist"?"white":"#555"),flex:1}}>Price List</button>
            <button onClick={()=>setMode("invoice")} style={{...btn(mode==="invoice"?"#003584":"#EEE",mode==="invoice"?"white":"#555"),flex:1}}>Invoice</button>
          </div>
          <div style={{marginBottom:14}}>
            <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Vendor</div>
            <select style={inp} value={vendorId} onChange={e=>setVendorId(e.target.value)}>
              {vendors.map(v=><option key={v.id} value={v.id}>{v.name}</option>)}
            </select>
          </div>
          <div style={{marginBottom:14}}>
            <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Paste {mode==="pricelist"?"price list":"invoice"} here, or drag a file in</div>
            <div
              onDragOver={e=>{e.preventDefault();setDragOver(true);}}
              onDragLeave={()=>setDragOver(false)}
              onDrop={e=>{e.preventDefault();setDragOver(false);handleDroppedFiles(e.dataTransfer.files);}}
              style={{position:"relative"}}
            >
              <textarea
                onDragOver={e=>{e.preventDefault();e.stopPropagation();setDragOver(true);}}
                onDragLeave={()=>setDragOver(false)}
                onDrop={e=>{e.preventDefault();e.stopPropagation();setDragOver(false);handleDroppedFiles(e.dataTransfer.files);}}
                style={{...inp,height:200,resize:"vertical",fontFamily:"monospace",fontSize:12,border:dragOver?"2px dashed #003584":inp.border}}
                value={text} onChange={e=>setText(e.target.value)}
                placeholder="Copy from Excel, email, PDF — paste here, or drag a file..." />
              {fileBusy&&<div style={{position:"absolute",inset:0,background:"rgba(255,255,255,0.85)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,color:"#003584",fontWeight:700,borderRadius:8}}>Reading file...</div>}
              <input type="file" accept=".csv,.txt,.tsv,.xlsx,.xls,.pdf"
                onChange={e=>handleDroppedFiles(e.target.files)}
                style={{marginTop:8,fontSize:12}} />
            </div>
            {rawFile&&<div style={{fontSize:11,color:"#888",marginTop:4}}>📎 {rawFile.name} {mode==="invoice"?"— original will be saved":""}</div>}
          </div>
          <button onClick={doParse} disabled={text.length<10} style={{...btn("#003584"),width:"100%"}}>
            Parse ({text.split("\n").filter(l=>l.trim()).length} lines)
          </button>
        </>}

        {step===2&&<>
          <div style={{background:"#E8F5E9",padding:"10px 14px",borderRadius:8,marginBottom:14,fontSize:13}}>
            Found {parsed.length} items — review and confirm
          </div>
          <div style={{maxHeight:280,overflowY:"auto",marginBottom:14}}>
            {parsed.map((row,i)=>(
              <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:"1px solid #F0F0F0",fontSize:13}}>
                <div>
                  {row.code&&<span style={{color:"#888",marginRight:8,fontFamily:"monospace",fontSize:11}}>{row.code}</span>}
                  <span>{row.description}</span>
                  {row.packSize&&<span style={{color:"#AAA",marginLeft:6,fontSize:11}}>{row.packSize}</span>}
                </div>
                <span style={{fontWeight:700,flexShrink:0,marginLeft:8}}>${row.price.toFixed(2)}</span>
              </div>
            ))}
          </div>
          {skipped.length>0&&(
            <div style={{background:"#FFF3E0",padding:"10px 14px",borderRadius:8,marginBottom:14,fontSize:12}}>
              <div style={{fontWeight:700,color:"#E65100",marginBottom:6}}>{skipped.length} line{skipped.length>1?"s":""} couldn't be read — skipped, not saved</div>
              <div style={{maxHeight:120,overflowY:"auto"}}>
                {skipped.map((s,i)=>(
                  <div key={i} style={{color:"#999",marginBottom:3,fontFamily:"monospace",fontSize:11}}>
                    "{s.line.slice(0,60)}" — {s.reason}
                  </div>
                ))}
              </div>
            </div>
          )}
          <div style={{display:"flex",gap:8}}>
            <button onClick={()=>setStep(1)} style={{...btn("#EEE","#555"),flex:1}}>← Back</button>
            <button onClick={doSave} disabled={loading||!parsed.length} style={{...btn("#003584"),flex:2}}>
              {loading?"Saving...":"Save "+parsed.length+" items"}
            </button>
          </div>
        </>}

        {step===3&&result&&(
          <div style={{textAlign:"center",padding:"20px 0"}}>
            <div style={{fontSize:40,marginBottom:12}}>{result.error?"⚠️":"✅"}</div>
            <h3 style={{margin:"0 0 8px"}}>{result.vendor}</h3>
            {result.mode==="pricelist"
              ?<p style={{color:"#666",fontSize:14}}>{result.updated} items updated · {result.created} new items added</p>
              :<p style={{color:"#666",fontSize:14}}>{result.count} lines · ${result.invoiceTotal?.toFixed(2)}</p>}
            {result.error&&<div style={{background:"#FFF3E0",color:"#E65100",padding:"10px 12px",borderRadius:8,fontSize:13,marginTop:12,textAlign:"left"}}>{result.error}</div>}
            <button onClick={()=>{onDone();onClose();}} style={{...btn("#003584"),marginTop:16}}>Done</button>
          </div>
        )}
      </div>
    </div>
  );
}

// ── MAIN APP ──────────────────────────────────────────────────────────
export default function App() {
  const [session,setSession]=useState(undefined);
  const [org,setOrg]=useState(null);
  const [vendors,setVendors]=useState([]);
  const [catalogItems,setCatalogItems]=useState([]);
  const [vendorItems,setVendorItems]=useState([]);
  const [mappings,setMappings]=useState([]);
  const [invoices,setInvoices]=useState([]);
  const [quantities,setQuantities]=useState({}); // {catalogItemId_case: n, catalogItemId_each: n}
  const [tab,setTab]=useState("home");
  const [showPaste,setShowPaste]=useState(false);
  const [selectedVendorId,setSelectedVendorId]=useState(null);
  const [importMode,setImportMode]=useState("pricelist");
  const [vendorDetailId,setVendorDetailId]=useState(null);
  const [logoUrl,setLogoUrl]=useState(null);
  const [logoUploading,setLogoUploading]=useState(false);
  const [expandedItem,setExpandedItem]=useState(null);
  const [search,setSearch]=useState("");
  const [loading,setLoading]=useState(true);

  useEffect(()=>{
    supabase.auth.getSession().then(({data:{session}})=>setSession(session));
    const {data:{subscription}}=supabase.auth.onAuthStateChange((_,s)=>setSession(s));
    return ()=>subscription.unsubscribe();
  },[]);

  useEffect(()=>{
    if(session===undefined) return;
    if(!session){setLoading(false);return;}
    loadData();
  },[session]);

  async function loadData(){
    setLoading(true);
    const {data:mem}=await supabase.from("organization_members").select("organization_id,role,organizations(*)").eq("user_id",session.user.id);
    if(!mem?.length){setOrg(null);setLoading(false);return;}
    const o={...mem[0].organizations,role:mem[0].role};
    setOrg(o);
    if(o.logo_url) getSignedUrl(o.logo_url).then(setLogoUrl); else setLogoUrl(null);
    const id=o.id;
    const [vr,cr,vir,mr,ir]=await Promise.all([
      supabase.from("vendors").select("*").eq("organization_id",id).eq("is_active",true).order("name"),
      supabase.from("catalog_items").select("*,catalog_categories(name)").eq("organization_id",id).order("master_item_number"),
      supabase.from("vendor_items").select("*").eq("organization_id",id),
      supabase.from("item_mappings").select("*").eq("organization_id",id),
      supabase.from("invoices").select("*,vendors(name)").eq("organization_id",id).order("created_at",{ascending:false}).limit(30),
    ]);
    setVendors(vr.data||[]);
    setCatalogItems(cr.data||[]);
    setVendorItems(vir.data||[]);
    setMappings(mr.data||[]);
    setInvoices(ir.data||[]);
    setLoading(false);
  }

  useEffect(()=>{
    if(!org) return;
    const ch=supabase.channel("kerdos")
      .on("postgres_changes",{event:"*",schema:"public",table:"vendor_items",filter:`organization_id=eq.${org.id}`},loadData)
      .on("postgres_changes",{event:"*",schema:"public",table:"invoices",filter:`organization_id=eq.${org.id}`},loadData)
      .subscribe();
    return ()=>supabase.removeChannel(ch);
  },[org?.id]);

  const vendorColors=useMemo(()=>new Map(vendors.map((v,i)=>[v.id,PALETTE[i%PALETTE.length]])),[vendors]);

  // Build unified product list — one ranked list per catalog item
  // showing all vendor options cheapest first, with case AND each pricing
  const productList=useMemo(()=>{
    if(!catalogItems.length||!vendorItems.length||!mappings.length) return [];
    const viMap=new Map(vendorItems.map(vi=>[vi.id,vi]));
    const vMap=new Map(vendors.map(v=>[v.id,v]));

    return catalogItems.map(ci=>{
      const ciMappings=mappings.filter(m=>m.catalog_item_id===ci.id);
      const options=ciMappings.map(m=>{
        const vi=viMap.get(m.vendor_item_id);
        const v=vi?vMap.get(vi.vendor_id):null;
        if(!vi||!v||!vi.price) return null;
        const price=parseFloat(vi.price);
        const pack=vi.pack_size;
        const norm=normalizedPrice(price,pack);
        const each=eachPrice(price,pack);
        return {
          vendorId:v.id, vendorName:v.name,
          vendorItemId:vi.id, vendorItemCode:vi.vendor_item_code,
          brand:vi.brand, packSize:pack,
          casePrice:price,
          eachPrice:each?.price||null, eachSize:each?.size||null,
          normalizedPrice:norm?.price||null, normalizedUnit:norm?.unit||null,
          isExactMatch:m.comparison_track==="exact",
          color:vendorColors.get(v.id)||PALETTE[0],
        };
      }).filter(Boolean).sort((a,b)=>a.casePrice-b.casePrice);

      return {
        catalogItemId:ci.id,
        masterItemNumber:ci.master_item_number,
        name:ci.name,
        category:ci.catalog_categories?.name||"General",
        brandLocked:ci.brand_locked||false,
        lockedBrand:ci.locked_brand||null,
        options,
      };
    }).filter(c=>c.options.length>0);
  },[catalogItems,vendorItems,mappings,vendors,vendorColors]);

  const setQty=(key,val)=>setQuantities(p=>({...p,[key]:Math.max(0,val)}));

  // Build cart items from quantities
  const cartItems=useMemo(()=>{
    const items=[];
    for(const prod of productList){
      const caseKey=`${prod.catalogItemId}_case`;
      const eachKey=`${prod.catalogItemId}_each`;
      const caseQty=quantities[caseKey]||0;
      const eachQty=quantities[eachKey]||0;
      if(caseQty>0){
        items.push({...prod,quantity:caseQty,orderUnit:"case",
          options:prod.options.map(o=>({...o,price:o.casePrice,orderUnit:"case"}))});
      }
      if(eachQty>0&&prod.options.some(o=>o.eachPrice)){
        items.push({...prod,catalogItemId:`${prod.catalogItemId}_each`,quantity:eachQty,orderUnit:"each",
          options:prod.options.filter(o=>o.eachPrice).map(o=>({...o,price:o.eachPrice,packSize:o.eachSize,orderUnit:"each"}))});
      }
    }
    return items;
  },[productList,quantities]);

  const assignments=useMemo(()=>solve(cartItems,vendors),[cartItems,vendors]);
  const assignMap=useMemo(()=>new Map(assignments.map(a=>[a.catalogItemId,a])),[assignments]);

  const baskets=useMemo(()=>{
    const map=new Map();
    for(const a of assignments){
      const b=map.get(a.assignedVendorId)||{vendorId:a.assignedVendorId,vendorName:a.assignedVendorName,items:[],dollar:0,units:0};
      b.items.push(a);b.dollar=r2(b.dollar+a.lineTotal);b.units+=a.quantity;
      map.set(a.assignedVendorId,b);
    }
    return Array.from(map.values());
  },[assignments]);

  const baselineSpend=useMemo(()=>cartItems.reduce((s,i)=>{
    const cheapest=[...i.options].sort((a,b)=>a.price-b.price)[0];
    return s+(cheapest?.price||0)*i.quantity;
  },0),[cartItems]);
  const totalSpend=baskets.reduce((s,b)=>s+b.dollar,0);

  const filtered=useMemo(()=>
    productList.filter(i=>!search||i.name.toLowerCase().includes(search.toLowerCase())),
    [productList,search]);

  async function handleLogoUpload(file) {
    if (!file) return;
    setLogoUploading(true);
    const result = await uploadOrgLogo(org.id, file);
    if (result.error) {
      alert("Couldn't upload logo: " + result.error.message);
    } else {
      await supabase.from("organizations").update({ logo_url: result.path }).eq("id", org.id);
      const url = await getSignedUrl(result.path);
      setLogoUrl(url);
      setOrg(o => ({ ...o, logo_url: result.path }));
    }
    setLogoUploading(false);
  }

  async function submitOrders(){
    for(const basket of baskets){
      const {data:order}=await supabase.from("purchase_orders").insert({
        organization_id:org.id,vendor_id:basket.vendorId,
        created_by:session.user.id,status:"submitted",total_amount:basket.dollar,
      }).select().single();
      if(order){
        await supabase.from("purchase_order_lines").insert(basket.items.map(item=>({
          purchase_order_id:order.id,catalog_item_id:item.catalogItemId.replace("_each",""),
          vendor_item_id:item.vendorItemId,quantity:item.quantity,
          unit_price:item.price,line_total:item.lineTotal,
        })));
      }
    }
    alert(`${baskets.length} order${baskets.length>1?"s":""} submitted!`);
    setQuantities({});
  }

  if(session===undefined||loading) return (
    <div style={{minHeight:"100vh",background:"#003584",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{color:"white",textAlign:"center"}}>
        <div style={{display:"flex",justifyContent:"center",marginBottom:4}}><KerdosMark size={64}/></div>
        <div style={{fontWeight:900,fontSize:20,letterSpacing:"0.18em",color:"#4A90D9",marginTop:8}}>KERDOS</div>
        <div style={{marginTop:12,opacity:0.6,fontSize:13}}>Loading...</div>
      </div>
    </div>
  );
  if(!session) return <Login />;
  if(!org) return <OrgGate user={session.user} onComplete={o=>{setOrg(o);loadData();}} />;

  return (
    <div style={{fontFamily:"'Inter',-apple-system,sans-serif",minHeight:"100vh",background:"#F0F2F5"}}>

      {/* HEADER */}
      <header style={{background:"#003584",color:"white",padding:"0 16px",height:52,
        display:"flex",alignItems:"center",justifyContent:"space-between",
        position:"sticky",top:0,zIndex:200,boxShadow:"0 2px 8px rgba(0,0,0,0.3)"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer"}} onClick={()=>setTab("home")}>
          <KerdosMark size={30}/>
          <div style={{fontWeight:900,fontSize:15,letterSpacing:"0.18em",color:"#4A90D9"}}>KERDOS</div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:14}}>
          {totalSpend>0&&(
            <div style={{textAlign:"right"}}>
              <div style={{fontWeight:800,fontSize:17}}>${totalSpend.toFixed(2)}</div>
              {totalSpend>baselineSpend+0.01&&<div style={{fontSize:10,color:"#FF9800"}}>+${(totalSpend-baselineSpend).toFixed(2)} vs cheapest</div>}
            </div>
          )}
          <div style={{fontSize:11,opacity:0.6,textAlign:"right"}}>
            <div>{org.name}</div>
            <button onClick={()=>supabase.auth.signOut()} style={{background:"none",border:"none",color:"#4A90D9",cursor:"pointer",fontSize:11,padding:0}}>Sign out</button>
          </div>
        </div>
      </header>

      {/* TABS */}
      <div style={{background:"white",display:"flex",borderBottom:"1px solid #EEE",position:"sticky",top:52,zIndex:100}}>
        {[["home","🏠 Home"],["order","📋 Order"],
          ...(org.role!=="employee"?[["import","📥 Import"]]:[]),
          ["invoices","🧾 Invoices"],
          ...(org.role==="owner"||org.role==="manager"?[["team","👥 Team"]]:[])].map(([id,label])=>(
          <button key={id} onClick={()=>setTab(id)}
            style={{flex:1,padding:"12px 4px",border:"none",background:"none",cursor:"pointer",
              fontSize:12,fontWeight:600,
              color:tab===id?"#003584":"#888",
              borderBottom:tab===id?"2px solid #003584":"2px solid transparent"}}>
            {label}
          </button>
        ))}
      </div>

      <div style={{maxWidth:1200,margin:"0 auto",padding:"12px 12px 80px"}}>

        {/* HOME TAB */}
        {tab==="home"&&(
          <div>
            <div style={{textAlign:"center",padding:"32px 16px 24px"}}>
              {logoUrl?(
                <img src={logoUrl} alt={org.name} style={{maxHeight:72,maxWidth:260,objectFit:"contain"}} />
              ):org.role==="owner"?(
                <label style={{display:"inline-flex",flexDirection:"column",alignItems:"center",cursor:"pointer",padding:"14px 20px",borderRadius:12,border:"2px dashed #CCC",background:"white"}}>
                  <KerdosMark size={40}/>
                  <span style={{fontSize:12,color:"#888",marginTop:8}}>{logoUploading?"Uploading...":"+ Add your logo"}</span>
                  <input type="file" accept="image/*" style={{display:"none"}} disabled={logoUploading}
                    onChange={e=>handleLogoUpload(e.target.files[0])} />
                </label>
              ):(
                <div style={{display:"flex",justifyContent:"center"}}><KerdosMark size={48}/></div>
              )}
            </div>
            <div style={{display:"flex",gap:14,flexWrap:"wrap",justifyContent:"center"}}>
              <div onClick={()=>setTab("order")} style={{cursor:"pointer",background:"white",borderRadius:14,padding:"28px 24px",width:230,textAlign:"center",boxShadow:"0 1px 4px rgba(0,0,0,0.08)",border:"1px solid #EEE"}}>
                <div style={{fontSize:34,marginBottom:10}}>📋</div>
                <div style={{fontWeight:800,fontSize:16,marginBottom:4}}>Place an Order</div>
                <div style={{color:"#888",fontSize:12}}>Browse items and build vendor baskets</div>
              </div>
              {org.role!=="employee"&&(
                <div onClick={()=>{setSelectedVendorId(null);setImportMode("pricelist");setTab("import");}} style={{cursor:"pointer",background:"white",borderRadius:14,padding:"28px 24px",width:230,textAlign:"center",boxShadow:"0 1px 4px rgba(0,0,0,0.08)",border:"1px solid #EEE"}}>
                  <div style={{fontSize:34,marginBottom:10}}>📥</div>
                  <div style={{fontWeight:800,fontSize:16,marginBottom:4}}>Import Price List</div>
                  <div style={{color:"#888",fontSize:12}}>Update vendor pricing from a file</div>
                </div>
              )}
              <div onClick={()=>{setSelectedVendorId(null);setImportMode("invoice");setTab("invoices");}} style={{cursor:"pointer",background:"white",borderRadius:14,padding:"28px 24px",width:230,textAlign:"center",boxShadow:"0 1px 4px rgba(0,0,0,0.08)",border:"1px solid #EEE"}}>
                <div style={{fontSize:34,marginBottom:10}}>🧾</div>
                <div style={{fontWeight:800,fontSize:16,marginBottom:4}}>Invoices</div>
                <div style={{color:"#888",fontSize:12}}>Record or review past invoices</div>
              </div>
              {(org.role==="owner"||org.role==="manager")&&(
                <div onClick={()=>setTab("team")} style={{cursor:"pointer",background:"white",borderRadius:14,padding:"28px 24px",width:230,textAlign:"center",boxShadow:"0 1px 4px rgba(0,0,0,0.08)",border:"1px solid #EEE"}}>
                  <div style={{fontSize:34,marginBottom:10}}>👥</div>
                  <div style={{fontWeight:800,fontSize:16,marginBottom:4}}>Team</div>
                  <div style={{color:"#888",fontSize:12}}>Invite people, manage access</div>
                </div>
              )}
            </div>
            {vendors.length>0&&(
              <div style={{marginTop:32,textAlign:"center"}}>
                <div style={{fontSize:11,fontWeight:700,color:"#AAA",letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:10}}>Your Vendors — tap for details</div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap",justifyContent:"center"}}>
                  {vendors.map(v=>{
                    const vc=vendorColors.get(v.id)||PALETTE[0];
                    return (
                      <button key={v.id} onClick={()=>{setVendorDetailId(v.id);setTab("vendorDetail");}}
                        style={{fontSize:12,fontWeight:700,padding:"6px 12px",borderRadius:20,background:vc.bg,color:vc.accent,border:"none",cursor:"pointer"}}>
                        {v.name}
                      </button>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        )}

        {/* ORDER TAB */}
        {tab==="order"&&(
          <div style={{display:"flex",gap:12,alignItems:"flex-start"}}>
            <main style={{flex:1,minWidth:0}}>
              <input value={search} onChange={e=>setSearch(e.target.value)}
                placeholder={`🔍 Search ${productList.length} items...`}
                style={{...inp,marginBottom:10}} />

              {productList.length===0&&(
                <div style={{background:"white",borderRadius:10,padding:32,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
                  <div style={{fontSize:32,marginBottom:8}}>📋</div>
                  <h3 style={{margin:"0 0 8px"}}>No items yet</h3>
                  <p style={{color:"#888",fontSize:14,margin:"0 0 16px"}}>Import a vendor price list to get started</p>
                  <button onClick={()=>setTab("import")} style={{...btn("#003584")}}>Import Price List</button>
                </div>
              )}

              {filtered.map(item=>{
                const caseKey=`${item.catalogItemId}_case`;
                const eachKey=`${item.catalogItemId}_each`;
                const caseQty=quantities[caseKey]||0;
                const eachQty=quantities[eachKey]||0;
                const hasEach=item.options.some(o=>o.eachPrice);
                const cheapest=item.options[0];
                const assignment=assignMap.get(item.catalogItemId);
                const displayVendorId=assignment?.assignedVendorId||cheapest?.vendorId;
                const vc=vendorColors.get(displayVendorId)||PALETTE[0];
                const isActive=caseQty>0||eachQty>0;
                const isExpanded=expandedItem===item.catalogItemId;

                return (
                  <div key={item.catalogItemId} style={{
                    background:"white",borderRadius:8,marginBottom:6,
                    border:isActive?`2px solid ${vc.accent}`:"1px solid #E8E8E8",
                    boxShadow:isActive?`0 2px 8px ${vc.accent}20`:"none",
                  }}>
                    {/* MAIN ROW */}
                    <div style={{padding:"10px 12px"}}>
                      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:hasEach?8:0}}>
                        <div style={{flex:1,minWidth:0}}>
                          <div style={{fontWeight:600,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                            {item.name}
                            {item.brandLocked&&<span style={{marginLeft:6,fontSize:10,background:"#E3F2FD",color:"#1565C0",padding:"1px 5px",borderRadius:4,fontWeight:700}}>🔒 {item.lockedBrand}</span>}
                          </div>
                          <div style={{fontSize:10,color:"#AAA",marginTop:1}}>{item.category}</div>
                        </div>
                        {item.options.length>1&&(
                          <button onClick={()=>setExpandedItem(isExpanded?null:item.catalogItemId)}
                            style={{background:"none",border:"none",cursor:"pointer",color:"#CCC",fontSize:13,flexShrink:0}}>
                            {isExpanded?"▲":"▼"}
                          </button>
                        )}
                      </div>

                      {/* CASE ROW */}
                      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:hasEach?4:0}}>
                        <div style={{display:"flex",alignItems:"center",gap:5,flexShrink:0}}>
                          <button onClick={()=>setQty(caseKey,caseQty-1)}
                            style={{width:28,height:28,borderRadius:6,border:"1px solid #DDD",background:"#F5F5F5",cursor:"pointer",fontSize:16,fontWeight:700,color:"#444"}}>−</button>
                          <span style={{width:22,textAlign:"center",fontWeight:700,fontSize:13,color:caseQty>0?vc.accent:"#CCC"}}>{caseQty||"·"}</span>
                          <button onClick={()=>setQty(caseKey,caseQty+1)}
                            style={{width:28,height:28,borderRadius:6,border:"none",background:"#003584",cursor:"pointer",fontSize:16,fontWeight:700,color:"white"}}>+</button>
                        </div>
                        <div style={{fontSize:11,color:"#888",flex:1}}>
                          Case · {cheapest?.packSize||""}
                        </div>
                        <div style={{textAlign:"right",flexShrink:0}}>
                          <div style={{fontWeight:800,fontSize:14}}>${cheapest?.casePrice?.toFixed(2)}</div>
                          {cheapest?.normalizedPrice&&<div style={{fontSize:9,color:"#AAA"}}>${cheapest.normalizedPrice.toFixed(4)}/{cheapest.normalizedUnit}</div>}
                        </div>
                        <span style={{fontSize:10,fontWeight:700,padding:"2px 6px",borderRadius:8,background:vc.bg,color:vc.accent,flexShrink:0}}>
                          {cheapest?.vendorName}
                        </span>
                      </div>

                      {/* EACH ROW — only shown when each pricing available */}
                      {hasEach&&(
                        <div style={{display:"flex",alignItems:"center",gap:8}}>
                          <div style={{display:"flex",alignItems:"center",gap:5,flexShrink:0}}>
                            <button onClick={()=>setQty(eachKey,eachQty-1)}
                              style={{width:28,height:28,borderRadius:6,border:"1px solid #E8D5FF",background:"#FAF0FF",cursor:"pointer",fontSize:16,fontWeight:700,color:"#6A1B9A"}}>−</button>
                            <span style={{width:22,textAlign:"center",fontWeight:700,fontSize:13,color:eachQty>0?"#6A1B9A":"#CCC"}}>{eachQty||"·"}</span>
                            <button onClick={()=>setQty(eachKey,eachQty+1)}
                              style={{width:28,height:28,borderRadius:6,border:"none",background:"#6A1B9A",cursor:"pointer",fontSize:16,fontWeight:700,color:"white"}}>+</button>
                          </div>
                          <div style={{fontSize:11,color:"#888",flex:1}}>
                            Each · {cheapest?.eachSize||""}
                          </div>
                          <div style={{textAlign:"right",flexShrink:0}}>
                            <div style={{fontWeight:700,fontSize:13,color:"#6A1B9A"}}>${cheapest?.eachPrice?.toFixed(2)}</div>
                          </div>
                          <span style={{fontSize:10,fontWeight:700,padding:"2px 6px",borderRadius:8,background:"#F3E5F5",color:"#6A1B9A",flexShrink:0}}>
                            {cheapest?.vendorName}
                          </span>
                        </div>
                      )}
                    </div>

                    {/* EXPANDED — all vendor options ranked cheapest first */}
                    {isExpanded&&(
                      <div style={{borderTop:"1px solid #F5F5F5",padding:"8px 12px 10px",background:"#FAFAFA"}}>
                        <div style={{fontSize:10,color:"#BBB",fontWeight:700,letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:6}}>
                          All options — cheapest first
                        </div>
                        {item.options.map((opt,i)=>{
                          const ovc=vendorColors.get(opt.vendorId)||PALETTE[0];
                          return (
                            <div key={opt.vendorItemId} style={{
                              display:"flex",alignItems:"center",justifyContent:"space-between",
                              padding:"6px 8px",borderRadius:6,marginBottom:3,
                              background:i===0?ovc.bg:"white",
                              border:`1px solid ${i===0?ovc.light:"#F0F0F0"}`,
                            }}>
                              <div style={{display:"flex",alignItems:"center",gap:6,flex:1}}>
                                {i===0&&<span style={{fontSize:9,fontWeight:800,color:ovc.accent,background:"white",padding:"1px 5px",borderRadius:4}}>BEST</span>}
                                {opt.isExactMatch&&<span style={{fontSize:9,fontWeight:700,color:"#2E7D32",background:"#E8F5E9",padding:"1px 5px",borderRadius:4}}>EXACT</span>}
                                <span style={{fontSize:12,fontWeight:i===0?600:400}}>{opt.vendorName}</span>
                                {opt.brand&&<span style={{fontSize:10,color:"#888"}}>{opt.brand}</span>}
                                <span style={{fontSize:10,color:"#AAA"}}>{opt.packSize}</span>
                              </div>
                              <div style={{textAlign:"right",flexShrink:0}}>
                                <div style={{fontWeight:700,fontSize:13,color:i===0?ovc.accent:"#888"}}>${opt.casePrice.toFixed(2)}</div>
                                {opt.eachPrice&&<div style={{fontSize:10,color:"#AAA"}}>${opt.eachPrice.toFixed(2)}/each</div>}
                                {opt.normalizedPrice&&<div style={{fontSize:9,color:"#CCC"}}>${opt.normalizedPrice.toFixed(4)}/{opt.normalizedUnit}</div>}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                );
              })}
            </main>

            {/* BASKETS */}
            {baskets.length>0&&(
              <aside style={{width:290,flexShrink:0}}>
                <div style={{position:"sticky",top:110}}>
                  <div style={{fontSize:11,fontWeight:800,color:"#999",letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:10}}>Order Baskets</div>

                  {vendors.map(vendor=>{
                    const basket=baskets.find(b=>b.vendorId===vendor.id);
                    if(!basket) return null;
                    const vc=vendorColors.get(vendor.id)||PALETTE[0];
                    const meetsDollar=!vendor.delivery_minimum_dollar||basket.dollar>=vendor.delivery_minimum_dollar;
                    const meetsUnits=!vendor.delivery_minimum_units||basket.units>=vendor.delivery_minimum_units;
                    const meetsAll=meetsDollar&&meetsUnits;
                    return (
                      <div key={vendor.id} style={{background:"white",borderRadius:10,marginBottom:10,
                        border:`2px solid ${meetsAll?vc.accent:"#FFB74D"}`,overflow:"hidden"}}>
                        <div style={{background:meetsAll?vc.bg:"#FFF8E1",padding:"9px 12px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                          <div>
                            <div style={{fontWeight:800,fontSize:13,color:vc.accent}}>{vendor.name}</div>
                            <div style={{fontSize:10,color:"#999"}}>{basket.units} items</div>
                          </div>
                          <div style={{textAlign:"right"}}>
                            <div style={{fontWeight:800,fontSize:15}}>${basket.dollar.toFixed(2)}</div>
                            {meetsAll&&<div style={{fontSize:9,color:vc.accent,fontWeight:700}}>✓ READY</div>}
                          </div>
                        </div>
                        {vendor.delivery_minimum_dollar&&(
                          <div style={{padding:"6px 12px"}}>
                            <div style={{display:"flex",justifyContent:"space-between",fontSize:9,color:"#AAA",marginBottom:2}}>
                              <span>Min ${vendor.delivery_minimum_dollar}</span>
                              <span style={{color:meetsDollar?vc.accent:"#FF9800",fontWeight:600}}>
                                {meetsDollar?"✓":`$${(vendor.delivery_minimum_dollar-basket.dollar).toFixed(2)} to go`}
                              </span>
                            </div>
                            <div style={{height:3,background:"#EEE",borderRadius:2}}>
                              <div style={{height:"100%",borderRadius:2,transition:"width 0.3s",
                                background:meetsDollar?vc.accent:"#FF9800",
                                width:`${Math.min(100,(basket.dollar/vendor.delivery_minimum_dollar)*100)}%`}} />
                            </div>
                          </div>
                        )}
                        <div style={{padding:"4px 12px 8px",maxHeight:180,overflowY:"auto"}}>
                          {basket.items.map(item=>(
                            <div key={item.catalogItemId} style={{display:"flex",justifyContent:"space-between",padding:"2px 0",fontSize:11,borderBottom:"1px solid #F8F8F8"}}>
                              <span style={{color:"#444",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",flex:1}}>
                                {item.quantity}× {item.name} {item.orderUnit==="each"?"(each)":""}
                              </span>
                              <span style={{fontWeight:700,flexShrink:0,marginLeft:6,color:item.premiumPaid>0.005?"#FF9800":"#333"}}>${item.lineTotal.toFixed(2)}</span>
                            </div>
                          ))}
                        </div>
                      </div>
                    );
                  })}

                  <div style={{background:"#003584",borderRadius:10,padding:"14px 16px",color:"white"}}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                      <span style={{fontSize:11,opacity:0.6}}>Cheapest possible</span>
                      <span style={{fontWeight:600}}>${baselineSpend.toFixed(2)}</span>
                    </div>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}>
                      <span style={{fontSize:11,opacity:0.6}}>Optimized total</span>
                      <span style={{fontWeight:800,fontSize:16,color:totalSpend>baselineSpend+0.01?"#FF9800":"#69F0AE"}}>${totalSpend.toFixed(2)}</span>
                    </div>
                    <button onClick={submitOrders}
                      style={{...btn("#4A90D9"),width:"100%"}}>
                      Submit {baskets.length} Order{baskets.length>1?"s":""}
                    </button>
                  </div>
                </div>
              </aside>
            )}
          </div>
        )}

        {/* IMPORT TAB */}
        {tab==="import"&&(
          <div>
            <div style={{background:"white",borderRadius:10,padding:24,textAlign:"center",marginBottom:12,boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
              <div style={{fontSize:36,marginBottom:8}}>📋</div>
              <h3 style={{margin:"0 0 6px"}}>Import Price List or Invoice</h3>
              <p style={{color:"#888",fontSize:13,margin:"0 0 16px"}}>Copy from email, Excel, or any format and paste it in</p>
              <button onClick={()=>{setSelectedVendorId(null);setImportMode("pricelist");setShowPaste(true);}} style={{...btn("#003584")}}>Open Import Tool</button>
            </div>
            <div style={{background:"white",borderRadius:10,padding:16,boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
              <h4 style={{margin:"0 0 12px",fontSize:14}}>Your Vendors</h4>
              {vendors.map(v=>{
                const vc=vendorColors.get(v.id)||PALETTE[0];
                const count=vendorItems.filter(vi=>vi.vendor_id===v.id).length;
                return (
                  <div key={v.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",
                    padding:"10px 12px",borderRadius:8,marginBottom:6,background:vc.bg}}>
                    <div>
                      <div style={{fontWeight:700,color:vc.accent}}>{v.name}</div>
                      <div style={{fontSize:11,color:"#888"}}>{count} items · Min ${v.delivery_minimum_dollar||0}</div>
                    </div>
                    <button onClick={()=>{setSelectedVendorId(v.id);setImportMode("pricelist");setShowPaste(true);}} style={{...btn(vc.accent,"white",{fontSize:12,padding:"6px 12px"})}}>Import</button>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* INVOICES TAB */}
        {tab==="invoices"&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
              <h3 style={{margin:0,fontSize:16}}>Invoice History</h3>
              <button onClick={()=>{setSelectedVendorId(null);setImportMode("invoice");setShowPaste(true);}} style={{...btn("#003584","white",{fontSize:12,padding:"8px 14px"})}}>+ Record Invoice</button>
            </div>
            {invoices.length===0?(
              <div style={{background:"white",borderRadius:10,padding:32,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
                <div style={{fontSize:32,marginBottom:8}}>🧾</div>
                <p style={{color:"#888"}}>No invoices recorded yet</p>
              </div>
            ):invoices.map(inv=>(
              <div key={inv.id} style={{background:"white",borderRadius:8,padding:14,marginBottom:8,
                display:"flex",justifyContent:"space-between",alignItems:"center",
                boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
                <div>
                  <div style={{fontWeight:700,fontSize:14}}>{inv.vendors?.name||"Unknown"}</div>
                  <div style={{fontSize:11,color:"#888"}}>{formatDateMDY(inv.invoice_date)||new Date(inv.created_at).toLocaleDateString()} · {inv.status}</div>
                </div>
                <div style={{display:"flex",alignItems:"center",gap:10}}>
                  <div style={{fontWeight:800,fontSize:16}}>${parseFloat(inv.total_amount||0).toFixed(2)}</div>
                  {inv.file_path&&(
                    <button onClick={()=>viewStoredFile(inv.file_path)}
                      style={{...btn("#003584","white",{fontSize:11,padding:"5px 10px"})}}>View</button>
                  )}
                </div>
              </div>
            ))}
          </div>
        )}

        {/* VENDOR DETAIL TAB */}
        {tab==="vendorDetail"&&vendorDetailId&&(()=>{
          const v=vendors.find(x=>x.id===vendorDetailId);
          if(!v) return <p>Vendor not found.</p>;
          const vc=vendorColors.get(v.id)||PALETTE[0];
          const itemCount=vendorItems.filter(vi=>vi.vendor_id===v.id).length;
          const vendorInvoices=invoices.filter(inv=>inv.vendor_id===v.id);
          return (
            <div>
              <button onClick={()=>setTab("home")} style={{background:"none",border:"none",cursor:"pointer",color:"#888",fontSize:13,marginBottom:14,padding:0}}>← Back</button>
              <div style={{background:"white",borderRadius:12,padding:20,marginBottom:14,boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}>
                  <div>
                    <div style={{fontWeight:900,fontSize:20,color:vc.accent}}>{v.name}</div>
                    <div style={{fontSize:12,color:"#888",marginTop:2}}>{itemCount} items in catalog · Min ${v.delivery_minimum_dollar||0}</div>
                  </div>
                  <a href={`https://www.google.com/search?q=${encodeURIComponent(v.name)}`} target="_blank" rel="noreferrer"
                    style={{fontSize:12,color:"#003584",textDecoration:"none",fontWeight:700}}>Research ↗</a>
                </div>
                <div style={{display:"flex",gap:8}}>
                  {org.role!=="employee"&&(
                    <button onClick={()=>{setSelectedVendorId(v.id);setImportMode("pricelist");setShowPaste(true);}}
                      style={{...btn(vc.accent,"white",{flex:1})}}>📥 Import Price List</button>
                  )}
                  <button onClick={()=>{setSelectedVendorId(v.id);setImportMode("invoice");setShowPaste(true);}}
                    style={{...btn("#003584","white",{flex:1})}}>🧾 Record Invoice</button>
                </div>
              </div>
              <div style={{fontSize:11,fontWeight:800,color:"#999",letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:8}}>Invoice history</div>
              {vendorInvoices.length===0?(
                <div style={{background:"white",borderRadius:10,padding:24,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
                  <p style={{color:"#888",fontSize:13,margin:0}}>No invoices recorded for {v.name} yet</p>
                </div>
              ):vendorInvoices.map(inv=>(
                <div key={inv.id} style={{background:"white",borderRadius:8,padding:14,marginBottom:8,
                  display:"flex",justifyContent:"space-between",alignItems:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
                  <div style={{fontSize:12,color:"#888"}}>{formatDateMDY(inv.invoice_date)||new Date(inv.created_at).toLocaleDateString()}</div>
                  <div style={{display:"flex",alignItems:"center",gap:10}}>
                    <div style={{fontWeight:800,fontSize:15}}>${parseFloat(inv.total_amount||0).toFixed(2)}</div>
                    {inv.file_path&&<button onClick={()=>viewStoredFile(inv.file_path)} style={{...btn("#003584","white",{fontSize:11,padding:"5px 10px"})}}>View</button>}
                  </div>
                </div>
              ))}
            </div>
          );
        })()}

        {/* TEAM TAB */}
        {tab==="team"&&(org.role==="owner"||org.role==="manager")&&(
          <TeamPanel orgId={org.id} myRole={org.role} currentUserId={session.user.id} />
        )}
      </div>

      {showPaste&&<PasteModal vendors={vendors} orgId={org.id} onClose={()=>setShowPaste(false)} onDone={loadData} initialVendorId={selectedVendorId} initialMode={importMode} />}
    </div>
  );
}
