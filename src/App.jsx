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
          <div style={{fontSize:36}}>🦉</div>
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
        <div style={{fontWeight:900,fontSize:16,letterSpacing:"0.18em",color:"#003584",marginBottom:4}}>🦉 KERDOS</div>
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
        <div style={{fontSize:48}}>🦉</div>
        <div style={{fontWeight:900,fontSize:20,letterSpacing:"0.18em",color:"#4A90D9",marginTop:8}}>KERDOS</div>
        <div style={{marginTop:12,opacity:0.6,fontSize:13}}>Loading...</div>
      </div>
    </div>
  );
  if(!session) return <Login />;
  if(!org) return <Setup user={session.user} onComplete={o=>{setOrg(o);loadData();}} />;

  return (
    <div style={{fontFamily:"'Inter',-apple-system,sans-serif",minHeight:"100vh",background:"#F0F2F5"}}>

      {/* HEADER */}
      <header style={{background:"#003584",color:"white",padding:"0 16px",height:52,
        display:"flex",alignItems:"center",justifyContent:"space-between",
        position:"sticky",top:0,zIndex:200,boxShadow:"0 2px 8px rgba(0,0,0,0.3)"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer"}} onClick={()=>setTab("home")}>
          <span style={{fontSize:24}}>🦉</span>
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
        {[["home","🏠 Home"],["order","📋 Order"],["import","📥 Import"],["invoices","🧾 Invoices"]].map(([id,label])=>(
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
              <div style={{fontSize:40,marginBottom:8}}>🦉</div>
              <h2 style={{margin:"0 0 4px",fontSize:22}}>Welcome back, {org.name}</h2>
              <p style={{color:"#888",fontSize:14,margin:0}}>What would you like to do?</p>
            </div>
            <div style={{display:"flex",gap:14,flexWrap:"wrap",justifyContent:"center"}}>
              <div onClick={()=>setTab("order")} style={{cursor:"pointer",background:"white",borderRadius:14,padding:"28px 24px",width:230,textAlign:"center",boxShadow:"0 1px 4px rgba(0,0,0,0.08)",border:"1px solid #EEE"}}>
                <div style={{fontSize:34,marginBottom:10}}>📋</div>
                <div style={{fontWeight:800,fontSize:16,marginBottom:4}}>Place an Order</div>
                <div style={{color:"#888",fontSize:12}}>Browse items and build vendor baskets</div>
              </div>
              <div onClick={()=>{setSelectedVendorId(null);setImportMode("pricelist");setTab("import");}} style={{cursor:"pointer",background:"white",borderRadius:14,padding:"28px 24px",width:230,textAlign:"center",boxShadow:"0 1px 4px rgba(0,0,0,0.08)",border:"1px solid #EEE"}}>
                <div style={{fontSize:34,marginBottom:10}}>📥</div>
                <div style={{fontWeight:800,fontSize:16,marginBottom:4}}>Import Price List</div>
                <div style={{color:"#888",fontSize:12}}>Update vendor pricing from a file</div>
              </div>
              <div onClick={()=>{setSelectedVendorId(null);setImportMode("invoice");setTab("invoices");}} style={{cursor:"pointer",background:"white",borderRadius:14,padding:"28px 24px",width:230,textAlign:"center",boxShadow:"0 1px 4px rgba(0,0,0,0.08)",border:"1px solid #EEE"}}>
                <div style={{fontSize:34,marginBottom:10}}>🧾</div>
                <div style={{fontWeight:800,fontSize:16,marginBottom:4}}>Invoices</div>
                <div style={{color:"#888",fontSize:12}}>Record or review past invoices</div>
              </div>
            </div>
            {vendors.length>0&&(
              <div style={{marginTop:32,textAlign:"center"}}>
                <div style={{fontSize:11,fontWeight:700,color:"#AAA",letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:10}}>Your Vendors</div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap",justifyContent:"center"}}>
                  {vendors.map(v=>{
                    const vc=vendorColors.get(v.id)||PALETTE[0];
                    return (
                      <span key={v.id} style={{fontSize:12,fontWeight:700,padding:"6px 12px",borderRadius:20,background:vc.bg,color:vc.accent}}>
                        {v.name}
                      </span>
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
                  <div style={{fontSize:11,color:"#888"}}>{inv.invoice_date||new Date(inv.created_at).toLocaleDateString()} · {inv.status}</div>
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
      </div>

      {showPaste&&<PasteModal vendors={vendors} orgId={org.id} onClose={()=>setShowPaste(false)} onDone={loadData} initialVendorId={selectedVendorId} initialMode={importMode} />}
    </div>
  );
}
