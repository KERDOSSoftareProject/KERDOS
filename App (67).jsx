import { useState, useEffect, useMemo } from "react";
import * as XLSXLib from "xlsx";
import * as pdfjsLib from "pdfjs-dist";
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.js?url";
import { data as supabase, sessionController, backendInfo } from "./data.js";
import { parseDocument, findDate, findInvoiceNumber } from "./ingestion.js";
import { classifyCategory, nextCategoryRange, normalizedPrice, eachPrice, safeProductScore, bestInvoiceMatch, MATCH_POLICY } from "./procurement.js";


// Bounds an async action to a maximum wait, so a hung network call (bad
// connection, a backend outage, a request that never resolves either
// way) becomes a clear, visible error after a fixed wait instead of
// leaving a button stuck on "Loading..." forever with no feedback and no
// way for the person to know whether to keep waiting or try again.
function withTimeout(promise, ms, label) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error(`${label} timed out after ${ms/1000}s - check your connection and try again`)), ms)),
  ]);
}

// XLSX and pdf.js are build dependencies, not runtime CDN fetches, so
// parsing a price sheet or invoice cannot fail because a third-party CDN
// is unreachable, blocked on a restaurant's network, or serving a
// different build than the one this app was tested against. Kept as an
// async function so every existing call site is unchanged.
async function loadFileLib() {
  return XLSXLib;
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
  pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;
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

  // OCR is the only remaining runtime-loaded library, and deliberately so:
  // it is ~2MB and only needed for image-only PDFs with no extractable
  // text, so bundling it would slow every page load for a path most
  // imports never reach. Everything else (Excel, CSV, text PDFs) is
  // bundled and works with no network access to a third party. If OCR
  // can't load, say so plainly rather than failing with a cryptic error.
  try {
    await loadScript(
      "https://cdnjs.cloudflare.com/ajax/libs/tesseract.js/4.1.1/tesseract.min.js", "Tesseract"
    );
  } catch {
    throw new Error("This PDF appears to be a scan with no readable text, and the text-recognition library could not be loaded. Check the network connection, or upload this vendor's file as a CSV/Excel export instead.");
  }
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

// ── CATALOG MATCHING ────────────────────────────────────────────────
// This is the actual "master catalog" engine: every vendor uses its own
// item numbers and wording, so the only way to let someone compare "the
// same product" across vendors is to link each vendor's item to one
// shared catalog entry. Deliberately rule-based (no AI at runtime, same
// as the rest of ingestion) — uses the size-aware safeProductScore
// engine imported from procurement.js (shared with invoice matching,
// and free of any industry vocabulary) rather than plain
// word overlap, since plain overlap was confirmed to misjudge different
// cuts/pack sizes as the same product. A strong match reuses an existing
// catalog item (comparison_track "exact"); a partial match still reuses
// it but is marked "similar" for lower-confidence display; no reasonable
// match creates a brand-new catalog item so the product is at least
// visible and orderable, ready to pick up a second vendor later.

async function ensureUncategorized(orgId, categories) {
  // "Uncategorized" is its own real category now, not interchangeable
  // with "General" - General (with its own keyword list) is a normal,
  // auto-classifiable category like Produce or Meat: confirmed to be
  // genuinely miscellaneous. Uncategorized is the holding pen: anything
  // that didn't match ANY category's keywords - General's included - and
  // needs a person to actually allocate it. Matching on "General" too
  // here would misfile real, correctly-classified General items into
  // this catch-all, so this only ever looks for "Uncategorized" by name.
  const existing = categories.find(c => c.name === "Uncategorized");
  if (existing) return existing;
  const { range_start, range_end } = nextCategoryRange(categories);
  const { data } = await supabase.from("catalog_categories").insert({
    organization_id: orgId, name: "Uncategorized", range_start, range_end, keywords: [],
  }).select().single();
  if (data) categories.push(data);
  return data;
}

// Reads global (not org-specific) starter templates for a given industry
// string. Adding support for a new industry is a data insert into
// industry_templates, never a code change — see catalog_categories_migration.sql.
async function loadIndustryTemplates(industry) {
  if (!industry || !industry.trim()) return [];
  const { data } = await supabase.from("industry_templates")
    .select("*").ilike("industry", industry.trim()).order("sort_order");
  return data || [];
}

// Actually creates the starter categories for a given industry as real
// catalog_categories rows for this org - shared by TeamPanel (auto-load
// right after picking an industry, with an opt-out) and CatalogPanel's
// manual "Load starter categories" button, so there's exactly one place
// that does this instead of two copies that could drift apart. Skips any
// category name this org already has (so calling it twice, or once from
// each caller, never creates duplicates).
async function loadStarterCategoriesForIndustry(orgId, industry, categories) {
  const templates = await loadIndustryTemplates(industry);
  if (!templates.length) return { added: 0, found: false };
  const existingNames = new Set(categories.map(c => c.name.toLowerCase()));
  const toAdd = templates.filter(t => !existingNames.has(t.category_name.toLowerCase()));
  const workingCategories = [...categories];
  const toInsert = toAdd.map(t => {
    const { range_start, range_end } = nextCategoryRange(workingCategories);
    workingCategories.push({ range_start, range_end }); // reserve this block before allocating the next
    return { organization_id: orgId, name: t.category_name, keywords: t.keywords, range_start, range_end };
  });
  if (toInsert.length) await supabase.from("catalog_categories").insert(toInsert);
  return { added: toInsert.length, found: true };
}

// Finds the best existing catalog item to attach a newly-imported vendor
// item to, or creates a new one when nothing reasonably matches. Mutates
// `workingCatalogItems` in place so multiple rows in the same import batch
// correctly match against catalog items created earlier in that same batch.
// `categories` is this org's own catalog_categories rows (with keywords);
// mutated in place the same way when the Uncategorized fallback gets
// created on first use.
async function matchOrCreateCatalogItem(orgId, description, workingCatalogItems, categories) {
  let best = null, bestScore = 0;
  for (const ci of workingCatalogItems) {
    const score = safeProductScore(description, ci.name);
    if (score > bestScore) { bestScore = score; best = ci; }
  }
  if (best && bestScore >= 0.5) {
    return { catalogItemId: best.id, track: bestScore >= MATCH_POLICY.autoLink ? "exact" : "similar", score: bestScore };
  }
  const category = classifyCategory(description, categories) || await ensureUncategorized(orgId, categories);
  // Numbered within the category's own block, not one global counter -
  // this is what actually makes the numbers read as a series per
  // category (e.g. everything in the 3000s is one category) instead of
  // one running count across the whole catalog.
  const itemsInCategory = workingCatalogItems.filter(ci => ci.category_id === category?.id);
  const nextNumber = itemsInCategory.length
    ? Math.max(...itemsInCategory.map(ci => ci.master_item_number || 0)) + 1
    : (category?.range_start || 1);
  const { data: created, error } = await supabase.from("catalog_items").insert({
    organization_id: orgId, category_id: category?.id || null,
    master_item_number: nextNumber, name: description.slice(0, 120),
    matching_behavior: "flexible", canonical_unit: null, brand_locked: false,
  }).select().single();
  if (error || !created) return null;
  workingCatalogItems.push(created);
  return { catalogItemId: created.id, track: "new", score: null };
}


// Every org sets its own refresh cadence (organizations.settings.price_refresh_days) —
// a deli might want weekly, a building-supply distributor might want monthly. No
// schedule is assumed here; if the org hasn't set one, prices simply never expire.
function isPriceExpired(lastUpdated, refreshDays) {
  if (!refreshDays || !lastUpdated) return false;
  const ms = Date.now() - new Date(lastUpdated).getTime();
  return ms / (1000 * 60 * 60 * 24) > refreshDays;
}

// ── FORMAL DOCUMENT OUTPUT ───────────────────────────────────────────
// Everything above this line reads, understands, matches, and cleans up
// whatever comes in - but none of it is worth anything if the cleaned
// result never leaves the app as something a person can actually use:
// open in Excel, print, email to a vendor, hand to an accountant, file
// away. CSV is the one format that's universally usable everywhere
// (Excel, Sheets, Numbers, a text editor) without needing any special
// software - so it's the baseline output format for every formal
// document this app produces, regardless of what industry the org is in.
function csvEscape(val) {
  const s = String(val ?? "");
  if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function rowsToCSV(rows) {
  return rows.map(r => r.map(csvEscape).join(",")).join("\r\n");
}

function downloadTextFile(filename, content, mimeType) {
  const blob = new Blob([content], { type: mimeType + ";charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// One row per catalog item: master number, name, category, best price,
// best vendor, a plain-language status derived from the same confidence
// tiers used everywhere else in the app, and then one column per vendor
// so every vendor's price sits side by side - the actual "compare prices
// across vendors" promise, as a document instead of only a live screen.
// Column set is built from whichever vendors this org actually has, not
// any fixed list - works identically for any industry.
function buildCatalogExportCSV(productList, vendors) {
  const header = ["Master #","Item","Category","Best Price","Best Vendor","Status","Confidence",
    ...vendors.map(v => v.name)];
  const rows = [header];
  for (const item of productList) {
    const usable = item.options.filter(o => !o.expired);
    const cheapest = usable[0] || item.options[0] || null;
    let status = "No price on file";
    let confidence = "";
    if (cheapest) {
      if (cheapest.expired) status = "Stale — needs refresh";
      else if (cheapest.matchTrack === "similar") { status = "Needs review"; confidence = `${cheapest.matchConfidence}%`; }
      else status = "100% matched";
    }
    const vendorCells = vendors.map(v => {
      const opt = item.options.find(o => o.vendorId === v.id && !o.expired);
      return opt ? formatMoney(opt.casePrice) : "";
    });
    rows.push([
      item.masterItemNumber, item.name, item.category,
      cheapest && !cheapest.expired ? formatMoney(cheapest.casePrice) : "",
      cheapest && !cheapest.expired ? cheapest.vendorName : "",
      status, confidence,
      ...vendorCells,
    ]);
  }
  return rowsToCSV(rows);
}

// One row per invoice line with a real price variance, across whatever
// date range is passed in - a clean, exportable record of "here's every
// time we were charged something different from what we were quoted",
// suitable for sending back to a vendor as backup or keeping for
// accounting. Lines with no variance (or nothing to compare against)
// are intentionally left out - this document is specifically the
// discrepancy record, not a full invoice dump.
function buildVarianceReportCSV(invoices, vendors) {
  const vMap = new Map(vendors.map(v => [v.id, v]));
  const header = ["Date","Vendor","Invoice #","Item","Quoted Price","Charged Price","Difference","Line Total Impact"];
  const rows = [header];
  for (const inv of invoices) {
    const v = vMap.get(inv.vendor_id);
    for (const line of (inv.invoice_lines || [])) {
      if (line.price_variance == null || Math.abs(line.price_variance) < 0.005) continue;
      rows.push([
        inv.invoice_date || "", v?.name || "", inv.invoice_number || "",
        line.description, formatMoney(line.unit_price - line.price_variance),
        formatMoney(line.unit_price), formatMoney(line.price_variance),
        formatMoney(r2(line.price_variance * (line.line_total && line.unit_price ? line.line_total / line.unit_price : 1))),
      ]);
    }
  }
  return rowsToCSV(rows);
}

function formatMoney(n) {
  const v = parseFloat(n || 0);
  return (v < 0 ? "-$" : "$") + Math.abs(v).toFixed(2);
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
  let assignments = cartItems.map(item => {
    // Expired (stale-price) options are shown to the person so they know a
    // vendor's price needs refreshing, but must never be picked as the
    // "cheapest" option — their price is forced to 0 purely as a visual
    // signal, not a real price to order at.
    const validOptions = item.options.filter(o=>!o.expired);
    const sorted = [...validOptions].sort((a,b)=>a.price-b.price);
    const cheapestOption = sorted[0] || item.options[0] || null;
    if (!cheapestOption) {
      return {...item, assignedVendorId:null, assignedVendorName:null, vendorItemId:null,
        price:0, packSize:null, orderUnit:item.orderUnit, lineTotal:0,
        cheapestPrice:0, premiumPaid:0, locked:false, allExpired:true};
    }
    const best = item.forcedVendorId
      ? (validOptions.find(o=>o.vendorId===item.forcedVendorId) || cheapestOption)
      : cheapestOption;
    const effectivePrice = item.forcedPrice!=null ? item.forcedPrice : best.price;
    const lineTotal = r2(effectivePrice*item.quantity);
    return {...item, assignedVendorId:best.vendorId, assignedVendorName:best.vendorName,
      vendorItemId:best.vendorItemId, price:effectivePrice, packSize:best.packSize,
      orderUnit:best.orderUnit, lineTotal,
      cheapestPrice:cheapestOption.price, premiumPaid:r2(Math.max(0,lineTotal-cheapestOption.price*item.quantity)),
      locked:!!item.forcedVendorId, allExpired:!!cheapestOption.expired};
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
      // Locked (manually-assigned) items are never used to help fill a
      // shortfall elsewhere, and are never moved away from their assigned
      // vendor even if that vendor itself is short — a manual choice stays put.
      const fills = assignments
        .filter(a=>a.assignedVendorId!==req.id&&!a.locked)
        .map(a=>{
          const opt=a.options.find(o=>o.vendorId===req.id&&!o.expired);
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
        if(a.assignedVendorId!==req.id||a.locked) return a;
        const alt=[...a.options].filter(o=>o.vendorId!==req.id&&!o.expired).sort((x,y)=>x.price-y.price)[0];
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

// Filter/sort pill buttons (category chips, "Full List", sort-mode
// toggles) sit directly on the app's dark blue page background. The
// SELECTED state is the one that must visually pop (solid white); the
// unselected state stays legible without competing for attention (a
// translucent outline reads clearly against dark blue). Never colour a
// selected state the same as the page background - it reads as
// unselected.
const chipStyle = (isSelected, size="md") => ({
  fontSize:size==="sm"?11:12, fontWeight:700, cursor:"pointer",
  padding:size==="sm"?"5px 12px":"6px 14px", borderRadius:size==="sm"?16:20,
  background:isSelected?"white":"rgba(255,255,255,0.14)",
  color:isSelected?"#003584":"white",
  border:isSelected?"2px solid white":"2px solid rgba(255,255,255,0.3)",
});

// ── LANDING ───────────────────────────────────────────────────────────
function LandingGate() {
  const [mode,setMode]=useState("login");
  const [email,setEmail]=useState("");
  const [password,setPassword]=useState("");
  const [loading,setLoading]=useState(false);
  const [error,setError]=useState("");

  async function submit(e) {
    e.preventDefault(); setLoading(true); setError("");
    const {error:err} = mode==="login" ? await supabase.auth.signInWithPassword({email,password}) : await supabase.auth.signUp({email,password});
    if(err) setError(err.message);
    else if(mode==="signup") setError("Check your email to confirm, then sign in.");
    setLoading(false);
  }

  const SERVICES=[
    {icon:"💰",color:"#2E7D32",bg:"#E8F5E9",name:"Price Comparison",tag:null,
      text:"See every vendor's price for the same product side-by-side, ranked cheapest first.",
      value:"Save money on every order."},
    {icon:"✅",color:"#1565C0",bg:"#E3F2FD",name:"Price Verification",tag:null,
      text:"Every invoice is checked against what was quoted, automatically — mismatches are flagged.",
      value:"Ensures the price you're quoted is the price you pay."},
    {icon:"🗄️",color:"#6A1B9A",bg:"#F3E5F5",name:"Invoice Retention",tag:null,
      text:"Every invoice kept and organized, never lost in a shoebox or a shared drive.",
      value:"Built for IRS-ready record keeping."},
    {icon:"🔗",color:"#888",bg:"#F0F0F0",name:"QuickBooks Integration",tag:"Coming Soon",
      text:"Send recorded invoices straight to QuickBooks — no re-entry.",
      value:"Easier accounting, one click away."},
    {icon:"📦",color:"#E65100",bg:"#FFF3E0",name:"Inventory Management",tag:null,
      text:"Every purchase already tracked — set a par level per item and know when you're running low.",
      value:"Never run out, never over-order."},
  ];

  return (
    <div style={{minHeight:"100vh",background:"linear-gradient(160deg,#003584 0%,#00204F 100%)",padding:"48px 20px 60px"}}>
      <div style={{textAlign:"center",marginBottom:40}}>
        <div style={{fontSize:64,marginBottom:14}}>🦉</div>
        <div style={{fontWeight:900,fontSize:34,letterSpacing:"0.2em",color:"white",marginBottom:14}}>KERDOS</div>
        <div style={{fontWeight:900,fontSize:22,color:"white",marginBottom:10,maxWidth:480,marginLeft:"auto",marginRight:"auto",lineHeight:1.3}}>
          Stop overpaying because you didn't check the other vendor.
        </div>
        <div style={{color:"rgba(255,255,255,0.7)",fontSize:15,maxWidth:480,margin:"0 auto"}}>
          Procurement software that compares vendor prices for you, keeps your ordering and records in one place.
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit, minmax(260px, 1fr))",gap:16,maxWidth:1000,margin:"0 auto 40px"}}>
        {SERVICES.map((s,i)=>(
          <div key={i} style={{background:"rgba(255,255,255,0.08)",borderRadius:14,padding:"20px",position:"relative"}}>
            {s.tag&&<div style={{position:"absolute",top:14,right:14,fontSize:10,fontWeight:800,color:"white",background:"rgba(255,255,255,0.2)",padding:"3px 8px",borderRadius:20}}>{s.tag}</div>}
            <div style={{width:44,height:44,borderRadius:12,background:s.bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,marginBottom:12}}>{s.icon}</div>
            <div style={{color:"white",fontWeight:800,fontSize:16,marginBottom:6}}>{s.name}</div>
            <div style={{color:"rgba(255,255,255,0.75)",fontSize:13,lineHeight:1.45,marginBottom:8}}>{s.text}</div>
            <div style={{color:s.tag?"rgba(255,255,255,0.5)":"#69F0AE",fontWeight:700,fontSize:12}}>{s.value}</div>
          </div>
        ))}
      </div>

      <div style={{background:"white",borderRadius:12,padding:28,width:"100%",maxWidth:380,margin:"0 auto",boxShadow:"0 4px 20px rgba(0,0,0,0.3)"}}>
        <div style={{display:"flex",background:"#F0F2F5",borderRadius:8,padding:3,marginBottom:20}}>
          <button onClick={()=>{setMode("login");setError("");}}
            style={{flex:1,padding:"9px",borderRadius:6,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,
              background:mode==="login"?"white":"transparent",color:mode==="login"?"#003584":"#888",
              boxShadow:mode==="login"?"0 1px 3px rgba(0,0,0,0.1)":"none"}}>
            Sign In
          </button>
          <button onClick={()=>{setMode("signup");setError("");}}
            style={{flex:1,padding:"9px",borderRadius:6,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,
              background:mode==="signup"?"white":"transparent",color:mode==="signup"?"#003584":"#888",
              boxShadow:mode==="signup"?"0 1px 3px rgba(0,0,0,0.1)":"none"}}>
            Sign Up
          </button>
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
      </div>
    </div>
  );
}

// ── SETUP WIZARD ─────────────────────────────────────────────────────
const SETUP_DRAFT_KEY = "kerdos_setup_draft";

function Setup({user,onComplete}) {
  // Nothing here is saved to the database until the final "Get Started"
  // click - it's a multi-step form, so a refresh, an accidental
  // navigation, or just stepping away mid-fill would otherwise lose
  // everything typed. Autosaving a local draft (keyed to this browser
  // AND this specific user, so it can't leak to someone else signing up
  // on the same shared computer) means coming back restores exactly
  // where you left off, without needing a half-created organization
  // sitting in the database in the meantime.
  const draftKey = SETUP_DRAFT_KEY + "_" + user.id;
  const [step,setStep]=useState(1);
  const [orgName,setOrgName]=useState("");
  const [industry,setIndustry]=useState("");
  const [vendors,setVendors]=useState([{name:"",minDollar:"",minUnits:""}]);
  const [loading,setLoading]=useState(false);
  const [error,setError]=useState("");
  const [draftRestored,setDraftRestored]=useState(false);

  useEffect(()=>{
    try{
      const saved=localStorage.getItem(draftKey);
      if(saved){
        const d=JSON.parse(saved);
        if(d.orgName) setOrgName(d.orgName);
        if(d.industry) setIndustry(d.industry);
        if(d.vendors&&d.vendors.length) setVendors(d.vendors);
        if(d.step) setStep(d.step);
        if(d.orgName||d.industry) setDraftRestored(true);
      }
    }catch(e){/* corrupted or unavailable draft - just start fresh */}
  },[]);

  useEffect(()=>{
    try{
      localStorage.setItem(draftKey,JSON.stringify({step,orgName,industry,vendors}));
    }catch(e){/* storage full/unavailable - draft save is best-effort, never blocks typing */}
  },[step,orgName,industry,vendors]);

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
      try{localStorage.removeItem(draftKey);}catch(e){}
      onComplete(org);
    } catch(err){setError(err.message);}
    setLoading(false);
  }

  return (
    <div style={{minHeight:"100vh",background:"#F0F2F5",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:"white",borderRadius:12,padding:28,width:"100%",maxWidth:480,boxShadow:"0 2px 8px rgba(0,0,0,0.1)"}}>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
          <span style={{fontSize:20}}>🦉</span>
          <span style={{fontWeight:900,fontSize:16,letterSpacing:"0.18em",color:"#003584"}}>KERDOS</span>
        </div>
        <h2 style={{margin:"0 0 4px",fontSize:18}}>Welcome — let's get set up</h2>
        <p style={{color:"#888",fontSize:13,margin:"0 0 20px"}}>Takes about 2 minutes.</p>
        {draftRestored&&(
          <div style={{background:"#E8F5E9",color:"#2E7D32",fontSize:12,fontWeight:600,borderRadius:6,padding:"8px 10px",marginBottom:14}}>
            ✓ Picked up where you left off
          </div>
        )}

        {step===1&&<>
          <div style={{marginBottom:12}}>
            <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Organization name</div>
            <input style={inp} value={orgName} onChange={e=>setOrgName(e.target.value)} placeholder="Your organization's name" />
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
        <div style={{display:"flex",justifyContent:"center",marginBottom:8}}><span style={{fontSize:36}}>🦉</span></div>
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
          <span style={{fontSize:20}}>🦉</span>
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

// Common industry presets shown as quick-pick chips in Admin - purely a
// UI convenience list (the underlying org.industry column stays free
// text, and "Load starter categories" still works for anything typed in,
// preset or custom). Not an exhaustive or hardcoded-logic list; an org
// can always type something else instead.
const INDUSTRY_PRESETS = ["Restaurant","Deli & Prepared Foods","Bar & Pub","Bakery","Catering","Grocery & Convenience","Food Truck","Building Supply","Retail"];

// Common price-refresh cadences shown as quick-pick chips in Admin -
// same idea as INDUSTRY_PRESETS: a convenience shortcut for the most
// common choices, with a custom number field always available alongside
// for anything else (or "Never" to turn the whole feature off).
const REFRESH_PRESETS = [[3,"3 days"],[7,"Weekly"],[14,"Every 2 weeks"],[30,"Monthly"]];

function TeamPanel({orgId,orgName,orgIndustry,orgSettings,categories,myRole,currentUserId,currentUserEmail,onOrgUpdated,logoUrl,onLogoUpload,logoUploading}) {
  const [editingOrgName,setEditingOrgName]=useState(false);
  const [orgNameInput,setOrgNameInput]=useState(orgName);
  const [savingOrgName,setSavingOrgName]=useState(false);
  const [editingIndustry,setEditingIndustry]=useState(false);
  const [industryInput,setIndustryInput]=useState(orgIndustry||"");
  const [savingIndustry,setSavingIndustry]=useState(false);
  // Checked by default: picking an industry and getting that industry's
  // starter categories is the whole point of the Industry field for most
  // people - but it's still a real opt-out, not just an FYI, for anyone
  // who wants to build their category list by hand instead.
  const [autoLoadCategories,setAutoLoadCategories]=useState(true);
  const [industryMsg,setIndustryMsg]=useState("");
  const [editingRefresh,setEditingRefresh]=useState(false);
  const [refreshInput,setRefreshInput]=useState(orgSettings?.price_refresh_days!=null?String(orgSettings.price_refresh_days):"");
  const [savingRefresh,setSavingRefresh]=useState(false);
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
    const memberRows = mr.data||[];
    let profileMap = {};
    if(memberRows.length){
      const ids = memberRows.map(m=>m.user_id);
      const {data:profs} = await supabase.from("profiles").select("id,email").in("id",ids);
      (profs||[]).forEach(p=>{ profileMap[p.id]=p.email; });
    }
    setMembers(memberRows.map(m=>({...m,email:profileMap[m.user_id]||null})));
    setCodes(cr.data||[]);
    setLoading(false);
  }

  useEffect(()=>{ load(); },[orgId]);

  async function saveOrgName(){
    if(!orgNameInput.trim()) return;
    setSavingOrgName(true);
    await supabase.from("organizations").update({name:orgNameInput.trim()}).eq("id",orgId);
    setSavingOrgName(false);
    setEditingOrgName(false);
    onOrgUpdated();
  }

  async function saveIndustry(){
    setSavingIndustry(true);
    setIndustryMsg("");
    await supabase.from("organizations").update({industry:industryInput.trim()||null}).eq("id",orgId);
    if(autoLoadCategories&&industryInput.trim()){
      const {added,found}=await loadStarterCategoriesForIndustry(orgId,industryInput.trim(),categories);
      setIndustryMsg(!found?`Saved. No starter template found for "${industryInput.trim()}" yet — add categories manually in Catalog Categories below.`:
        added?`Saved — added ${added} starter categor${added===1?"y":"ies"} for ${industryInput.trim()}.`:
        "Saved. Starter categories for this industry were already all present.");
    }
    setSavingIndustry(false);
    setEditingIndustry(false);
    onOrgUpdated();
  }

  // Every org sets its OWN refresh cadence — a deli might want prices
  // treated as stale after a week, a building supply distributor with
  // monthly price sheets might want 30+ days. Nothing here assumes any
  // particular schedule; it's just a number stored per org. Merging into
  // existing settings so other keys (if any get added later) aren't wiped.
  async function saveRefreshDays(){
    setSavingRefresh(true);
    const days=refreshInput.trim()===""?null:Math.max(1,parseInt(refreshInput,10)||0)||null;
    await supabase.from("organizations").update({settings:{...(orgSettings||{}),price_refresh_days:days}}).eq("id",orgId);
    setSavingRefresh(false);
    setEditingRefresh(false);
    onOrgUpdated();
  }

  async function createInvite() {
    setError("");
    const code=generateInviteCode();
    const {error:e}=await supabase.from("invite_codes").insert({organization_id:orgId,code,role:inviteRole,created_by:currentUserId});
    if(e){ setError(e.message); return; }
    setNewCode(code);
    load();
  }

  async function revokeCode(c){
    if(!window.confirm("Revoke this invite code? It can no longer be used to join.")) return;
    await supabase.from("invite_codes").delete().eq("id",c.id);
    load();
  }

  async function changeRole(m,newRole){
    if(m.role==="owner"&&newRole!=="owner"&&roleCounts.owner<=1){
      alert("You can't change the only owner's role — make someone else an owner first.");
      return;
    }
    await supabase.from("organization_members").update({role:newRole}).eq("organization_id",orgId).eq("user_id",m.user_id);
    load();
  }

  async function removeMember(m){
    if(m.user_id===currentUserId){ alert("You can't remove yourself from the team."); return; }
    if(m.role==="owner"&&roleCounts.owner<=1){ alert("You can't remove the only owner."); return; }
    if(!window.confirm("Remove this person from your team? They'll lose access immediately.")) return;
    await supabase.from("organization_members").delete().eq("organization_id",orgId).eq("user_id",m.user_id);
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

  if(loading) return <p style={{color:"rgba(255,255,255,0.7)"}}>Loading team...</p>;

  return (
    <div>
      {myRole==="owner"&&(
        <div style={{background:"white",borderRadius:10,padding:16,marginBottom:16,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <div style={{fontSize:11,fontWeight:800,color:"#999",letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:12}}>Organization</div>

          <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:16}}>
            {logoUrl?(
              <img src={logoUrl} alt={orgName} style={{height:56,maxWidth:120,objectFit:"contain",borderRadius:6}} />
            ):(
              <div style={{fontSize:36,width:56,height:56,display:"flex",alignItems:"center",justifyContent:"center",background:"#F0F2F5",borderRadius:8}}>🦉</div>
            )}
            <label style={{cursor:"pointer",fontSize:12,fontWeight:700,color:"#003584"}}>
              {logoUploading?"Uploading...":logoUrl?"Change logo":"+ Add logo"}
              <input type="file" accept="image/*" style={{display:"none"}} disabled={logoUploading}
                onChange={e=>onLogoUpload(e.target.files[0])} />
            </label>
          </div>

          <div style={{fontSize:10,fontWeight:700,color:"#AAA",letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:4}}>Name</div>
          {!editingOrgName?(
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
              <div style={{fontWeight:700,fontSize:15}}>{orgName}</div>
              <button onClick={()=>{setOrgNameInput(orgName);setEditingOrgName(true);}} style={{background:"none",border:"none",cursor:"pointer",color:"#888",fontSize:12,padding:0}}>✎ Rename</button>
            </div>
          ):(
            <div style={{display:"flex",gap:8,marginBottom:14}}>
              <input style={{...inp,flex:1}} value={orgNameInput} onChange={e=>setOrgNameInput(e.target.value)} />
              <button onClick={()=>setEditingOrgName(false)} style={{...btn("#EEE","#555",{padding:"10px 14px"})}}>Cancel</button>
              <button onClick={saveOrgName} disabled={savingOrgName} style={{...btn("#003584",undefined,{padding:"10px 14px"})}}>{savingOrgName?"...":"Save"}</button>
            </div>
          )}

          <div style={{fontSize:10,fontWeight:700,color:"#AAA",letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:4}}>Industry</div>
          {!editingIndustry?(
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div style={{fontWeight:700,fontSize:15,color:orgIndustry?"#111":"#BBB"}}>{orgIndustry||"Not set"}</div>
              <button onClick={()=>{setIndustryInput(orgIndustry||"");setEditingIndustry(true);}} style={{background:"none",border:"none",cursor:"pointer",color:"#888",fontSize:12,padding:0}}>✎ Edit</button>
            </div>
          ):(
            <div>
              <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:8}}>
                {INDUSTRY_PRESETS.map(p=>(
                  <button key={p} onClick={()=>setIndustryInput(p)}
                    style={{fontSize:11,fontWeight:700,padding:"5px 11px",borderRadius:14,cursor:"pointer",
                      background:industryInput===p?"#003584":"#EEF2F8",color:industryInput===p?"white":"#003584",
                      border:industryInput===p?"2px solid #003584":"2px solid transparent"}}>
                    {p}
                  </button>
                ))}
              </div>
              <label style={{display:"flex",alignItems:"center",gap:7,fontSize:12,color:"#555",margin:"10px 0"}}>
                <input type="checkbox" checked={autoLoadCategories} onChange={e=>setAutoLoadCategories(e.target.checked)} />
                Also load starter categories for this industry when I save
              </label>
              <div style={{display:"flex",gap:8}}>
                <input style={{...inp,flex:1}} value={industryInput} onChange={e=>setIndustryInput(e.target.value)} placeholder="Or type your own..." />
                <button onClick={()=>setEditingIndustry(false)} style={{...btn("#EEE","#555",{padding:"10px 14px"})}}>Cancel</button>
                <button onClick={saveIndustry} disabled={savingIndustry} style={{...btn("#003584",undefined,{padding:"10px 14px"})}}>{savingIndustry?"...":"Save"}</button>
              </div>
              <div style={{fontSize:11,color:"#999",marginTop:6}}>
                {autoLoadCategories?"Saving will pull in a starter category set for whatever you pick above - fully editable after (rename, add, delete, edit keywords) in Catalog Categories below.":
                  "Auto-load is off - you can still load a starter set manually from Catalog Categories below any time, or build categories by hand."}
              </div>
            </div>
          )}
          {industryMsg&&<div style={{fontSize:12,color:"#2E7D32",marginTop:6}}>{industryMsg}</div>}

          <div style={{fontSize:10,fontWeight:700,color:"#AAA",letterSpacing:"0.06em",textTransform:"uppercase",margin:"14px 0 4px"}}>Price refresh period</div>
          {!editingRefresh?(
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div style={{fontWeight:700,fontSize:15,color:orgSettings?.price_refresh_days?"#111":"#BBB"}}>
                {orgSettings?.price_refresh_days?`${orgSettings.price_refresh_days} day${orgSettings.price_refresh_days===1?"":"s"}`:"Off — prices never expire"}
              </div>
              <button onClick={()=>{setRefreshInput(orgSettings?.price_refresh_days!=null?String(orgSettings.price_refresh_days):"");setEditingRefresh(true);}} style={{background:"none",border:"none",cursor:"pointer",color:"#888",fontSize:12,padding:0}}>✎ Edit</button>
            </div>
          ):(
            <div>
              <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:8}}>
                <button onClick={()=>setRefreshInput("")}
                  style={{fontSize:11,fontWeight:700,padding:"5px 11px",borderRadius:14,cursor:"pointer",
                    background:refreshInput===""?"#003584":"#EEF2F8",color:refreshInput===""?"white":"#003584",
                    border:refreshInput===""?"2px solid #003584":"2px solid transparent"}}>
                  Never expire
                </button>
                {REFRESH_PRESETS.map(([days,label])=>(
                  <button key={days} onClick={()=>setRefreshInput(String(days))}
                    style={{fontSize:11,fontWeight:700,padding:"5px 11px",borderRadius:14,cursor:"pointer",
                      background:refreshInput===String(days)?"#003584":"#EEF2F8",color:refreshInput===String(days)?"white":"#003584",
                      border:refreshInput===String(days)?"2px solid #003584":"2px solid transparent"}}>
                    {label}
                  </button>
                ))}
              </div>
              <div style={{display:"flex",gap:8}}>
                <input type="number" min="1" style={{...inp,flex:1}} value={refreshInput} onChange={e=>setRefreshInput(e.target.value)} placeholder="Or enter a custom number of days..." />
                <button onClick={()=>setEditingRefresh(false)} style={{...btn("#EEE","#555",{padding:"10px 14px"})}}>Cancel</button>
                <button onClick={saveRefreshDays} disabled={savingRefresh} style={{...btn("#003584",undefined,{padding:"10px 14px"})}}>{savingRefresh?"...":"Save"}</button>
              </div>
              <div style={{fontSize:11,color:"#999",marginTop:6}}>
                If a vendor's price hasn't been refreshed by a new import within this many days, Order Guide shows it as needing a refresh instead of the last-known price — historical prices in Invoices are never affected.
              </div>
            </div>
          )}
        </div>
      )}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
        <h3 style={{margin:0,fontSize:16,color:"white"}}>Team</h3>
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

      <div style={{fontSize:11,fontWeight:800,color:"rgba(255,255,255,0.7)",letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:8}}>Members</div>
      {members.map(m=>(
        <div key={m.user_id} style={{background:"white",borderRadius:8,padding:"10px 14px",marginBottom:6,
          display:"flex",justifyContent:"space-between",alignItems:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <div style={{minWidth:0,flex:1}}>
            <div style={{fontWeight:700,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
              {m.email||(m.user_id===currentUserId?currentUserEmail:null)||"Pending — hasn't logged in yet"}{m.user_id===currentUserId?" (you)":""}
            </div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:8,flexShrink:0}}>
            {myRole==="owner"&&m.user_id!==currentUserId?(
              <select value={m.role} onChange={e=>changeRole(m,e.target.value)} style={{fontSize:11,fontWeight:700,padding:"3px 6px",borderRadius:8,border:"1px solid #DDD"}}>
                <option value="owner">Owner</option>
                <option value="manager">Manager</option>
                <option value="employee">Employee</option>
              </select>
            ):roleBadge(m.role)}
            {myRole==="owner"&&m.user_id!==currentUserId&&(
              <button onClick={()=>removeMember(m)} style={{background:"none",border:"none",cursor:"pointer",color:"#E65100",fontSize:16,padding:0}} title="Remove">×</button>
            )}
          </div>
        </div>
      ))}

      <div style={{fontSize:11,fontWeight:800,color:"rgba(255,255,255,0.7)",letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:8,marginTop:16}}>Invite codes</div>
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
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            {roleBadge(c.role)}
            {!c.used_by&&<button onClick={()=>revokeCode(c)} style={{background:"none",border:"none",cursor:"pointer",color:"#E65100",fontSize:11,fontWeight:700,padding:0}}>Revoke</button>}
          </div>
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

// Lets an org manage its own catalog categories directly — see, add,
// rename keyword lists, delete unused categories, or pull a starter set
// based on the org's industry. Nothing here is specific to any one
// client: the same panel runs for a deli org or a building-supply org,
// the only difference is which rows exist in THEIR catalog_categories.

// Confirms a fuzzy-matched mapping is correct - bumps it to a full,
// exact match so it stops showing up for review.
async function confirmMapping(mappingId) {
  return supabase.from("item_mappings").update({comparison_track:"exact",confidence_score:100}).eq("id",mappingId);
}

// Re-points a mapping at a different, existing catalog item - for when
// the auto-match picked the wrong one.
async function remapToExistingItem(mappingId, newCatalogItemId) {
  return supabase.from("item_mappings").update({catalog_item_id:newCatalogItemId,comparison_track:"exact",confidence_score:100}).eq("id",mappingId);
}

// Splits a mapping out into its own brand-new catalog item - for when
// the auto-match merged it into something that isn't actually the same
// product at all.
async function remapToNewItem(orgId, mappingId, description, catalogItems, categories) {
  const category=classifyCategory(description,categories)||await ensureUncategorized(orgId,categories);
  const itemsInCategory=catalogItems.filter(ci=>ci.category_id===category?.id);
  const nextNumber=itemsInCategory.length
    ? Math.max(...itemsInCategory.map(ci=>ci.master_item_number||0))+1
    : (category?.range_start||1);
  const {data:created,error}=await supabase.from("catalog_items").insert({
    organization_id:orgId, category_id:category?.id||null, master_item_number:nextNumber,
    name:description.slice(0,120), matching_behavior:"flexible", canonical_unit:null, brand_locked:false,
  }).select().single();
  if(error||!created) return {error};
  return supabase.from("item_mappings").update({catalog_item_id:created.id,comparison_track:"exact",confidence_score:100}).eq("id",mappingId);
}

// Merges two catalog items that a person has manually identified as the
// same product (via drag-and-drop in the Item Catalog) - every vendor
// mapping pointing at the dragged (source) item gets re-pointed at the
// drop target, marked exact/100% since a human just confirmed it
// directly, and the now-empty source item is removed. The target item's
// identity (name, master number, category) survives; the source does not.
async function mergeCatalogItems(sourceCatalogItemId, targetCatalogItemId) {
  if(sourceCatalogItemId===targetCatalogItemId) return {error:null};
  const {data:sourceMappings,error:selErr}=await supabase.from("item_mappings").select("id").eq("catalog_item_id",sourceCatalogItemId);
  if(selErr) return {error:selErr};
  for(const m of (sourceMappings||[])){
    const {error}=await supabase.from("item_mappings").update({catalog_item_id:targetCatalogItemId,comparison_track:"exact",confidence_score:100}).eq("id",m.id);
    if(error) return {error};
  }
  return supabase.from("catalog_items").delete().eq("id",sourceCatalogItemId);
}

// The catalog item's name is the CLIENT's identity for that product, not
// whichever vendor's raw wording happened to trigger its creation first -
// this is what lets the client actually own that naming instead of being
// stuck with it.
async function renameCatalogItem(catalogItemId, newName) {
  return supabase.from("catalog_items").update({name:newName.slice(0,120)}).eq("id",catalogItemId);
}

// Lets the client create their own catalog item directly, ahead of any
// vendor data - future vendor imports match INTO it the same way they'd
// match into any other existing item, via the normal matching engine.
// This is the client-initiated counterpart to a vendor import
// auto-creating one reactively.
async function createClientCatalogItem(orgId, name, categoryId, catalogItems, categories) {
  const category = categories.find(c => c.id === categoryId) || null;
  const itemsInCategory = catalogItems.filter(ci => ci.category_id === categoryId);
  const nextNumber = itemsInCategory.length
    ? Math.max(...itemsInCategory.map(ci => ci.master_item_number || 0)) + 1
    : (category?.range_start || 1);
  return supabase.from("catalog_items").insert({
    organization_id: orgId, category_id: categoryId || null, master_item_number: nextNumber,
    name: name.slice(0, 120), matching_behavior: "flexible", canonical_unit: null, brand_locked: false,
  }).select().single();
}

// Directly links one specific vendor's item to a specific client item
// number - not limited to items the matching engine happened to flag.
// Works whether this vendor item has never been mapped at all, or
// already has a mapping that needs correcting. Marked exact/100% since
// a person just typed the exact number themselves - the most certain
// kind of match there is.
async function assignVendorItemMapping(orgId, vendorItemId, targetCatalogItemId, existingMappingId) {
  if (existingMappingId) {
    return supabase.from("item_mappings").update({
      catalog_item_id: targetCatalogItemId, comparison_track: "exact", confidence_score: 100, match_method: "manual",
    }).eq("id", existingMappingId);
  }
  return supabase.from("item_mappings").insert({
    organization_id: orgId, catalog_item_id: targetCatalogItemId, vendor_item_id: vendorItemId,
    comparison_track: "exact", confidence_score: 100, match_method: "manual",
  });
}

// The actual browsable master item list, AND where review happens - not
// a separate hidden admin section. Search, filter by category, see
// every item's status and every vendor's price side by side, with the
// same export this data already has elsewhere. A single vendor
// introducing a product for the first time has nothing to be uncertain
// about (there's no second vendor's wording to conflict with), so it's
// shown as a clean 100% match, not flagged as "unconfirmed" - only a
// genuine fuzzy merge between two different vendors' wording gets
// flagged, with the real percentage and a way to confirm or correct it.
const REVIEW_FILTER="__needs_review__";

// Three ways to order items within a category (or a full/unfiltered list):
// alphabetical by name, this org's own client item-number sequence, or the
// vendor's own item code (taken from that item's cheapest/first-listed
// vendor option, since one client item can carry several vendors' different
// codes - there's no single "the" vendor code). Items missing whatever key
// the current mode needs sink to the end rather than disappearing. Shared
// by Item Catalog (mapping work) and Order Guide (placing orders) - same
// browsing idea, different job each screen is doing with the result.
function compareItems(a,b,mode){
  if(mode==="itemNumber"){
    const na=a.masterItemNumber,nb=b.masterItemNumber;
    if(na==null&&nb==null) return a.name.localeCompare(b.name);
    if(na==null) return 1; if(nb==null) return -1;
    return na-nb;
  }
  if(mode==="vendorCode"){
    const ca=a.options[0]?.vendorItemCode,cb=b.options[0]?.vendorItemCode;
    if(!ca&&!cb) return a.name.localeCompare(b.name);
    if(!ca) return 1; if(!cb) return -1;
    return String(ca).localeCompare(String(cb),undefined,{numeric:true});
  }
  if(mode==="added"){
    // Chronological = the order items actually entered the catalog, oldest
    // first - not alphabetical, not the numbering scheme. Falls back to
    // name order for anything missing a timestamp rather than dropping it.
    const ta=a.createdAt?new Date(a.createdAt).getTime():null,tb=b.createdAt?new Date(b.createdAt).getTime():null;
    if(ta==null&&tb==null) return a.name.localeCompare(b.name);
    if(ta==null) return 1; if(tb==null) return -1;
    return ta-tb;
  }
  return a.name.localeCompare(b.name);
}

function daysAgo(iso){
  if(!iso) return "unknown";
  const days=Math.floor((Date.now()-new Date(iso).getTime())/(1000*60*60*24));
  return days<=0?"today":days===1?"1 day ago":`${days} days ago`;
}

// Small labeled block used by the review sections on Item Catalog, Price
// Sheets, and Invoices - each tab only ever shows the review data that's
// actually ITS OWN (mapping issues on Item Catalog, price-sheet health on
// Price Sheets, invoice-line issues on Invoices), so this is shared
// purely for the consistent look, not because any data crosses tabs.
function Section({title,count,emptyText,children}){
  return (
    <div style={{marginBottom:20}}>
      <div style={{fontWeight:800,fontSize:13,color:"#003584",marginBottom:8}}>{title} {count>0&&<span style={{color:"#E65100"}}>({count})</span>}</div>
      {count===0?(
        <div style={{fontSize:12,color:"#AAA",background:"white",borderRadius:8,padding:"10px 12px"}}>{emptyText}</div>
      ):children}
    </div>
  );
}

function ItemCatalogPanel({orgId,productList,vendors,catalogItems,mappings,vendorItems,categories,onOpenVendor,onUpdated}) {
  const [search,setSearch]=useState("");
  const [categoryFilter,setCategoryFilter]=useState("");
  const [remapOpenFor,setRemapOpenFor]=useState(null);
  const [remapSearch,setRemapSearch]=useState("");
  const [busyMappingId,setBusyMappingId]=useState(null);
  const [draggedItemId,setDraggedItemId]=useState(null);
  const [dragOverItemId,setDragOverItemId]=useState(null);
  const [merging,setMerging]=useState(false);
  const [renamingId,setRenamingId]=useState(null);
  const [renameValue,setRenameValue]=useState("");
  const [categoryEditId,setCategoryEditId]=useState(null);
  const [categoryEditBusy,setCategoryEditBusy]=useState(false);
  // Which catalog item's "map this item's vendors" panel is open. This
  // is separate from remapOpenFor (which mapping's remap-search box is
  // open, below) - a person opens the item's panel first, then may open
  // remap-search on one specific vendor line inside it. Available for
  // EVERY item, not just ones flagged for review - re-pointing a vendor's
  // price to a different catalog item shouldn't require waiting for the
  // system to flag it first.
  const [mapPanelOpenFor,setMapPanelOpenFor]=useState(null);
  // Bulk allocation of unclassified items. One-at-a-time reassignment is
  // fine for a stray item, but an import that leaves dozens unmatched
  // needs to be workable in one pass, not dozens of separate dropdowns.
  const [selectedIds,setSelectedIds]=useState(new Set());
  const [bulkBusy,setBulkBusy]=useState(false);
  // After a manual allocation, the words that would have matched are
  // offered back to the destination category so the SAME correction is
  // never needed twice. This is the dictionary learning from real use
  // rather than waiting on someone to hand-edit keywords.
  const [teach,setTeach]=useState(null); // {categoryId, categoryName, words:[], chosen:Set}
  const [teachBusy,setTeachBusy]=useState(false);
  const [addingItem,setAddingItem]=useState(false);
  const [newItemName,setNewItemName]=useState("");
  const [newItemCategoryId,setNewItemCategoryId]=useState("");
  const [addingBusy,setAddingBusy]=useState(false);
  // How items are ordered inside a category (or the full list): by when
  // they were actually added to the catalog (the default - "chronological"),
  // pure alphabetical, this org's own client item number sequence, or the
  // vendor's own item code on the cheapest/first-listed vendor for that
  // item. Independent of which category chip (or Full List) is selected -
  // all four sort modes work in any filter combination.
  const [sortMode,setSortMode]=useState("added"); // "added" | "alpha" | "itemNumber" | "vendorCode"

  const vMap=useMemo(()=>new Map(vendors.map(v=>[v.id,v])),[vendors]);
  const viMap=useMemo(()=>new Map(vendorItems.map(vi=>[vi.id,vi])),[vendorItems]);
  const ciMap=useMemo(()=>new Map(catalogItems.map(ci=>[ci.id,ci])),[catalogItems]);

  // Categories are ordered alphabetically (Dairy, Meat, Paper Goods,
  // Produce, Uncategorized...) - same as Order Guide, so both screens'
  // category chips read the same way. Item numbering itself still uses
  // the 10000/20000/30000 blocks; that's a numbering scheme, not a
  // display order.
  const categoryList=useMemo(()=>{
    const set=new Set(productList.map(p=>p.category));
    return [...set].sort((a,b)=>a.localeCompare(b));
  },[productList]);

  // A dedicated hot button for "show me only what needs review" -
  // reuses the categoryFilter slot with a sentinel value rather than a
  // second piece of state, so category chips and this button stay
  // mutually exclusive the same simple way.
  const needsReviewItems=useMemo(()=>
    new Set(productList.filter(item=>item.options.some(o=>o.matchTrack==="similar")).map(item=>item.catalogItemId)),
  [productList]);

  // Grouped by category (Produce, Meat, Dairy, Paper Goods, Uncategorized,
  // whatever this org's own categories are) - these headers ARE the
  // browsing structure, alphabetical top to bottom. Ordered within each
  // by whichever sortMode is active (chronological by default).
  const groupedItems=useMemo(()=>{
    const q=search.trim().toLowerCase();
    const items=productList.filter(item=>{
      if(categoryFilter===REVIEW_FILTER){ if(!needsReviewItems.has(item.catalogItemId)) return false; }
      else if(categoryFilter&&item.category!==categoryFilter) return false;
      if(q&&!item.name.toLowerCase().includes(q)) return false;
      return true;
    });
    const byCategory=new Map();
    for(const item of items){
      if(!byCategory.has(item.category)) byCategory.set(item.category,[]);
      byCategory.get(item.category).push(item);
    }
    return [...byCategory.entries()]
      .map(([category,catItems])=>({category,items:catItems.sort((a,b)=>compareItems(a,b,sortMode))}))
      .sort((a,b)=>a.category.localeCompare(b.category));
  },[productList,search,categoryFilter,needsReviewItems,sortMode]);

  // A single-vendor "new" item has nothing to compare against, so it's
  // a clean 100% match by definition - only "similar" is a real fuzzy
  // merge worth reviewing.
  function statusFor(item){
    const usable=item.options.filter(o=>!o.expired);
    const cheapest=usable[0]||item.options[0]||null;
    if(!cheapest) return {label:"No price on file",color:"#999",bg:"#F5F5F5"};
    if(cheapest.expired) return {label:"Stale — needs refresh",color:"#B26A00",bg:"#FFF3E0"};
    if(cheapest.matchTrack==="similar") return {label:`Needs review — ${cheapest.matchConfidence}%`,color:"#B26A00",bg:"#FFF3E0"};
    return {label:"✓ 100% matched",color:"#2E7D32",bg:"#E8F5E9"};
  }

  async function handleConfirm(mappingId){
    setBusyMappingId(mappingId);
    await confirmMapping(mappingId);
    setBusyMappingId(null);
    onUpdated();
  }
  async function handleRemapExisting(mappingId,newCatalogItemId){
    setBusyMappingId(mappingId);
    await remapToExistingItem(mappingId,newCatalogItemId);
    setBusyMappingId(null);
    setRemapOpenFor(null); setRemapSearch("");
    onUpdated();
  }
  async function handleRemapNew(mappingId,description){
    setBusyMappingId(mappingId);
    await remapToNewItem(orgId,mappingId,description,catalogItems,categories);
    setBusyMappingId(null);
    setRemapOpenFor(null); setRemapSearch("");
    onUpdated();
  }

  function handleDrop(targetItem){
    const sourceId=draggedItemId;
    setDraggedItemId(null); setDragOverItemId(null);
    if(!sourceId||sourceId===targetItem.catalogItemId) return;
    const sourceItem=productList.find(p=>p.catalogItemId===sourceId);
    if(!sourceItem) return;
    if(!window.confirm(`Merge "${sourceItem.name}" into "${targetItem.name}"?\n\nAll of "${sourceItem.name}"'s vendor prices will move under "${targetItem.name}", and "${sourceItem.name}" will be removed as its own item. This can't be undone automatically.`)) return;
    setMerging(true);
    mergeCatalogItems(sourceId,targetItem.catalogItemId).then(()=>{
      setMerging(false);
      onUpdated();
    });
  }

  async function saveRename(catalogItemId){
    const trimmed=renameValue.trim();
    setRenamingId(null);
    if(!trimmed) return;
    await renameCatalogItem(catalogItemId,trimmed);
    onUpdated();
  }

  // Words shared by the just-allocated items that the destination
  // category does not already match. These are exactly the words whose
  // absence caused the miss, so offering them back closes the gap at its
  // source. Deterministic word frequency - no model, no guessing.
  function suggestKeywords(items,targetCategory){
    const existing=(targetCategory?.keywords||[]).map(k=>String(k).toLowerCase());
    const freq=new Map();
    for(const it of items){
      const seen=new Set();
      for(const w of String(it.name||"").toLowerCase().split(/[^a-z0-9]+/)){
        if(w.length<3||seen.has(w)) continue;
        seen.add(w);
        freq.set(w,(freq.get(w)||0)+1);
      }
    }
    // Drop anything the category already covers, and anything that is
    // just a number or a size token.
    return [...freq.entries()]
      .filter(([w])=>!existing.some(k=>k===w||k.includes(w)||w.includes(k)))
      .filter(([w])=>!/^\d/.test(w))
      .sort((a,b)=>b[1]-a[1])
      .slice(0,8)
      .map(([w,n])=>({word:w,count:n}));
  }

  async function handleBulkAssign(newCategoryId){
    if(!newCategoryId||selectedIds.size===0) return;
    setBulkBusy(true);
    setError("");
    const target=categories.find(c=>c.id===newCategoryId);
    const chosenItems=productList.filter(p=>selectedIds.has(p.catalogItemId));
    try{
      // Numbering must stay collision-free, so assignItemCategory is
      // called in sequence against a working copy - same rule the
      // reclassify pass uses.
      const working=[...catalogItems];
      for(const item of chosenItems){
        const n=await assignItemCategory(item.catalogItemId,newCategoryId,working,categories);
        const i=working.findIndex(ci=>ci.id===item.catalogItemId);
        if(i>=0&&n!=null) working[i]={...working[i],category_id:newCategoryId,master_item_number:n};
      }
      const words=suggestKeywords(chosenItems,target);
      setSelectedIds(new Set());
      if(words.length) setTeach({categoryId:newCategoryId,categoryName:target?.name||"",words,chosen:new Set()});
      onUpdated();
    }catch(err){
      setError(`Could not move those items: ${err.message||String(err)}`);
    }finally{
      setBulkBusy(false);
    }
  }

  async function saveTeachedKeywords(){
    if(!teach||teach.chosen.size===0){ setTeach(null); return; }
    setTeachBusy(true);
    try{
      const cat=categories.find(c=>c.id===teach.categoryId);
      const merged=[...(cat?.keywords||[]),...[...teach.chosen]];
      const {error:e}=await supabase.from("catalog_categories")
        .update({keywords:merged}).eq("id",teach.categoryId);
      if(e) setError(`Could not save keywords: ${e.message}`);
      setTeach(null);
      onUpdated();
    }finally{
      setTeachBusy(false);
    }
  }

  async function handleAssignCategory(catalogItemId,newCategoryId){
    setCategoryEditId(null);
    if(!newCategoryId) return;
    setCategoryEditBusy(true);
    await assignItemCategory(catalogItemId,newCategoryId,catalogItems,categories);
    setCategoryEditBusy(false);
    onUpdated();
  }

  async function handleAddItem(){
    if(!newItemName.trim()) return;
    setAddingBusy(true);
    const {error}=await createClientCatalogItem(orgId,newItemName.trim(),newItemCategoryId||null,catalogItems,categories);
    setAddingBusy(false);
    if(error){ alert("Couldn't create item: "+error.message); return; }
    setNewItemName(""); setNewItemCategoryId(""); setAddingItem(false);
    onUpdated();
  }

  const lowConfidenceMatches=useMemo(()=>
    mappings.filter(m=>m.comparison_track==="similar").map(m=>{
      const vi=viMap.get(m.vendor_item_id); const ci=ciMap.get(m.catalog_item_id);
      const v=vi?vMap.get(vi.vendor_id):null;
      return {mappingId:m.id, catalogItemId:m.catalog_item_id, catalogName:ci?.name||"(deleted item)", vendorName:v?.name||"—",
        vendorDescription:vi?.description||"—", confidence:m.confidence_score};
    }).sort((a,b)=>(a.confidence??0)-(b.confidence??0)),
  [mappings,viMap,ciMap,vMap]);

  // Item Catalog's review badge is scoped to CATALOG MAPPING issues only
  // (fuzzy vendor-item matches) - invoice-line issues live on Invoices,
  // price-sheet health (unavailable/stale prices) lives on Price Sheets.
  // Each tab shows only what's actually its own job to review.
  const totalCount=lowConfidenceMatches.length;

  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
        <h3 style={{margin:0,fontSize:16,color:"white"}}>Item Catalog</h3>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          {totalCount>0&&<span style={{background:"#E65100",color:"white",fontSize:12,fontWeight:700,padding:"3px 10px",borderRadius:12}}>{totalCount} to review</span>}
          <button onClick={()=>setAddingItem(true)} style={{...btn("#003584","white",{fontSize:12,padding:"8px 14px"})}}>+ Add Item</button>
          <button onClick={()=>downloadTextFile(`catalog-export-${new Date().toISOString().split("T")[0]}.csv`,buildCatalogExportCSV(productList,vendors),"text/csv")}
            disabled={!productList.length} style={{...btn("#2E7D32","white",{fontSize:12,padding:"8px 14px"})}}>
            📄 Export (CSV)
          </button>
        </div>
      </div>

      {addingItem&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div style={{background:"white",borderRadius:12,padding:24,width:"100%",maxWidth:380}}>
            <h3 style={{margin:"0 0 4px",fontSize:16}}>Add a catalog item</h3>
            <p style={{margin:"0 0 14px",fontSize:12,color:"#888"}}>This is your item — name it however makes sense to you. Vendor prices get matched into it as they come in.</p>
            <input style={{...inp,width:"100%",marginBottom:10,boxSizing:"border-box"}} placeholder="Item name" value={newItemName} onChange={e=>setNewItemName(e.target.value)} autoFocus />
            <select style={{...inp,width:"100%",marginBottom:14,boxSizing:"border-box"}} value={newItemCategoryId} onChange={e=>setNewItemCategoryId(e.target.value)}>
              <option value="">No category (Uncategorized)</option>
              {categories.filter(c=>c.name!=="Uncategorized").map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
            <div style={{display:"flex",gap:8}}>
              <button onClick={()=>{setAddingItem(false);setNewItemName("");setNewItemCategoryId("");}} style={{...btn("#EEE","#555"),flex:1}}>Cancel</button>
              <button onClick={handleAddItem} disabled={addingBusy||!newItemName.trim()} style={{...btn("#003584"),flex:2}}>{addingBusy?"Adding...":"Add Item"}</button>
            </div>
          </div>
        </div>
      )}

      {totalCount>0&&(
        <div style={{marginBottom:24,paddingBottom:4}}>
          <Section title="Fuzzy-matched items — confirm or remap" count={lowConfidenceMatches.length} emptyText="Nothing flagged — every mapping is an exact or single-vendor match.">
            {lowConfidenceMatches.map(m=>(
              <div key={m.mappingId} style={{background:"white",borderRadius:8,padding:"10px 12px",marginBottom:6,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
                <div style={{display:"flex",justifyContent:"space-between"}}>
                  <div style={{fontWeight:600,fontSize:13}}>{m.catalogName}</div>
                  <div style={{fontSize:11,fontWeight:700,color:"#B26A00"}}>🔍 {m.confidence}% match</div>
                </div>
                <div style={{fontSize:11,color:"#999",marginTop:2,marginBottom:8}}>{m.vendorName} — "{m.vendorDescription}"</div>
                {remapOpenFor===m.mappingId?(
                  <div style={{background:"#F7F9FC",borderRadius:6,padding:8}}>
                    <input style={{...inp,marginBottom:6,fontSize:12,padding:"7px 9px"}} placeholder="Search by name, or enter item #..." value={remapSearch} onChange={e=>setRemapSearch(e.target.value)} autoFocus />
                    <div style={{maxHeight:140,overflowY:"auto"}}>
                      {catalogItems.filter(ci=>{
                        if(ci.id===m.catalogItemId) return false;
                        const q=remapSearch.trim().toLowerCase();
                        if(!q) return true;
                        return ci.name.toLowerCase().includes(q) || String(ci.master_item_number)===remapSearch.trim();
                      }).slice(0,8).map(ci=>(
                        <button key={ci.id} disabled={busyMappingId===m.mappingId} onClick={()=>handleRemapExisting(m.mappingId,ci.id)}
                          style={{display:"block",width:"100%",textAlign:"left",background:"white",border:"1px solid #EEE",borderRadius:5,padding:"6px 8px",marginBottom:4,fontSize:12,cursor:"pointer"}}>
                          <span style={{color:"#AAA",fontFamily:"monospace"}}>#{ci.master_item_number}</span> {ci.name}
                        </button>
                      ))}
                    </div>
                    <div style={{display:"flex",gap:6,marginTop:6}}>
                      <button disabled={busyMappingId===m.mappingId} onClick={()=>handleRemapNew(m.mappingId,m.vendorDescription)}
                        style={{...btn("#EEE","#555",{fontSize:11,padding:"6px 10px",flex:1})}}>None of these — make separate item</button>
                      <button onClick={()=>{setRemapOpenFor(null);setRemapSearch("");}} style={{...btn("#EEE","#555",{fontSize:11,padding:"6px 10px"})}}>Cancel</button>
                    </div>
                  </div>
                ):(
                  <div style={{display:"flex",gap:6}}>
                    <button disabled={busyMappingId===m.mappingId} onClick={()=>handleConfirm(m.mappingId)}
                      style={{...btn("#2E7D32","white",{fontSize:11,padding:"6px 12px"})}}>{busyMappingId===m.mappingId?"...":"✓ Confirm match"}</button>
                    <button disabled={busyMappingId===m.mappingId} onClick={()=>setRemapOpenFor(m.mappingId)}
                      style={{...btn("#EEE","#555",{fontSize:11,padding:"6px 12px"})}}>✎ Remap</button>
                  </div>
                )}
              </div>
            ))}
          </Section>

          <div style={{fontSize:10,fontWeight:700,color:"#AAA",letterSpacing:"0.06em",textTransform:"uppercase",margin:"18px 0 10px"}}>Full catalog</div>
        </div>
      )}

      <input style={{...inp,marginBottom:10}} placeholder="Search items..." value={search} onChange={e=>setSearch(e.target.value)} />

      <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",marginBottom:10}}>
        <span style={{fontSize:10,fontWeight:700,color:"rgba(255,255,255,0.6)",textTransform:"uppercase",letterSpacing:"0.05em"}}>Sort:</span>
        {[["added","Date Added"],["alpha","A–Z"],["itemNumber","Item #"],["vendorCode","Vendor Code"]].map(([id,label])=>{
          const isSelected=sortMode===id;
          return (
            <button key={id} onClick={()=>setSortMode(id)} style={chipStyle(isSelected,"sm")}>
              {label}
            </button>
          );
        })}
      </div>

      <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:18}}>
        <button onClick={()=>setCategoryFilter("")} style={chipStyle(!categoryFilter)}>
          Full List
        </button>
        {needsReviewItems.size>0&&(
          <button onClick={()=>setCategoryFilter(categoryFilter===REVIEW_FILTER?"":REVIEW_FILTER)}
            style={{fontSize:12,fontWeight:700,padding:"6px 14px",borderRadius:20,cursor:"pointer",
              background:categoryFilter===REVIEW_FILTER?"#E65100":"#FFF3E0",color:categoryFilter===REVIEW_FILTER?"white":"#B26A00",
              border:categoryFilter===REVIEW_FILTER?"2px solid #E65100":"2px solid transparent"}}>
            ⚠ Needs Review ({needsReviewItems.size})
          </button>
        )}
        {categoryList.map(c=>{
          const isSelected=categoryFilter===c;
          return (
            <button key={c} onClick={()=>setCategoryFilter(isSelected?"":c)} style={chipStyle(isSelected)}>
              {c}
            </button>
          );
        })}
      </div>

      {teach&&(
        <div style={{background:"#E8F5E9",border:"2px solid #2E7D32",borderRadius:10,padding:14,marginBottom:12}}>
          <div style={{fontSize:13,fontWeight:700,color:"#1B5E20",marginBottom:4}}>Teach the catalog so this doesn't happen again</div>
          <div style={{fontSize:12,color:"#33691E",marginBottom:8}}>
            Add any of these words to <b>{teach.categoryName}</b> and items like these will classify automatically from now on, on every future import.
          </div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:10}}>
            {teach.words.map(({word,count})=>{
              const on=teach.chosen.has(word);
              return (
                <button key={word} onClick={()=>setTeach(t=>{const c=new Set(t.chosen);on?c.delete(word):c.add(word);return {...t,chosen:c};})}
                  style={{fontSize:12,fontWeight:700,padding:"5px 11px",borderRadius:14,cursor:"pointer",
                    background:on?"#2E7D32":"white",color:on?"white":"#33691E",
                    border:on?"2px solid #2E7D32":"2px solid #A5D6A7"}}>
                  {on?"✓ ":""}{word}{count>1?` (${count})`:""}
                </button>
              );
            })}
          </div>
          <div style={{display:"flex",gap:8}}>
            <button onClick={saveTeachedKeywords} disabled={teachBusy||teach.chosen.size===0}
              style={{...btn("#2E7D32","white",{fontSize:12,padding:"8px 14px",opacity:teach.chosen.size===0?0.5:1})}}>
              {teachBusy?"Saving...":`Add ${teach.chosen.size||""} word${teach.chosen.size===1?"":"s"} to ${teach.categoryName}`}
            </button>
            <button onClick={()=>setTeach(null)} style={{...btn("#EEE","#555",{fontSize:12,padding:"8px 14px"})}}>Not now</button>
          </div>
        </div>
      )}

      {selectedIds.size>0&&(
        <div style={{background:"white",borderRadius:10,padding:12,marginBottom:12,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",
          display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",position:"sticky",top:8,zIndex:5}}>
          <span style={{fontSize:13,fontWeight:700,color:"#003584"}}>{selectedIds.size} selected</span>
          <select defaultValue="" disabled={bulkBusy} onChange={e=>{handleBulkAssign(e.target.value);e.target.value="";}}
            style={{...inp,fontSize:12,padding:"7px 9px",width:"auto",flex:"0 1 220px"}}>
            <option value="" disabled>{bulkBusy?"Moving...":"Move all to..."}</option>
            {categories.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
          <button onClick={()=>setSelectedIds(new Set())} style={{...btn("#EEE","#555",{fontSize:12,padding:"7px 12px"})}}>Clear</button>
        </div>
      )}

      {productList.length===0?(
        <div style={{background:"white",borderRadius:10,padding:20,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <p style={{color:"#888",fontSize:13,margin:0}}>Nothing in the catalog yet — import a price list to get started.</p>
        </div>
      ):groupedItems.length===0?(
        <div style={{background:"white",borderRadius:10,padding:20,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <p style={{color:"#888",fontSize:13,margin:0}}>No items match that search.</p>
        </div>
      ):groupedItems.map(group=>(
        <div key={group.category} style={{marginBottom:20}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,paddingLeft:2}}>
            <input type="checkbox"
              checked={group.items.length>0&&group.items.every(i=>selectedIds.has(i.catalogItemId))}
              onChange={e=>{
                const all=e.target.checked;
                setSelectedIds(prev=>{
                  const next=new Set(prev);
                  group.items.forEach(i=>all?next.add(i.catalogItemId):next.delete(i.catalogItemId));
                  return next;
                });
              }}
              title={`Select every item shown under ${group.category}`} />
            <span style={{fontWeight:800,fontSize:13,color:"rgba(255,255,255,0.85)"}}>{group.category}</span>
            <span style={{fontSize:11,color:"rgba(255,255,255,0.5)"}}>({group.items.length})</span>
          </div>
          {group.items.map(item=>{
            const status=statusFor(item);
            const isDragOver=dragOverItemId===item.catalogItemId;
            return (
              <div key={item.catalogItemId}
                draggable
                onDragStart={()=>setDraggedItemId(item.catalogItemId)}
                onDragEnd={()=>{setDraggedItemId(null);setDragOverItemId(null);}}
                onDragOver={(e)=>{e.preventDefault();if(draggedItemId&&draggedItemId!==item.catalogItemId) setDragOverItemId(item.catalogItemId);}}
                onDragLeave={()=>{if(dragOverItemId===item.catalogItemId) setDragOverItemId(null);}}
                onDrop={(e)=>{e.preventDefault();handleDrop(item);}}
                title="Drag onto another item to merge them as the same product"
                style={{background:isDragOver?"#E8F5E9":"white",borderRadius:8,padding:"12px 14px",marginBottom:8,
                  boxShadow:isDragOver?"0 0 0 2px #2E7D32":"0 1px 3px rgba(0,0,0,0.06)",cursor:"grab",
                  opacity:draggedItemId===item.catalogItemId?0.4:1}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                  <input type="checkbox" checked={selectedIds.has(item.catalogItemId)}
                    onClick={e=>e.stopPropagation()}
                    onChange={e=>setSelectedIds(prev=>{
                      const next=new Set(prev);
                      e.target.checked?next.add(item.catalogItemId):next.delete(item.catalogItemId);
                      return next;
                    })}
                    style={{marginRight:10,marginTop:3}} />
                  <div style={{flex:1,minWidth:0}}>
                    {renamingId===item.catalogItemId?(
                      <div style={{display:"flex",gap:6}} onClick={e=>e.stopPropagation()}>
                        <input autoFocus value={renameValue} onChange={e=>setRenameValue(e.target.value)}
                          onKeyDown={e=>{if(e.key==="Enter") saveRename(item.catalogItemId); if(e.key==="Escape") setRenamingId(null);}}
                          style={{...inp,fontSize:13,padding:"5px 8px",flex:1}} />
                        <button onClick={()=>saveRename(item.catalogItemId)} style={{...btn("#003584",undefined,{fontSize:11,padding:"5px 10px"})}}>Save</button>
                        <button onClick={()=>setRenamingId(null)} style={{...btn("#EEE","#555",{fontSize:11,padding:"5px 10px"})}}>✕</button>
                      </div>
                    ):(
                      <div style={{fontWeight:700,fontSize:14,display:"flex",alignItems:"center",gap:6}}>
                        {item.name}
                        <button onClick={()=>{setRenamingId(item.catalogItemId);setRenameValue(item.name);}}
                          title="Rename - this is your item, name it however makes sense to you"
                          style={{background:"none",border:"none",cursor:"pointer",color:"#BBB",fontSize:12,padding:0}}>✎</button>
                      </div>
                    )}
                    {categoryEditId===item.catalogItemId?(
                      <div style={{display:"flex",gap:6,alignItems:"center",marginTop:2}} onClick={e=>e.stopPropagation()}>
                        <span style={{fontSize:11,color:"#AAA"}}>#{item.masterItemNumber} ·</span>
                        <select autoFocus defaultValue="" onChange={e=>handleAssignCategory(item.catalogItemId,e.target.value)}
                          onBlur={()=>setCategoryEditId(null)}
                          style={{...inp,fontSize:11,padding:"3px 6px",width:"auto"}}>
                          <option value="" disabled>Move to...</option>
                          {categories.filter(c=>c.name!==item.category).map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
                        </select>
                      </div>
                    ):(
                      <div style={{fontSize:11,color:"#AAA",marginTop:2,display:"flex",alignItems:"center",gap:4}}>
                        #{item.masterItemNumber} · {item.category}
                        <button onClick={()=>setCategoryEditId(item.catalogItemId)} disabled={categoryEditBusy}
                          title="Move this item to a different category"
                          style={{background:"none",border:"none",cursor:"pointer",color:"#BBB",fontSize:11,padding:0}}>✎</button>
                      </div>
                    )}
                  </div>
                  <span style={{fontSize:10,fontWeight:700,color:status.color,background:status.bg,padding:"3px 8px",borderRadius:5,whiteSpace:"nowrap"}}>{status.label}</span>
                </div>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:8}}>
                  <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
                    {item.options.map(o=>(
                      <span key={o.vendorId} style={{fontSize:11,background:"#F5F7FA",borderRadius:5,padding:"3px 8px"}}>
                        {o.vendorName}: <b>{o.expired?"stale":formatMoney(o.casePrice)}</b>
                      </span>
                    ))}
                    {!item.options.length&&<span style={{fontSize:11,color:"#CCC"}}>No vendor price linked yet</span>}
                  </div>
                  <button onClick={()=>setMapPanelOpenFor(mapPanelOpenFor===item.catalogItemId?null:item.catalogItemId)}
                    title="Map this item's vendor prices - link, unlink, or re-point any of them, any time"
                    style={{background:"none",border:"none",cursor:"pointer",color:"#003584",fontSize:11,fontWeight:700,padding:0,whiteSpace:"nowrap",marginLeft:8}}>
                    🔗 Map {mapPanelOpenFor===item.catalogItemId?"▲":"▾"}
                  </button>
                </div>

                {mapPanelOpenFor===item.catalogItemId&&(
                  <div style={{marginTop:8,background:"#F7F9FC",borderRadius:6,padding:8}} onClick={e=>e.stopPropagation()}>
                    {item.options.length===0&&<div style={{fontSize:11,color:"#AAA",marginBottom:6}}>No vendor is currently linked to this item — it'll pick one up automatically the next time a price sheet mentions it, or wait for a manual link once a vendor item exists to point at.</div>}
                    {item.options.map(o=>(
                      <div key={o.vendorItemId} style={{background:"white",borderRadius:6,padding:8,marginBottom:6,border:"1px solid #EEE"}}>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                          <div style={{fontSize:12,fontWeight:600}}>{o.vendorName}</div>
                          <div style={{fontSize:11,fontWeight:700,color:o.matchTrack==="similar"?"#B26A00":"#2E7D32"}}>
                            {o.matchTrack==="similar"?`🔍 ${o.matchConfidence}%`:"✓ exact"}
                          </div>
                        </div>
                        <div style={{fontSize:11,color:"#999",margin:"2px 0 6px"}}>"{o.description}" — {o.expired?"stale":formatMoney(o.casePrice)}</div>
                        {remapOpenFor===o.mappingId?(
                          <div>
                            <input style={{...inp,marginBottom:6,fontSize:12,padding:"7px 9px"}} placeholder="Search by name, or enter item #..." value={remapSearch} onChange={e=>setRemapSearch(e.target.value)} autoFocus />
                            <div style={{maxHeight:140,overflowY:"auto"}}>
                              {catalogItems.filter(ci=>{
                                if(ci.id===item.catalogItemId) return false;
                                const q=remapSearch.trim().toLowerCase();
                                if(!q) return true;
                                return ci.name.toLowerCase().includes(q) || String(ci.master_item_number)===remapSearch.trim();
                              }).slice(0,8).map(ci=>(
                                <button key={ci.id} disabled={busyMappingId===o.mappingId} onClick={()=>handleRemapExisting(o.mappingId,ci.id)}
                                  style={{display:"block",width:"100%",textAlign:"left",background:"white",border:"1px solid #EEE",borderRadius:5,padding:"6px 8px",marginBottom:4,fontSize:12,cursor:"pointer"}}>
                                  <span style={{color:"#AAA",fontFamily:"monospace"}}>#{ci.master_item_number}</span> {ci.name}
                                </button>
                              ))}
                            </div>
                            <div style={{display:"flex",gap:6,marginTop:6}}>
                              <button disabled={busyMappingId===o.mappingId} onClick={()=>handleRemapNew(o.mappingId,o.description)}
                                style={{...btn("#EEE","#555",{fontSize:11,padding:"6px 10px",flex:1})}}>None of these — make separate item</button>
                              <button onClick={()=>{setRemapOpenFor(null);setRemapSearch("");}} style={{...btn("#EEE","#555",{fontSize:11,padding:"6px 10px"})}}>Cancel</button>
                            </div>
                          </div>
                        ):(
                          <button disabled={busyMappingId===o.mappingId} onClick={()=>{setRemapOpenFor(o.mappingId);setRemapSearch("");}}
                            style={{...btn("#EEE","#555",{fontSize:11,padding:"6px 10px"})}}>✎ Point at a different item</button>
                        )}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      ))}
      {merging&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.3)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:"white",borderRadius:10,padding:"14px 20px",fontSize:13,fontWeight:700}}>Merging items...</div>
        </div>
      )}
    </div>
  );
}

// Classification only runs ONCE, at the moment a catalog item is first
// created (see matchOrCreateCatalogItem) - so an item imported before an
// org had real categories set up (or before a category had the right
// keyword) stays in Uncategorized forever unless something re-checks it.
// This re-runs classifyCategory against every current Uncategorized item
// using whatever categories/keywords exist NOW, and moves anything that
// now has a real match - same numbering rule as a brand-new item (next
// number in the target category's own block). Items that still don't
// match anything stay in Uncategorized and genuinely need a person to
// allocate them - unlike a real "General" category (if an org adds one),
// Uncategorized is a queue that SHOULD tend toward empty.
async function reclassifyUncategorizedItems(orgId, catalogItems, categories) {
  const uncategorized = categories.find(c => c.name === "Uncategorized");
  if (!uncategorized) return { moved: 0, checked: 0 };
  const stuck = catalogItems.filter(ci => ci.category_id === uncategorized.id);
  const otherCategories = categories.filter(c => c.id !== uncategorized.id);

  // Pass 1: keywords, for every item. Pure local computation, no network
  // calls at all.
  const keywordHits = stuck.map(item => ({ item, target: classifyCategory(item.name, otherCategories) }));

  // Pass 2: numbering. Still done one at a time (needed to avoid two
  // items landing on the same number in the same target category), but
  // this is local computation, not a network round-trip, so it costs
  // microseconds regardless of how many items there are.
  const working = [...catalogItems];
  const assignments = [];
  for (const { item, target } of keywordHits) {
    if (!target) continue;
    const itemsInCategory = working.filter(ci => ci.category_id === target.id);
    const nextNumber = itemsInCategory.length
      ? Math.max(...itemsInCategory.map(ci => ci.master_item_number || 0)) + 1
      : (target.range_start || 1);
    const idx = working.findIndex(ci => ci.id === item.id);
    if (idx >= 0) working[idx] = { ...working[idx], category_id: target.id, master_item_number: nextNumber };
    assignments.push({ itemId: item.id, categoryId: target.id, masterItemNumber: nextNumber });
  }

  // Pass 3: now that numbering is already decided and collision-free,
  // the actual writes are independent of each other - fire them all at
  // once instead of one sequential round-trip per item. With dozens of
  // items stuck in Uncategorized, sequential awaits here was the real
  // cause of the button appearing to hang: N items x one network
  // round-trip each, one at a time, easily 20-60+ seconds total.
  const results = await Promise.all(assignments.map(a =>
    supabase.from("catalog_items").update({
      category_id: a.categoryId, master_item_number: a.masterItemNumber,
    }).eq("id", a.itemId)
  ));
  const moved = results.filter(r => !r.error).length;
  // Matched-but-failed-to-save is a COMPLETELY different problem from
  // matched-nothing (a real database/permissions error, not a keyword
  // gap) and must never be reported with the same "none matched"
  // wording - that would hide a genuine write failure behind a message
  // that points at the wrong cause (the dictionary) entirely.
  const firstWriteError = results.find(r => r.error)?.error?.message || null;

  return { moved, checked: stuck.length, matchedButFailed: assignments.length - moved, firstWriteError };
}

// The manual counterpart to reclassifyUncategorizedItems: automatic
// classification only ever finds what its keywords can find, so anything
// genuinely ambiguous - or anything an org just wants somewhere else -
// needs a person to move it directly. Same numbering rule as every other
// path that assigns a category (next number in the target's own block),
// so a manually-moved item is indistinguishable from one the system
// filed correctly the first time.
async function assignItemCategory(catalogItemId, newCategoryId, catalogItems, categories) {
  const target = categories.find(c => c.id === newCategoryId);
  if (!target) return null;
  const itemsInCategory = catalogItems.filter(ci => ci.category_id === target.id);
  const nextNumber = itemsInCategory.length
    ? Math.max(...itemsInCategory.map(ci => ci.master_item_number || 0)) + 1
    : (target.range_start || 1);
  const { error } = await supabase.from("catalog_items").update({
    category_id: target.id, master_item_number: nextNumber,
  }).eq("id", catalogItemId);
  return error ? null : nextNumber;
}

function CatalogPanel({orgId,orgIndustry,categories,catalogItems,onUpdated}) {
  const [adding,setAdding]=useState(false);
  const [newName,setNewName]=useState("");
  const [newKeywords,setNewKeywords]=useState("");
  const [editingId,setEditingId]=useState(null);
  const [editKeywords,setEditKeywords]=useState("");
  const [loadingTemplate,setLoadingTemplate]=useState(false);
  const [templateMsg,setTemplateMsg]=useState("");
  const [reclassifying,setReclassifying]=useState(false);
  const [reclassifyMsg,setReclassifyMsg]=useState("");
  const [error,setError]=useState("");

  const uncategorizedCategory=useMemo(()=>categories.find(c=>c.name==="Uncategorized"),[categories]);

  async function reclassifyUncategorized(){
    setReclassifying(true);
    setReclassifyMsg("");
    setError("");
    try{
      const {moved,checked,matchedButFailed,firstWriteError}=await withTimeout(reclassifyUncategorizedItems(orgId,catalogItems,categories),20000,"Reclassify");
      if(matchedButFailed>0){
        // These DID match a category - the write itself failed. Almost
        // certainly a database permissions (RLS) problem, not a
        // dictionary gap - showing this as "none matched" would point
        // straight at the wrong cause.
        setError(`${matchedButFailed} item(s) matched a category but failed to save${firstWriteError?`: ${firstWriteError}`:" (unknown database error)"}.`);
        setReclassifyMsg(moved>0?`Moved ${moved} of ${checked} item(s) — the rest failed to save (see error above).`:"");
      }else{
        setReclassifyMsg(checked===0?"Nothing in Uncategorized right now.":
          moved===0?`Checked ${checked} item(s) in Uncategorized — none matched a current category's keywords.`:
          `Moved ${moved} of ${checked} item(s) out of Uncategorized into a matching category.`);
      }
      onUpdated();
    }catch(err){
      setError(`Reclassify failed: ${err.message||String(err)}`);
    }finally{
      setReclassifying(false);
    }
  }

  const itemCounts=useMemo(()=>{
    const m={};
    catalogItems.forEach(ci=>{ if(ci.category_id) m[ci.category_id]=(m[ci.category_id]||0)+1; });
    return m;
  },[catalogItems]);

  const uncategorizedCount=useMemo(()=>uncategorizedCategory?itemCounts[uncategorizedCategory.id]||0:0,[uncategorizedCategory,itemCounts]);


  async function addCategory(){
    if(!newName.trim()) return;
    const keywords=newKeywords.split(",").map(k=>k.trim()).filter(Boolean);
    const {range_start,range_end}=nextCategoryRange(categories);
    const {error:e}=await supabase.from("catalog_categories").insert({
      organization_id:orgId,name:newName.trim(),keywords,range_start,range_end,
    });
    if(e){ setError(e.message); return; }
    setNewName("");setNewKeywords("");setAdding(false);setError("");
    onUpdated();
  }

  async function saveKeywords(cat){
    const keywords=editKeywords.split(",").map(k=>k.trim()).filter(Boolean);
    await supabase.from("catalog_categories").update({keywords}).eq("id",cat.id);
    setEditingId(null);
    onUpdated();
  }

  async function deleteCategory(cat){
    if(itemCounts[cat.id]>0){
      alert(`Can't delete "${cat.name}" — ${itemCounts[cat.id]} catalog item(s) are still assigned to it. Reassign them first.`);
      return;
    }
    if(!window.confirm(`Delete category "${cat.name}"?`)) return;
    await supabase.from("catalog_categories").delete().eq("id",cat.id);
    onUpdated();
  }

  async function loadStarterTemplate(){
    setLoadingTemplate(true);
    setTemplateMsg("");
    setError("");
    if(!orgIndustry){
      setTemplateMsg("Set an industry for this org above first.");
      setLoadingTemplate(false);
      return;
    }
    try{
      const {added,found}=await withTimeout(loadStarterCategoriesForIndustry(orgId,orgIndustry,categories),20000,"Load starter categories");
      setTemplateMsg(!found?`No starter template found for "${orgIndustry}" yet — add categories manually below, or ask for that industry to be added to industry_templates.`:
        added?`Added ${added} starter categor${added===1?"y":"ies"}.`:"Starter categories for this industry are already all present.");
      onUpdated();
    }catch(err){
      setError(`Load starter categories failed: ${err.message||String(err)}`);
    }finally{
      setLoadingTemplate(false);
    }
  }

  return (
    <div>
      {!backendInfo.configured&&(
        <div style={{background:"#FFF8E1",border:"1px solid #FFE082",borderRadius:8,padding:"8px 11px",fontSize:11,color:"#8D6E63",marginBottom:12}}>
          This build is using the built-in backend credentials. To point a deployment at a
          different database (staging, or another tenancy), set <b>VITE_SUPABASE_URL</b> and
          <b> VITE_SUPABASE_ANON_KEY</b> at build time.
        </div>
      )}

      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
        <h3 style={{margin:0,fontSize:16,color:"white"}}>Catalog Categories</h3>
        <button onClick={()=>setAdding(true)} style={{...btn("#003584","white",{fontSize:12,padding:"8px 14px"})}}>+ Category</button>
      </div>

      {error&&<div style={{background:"#FFF3E0",color:"#E65100",padding:"10px 12px",borderRadius:8,fontSize:13,marginBottom:14}}>{error}</div>}

      <div style={{background:"white",borderRadius:10,padding:14,marginBottom:16,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        <div style={{fontSize:12,color:"#666",marginBottom:8}}>
          Pull a starter set of categories based on this org's industry{orgIndustry?<> (<b>{orgIndustry}</b>)</>:""}. Fully editable after — nothing here is locked to any one client.
        </div>
        <button onClick={loadStarterTemplate} disabled={loadingTemplate} style={{...btn("#2E7D32","white",{fontSize:12,padding:"8px 14px"})}}>
          {loadingTemplate?"Loading...":"Load starter categories"}
        </button>
        {templateMsg&&<div style={{fontSize:11,color:"#888",marginTop:8}}>{templateMsg}</div>}
      </div>

      {uncategorizedCount>0&&(
        <div style={{background:"white",borderRadius:10,padding:14,marginBottom:16,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <div style={{fontSize:12,color:"#666",marginBottom:8}}>
            <b>{uncategorizedCount}</b> item{uncategorizedCount===1?"":"s"} currently sitting in Uncategorized and need to be allocated — these didn't match any category's keywords (either at import time, or before the category existed at all). Re-check now against the categories and keywords you have today.
          </div>
          <button onClick={reclassifyUncategorized} disabled={reclassifying} style={{...btn("#2E7D32","white",{fontSize:12,padding:"8px 14px"})}}>
            {reclassifying?"Checking...":"Reclassify Uncategorized items"}
          </button>
          {reclassifyMsg&&<div style={{fontSize:11,color:"#888",marginTop:8}}>{reclassifyMsg}</div>}
        </div>
      )}

      {categories.length===0?(
        <div style={{background:"white",borderRadius:10,padding:20,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <p style={{color:"#888",fontSize:13,margin:0}}>No categories yet — everything imported so far lands in "Uncategorized" automatically. Add categories or load a starter set above.</p>
        </div>
      ):categories.map(cat=>{
        const isEditing=editingId===cat.id;
        return (
          <div key={cat.id} style={{background:"white",borderRadius:8,padding:"12px 14px",marginBottom:8,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:isEditing?8:0}}>
              <div style={{fontWeight:700,fontSize:14}}>{cat.name} <span style={{fontWeight:400,fontSize:11,color:"#AAA"}}>({itemCounts[cat.id]||0} item{itemCounts[cat.id]===1?"":"s"})</span></div>
              <div style={{display:"flex",gap:10,alignItems:"center"}}>
                {!isEditing&&<button onClick={()=>{setEditingId(cat.id);setEditKeywords((cat.keywords||[]).join(", "));}} style={{background:"none",border:"none",cursor:"pointer",color:"#888",fontSize:12,padding:0}}>✎ Keywords</button>}
                {cat.name!=="Uncategorized"&&<button onClick={()=>deleteCategory(cat)} style={{background:"none",border:"none",cursor:"pointer",color:"#E65100",fontSize:16,padding:0}} title="Delete">×</button>}
              </div>
            </div>
            {isEditing?(
              <div>
                <textarea value={editKeywords} onChange={e=>setEditKeywords(e.target.value)}
                  placeholder="comma, separated, keywords"
                  style={{...inp,width:"100%",minHeight:60,fontFamily:"inherit",boxSizing:"border-box"}} />
                <div style={{display:"flex",gap:8,marginTop:8}}>
                  <button onClick={()=>setEditingId(null)} style={{...btn("#EEE","#555",{padding:"8px 14px",fontSize:12}),flex:1}}>Cancel</button>
                  <button onClick={()=>saveKeywords(cat)} style={{...btn("#003584",undefined,{padding:"8px 14px",fontSize:12}),flex:1}}>Save</button>
                </div>
              </div>
            ):(
              <div style={{fontSize:11,color:"#AAA"}}>{(cat.keywords||[]).length?(cat.keywords||[]).join(", "):"No keywords set — new items won't auto-match into this category yet."}</div>
            )}
          </div>
        );
      })}

      {adding&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div style={{background:"white",borderRadius:12,padding:24,width:"100%",maxWidth:380}}>
            <h3 style={{margin:"0 0 14px",fontSize:16}}>New category</h3>
            <input style={{...inp,width:"100%",marginBottom:10,boxSizing:"border-box"}} placeholder="Category name" value={newName} onChange={e=>setNewName(e.target.value)} />
            <textarea style={{...inp,width:"100%",minHeight:60,marginBottom:10,fontFamily:"inherit",boxSizing:"border-box"}} placeholder="comma, separated, keywords (optional)" value={newKeywords} onChange={e=>setNewKeywords(e.target.value)} />
            {error&&<div style={{color:"#E65100",fontSize:12,marginBottom:10}}>{error}</div>}
            <div style={{display:"flex",gap:8}}>
              <button onClick={()=>{setAdding(false);setError("");}} style={{...btn("#EEE","#555"),flex:1}}>Cancel</button>
              <button onClick={addCategory} style={{...btn("#003584"),flex:2}}>Add</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function VendorDetail({vendor,vc,vendorItems,invoices,purchaseOrders,priceHistory,mappings,catalogItems,orgId,myRole,onBack,onOpenImport,onOpenRecordInvoice,onUpdated,onEditInvoice,onDeleteInvoice}) {
  const [editing,setEditing]=useState(false);
  const [name,setName]=useState(vendor.name);
  const [email,setEmail]=useState(vendor.email||"");
  const [minDollar,setMinDollar]=useState(vendor.delivery_minimum_dollar||"");
  const [minUnits,setMinUnits]=useState(vendor.delivery_minimum_units||"");
  const [saving,setSaving]=useState(false);
  const [error,setError]=useState("");
  const [priceEditId,setPriceEditId]=useState(null);
  const [priceEditValue,setPriceEditValue]=useState("");
  const [itemSearch,setItemSearch]=useState("");
  const [showItems,setShowItems]=useState(false);
  const [mapEditId,setMapEditId]=useState(null);
  const [mapEditValue,setMapEditValue]=useState("");
  const [mapError,setMapError]=useState("");
  const [mapBusy,setMapBusy]=useState(false);
  const [expandedOrderId,setExpandedOrderId]=useState(null);
  const [expandedPeriod,setExpandedPeriod]=useState(null);

  const canManage = myRole==="owner"||myRole==="manager";

  async function saveVendorDetails(){
    setSaving(true); setError("");
    const {error:e}=await supabase.from("vendors").update({
      name:name.trim(),
      email:email.trim()||null,
      delivery_minimum_dollar:minDollar?parseFloat(minDollar):null,
      delivery_minimum_units:minUnits?parseInt(minUnits):null,
    }).eq("id",vendor.id);
    if(e){ setError(e.message); setSaving(false); return; }
    setEditing(false); setSaving(false);
    onUpdated();
  }

  async function deactivateVendor(){
    if(!window.confirm(`Remove ${vendor.name} from your active vendors? This can only be undone from the database directly.`)) return;
    await supabase.from("vendors").update({is_active:false}).eq("id",vendor.id);
    onUpdated();
    onBack();
  }

  async function savePriceEdit(item){
    const newPrice=parseFloat(priceEditValue);
    if(isNaN(newPrice)||newPrice<=0){ setPriceEditId(null); return; }
    if(Math.abs((item.price||0)-newPrice)>0.001){
      await supabase.from("price_history").insert({vendor_item_id:item.id,organization_id:orgId,price:newPrice,source:"manual_edit"});
      await supabase.from("vendor_items").update({price:newPrice,last_updated:new Date().toISOString()}).eq("id",item.id);
    }
    setPriceEditId(null);
    onUpdated();
  }

  const items = vendorItems.filter(vi=>vi.vendor_id===vendor.id
    && (!itemSearch || vi.description.toLowerCase().includes(itemSearch.toLowerCase())));
  const vendorInvoices = invoices.filter(inv=>inv.vendor_id===vendor.id);
  const vendorOrders = (purchaseOrders||[]).filter(po=>po.vendor_id===vendor.id);

  // Looks up the current client-item mapping (if any) for a vendor item,
  // so the item list can show what it's linked to and let that be
  // corrected directly, not just items the matching engine happened to
  // flag as uncertain.
  const ciById=new Map((catalogItems||[]).map(ci=>[ci.id,ci]));
  function mappingFor(vendorItemId){
    const m=(mappings||[]).find(m=>m.vendor_item_id===vendorItemId);
    if(!m) return {mapping:null,catalogItem:null};
    return {mapping:m,catalogItem:ciById.get(m.catalog_item_id)||null};
  }

  async function saveItemMapping(vendorItem){
    setMapError("");
    const num=mapEditValue.trim();
    if(!num){ setMapEditId(null); return; }
    const target=(catalogItems||[]).find(ci=>String(ci.master_item_number)===num);
    if(!target){ setMapError(`No client item numbered ${num} was found.`); return; }
    setMapBusy(true);
    const {mapping}=mappingFor(vendorItem.id);
    await assignVendorItemMapping(orgId,vendorItem.id,target.id,mapping?.id||null);
    setMapBusy(false);
    setMapEditId(null); setMapEditValue("");
    onUpdated();
  }

  const vendorItemIds=new Set(vendorItems.filter(vi=>vi.vendor_id===vendor.id).map(vi=>vi.id));
  const vendorPricePeriods=(()=>{
    const groups=new Map();
    (priceHistory||[]).forEach(ph=>{
      if(!vendorItemIds.has(ph.vendor_item_id)) return;
      if(!groups.has(ph.effective_date)) groups.set(ph.effective_date,[]);
      groups.get(ph.effective_date).push(ph);
    });
    return [...groups.entries()]
      .map(([date,entries])=>({date,entries}))
      .sort((a,b)=>new Date(b.date)-new Date(a.date));
  })();

  return (
    <div>
      <button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",color:"rgba(255,255,255,0.75)",fontSize:13,marginBottom:14,padding:0}}>← Back</button>
      <div style={{background:"white",borderRadius:12,padding:20,marginBottom:14,boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
        {!editing?(<>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}>
            <div>
              <div style={{fontWeight:900,fontSize:20,color:vc.accent}}>{vendor.name}</div>
              <div style={{fontSize:12,color:"#888",marginTop:2}}>{vendorItems.filter(vi=>vi.vendor_id===vendor.id).length} items · Min ${vendor.delivery_minimum_dollar||0} · {vendor.delivery_minimum_units||0} units</div>
              {vendor.email&&<div style={{fontSize:12,color:"#888",marginTop:2}}>✉️ {vendor.email}</div>}
            </div>
            <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:6}}>
              <a href={`https://www.google.com/search?q=${encodeURIComponent(vendor.name)}`} target="_blank" rel="noreferrer"
                style={{fontSize:12,color:"#003584",textDecoration:"none",fontWeight:700}}>Research ↗</a>
              {canManage&&<button onClick={()=>setEditing(true)} style={{background:"none",border:"none",cursor:"pointer",color:"#888",fontSize:12,padding:0}}>✎ Edit</button>}
            </div>
          </div>
          <div style={{display:"flex",gap:8}}>
            {myRole!=="employee"&&(
              <button onClick={onOpenImport} style={{...btn(vc.accent,"white",{flex:1})}}>📥 Import Price Sheet</button>
            )}
            <button onClick={onOpenRecordInvoice} style={{...btn("#003584","white",{flex:1})}}>🧾 Record Invoice</button>
          </div>
        </>):(<>
          <div style={{marginBottom:10}}>
            <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Vendor name</div>
            <input style={inp} value={name} onChange={e=>setName(e.target.value)} />
          </div>
          <div style={{marginBottom:10}}>
            <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Email — where to send orders</div>
            <input style={inp} type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="orders@vendor.com" />
          </div>
          <div style={{display:"flex",gap:8,marginBottom:10}}>
            <div style={{flex:1}}>
              <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Min order ($)</div>
              <input style={inp} type="number" value={minDollar} onChange={e=>setMinDollar(e.target.value)} />
            </div>
            <div style={{flex:1}}>
              <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Min items</div>
              <input style={inp} type="number" value={minUnits} onChange={e=>setMinUnits(e.target.value)} />
            </div>
          </div>
          {error&&<div style={{color:"#E65100",fontSize:12,marginBottom:10}}>{error}</div>}
          <div style={{display:"flex",gap:8,marginBottom:10}}>
            <button onClick={()=>setEditing(false)} style={{...btn("#EEE","#555"),flex:1}}>Cancel</button>
            <button onClick={saveVendorDetails} disabled={saving} style={{...btn("#003584"),flex:2}}>{saving?"Saving...":"Save changes"}</button>
          </div>
          {myRole==="owner"&&(
            <button onClick={deactivateVendor} style={{background:"none",border:"none",cursor:"pointer",color:"#E65100",fontSize:12,padding:0}}>Remove this vendor</button>
          )}
        </>)}
      </div>

      <div style={{background:"white",borderRadius:12,marginBottom:14,boxShadow:"0 1px 3px rgba(0,0,0,0.08)",overflow:"hidden"}}>
        <button onClick={()=>setShowItems(!showItems)} style={{width:"100%",background:"none",border:"none",cursor:"pointer",padding:"14px 18px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{fontWeight:700,fontSize:14}}>Items & Pricing ({vendorItems.filter(vi=>vi.vendor_id===vendor.id).length})</span>
          <span style={{color:"#CCC"}}>{showItems?"▲":"▼"}</span>
        </button>
        {showItems&&(
          <div style={{padding:"0 18px 16px"}}>
            <input style={{...inp,marginBottom:10}} value={itemSearch} onChange={e=>setItemSearch(e.target.value)} placeholder="🔍 Search items..." />
            <div style={{maxHeight:320,overflowY:"auto"}}>
              {items.length===0?(
                <p style={{color:"#888",fontSize:13}}>No items match.</p>
              ):items.map(item=>{
                const {catalogItem}=mappingFor(item.id);
                return (
                <div key={item.id} style={{padding:"8px 0",borderBottom:"1px solid #F0F0F0"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <div style={{fontSize:13,flex:1,minWidth:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{item.description}</div>
                    {priceEditId===item.id?(
                      <div style={{display:"flex",gap:6,alignItems:"center",flexShrink:0}}>
                        <input style={{...inp,width:80,padding:"4px 8px",fontSize:13}} type="number" step="0.01" autoFocus
                          value={priceEditValue} onChange={e=>setPriceEditValue(e.target.value)} />
                        <button onClick={()=>savePriceEdit(item)} style={{...btn("#003584","white",{fontSize:11,padding:"5px 8px"})}}>✓</button>
                        <button onClick={()=>setPriceEditId(null)} style={{...btn("#EEE","#555",{fontSize:11,padding:"5px 8px"})}}>✕</button>
                      </div>
                    ):(
                      <div style={{display:"flex",gap:8,alignItems:"center",flexShrink:0}}>
                        <span style={{fontWeight:700,fontSize:13}}>${parseFloat(item.price||0).toFixed(2)}</span>
                        {canManage&&<button onClick={()=>{setPriceEditId(item.id);setPriceEditValue(String(item.price||""));}} style={{background:"none",border:"none",cursor:"pointer",color:"#888",fontSize:12}}>✎</button>}
                      </div>
                    )}
                  </div>
                  {canManage&&(
                    mapEditId===item.id?(
                      <div style={{display:"flex",gap:6,alignItems:"center",marginTop:6}}>
                        <span style={{fontSize:11,color:"#888"}}>Client item #:</span>
                        <input style={{...inp,width:90,padding:"4px 8px",fontSize:12}} autoFocus placeholder="e.g. 1000"
                          value={mapEditValue} onChange={e=>setMapEditValue(e.target.value)}
                          onKeyDown={e=>{if(e.key==="Enter") saveItemMapping(item); if(e.key==="Escape") setMapEditId(null);}} />
                        <button disabled={mapBusy} onClick={()=>saveItemMapping(item)} style={{...btn("#003584","white",{fontSize:11,padding:"4px 9px"})}}>Save</button>
                        <button onClick={()=>{setMapEditId(null);setMapError("");}} style={{...btn("#EEE","#555",{fontSize:11,padding:"4px 9px"})}}>✕</button>
                      </div>
                    ):(
                      <div style={{display:"flex",gap:6,alignItems:"center",marginTop:4}}>
                        <span style={{fontSize:11,color:"#AAA"}}>
                          {catalogItem?<>Mapped to <b style={{color:"#666"}}>#{catalogItem.master_item_number} {catalogItem.name}</b></>:"Not linked to a client item"}
                        </span>
                        <button onClick={()=>{setMapEditId(item.id);setMapEditValue(catalogItem?String(catalogItem.master_item_number):"");setMapError("");}}
                          style={{background:"none",border:"none",cursor:"pointer",color:"#888",fontSize:11,padding:0}}>✎</button>
                      </div>
                    )
                  )}
                  {mapEditId===item.id&&mapError&&<div style={{color:"#E65100",fontSize:11,marginTop:4}}>{mapError}</div>}
                </div>
                );
              })}
            </div>
          </div>
        )}
      </div>

      <div style={{fontSize:11,fontWeight:800,color:"rgba(255,255,255,0.7)",letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:8}}>Invoice history</div>
      {vendorInvoices.length===0?(
        <div style={{background:"white",borderRadius:10,padding:24,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
          <p style={{color:"#888",fontSize:13,margin:0}}>No invoices recorded for {vendor.name} yet</p>
        </div>
      ):vendorInvoices.map(inv=>(
        <div key={inv.id} style={{background:"white",borderRadius:8,padding:14,marginBottom:8,
          display:"flex",justifyContent:"space-between",alignItems:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <div style={{fontSize:12,color:"#888"}}>{formatDateMDY(inv.invoice_date)||new Date(inv.created_at).toLocaleDateString()}{inv.invoice_number?` · #${inv.invoice_number}`:""}</div>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <div style={{fontWeight:800,fontSize:15}}>{formatMoney(inv.total_amount)}</div>
            {inv.file_path&&<button onClick={()=>viewStoredFile(inv.file_path)} style={{...btn("#003584","white",{fontSize:11,padding:"5px 10px"})}}>View</button>}
            <button onClick={()=>onEditInvoice(inv)} style={{background:"none",border:"none",cursor:"pointer",color:"#888",fontSize:14,padding:0}} title="Edit">✎</button>
            <button onClick={()=>onDeleteInvoice(inv)} style={{background:"none",border:"none",cursor:"pointer",color:"#E65100",fontSize:16,padding:0}} title="Delete">×</button>
          </div>
        </div>
      ))}

      <div style={{fontSize:11,fontWeight:800,color:"rgba(255,255,255,0.7)",letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:8,marginTop:20}}>Order history</div>
      {vendorOrders.length===0?(
        <div style={{background:"white",borderRadius:10,padding:24,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
          <p style={{color:"#888",fontSize:13,margin:0}}>No orders placed with {vendor.name} yet</p>
        </div>
      ):vendorOrders.map(po=>{
        const isOpen = expandedOrderId===po.id;
        const lines = po.purchase_order_lines||[];
        return (
          <div key={po.id} style={{background:"white",borderRadius:8,marginBottom:8,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",overflow:"hidden"}}>
            <button onClick={()=>setExpandedOrderId(isOpen?null:po.id)}
              style={{width:"100%",background:"none",border:"none",cursor:"pointer",padding:"14px 14px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div style={{textAlign:"left"}}>
                <div style={{fontSize:12,color:"#888"}}>{new Date(po.created_at).toLocaleDateString()} · {lines.length} item{lines.length===1?"":"s"}</div>
                <div style={{fontSize:11,fontWeight:700,color:po.status==="submitted"?"#0A8A4B":"#888",textTransform:"capitalize"}}>{po.status||"submitted"}</div>
              </div>
              <div style={{display:"flex",alignItems:"center",gap:10}}>
                <div style={{fontWeight:800,fontSize:15}}>{formatMoney(po.total_amount)}</div>
                <span style={{color:"#CCC"}}>{isOpen?"▲":"▼"}</span>
              </div>
            </button>
            {isOpen&&(
              <div style={{borderTop:"1px solid #F0F0F0",padding:"10px 14px"}}>
                {lines.length===0?(
                  <div style={{color:"#AAA",fontSize:12}}>No line items recorded for this order.</div>
                ):lines.map(line=>{
                  const vi = vendorItems.find(v=>v.id===line.vendor_item_id);
                  return (
                    <div key={line.id} style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"5px 0",borderBottom:"1px solid #FAFAFA"}}>
                      <div>{line.quantity}× {vi?.description||"Item"}</div>
                      <div style={{fontWeight:700}}>${parseFloat(line.line_total||0).toFixed(2)}</div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        );
      })}

      <div style={{fontSize:11,fontWeight:800,color:"rgba(255,255,255,0.7)",letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:8,marginTop:20}}>Price history by period</div>
      {vendorPricePeriods.length===0?(
        <div style={{background:"white",borderRadius:10,padding:24,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
          <p style={{color:"#888",fontSize:13,margin:0}}>No price sheets imported for {vendor.name} yet</p>
        </div>
      ):vendorPricePeriods.map(period=>{
        const isOpen=expandedPeriod===period.date;
        return (
          <div key={period.date} style={{background:"white",borderRadius:8,marginBottom:8,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",overflow:"hidden"}}>
            <button onClick={()=>setExpandedPeriod(isOpen?null:period.date)}
              style={{width:"100%",background:"none",border:"none",cursor:"pointer",padding:"14px 14px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div style={{textAlign:"left"}}>
                <div style={{fontWeight:700,fontSize:13}}>Week of {new Date(period.date).toLocaleDateString()}</div>
                <div style={{fontSize:12,color:"#888"}}>{period.entries.length} item{period.entries.length===1?"":"s"} in this sheet</div>
              </div>
              <span style={{color:"#CCC"}}>{isOpen?"▲":"▼"}</span>
            </button>
            {isOpen&&(
              <div style={{borderTop:"1px solid #F0F0F0",padding:"10px 14px"}}>
                {period.entries.map(entry=>{
                  const vi=vendorItems.find(v=>v.id===entry.vendor_item_id);
                  return (
                    <div key={entry.id} style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"5px 0",borderBottom:"1px solid #FAFAFA"}}>
                      <div>{vi?.description||"Item"}</div>
                      <div style={{fontWeight:700}}>{formatMoney(entry.price)}</div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

function InvoiceEditModal({invoice,vendors,onClose,onDone}) {
  const [vendorId,setVendorId]=useState(invoice.vendor_id);
  const [date,setDate]=useState(invoice.invoice_date||"");
  const [invoiceNumber,setInvoiceNumber]=useState(invoice.invoice_number||"");
  const [totalAmount,setTotalAmount]=useState(invoice.total_amount||"");
  const [saving,setSaving]=useState(false);
  const [error,setError]=useState("");

  async function save(){
    setSaving(true); setError("");
    const {error:e}=await supabase.from("invoices").update({
      vendor_id:vendorId,
      invoice_date:date||null,
      invoice_number:invoiceNumber.trim()||null,
      total_amount:parseFloat(totalAmount)||0,
    }).eq("id",invoice.id);
    if(e){ setError(e.message); setSaving(false); return; }
    onDone();
    onClose();
  }

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:"white",borderRadius:12,padding:24,width:"100%",maxWidth:400}}>
        <h3 style={{margin:"0 0 16px",fontSize:16}}>Edit invoice</h3>
        <div style={{marginBottom:10}}>
          <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Vendor</div>
          <select style={inp} value={vendorId} onChange={e=>setVendorId(e.target.value)}>
            {vendors.map(v=><option key={v.id} value={v.id}>{v.name}</option>)}
          </select>
        </div>
        <div style={{display:"flex",gap:8,marginBottom:10}}>
          <div style={{flex:1}}>
            <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Invoice date</div>
            <input style={inp} type="date" value={date} onChange={e=>setDate(e.target.value)} />
          </div>
          <div style={{flex:1}}>
            <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Invoice #</div>
            <input style={inp} value={invoiceNumber} onChange={e=>setInvoiceNumber(e.target.value)} />
          </div>
        </div>
        <div style={{marginBottom:16}}>
          <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Total amount ($)</div>
          <input style={inp} type="number" step="0.01" value={totalAmount} onChange={e=>setTotalAmount(e.target.value)} />
        </div>
        {error&&<div style={{color:"#E65100",fontSize:12,marginBottom:12}}>{error}</div>}
        <div style={{display:"flex",gap:8}}>
          <button onClick={onClose} style={{...btn("#EEE","#555"),flex:1}}>Cancel</button>
          <button onClick={save} disabled={saving} style={{...btn("#003584"),flex:2}}>{saving?"Saving...":"Save changes"}</button>
        </div>
      </div>
    </div>
  );
}

function AddVendorModal({orgId,onClose,onDone}) {
  const [name,setName]=useState("");
  const [email,setEmail]=useState("");
  const [minDollar,setMinDollar]=useState("");
  const [minUnits,setMinUnits]=useState("");
  const [saving,setSaving]=useState(false);
  const [error,setError]=useState("");

  async function save(){
    if(!name.trim()) return;
    setSaving(true); setError("");
    const {error:e}=await supabase.from("vendors").insert({
      organization_id:orgId, name:name.trim(), email:email.trim()||null,
      delivery_minimum_dollar:minDollar?parseFloat(minDollar):null,
      delivery_minimum_units:minUnits?parseInt(minUnits):null,
    });
    if(e){ setError(e.message); setSaving(false); return; }
    onDone();
    onClose();
  }

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:"white",borderRadius:12,padding:24,width:"100%",maxWidth:400}}>
        <h3 style={{margin:"0 0 16px",fontSize:16}}>Add a vendor</h3>
        <div style={{marginBottom:10}}>
          <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Vendor name</div>
          <input style={inp} value={name} onChange={e=>setName(e.target.value)} placeholder="e.g. US Foods" />
        </div>
        <div style={{marginBottom:10}}>
          <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Email — where to send orders</div>
          <input style={inp} type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="orders@vendor.com" />
        </div>
        <div style={{display:"flex",gap:8,marginBottom:16}}>
          <div style={{flex:1}}>
            <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Min order ($)</div>
            <input style={inp} type="number" value={minDollar} onChange={e=>setMinDollar(e.target.value)} placeholder="500" />
          </div>
          <div style={{flex:1}}>
            <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Min items</div>
            <input style={inp} type="number" value={minUnits} onChange={e=>setMinUnits(e.target.value)} placeholder="20" />
          </div>
        </div>
        {error&&<div style={{color:"#E65100",fontSize:12,marginBottom:12}}>{error}</div>}
        <div style={{display:"flex",gap:8}}>
          <button onClick={onClose} style={{...btn("#EEE","#555"),flex:1}}>Cancel</button>
          <button onClick={save} disabled={saving||!name.trim()} style={{...btn("#003584"),flex:2}}>{saving?"Adding...":"Add vendor"}</button>
        </div>
      </div>
    </div>
  );
}

// ── PASTE MODAL ───────────────────────────────────────────────────────
function PasteModal({vendors,orgId,catalogItems,categories,onClose,onDone,initialVendorId,initialMode}) {
  const [vendorId,setVendorId]=useState(initialVendorId||vendors[0]?.id||"");
  const [mode,setMode]=useState(initialMode||"pricelist");
  const [pastedText,setPastedText]=useState("");
  const [fileGroups,setFileGroups]=useState([]); // [{id,file,name,text}] — one entry per dragged/selected file
  const [parsedGroups,setParsedGroups]=useState([]); // after Parse: fileGroups (+pasted text) each with rows/skipped attached
  const [step,setStep]=useState(1);
  const [loading,setLoading]=useState(false);
  const [result,setResult]=useState(null);
  const [dragOver,setDragOver]=useState(false);
  const [fileBusy,setFileBusy]=useState(false);

  async function handleDroppedFiles(files){
    const fileArr=Array.from(files||[]);
    if(!fileArr.length) return;
    setFileBusy(true);
    try{
      const newGroups=await Promise.all(fileArr.map(async file=>{
        const text=await fileToText(file);
        return {id:`${file.name}_${file.size}_${Date.now()}_${Math.random()}`,file,name:file.name,text};
      }));
      setFileGroups(prev=>[...prev,...newGroups]);
    }catch(err){
      alert(err.message);
    }
    setFileBusy(false);
  }

  function removeFileGroup(id){
    setFileGroups(prev=>prev.filter(g=>g.id!==id));
  }

  function doParse(){
    // Every dropped/selected file is parsed on its own — never merged into
    // one blob of raw text — since two different vendor documents can use
    // completely different table structures. A pasted blob (if any) is
    // treated as one more independent document, the same way.
    const docs=[...fileGroups];
    if(pastedText.trim().length>0){
      docs.push({id:"pasted",file:null,name:"Pasted text",text:pastedText});
    }
    const groups=docs.map(d=>{
      const r=parseDocument(d.text);
      return {...d,rows:r.rows,skipped:r.skipped};
    });
    setParsedGroups(groups);
    setStep(2);
  }

  const allRows=parsedGroups.flatMap(g=>g.rows.map(row=>({...row,_source:g.name})));
  const allSkipped=parsedGroups.flatMap(g=>g.skipped.map(s=>({...s,_source:g.name})));
  const allIncomplete=allRows.filter(r=>r.priceUnavailable);
  const showSourceTags=parsedGroups.length>1;

  async function doSave(){
    setLoading(true);
    const vendor=vendors.find(v=>v.id===vendorId);
    let updated=0,created=0,invoiceTotal=0,invoicesCreated=0,mapped=0;
    let saveError=null;

    if(mode==="pricelist"){
      const workingCatalogItems=[...catalogItems];
      const workingCategories=[...categories];
      const importBatchTime=new Date().toISOString();

      for(const row of allRows){
        const q=row.code
          ?supabase.from("vendor_items").select("id,price").eq("organization_id",orgId).eq("vendor_id",vendorId).eq("vendor_item_code",row.code).maybeSingle()
          :supabase.from("vendor_items").select("id,price").eq("organization_id",orgId).eq("vendor_id",vendorId).eq("description",row.description).maybeSingle();
        const {data:ex}=await q;
        let vendorItemId;
        if(ex){
          vendorItemId=ex.id;
          if(row.priceUnavailable){
            // Vendor listed this product with no price this time (N/A,
            // TBD, call, etc). Don't overwrite the last known real price
            // with null - that would destroy good data over a vendor
            // simply not quoting it this week. Still bump last_updated
            // (they DID check in on this item, just without a number) and
            // flag it so Orders can show it needs attention instead of
            // silently reusing a possibly-stale price.
            await supabase.from("vendor_items").update({pack_size:row.packSize,last_updated:importBatchTime,price_unavailable:true}).eq("id",ex.id);
          } else {
            // Always log this period's price for this item, even if it
            // didn't change — the price sheet needs to be reconstructable
            // as a whole for any given week, not just a log of changes.
            await supabase.from("price_history").insert({vendor_item_id:ex.id,organization_id:orgId,price:row.price,source:"price_list",effective_date:importBatchTime});
            // last_updated must bump every time this item appears in a price
            // sheet, even when the price is unchanged — the vendor just
            // CONFIRMED this price as of today. This is what Admin's price
            // refresh/staleness indicator keys off of; only bumping it when
            // the price literally moved would falsely flag a flat-priced
            // item as needing a refresh even after a brand-new sheet came in.
            await supabase.from("vendor_items").update({price:row.price,pack_size:row.packSize,last_updated:importBatchTime,price_unavailable:false,price_source:"price_list"}).eq("id",ex.id);
          }
          updated++;
        } else {
          const {data:ni}=await supabase.from("vendor_items").insert({organization_id:orgId,vendor_id:vendorId,vendor_item_code:row.code,description:row.description,pack_size:row.packSize,price:row.price,last_updated:importBatchTime,price_unavailable:!!row.priceUnavailable}).select().single();
          if(ni){
            vendorItemId=ni.id;
            if(!row.priceUnavailable){
              await supabase.from("price_history").insert({vendor_item_id:ni.id,organization_id:orgId,price:row.price,source:"price_list",effective_date:importBatchTime});
            }
          }
          created++;
        }

        // Whether this vendor item is brand new or was just updated, make
        // sure it's actually linked to a catalog item — this is the step
        // that was missing entirely: without it, an imported price never
        // shows up anywhere to order from or compare against other vendors.
        if(vendorItemId){
          const {data:existingMapping}=await supabase.from("item_mappings").select("id").eq("organization_id",orgId).eq("vendor_item_id",vendorItemId).maybeSingle();
          if(!existingMapping){
            const match=await matchOrCreateCatalogItem(orgId,row.description,workingCatalogItems,workingCategories);
            if(match){
              await supabase.from("item_mappings").insert({
                organization_id:orgId, catalog_item_id:match.catalogItemId, vendor_item_id:vendorItemId,
                confidence_score:Math.round((match.score??0)*100),
                match_method:"rule_based", comparison_track:match.track,
              });
              mapped++;
            }
          }
        }
      }
    } else {
      // Invoice mode: each source document is its own invoice — a dropped
      // batch of 3 invoice PDFs must become 3 separate invoice records,
      // each with its own original file and its own line items, never
      // merged into one.
      //
      // Fetch this vendor's full item list ONCE, reused for both exact and
      // fuzzy matching across every line of every invoice in this batch —
      // avoids a database round-trip per line, and gives the fuzzy fallback
      // below a full candidate pool to compare against.
      const {data:vendorItemsForMatch}=await supabase.from("vendor_items").select("id,vendor_item_code,description,price").eq("organization_id",orgId).eq("vendor_id",vendorId);
      const viList=vendorItemsForMatch||[];
      // Same working-copy pattern as price-sheet import, so an item that
      // genuinely matches nothing on file still ends up in the catalog
      // and orderable, instead of vanishing into a permanent "no match".
      const workingCatalogItems=[...catalogItems];
      const workingCategories=[...categories];

      for(const group of parsedGroups){
        if(!group.rows.length) continue;
        const groupTotal=group.rows.reduce((s,row)=>s+(row.amount!=null?row.amount:row.price),0);

        let filePath=null, fileName=null;
        if(group.file){
          const upload=await uploadOriginalFile(orgId,vendorId,group.file);
          if(upload.error){
            saveError=(saveError?saveError+" ":"")+`"${group.name}": data was extracted, but the original file couldn't be saved: `+upload.error.message;
          } else {
            filePath=upload.path;
            fileName=upload.name;
          }
        }

        const {data:inv,error:invErr}=await supabase.from("invoices").insert({
          organization_id:orgId, vendor_id:vendorId, total_amount:r2(groupTotal),
          raw_text:group.text, invoice_date:findDate(group.text)||new Date().toISOString().split("T")[0],
          invoice_number:findInvoiceNumber(group.text), status:"recorded",
          file_path:filePath, file_name:fileName,
        }).select().single();

        if(invErr){
          saveError=(saveError?saveError+" ":"")+`"${group.name}" couldn't be saved: `+invErr.message;
          continue;
        }
        invoicesCreated++;
        invoiceTotal+=groupTotal;

        const linesToInsert=[];
        for(const row of group.rows){
          // Price verification: look up what this item was actually quoted
          // at (its current vendor price) and compare to what was actually
          // paid on this invoice. A real mismatch here is exactly the kind
          // of thing worth catching — paying more than what was quoted.
          //
          // Matching order: exact item code, then exact description, then
          // — only when neither of those hits — a size-aware fuzzy
          // fallback against this vendor's item descriptions (see
          // bestInvoiceMatch above). That fallback matters most for
          // vendors with no item codes at all (their invoices and price
          // sheets are often separately-typed documents that word the
          // same product slightly differently), where exact matching
          // alone silently loses price verification.
          //
          // Every line gets a real confidence percentage and method
          // recorded — not just matched-or-not — so anything auto-matched
          // at less than full confidence, or not matched at all, can be
          // flagged and studied in Records rather than silently blending in.
          let matched=null, confidence=null, method=null;
          if(row.code){
            matched=viList.find(vi=>vi.vendor_item_code===row.code)||null;
            if(matched){ confidence=100; method="code"; }
          }
          if(!matched&&row.description){
            matched=viList.find(vi=>vi.description===row.description)||null;
            if(matched){ confidence=100; method="exact_description"; }
          }
          if(!matched&&row.description&&viList.length){
            const fuzzy=bestInvoiceMatch(row.description,viList,MATCH_POLICY.autoLink);
            if(fuzzy){ matched=fuzzy.vendorItem; confidence=Math.round(fuzzy.score*100); method="fuzzy"; }
          }
          // Nothing matched at all - rather than leave this line
          // permanently unmatched (no price ever tracked, never appears
          // in the order guide), create a vendor item from the invoice
          // itself. Clearly flagged price_source:"invoice" so it's never
          // confused with a real price-sheet-quoted price - an actual
          // vendor quote always takes over the moment one comes in,
          // since a later price-sheet import matches on description the
          // same way it always has. Erring toward "create a new entry"
          // rather than force a shaky match is the safer failure mode:
          // a duplicate is visible and fixable, a wrong price match is
          // silently misleading.
          if(!matched&&row.description){
            const {data:created}=await supabase.from("vendor_items").insert({
              organization_id:orgId, vendor_id:vendorId, vendor_item_code:row.code,
              description:row.description, pack_size:row.packSize, price:row.price,
              last_updated:inv.invoice_date, price_source:"invoice",
            }).select().single();
            if(created){
              viList.push(created);
              matched=created; method="created_from_invoice"; confidence=null;
              const catMatch=await matchOrCreateCatalogItem(orgId,row.description,workingCatalogItems,workingCategories);
              if(catMatch){
                await supabase.from("item_mappings").insert({
                  organization_id:orgId, catalog_item_id:catMatch.catalogItemId, vendor_item_id:created.id,
                  confidence_score:Math.round((catMatch.score??0)*100), match_method:"rule_based", comparison_track:catMatch.track,
                });
              }
            }
          }
          const quotedPrice=matched?.price!=null&&method!=="created_from_invoice"?parseFloat(matched.price):null;
          const variance=quotedPrice!=null?r2(row.price-quotedPrice):null;
          linesToInsert.push({
            invoice_id:inv.id, vendor_item_id:matched?.id||null, vendor_item_code:row.code, description:row.description,
            unit_price:row.price, line_total:(row.amount!=null?row.amount:row.price),
            price_variance:variance, match_confidence:confidence, match_method:method,
          });
        }
        const {error:linesErr}=await supabase.from("invoice_lines").insert(linesToInsert);
        if(linesErr){
          saveError=(saveError?saveError+" ":"")+`"${group.name}" was recorded, but its line items didn't save: `+linesErr.message;
        }
      }
    }

    setResult({mode,vendor:vendor?.name,updated,created,mapped,invoiceTotal:r2(invoiceTotal),invoicesCreated,count:allRows.length,error:saveError});
    setStep(3);setLoading(false);
  }

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:1000,display:"flex",alignItems:"flex-end",justifyContent:"center"}}>
      <div style={{background:"white",borderRadius:"16px 16px 0 0",padding:20,width:"100%",maxWidth:600,maxHeight:"90vh",overflowY:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <h3 style={{margin:0,fontSize:16}}>{mode==="pricelist"?"📋 Import Price Sheet":"🧾 Record Invoice"}</h3>
          <button onClick={onClose} style={{background:"none",border:"none",fontSize:24,cursor:"pointer",color:"#888"}}>×</button>
        </div>

        {step===1&&<>
          <div style={{display:"flex",gap:8,marginBottom:14}}>
            <button onClick={()=>setMode("pricelist")} style={{...btn(mode==="pricelist"?"#003584":"#EEE",mode==="pricelist"?"white":"#555"),flex:1}}>Import Price Sheet</button>
            <button onClick={()=>setMode("invoice")} style={{...btn(mode==="invoice"?"#003584":"#EEE",mode==="invoice"?"white":"#555"),flex:1}}>Import Invoice</button>
          </div>
          <div style={{marginBottom:14}}>
            <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Vendor</div>
            <select style={inp} value={vendorId} onChange={e=>setVendorId(e.target.value)}>
              {vendors.map(v=><option key={v.id} value={v.id}>{v.name}</option>)}
            </select>
          </div>
          <div style={{marginBottom:14}}>
            <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Drag in one or more {mode==="pricelist"?"price list":"invoice"} files</div>
            <div
              onDragOver={e=>{e.preventDefault();setDragOver(true);}}
              onDragLeave={()=>setDragOver(false)}
              onDrop={e=>{e.preventDefault();setDragOver(false);handleDroppedFiles(e.dataTransfer.files);}}
              style={{position:"relative",border:dragOver?"2px dashed #003584":"2px dashed #DDD",borderRadius:8,padding:16,textAlign:"center",background:dragOver?"#F0F6FF":"#FAFAFA"}}
            >
              <div style={{fontSize:13,color:"#888",marginBottom:8}}>Drop files here — any number at once</div>
              <input type="file" multiple accept=".csv,.txt,.tsv,.xlsx,.xls,.pdf"
                onChange={e=>{handleDroppedFiles(e.target.files);e.target.value="";}}
                style={{fontSize:12}} />
              {fileBusy&&<div style={{position:"absolute",inset:0,background:"rgba(255,255,255,0.85)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,color:"#003584",fontWeight:700,borderRadius:8}}>Reading file...</div>}
            </div>
            {fileGroups.length>0&&(
              <div style={{marginTop:8}}>
                {fileGroups.map(g=>(
                  <div key={g.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:12,color:"#555",padding:"6px 10px",background:"#F5F5F5",borderRadius:6,marginBottom:4}}>
                    <span>📎 {g.name} {mode==="invoice"?"— becomes its own invoice":""}</span>
                    <button onClick={()=>removeFileGroup(g.id)} style={{background:"none",border:"none",cursor:"pointer",color:"#E65100",fontSize:14,padding:0}}>✕</button>
                  </div>
                ))}
              </div>
            )}
          </div>
          <div style={{marginBottom:14}}>
            <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Or paste text directly</div>
            <textarea
              style={{...inp,height:120,resize:"vertical",fontFamily:"monospace",fontSize:12}}
              value={pastedText} onChange={e=>setPastedText(e.target.value)}
              placeholder="Copy from Excel, email, PDF — paste here..." />
          </div>
          <button onClick={doParse} disabled={fileGroups.length===0 && pastedText.trim().length<10} style={{...btn("#003584"),width:"100%"}}>
            Parse {fileGroups.length>0?`${fileGroups.length} file${fileGroups.length>1?"s":""}`+(pastedText.trim()?" + pasted text":""):"pasted text"}
          </button>
        </>}

        {step===2&&<>
          <div style={{background:"#E8F5E9",padding:"10px 14px",borderRadius:8,marginBottom:14,fontSize:13}}>
            Found {allRows.length} items across {parsedGroups.length} document{parsedGroups.length===1?"":"s"} — review and confirm
          </div>
          <div style={{maxHeight:280,overflowY:"auto",marginBottom:14}}>
            {allRows.map((row,i)=>(
              <div key={i} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:"1px solid #F0F0F0",fontSize:13}}>
                <div>
                  {row.code&&<span style={{color:"#888",marginRight:8,fontFamily:"monospace",fontSize:11}}>{row.code}</span>}
                  <span>{row.description}</span>
                  {row.packSize&&<span style={{color:"#AAA",marginLeft:6,fontSize:11}}>{row.packSize}</span>}
                  {showSourceTags&&<span style={{color:"#BBB",marginLeft:6,fontSize:10}}>· {row._source}</span>}
                </div>
                {row.priceUnavailable?(
                  <span style={{fontWeight:700,flexShrink:0,marginLeft:8,color:"#B26A00",fontSize:11}}>no price listed</span>
                ):(
                  <span style={{fontWeight:700,flexShrink:0,marginLeft:8}}>${row.price.toFixed(2)}</span>
                )}
              </div>
            ))}
          </div>
          {allIncomplete.length>0&&(
            <div style={{background:"#FFF3E0",padding:"10px 14px",borderRadius:8,marginBottom:14,fontSize:12}}>
              <div style={{fontWeight:700,color:"#B26A00"}}>{allIncomplete.length} product{allIncomplete.length>1?"s":""} listed with no price — not an error, just incomplete for now</div>
              <div style={{color:"#996600",marginTop:2}}>Saved with their last known price kept, flagged for follow-up with the vendor.</div>
            </div>
          )}
          {allSkipped.length>0&&(
            <div style={{background:"#FFF3E0",padding:"10px 14px",borderRadius:8,marginBottom:14,fontSize:12}}>
              <div style={{fontWeight:700,color:"#E65100",marginBottom:6}}>{allSkipped.length} line{allSkipped.length>1?"s":""} couldn't be read — skipped, not saved</div>
              <div style={{maxHeight:120,overflowY:"auto"}}>
                {allSkipped.map((s,i)=>(
                  <div key={i} style={{color:"#999",marginBottom:3,fontFamily:"monospace",fontSize:11}}>
                    "{s.line.slice(0,60)}" — {s.reason}{showSourceTags?` (${s._source})`:""}
                  </div>
                ))}
              </div>
            </div>
          )}
          <div style={{display:"flex",gap:8}}>
            <button onClick={()=>setStep(1)} style={{...btn("#EEE","#555"),flex:1}}>← Back</button>
            <button onClick={doSave} disabled={loading||!allRows.length} style={{...btn("#003584"),flex:2}}>
              {loading?"Saving...":mode==="invoice"?`Save ${parsedGroups.filter(g=>g.rows.length).length} invoice${parsedGroups.filter(g=>g.rows.length).length===1?"":"s"}`:"Save "+allRows.length+" items"}
            </button>
          </div>
        </>}

        {step===3&&result&&(
          <div style={{textAlign:"center",padding:"20px 0"}}>
            <div style={{fontSize:40,marginBottom:12}}>{result.error?"⚠️":"✅"}</div>
            <h3 style={{margin:"0 0 8px"}}>{result.vendor}</h3>
            {result.mode==="pricelist"
              ?<p style={{color:"#666",fontSize:14}}>{result.updated} items updated · {result.created} new items added · {result.mapped} linked to your catalog for ordering</p>
              :<p style={{color:"#666",fontSize:14}}>{result.invoicesCreated} invoice{result.invoicesCreated===1?"":"s"} recorded · {result.count} line{result.count===1?"":"s"} · {formatMoney(result.invoiceTotal)} total</p>}
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
  const [categories,setCategories]=useState([]);
  const [vendorItems,setVendorItems]=useState([]);
  const [mappings,setMappings]=useState([]);
  const [invoices,setInvoices]=useState([]);
  const [purchaseOrders,setPurchaseOrders]=useState([]);
  const [priceHistory,setPriceHistory]=useState([]);
  const [quantities,setQuantities]=useState({}); // {catalogItemId_case: n, catalogItemId_each: n}
  const [tab,setTab]=useState("order");
  const [showPaste,setShowPaste]=useState(false);
  const [selectedVendorId,setSelectedVendorId]=useState(null);
  const [importMode,setImportMode]=useState("pricelist");
  const [vendorDetailId,setVendorDetailId]=useState(null);
  const [logoUrl,setLogoUrl]=useState(null);
  const [logoUploading,setLogoUploading]=useState(false);
  const [showAddVendor,setShowAddVendor]=useState(false);
  const [editingInvoice,setEditingInvoice]=useState(null);
  const [unitSelection,setUnitSelection]=useState({});
  const [vendorOverride,setVendorOverride]=useState({});
  const [priceOverride,setPriceOverride]=useState({});
  const [openPriceMenu,setOpenPriceMenu]=useState(null);
  const [customPriceInput,setCustomPriceInput]=useState("");
  const [expandedOrderTabOrder,setExpandedOrderTabOrder]=useState(null);
  const [expandedInvoiceId,setExpandedInvoiceId]=useState(null);
  const [expandedPricePeriod,setExpandedPricePeriod]=useState(null);
  const [recordsVendorFilter,setRecordsVendorFilter]=useState(null);
  const [priceSheetVendorFilter,setPriceSheetVendorFilter]=useState(null);
  const [search,setSearch]=useState("");
  const [loading,setLoading]=useState(true);
  // Order Guide's own category filter + sort mode - separate state from
  // Item Catalog's (different tab, different job: this one is for
  // PLACING orders, so it only ever shows items with a vendor price -
  // but the same "browse by item type, order by code/alpha" idea applies.
  const [orderCategoryFilter,setOrderCategoryFilter]=useState("");
  const [orderSortMode,setOrderSortMode]=useState("alpha"); // "alpha" | "itemNumber" | "vendorCode"

  // Startup and session transitions are owned by sessionController (see
  // session.js), not decided here. It reports a single "entered" flag for
  // a REAL sign-in, so a token refresh - which the auth provider also
  // reports as SIGNED_IN, and which fires just from switching back to
  // this tab - can't yank someone back to Orders mid-work. It also
  // discards callbacks from a superseded startup, so a fast
  // sign-out/sign-in can't leave a stale session on screen.
  useEffect(()=>{
    let cancelled=false;
    sessionController.start(({session:s,entered})=>{
      if(cancelled) return;
      setSession(s);
      if(entered) setTab("order");
    }).catch(()=>{ if(!cancelled) setSession(null); });
    return ()=>{ cancelled=true; sessionController.stop(); };
  },[]);

  // Without this, dropping a file anywhere outside the exact import drop
  // zone — even one pixel off, or before the import modal is even open —
  // falls through to the browser's default behavior, which just opens the
  // file in the tab instead of letting our own drop handler run.
  useEffect(()=>{
    const preventDefault=e=>e.preventDefault();
    window.addEventListener("dragover",preventDefault);
    window.addEventListener("drop",preventDefault);
    return ()=>{
      window.removeEventListener("dragover",preventDefault);
      window.removeEventListener("drop",preventDefault);
    };
  },[]);

  useEffect(()=>{
    if(session===undefined) return;
    if(!session){setLoading(false);return;}
    supabase.from("profiles").upsert({id:session.user.id,email:session.user.email,updated_at:new Date().toISOString()})
      .then(({error})=>{ if(error) console.error("Profile sync failed:",error.message); });
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
    const [vr,cr,catr,vir,mr,ir,por,ph]=await Promise.all([
      supabase.from("vendors").select("*").eq("organization_id",id).eq("is_active",true).order("name"),
      supabase.from("catalog_items").select("*,catalog_categories(name)").eq("organization_id",id).order("master_item_number"),
      supabase.from("catalog_categories").select("*").eq("organization_id",id).order("name"),
      supabase.from("vendor_items").select("*").eq("organization_id",id),
      supabase.from("item_mappings").select("*").eq("organization_id",id),
      supabase.from("invoices").select("*,vendors(name),invoice_lines(*)").eq("organization_id",id).order("invoice_date",{ascending:false,nullsFirst:false}).order("created_at",{ascending:false}).limit(30),
      supabase.from("purchase_orders").select("*,purchase_order_lines(*)").eq("organization_id",id).order("created_at",{ascending:false}).limit(30),
      supabase.from("price_history").select("*").eq("organization_id",id).eq("source","price_list").order("effective_date",{ascending:false}).limit(2000),
    ]);
    setVendors(vr.data||[]);
    setCatalogItems(cr.data||[]);
    setCategories(catr.data||[]);
    setVendorItems(vir.data||[]);
    setMappings(mr.data||[]);
    setInvoices(ir.data||[]);
    setPurchaseOrders(por.data||[]);
    setPriceHistory(ph.data||[]);
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
  const unmappedCount=useMemo(()=>{
    const mappedIds=new Set(mappings.map(m=>m.vendor_item_id));
    return vendorItems.filter(vi=>!mappedIds.has(vi.id)).length;
  },[vendorItems,mappings]);

  const [backfilling,setBackfilling]=useState(false);
  async function backfillMappings(){
    setBackfilling(true);
    const mappedIds=new Set(mappings.map(m=>m.vendor_item_id));
    const unmapped=vendorItems.filter(vi=>!mappedIds.has(vi.id));
    const workingCatalogItems=[...catalogItems];
    const workingCategories=[...categories];
    for(const vi of unmapped){
      const match=await matchOrCreateCatalogItem(org.id,vi.description,workingCatalogItems,workingCategories);
      if(match){
        await supabase.from("item_mappings").insert({
          organization_id:org.id, catalog_item_id:match.catalogItemId, vendor_item_id:vi.id,
          confidence_score:Math.round((match.score??0)*100),
          match_method:"rule_based", comparison_track:match.track,
        });
      }
    }
    await loadData();
    setBackfilling(false);
  }

  const productList=useMemo(()=>{
    if(!catalogItems.length) return [];
    const viMap=new Map(vendorItems.map(vi=>[vi.id,vi]));
    const vMap=new Map(vendors.map(v=>[v.id,v]));
    const refreshDays=org?.settings?.price_refresh_days||null;

    return catalogItems.map(ci=>{
      const ciMappings=mappings.filter(m=>m.catalog_item_id===ci.id);
      const options=ciMappings.map(m=>{
        const vi=viMap.get(m.vendor_item_id);
        const v=vi?vMap.get(vi.vendor_id):null;
        if(!vi||!v||!vi.price) return null;
        // A price that hasn't been refreshed by a new vendor import within
        // the org's configured window is shown as $0 rather than the last
        // real number — a clear, impossible-to-miss signal on the Orders
        // screen that new pricing is needed. The real number is never
        // lost: vendor_items.price and price_history keep it untouched,
        // this only changes what's DISPLAYED and never lets this vendor
        // win the cheapest-price ranking while stale (see sort below and
        // the solver, which both exclude expired options from "cheapest").
        const expired=isPriceExpired(vi.last_updated,refreshDays);
        const price=expired?0:parseFloat(vi.price);
        const pack=vi.pack_size;
        const norm=expired?null:normalizedPrice(price,pack);
        const each=expired?null:eachPrice(price,pack);
        return {
          vendorId:v.id, vendorName:v.name,
          vendorItemId:vi.id, vendorItemCode:vi.vendor_item_code,
          brand:vi.brand, packSize:pack, description:vi.description,
          casePrice:price,
          eachPrice:each?.price||null, eachSize:each?.size||null,
          normalizedPrice:norm?.price||null, normalizedUnit:norm?.unit||null,
          isExactMatch:m.comparison_track==="exact",
          mappingId:m.id,
          matchConfidence:m.confidence_score, matchTrack:m.comparison_track,
          color:vendorColors.get(v.id)||PALETTE[0],
          expired, lastUpdated:vi.last_updated,
          priceUnavailable:!!vi.price_unavailable,
        };
      }).filter(Boolean).sort((a,b)=>{
        // Expired options always sink to the bottom regardless of their
        // (forced-to-0) price, so a stale vendor can never look "cheapest".
        if(a.expired!==b.expired) return a.expired?1:-1;
        return a.casePrice-b.casePrice;
      });

      return {
        catalogItemId:ci.id,
        masterItemNumber:ci.master_item_number,
        name:ci.name,
        category:ci.catalog_categories?.name||"Uncategorized",
        createdAt:ci.created_at,
        brandLocked:ci.brand_locked||false,
        options,
      };
    });
  },[catalogItems,vendorItems,mappings,vendors,vendorColors,org?.settings?.price_refresh_days]);

  const vMap=useMemo(()=>new Map(vendors.map(v=>[v.id,v])),[vendors]);

  // Invoices tab's own review data: lines with no match, or only a
  // fuzzy match, plus vendor items whose only price on file so far came
  // from an invoice rather than a confirmed price sheet. Lives here (not
  // Item Catalog) because it's specifically about invoice data.
  const flaggedInvoiceLines=useMemo(()=>{
    const out=[];
    for(const inv of invoices){
      const v=vMap.get(inv.vendor_id);
      for(const line of (inv.invoice_lines||[])){
        if(!line.vendor_item_id||line.match_method==="fuzzy"){
          out.push({lineId:line.id, vendorName:v?.name||"—", vendorId:inv.vendor_id,
            invoiceDate:inv.invoice_date, description:line.description,
            price:line.unit_price, confidence:line.match_confidence, noMatch:!line.vendor_item_id});
        }
      }
    }
    return out.sort((a,b)=>(a.confidence??-1)-(b.confidence??-1));
  },[invoices,vMap]);

  const invoiceDerivedItems=useMemo(()=>
    vendorItems.filter(vi=>vi.price_source==="invoice").map(vi=>{
      const v=vMap.get(vi.vendor_id);
      return {id:vi.id, vendorName:v?.name||"—", vendorId:vi.vendor_id, description:vi.description, price:vi.price};
    }),
  [vendorItems,vMap]);

  // Price Sheets tab's own review data: prices marked unavailable, or
  // past this org's refresh window. Lives here (not Item Catalog)
  // because it's specifically about price-sheet data health.
  const priceUnavailableItems=useMemo(()=>
    vendorItems.filter(vi=>vi.price_unavailable).map(vi=>{
      const v=vMap.get(vi.vendor_id);
      return {id:vi.id, vendorName:v?.name||"—", vendorId:vi.vendor_id, description:vi.description, lastUpdated:vi.last_updated};
    }),
  [vendorItems,vMap]);

  const expiredItems=useMemo(()=>{
    const refreshDays=org?.settings?.price_refresh_days||null;
    if(!refreshDays) return [];
    return vendorItems.filter(vi=>!vi.price_unavailable&&isPriceExpired(vi.last_updated,refreshDays)).map(vi=>{
      const v=vMap.get(vi.vendor_id);
      return {id:vi.id, vendorName:v?.name||"—", vendorId:vi.vendor_id, description:vi.description, lastUpdated:vi.last_updated};
    });
  },[vendorItems,vMap,org?.settings?.price_refresh_days]);

  // Item Catalog's nav badge is scoped to catalog MAPPING issues only
  // (fuzzy vendor-item matches) - invoice-line issues get their own
  // badge on Invoices, price-sheet health (unavailable/stale) gets its
  // own badge on Price Sheets. Each tab's badge reflects only what's
  // actually reviewable on that tab.
  const needsAttentionCount=useMemo(()=>
    mappings.filter(m=>m.comparison_track==="similar").length,
  [mappings]);
  const invoiceReviewCount=flaggedInvoiceLines.length;
  const priceSheetReviewCount=priceUnavailableItems.length+expiredItems.length;

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
          options:prod.options.map(o=>({...o,price:o.casePrice,orderUnit:"case"})),
          forcedVendorId:vendorOverride[caseKey]||null,
          forcedPrice:priceOverride[caseKey]!=null?priceOverride[caseKey]:null});
      }
      if(eachQty>0&&prod.options.some(o=>o.eachPrice)){
        items.push({...prod,catalogItemId:`${prod.catalogItemId}_each`,quantity:eachQty,orderUnit:"each",
          options:prod.options.filter(o=>o.eachPrice).map(o=>({...o,price:o.eachPrice,packSize:o.eachSize,orderUnit:"each"})),
          forcedVendorId:vendorOverride[eachKey]||null,
          forcedPrice:priceOverride[eachKey]!=null?priceOverride[eachKey]:null});
      }
    }
    return items;
  },[productList,quantities,vendorOverride,priceOverride]);

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
  const totalUnits=baskets.reduce((s,b)=>s+b.units,0);

  const orderCategoryList=useMemo(()=>{
    // Order Guide only ever lists items with a vendor price (see the
    // options.length>0 filter below) - the category chip list is scoped
    // to that same orderable set, not the full catalog, so a category
    // that's entirely unpriced doesn't show an empty chip here. Order
    // Guide's categories are alphabetical (Dairy, General, Meat,
    // Paper Goods, Produce...) - not the number-range order Item
    // Catalog uses, since this screen is for FINDING something to
    // order, not for working the numbering itself.
    const set=new Set(productList.filter(i=>i.options.length>0).map(p=>p.category||"Uncategorized"));
    return [...set].sort((a,b)=>a.localeCompare(b));
  },[productList]);

  const filtered=useMemo(()=>{
    // Order Guide is for ORDERING - a client-created item with no vendor
    // price mapped to it yet has nothing to order, so it's excluded here
    // even though it's fully visible in Item Catalog.
    const matches=productList.filter(i=>{
      if(i.options.length===0) return false;
      if(orderCategoryFilter&&(i.category||"Uncategorized")!==orderCategoryFilter) return false;
      if(search&&!i.name.toLowerCase().includes(search.toLowerCase())) return false;
      return true;
    });
    // Category alphabetical first, then ordered within each category by
    // whichever sort mode is active (defaults to alphabetical too, so
    // "alphabetize category, then alphabetize in the category" is the
    // out-of-the-box behavior).
    const groups=new Map();
    matches.forEach(item=>{
      const cat=item.category||"Uncategorized";
      if(!groups.has(cat)) groups.set(cat,[]);
      groups.get(cat).push(item);
    });
    return [...groups.entries()]
      .map(([category,items])=>({category,items:items.sort((a,b)=>compareItems(a,b,orderSortMode))}))
      .sort((a,b)=>a.category.localeCompare(b.category));
  },[productList,search,orderCategoryFilter,orderSortMode]);

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

  async function deleteInvoice(inv){
    if(!window.confirm(`Delete this ${inv.vendors?.name||""} invoice? This can't be undone.`)) return;
    await supabase.from("invoice_lines").delete().eq("invoice_id",inv.id);
    await supabase.from("invoices").delete().eq("id",inv.id);
    if(inv.file_path){
      await supabase.storage.from("documents").remove([inv.file_path]);
    }
    loadData();
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
        <div style={{fontSize:48}}>🦉</div>
        <div style={{fontWeight:900,fontSize:20,letterSpacing:"0.18em",color:"#4A90D9",marginTop:8}}>KERDOS</div>
        <div style={{marginTop:12,opacity:0.6,fontSize:13}}>Loading...</div>
      </div>
    </div>
  );
  if(!session) return <LandingGate />;
  if(!org) return <OrgGate user={session.user} onComplete={o=>{setOrg(o);loadData();}} />;

  return (
    <div style={{fontFamily:"'Inter',-apple-system,sans-serif",minHeight:"100vh",background:"#003584"}}>

      {/* HEADER */}
      <header style={{background:"#003584",color:"white",padding:"0 16px",height:52,
        display:"flex",alignItems:"center",justifyContent:"space-between",
        position:"sticky",top:0,zIndex:200,boxShadow:"0 2px 8px rgba(0,0,0,0.3)"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer"}} onClick={()=>setTab("order")}>
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
        {[["order","📋 Order Guide"],
          ["catalog",`🗂️ Item Catalog${needsAttentionCount>0?` (${needsAttentionCount})`:""}`],
          ["priceSheets",`📊 Price Sheets${priceSheetReviewCount>0?` (${priceSheetReviewCount})`:""}`],
          ["invoices",`📁 Invoices${invoiceReviewCount>0?` (${invoiceReviewCount})`:""}`],
          ...(org.role==="owner"||org.role==="manager"?[["team","👥 Admin"]]:[])].map(([id,label])=>(
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
        {/* ORDER TAB */}
        {tab==="order"&&(
          <div className="order-layout">
            <style>{`
              .order-layout { display:flex; gap:12px; align-items:flex-start; max-width:1040px; margin:0 auto; }
              .order-aside { width:290px; flex-shrink:0; }
              @media (max-width:720px) {
                .order-layout { flex-direction:column; }
                .order-aside { width:100%; }
              }
            `}</style>
            <main style={{flex:1,minWidth:0}}>
              {vendors.length>0&&(
                <div style={{marginBottom:14}}>
                  <div style={{fontSize:11,fontWeight:700,color:"rgba(255,255,255,0.65)",letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:8}}>Your Vendors — tap for details</div>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
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
              <input value={search} onChange={e=>setSearch(e.target.value)}
                placeholder={`🔍 Search ${productList.length} items...`}
                style={{...inp,marginBottom:10}} />

              {orderCategoryList.length>0&&(
                <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:10}}>
                  <button onClick={()=>setOrderCategoryFilter("")} style={chipStyle(!orderCategoryFilter)}>
                    Full List
                  </button>
                  {orderCategoryList.map(c=>{
                    const isSelected=orderCategoryFilter===c;
                    return (
                      <button key={c} onClick={()=>setOrderCategoryFilter(isSelected?"":c)} style={chipStyle(isSelected)}>
                        {c}
                      </button>
                    );
                  })}
                </div>
              )}

              {orderCategoryList.length>0&&(
                <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",marginBottom:14}}>
                  <span style={{fontSize:10,fontWeight:700,color:"rgba(255,255,255,0.65)",textTransform:"uppercase",letterSpacing:"0.05em"}}>Sort:</span>
                  {[["alpha","A–Z"],["added","Date Added"],["itemNumber","Item #"],["vendorCode","Vendor Code"]].map(([id,label])=>{
                    const isSelected=orderSortMode===id;
                    return (
                      <button key={id} onClick={()=>setOrderSortMode(id)} style={chipStyle(isSelected,"sm")}>
                        {label}
                      </button>
                    );
                  })}
                </div>
              )}

              {productList.length===0&&(
                <div style={{background:"white",borderRadius:10,padding:32,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
                  <div style={{fontSize:32,marginBottom:8}}>📋</div>
                  <h3 style={{margin:"0 0 8px"}}>No items yet</h3>
                  <p style={{color:"#888",fontSize:14,margin:"0 0 16px"}}>Import a vendor price list to get started</p>
                  <button onClick={()=>setTab("priceSheets")} style={{...btn("#003584")}}>Import Price Sheet</button>
                </div>
              )}

              {filtered.map(group=>(
                <div key={group.category} style={{marginBottom:8}}>
                  <div style={{fontSize:11,fontWeight:800,color:"rgba(255,255,255,0.75)",letterSpacing:"0.08em",textTransform:"uppercase",margin:"14px 0 6px"}}>{group.category}</div>
                  <div style={{background:"white",borderRadius:8,overflow:"hidden"}}>
                    <div style={{display:"grid",gridTemplateColumns:"minmax(110px,1fr) 78px 96px 76px",columnGap:8,rowGap:2,alignItems:"center",padding:"8px 10px"}}>
                      {group.items.flatMap(item=>{
                        const caseKey=`${item.catalogItemId}_case`;
                        const eachKey=`${item.catalogItemId}_each`;
                        const caseQty=quantities[caseKey]||0;
                        const eachQty=quantities[eachKey]||0;
                        const hasEach=item.options.some(o=>o.eachPrice);
                        const cheapest=item.options[0];
                        const isActive=caseQty>0||eachQty>0;
                        const selectedUnit=hasEach?(unitSelection[item.catalogItemId]||"case"):"case";
                        const activeKey=selectedUnit==="case"?caseKey:eachKey;
                        const activeQty=selectedUnit==="case"?caseQty:eachQty;
                        const activePackSize=selectedUnit==="case"?cheapest?.packSize:cheapest?.eachSize;
                        const cartItemIdForActive=selectedUnit==="case"?item.catalogItemId:item.catalogItemId+"_each";
                        const assignment=assignMap.get(cartItemIdForActive);
                        const vc=vendorColors.get(assignment?.assignedVendorId||cheapest?.vendorId)||PALETTE[0];
                        const activePrice=assignment?assignment.price:(selectedUnit==="case"?cheapest?.casePrice:cheapest?.eachPrice);
                        const activeVendorName=assignment?assignment.assignedVendorName:cheapest?.vendorName;
                        const isCustomPrice=priceOverride[activeKey]!=null;
                        const unitOptions=(selectedUnit==="case"
                          ?item.options.map(o=>({...o,unitPrice:o.casePrice}))
                          :item.options.filter(o=>o.eachPrice).map(o=>({...o,unitPrice:o.eachPrice})))
                          .sort((a,b)=>a.unitPrice-b.unitPrice);
                        const cheapestUnitPrice=unitOptions.find(o=>!o.expired)?.unitPrice;
                        const menuOpen=openPriceMenu===activeKey;
                        const activeExpired=assignment?assignment.allExpired:!!cheapest?.expired;
                        const activeOption=item.options.find(o=>o.vendorItemId===(assignment?.vendorItemId))||cheapest;
                        const activeMatchTrack=activeOption?.matchTrack;
                        const activeMatchConfidence=activeOption?.matchConfidence;

                        const cells=[
                          <div key={item.catalogItemId+"_name"} style={{minWidth:0,padding:"8px 8px 8px 2px",borderTop:"1px solid #F2F2F2"}}>
                            <div style={{fontWeight:600,fontSize:12.5,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                              {item.name}
                              {item.brandLocked&&<span style={{marginLeft:4,fontSize:9,background:"#E3F2FD",color:"#1565C0",padding:"1px 4px",borderRadius:4,fontWeight:700}}>🔒</span>}
                              {activeMatchTrack==="similar"&&(
                                <span title="Auto-matched to this product below full confidence - worth double-checking it's really the same item"
                                  style={{marginLeft:4,fontSize:9,background:"#FFF3E0",color:"#B26A00",padding:"1px 4px",borderRadius:4,fontWeight:700}}>
                                  🔍 {activeMatchConfidence}%
                                </span>
                              )}
                            </div>
                            <div style={{fontSize:10,color:"#AAA"}}>{activePackSize||""}</div>
                          </div>,
                          <div key={item.catalogItemId+"_unit"} style={{padding:"8px 6px",borderTop:"1px solid #F2F2F2",borderLeft:"1px solid #EEE",background:"#FAFBFC"}}>
                            {hasEach?(
                              <select value={selectedUnit} onChange={e=>setUnitSelection(prev=>({...prev,[item.catalogItemId]:e.target.value}))}
                                style={{width:"100%",fontSize:11,padding:"4px 2px",borderRadius:6,border:"1px solid #DDD",background:"white",color:"#444"}}>
                                <option value="case">Case</option>
                                <option value="each">Each</option>
                              </select>
                            ):(
                              <div style={{fontSize:11,color:"#AAA",textAlign:"center"}}>Case</div>
                            )}
                          </div>,
                          <div key={item.catalogItemId+"_qty"} style={{display:"flex",alignItems:"center",justifyContent:"center",gap:4,padding:"8px 6px",borderTop:"1px solid #F2F2F2",borderLeft:"1px solid #EEE",background:"#F5F8FF"}}>
                            <button onClick={()=>setQty(activeKey,activeQty-1)}
                              style={{width:26,height:26,borderRadius:7,border:"1px solid #FFCDD2",background:"#FFEBEE",cursor:"pointer",fontSize:15,fontWeight:800,color:"#D32F2F",flexShrink:0}}>−</button>
                            <span style={{width:18,textAlign:"center",fontWeight:800,fontSize:13,color:activeQty>0?vc.accent:"#CCC"}}>{activeQty||"·"}</span>
                            <button onClick={()=>{if(!activeExpired) setQty(activeKey,activeQty+1);}} disabled={activeExpired}
                              title={activeExpired?"This vendor's price needs refreshing before it can be ordered":undefined}
                              style={{width:26,height:26,borderRadius:7,border:"none",background:activeExpired?"#DDD":"#2E7D32",cursor:activeExpired?"not-allowed":"pointer",fontSize:15,fontWeight:800,color:"white",flexShrink:0}}>+</button>
                          </div>,
                          <div key={item.catalogItemId+"_price"} style={{textAlign:"right",padding:"6px",borderTop:"1px solid #F2F2F2",borderLeft:"1px solid #EEE",background:vc.bg}}>
                            <button onClick={()=>{
                                setOpenPriceMenu(menuOpen?null:activeKey);
                                setCustomPriceInput(activePrice!=null?String(activePrice):"");
                              }}
                              style={{background:"white",border:`1px solid ${activeExpired?"#FFCC80":vc.light}`,borderRadius:8,cursor:unitOptions.length?"pointer":"default",padding:"5px 8px",textAlign:"right",width:"100%"}}>
                              {activeExpired?(
                                <div style={{fontWeight:700,fontSize:11,color:"#E65100"}}>⚠ Refresh needed</div>
                              ):(
                                <div style={{fontWeight:800,fontSize:13,color:vc.accent}}>${activePrice?.toFixed(2)}{isCustomPrice&&<span title="Custom price" style={{marginLeft:2,fontSize:9}}>✎</span>}</div>
                              )}
                              <div style={{fontSize:9,fontWeight:700,color:activeExpired?"#E65100":vc.accent}}>{activeVendorName} {unitOptions.length>1&&(menuOpen?"▲":"▾")}</div>
                            </button>
                          </div>,
                        ];

                        if(menuOpen){
                          cells.push(
                            <div key={item.catalogItemId+"_menu"} style={{gridColumn:"1 / -1",background:"#FAFAFA",borderTop:"1px solid #F0F0F0",borderRadius:6,padding:"8px 10px",marginBottom:4}}>
                              <div style={{fontSize:10,color:"#BBB",fontWeight:700,letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:6}}>Power ranked — tap to select</div>
                              {unitOptions.map(opt=>{
                                const isSelected=assignment?.vendorItemId===opt.vendorItemId&&!isCustomPrice;
                                return (
                                  <button key={opt.vendorItemId} onClick={()=>{
                                      if(opt.expired) return;
                                      setVendorOverride(prev=>({...prev,[activeKey]:opt.vendorId}));
                                      setPriceOverride(prev=>{const n={...prev};delete n[activeKey];return n;});
                                      setOpenPriceMenu(null);
                                    }}
                                    disabled={opt.expired}
                                    title={opt.expired?"This vendor hasn't sent updated pricing yet":undefined}
                                    style={{width:"100%",display:"flex",justifyContent:"space-between",alignItems:"center",
                                      background:opt.expired?"#FAFAFA":isSelected?"#E8F5E9":"white",border:`1px solid ${opt.expired?"#F0F0F0":isSelected?"#A5D6A7":"#EEE"}`,
                                      borderRadius:6,padding:"7px 10px",marginBottom:4,cursor:opt.expired?"not-allowed":"pointer",textAlign:"left",opacity:opt.expired?0.6:1}}>
                                    <span style={{fontSize:12,fontWeight:isSelected?700:500}}>{isSelected&&"✓ "}{opt.vendorName}
                                      {!opt.expired&&opt.priceUnavailable&&(
                                        <span title="Vendor's latest price sheet listed no price for this item - showing the last known price instead" style={{marginLeft:6,fontSize:10,fontWeight:700,color:"#B26A00",background:"#FFF3E0",padding:"1px 5px",borderRadius:4}}>
                                          ⓘ last known price
                                        </span>
                                      )}
                                      {!opt.expired&&opt.matchTrack==="similar"&&(
                                        <span title="Auto-matched to this product below full confidence - worth double-checking" style={{marginLeft:6,fontSize:10,fontWeight:700,color:"#B26A00",background:"#FFF3E0",padding:"1px 5px",borderRadius:4}}>
                                          🔍 {opt.matchConfidence}% match
                                        </span>
                                      )}
                                    </span>
                                    {opt.expired?(
                                      <span style={{fontSize:11,fontWeight:700,color:"#E65100"}}>⚠ Refresh needed</span>
                                    ):(
                                      <span style={{fontSize:12,fontWeight:700}}>${opt.unitPrice.toFixed(2)} {cheapestUnitPrice!=null&&opt.unitPrice>cheapestUnitPrice&&<span style={{color:"#E65100",fontWeight:600}}>(+${(opt.unitPrice-cheapestUnitPrice).toFixed(2)})</span>}</span>
                                    )}
                                  </button>
                                );
                              })}
                              <div style={{display:"flex",gap:6,marginTop:8,alignItems:"center"}}>
                                <span style={{fontSize:11,color:"#888",flexShrink:0}}>Negotiated price:</span>
                                <input type="number" step="0.01" value={customPriceInput} onChange={e=>setCustomPriceInput(e.target.value)}
                                  style={{...inp,padding:"5px 8px",fontSize:12,flex:1}} placeholder="$0.00" />
                                <button onClick={()=>{
                                    const val=parseFloat(customPriceInput);
                                    if(isNaN(val)||val<0) return;
                                    setPriceOverride(prev=>({...prev,[activeKey]:val}));
                                    setVendorOverride(prev=>({...prev,[activeKey]:assignment?.assignedVendorId||cheapest.vendorId}));
                                    setOpenPriceMenu(null);
                                  }}
                                  style={{...btn("#003584","white",{fontSize:11,padding:"6px 10px"}),flexShrink:0}}>Apply</button>
                              </div>
                              {(vendorOverride[activeKey]||isCustomPrice)&&(
                                <button onClick={()=>{
                                    setVendorOverride(prev=>{const n={...prev};delete n[activeKey];return n;});
                                    setPriceOverride(prev=>{const n={...prev};delete n[activeKey];return n;});
                                    setOpenPriceMenu(null);
                                  }}
                                  style={{background:"none",border:"none",color:"#888",fontSize:11,cursor:"pointer",padding:"6px 0 0",textDecoration:"underline"}}>
                                  Reset to automatic
                                </button>
                              )}
                            </div>
                          );
                        }
                        return cells;
                      })}
                    </div>
                  </div>
                </div>
              ))}

              {purchaseOrders.length>0&&(
                <div style={{marginTop:24}}>
                  <div style={{fontSize:11,fontWeight:700,color:"rgba(255,255,255,0.7)",letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:10}}>Recent Orders</div>
                  {purchaseOrders.map(po=>{
                    const isOpen=expandedOrderTabOrder===po.id;
                    const lines=po.purchase_order_lines||[];
                    const vendorName=vendors.find(v=>v.id===po.vendor_id)?.name||"Unknown vendor";
                    return (
                      <div key={po.id} style={{background:"white",borderRadius:8,marginBottom:8,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",overflow:"hidden"}}>
                        <button onClick={()=>setExpandedOrderTabOrder(isOpen?null:po.id)}
                          style={{width:"100%",background:"none",border:"none",cursor:"pointer",padding:"14px 14px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                          <div style={{textAlign:"left"}}>
                            <div style={{fontWeight:700,fontSize:13}}>{vendorName}</div>
                            <div style={{fontSize:12,color:"#888"}}>{new Date(po.created_at).toLocaleDateString()} · {lines.length} item{lines.length===1?"":"s"}</div>
                            <div style={{fontSize:11,fontWeight:700,color:po.status==="submitted"?"#0A8A4B":"#888",textTransform:"capitalize"}}>{po.status||"submitted"}</div>
                          </div>
                          <div style={{display:"flex",alignItems:"center",gap:10}}>
                            <div style={{fontWeight:800,fontSize:15}}>{formatMoney(po.total_amount)}</div>
                            <span style={{color:"#CCC"}}>{isOpen?"▲":"▼"}</span>
                          </div>
                        </button>
                        {isOpen&&(
                          <div style={{borderTop:"1px solid #F0F0F0",padding:"10px 14px"}}>
                            {lines.length===0?(
                              <div style={{color:"#AAA",fontSize:12}}>No line items recorded for this order.</div>
                            ):lines.map(line=>{
                              const vi=vendorItems.find(v=>v.id===line.vendor_item_id);
                              return (
                                <div key={line.id} style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"5px 0",borderBottom:"1px solid #FAFAFA"}}>
                                  <div>{line.quantity}× {vi?.description||"Item"}</div>
                                  <div style={{fontWeight:700}}>{formatMoney(line.line_total)}</div>
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </main>

            {/* BASKETS */}
            <aside className="order-aside">
              <div className="order-aside-inner">
                <div style={{fontSize:11,fontWeight:800,color:"rgba(255,255,255,0.7)",letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:10}}>Order Baskets</div>

                {baskets.length===0&&(
                  <div style={{background:"white",borderRadius:10,padding:"24px 16px",textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
                    <div style={{fontSize:28,marginBottom:8}}>🧺</div>
                    <p style={{color:"#888",fontSize:13,margin:0}}>Add items to start a basket</p>
                  </div>
                )}

                {baskets.length>0&&(<>
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
                          <div style={{padding:"6px 12px 0"}}>
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
                        {vendor.delivery_minimum_units&&(
                          <div style={{padding:"6px 12px 8px"}}>
                            <div style={{display:"flex",justifyContent:"space-between",fontSize:9,color:"#AAA",marginBottom:2}}>
                              <span>Min {vendor.delivery_minimum_units} units</span>
                              <span style={{color:meetsUnits?vc.accent:"#FF9800",fontWeight:600}}>
                                {meetsUnits?"✓":`${vendor.delivery_minimum_units-basket.units} unit${vendor.delivery_minimum_units-basket.units===1?"":"s"} to go`}
                              </span>
                            </div>
                            <div style={{height:3,background:"#EEE",borderRadius:2}}>
                              <div style={{height:"100%",borderRadius:2,transition:"width 0.3s",
                                background:meetsUnits?vc.accent:"#FF9800",
                                width:`${Math.min(100,(basket.units/vendor.delivery_minimum_units)*100)}%`}} />
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
                      <span style={{fontSize:11,opacity:0.6}}>Total units</span>
                      <span style={{fontWeight:600}}>{totalUnits}</span>
                    </div>
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
                </>)}
              </div>
            </aside>
          </div>
        )}

        {/* IMPORT TAB */}
        {/* INVOICES TAB */}
        {tab==="invoices"&&(()=>{
          const invoicesWithVariance=invoices.map(inv=>{
            const lines=inv.invoice_lines||[];
            const flagged=lines.filter(l=>l.price_variance!=null&&Math.abs(l.price_variance)>0.009);
            const totalVariance=flagged.reduce((s,l)=>s+l.price_variance,0);
            return {...inv,_lines:lines,_flagged:flagged,_totalVariance:r2(totalVariance)};
          });
          const invoicesByVendor=new Map();
          for(const inv of invoicesWithVariance){
            const vid=inv.vendor_id||"unknown";
            if(!invoicesByVendor.has(vid)) invoicesByVendor.set(vid,[]);
            invoicesByVendor.get(vid).push(inv);
          }
          const vendorInvoiceGroups=[...invoicesByVendor.entries()]
            .map(([vendorId,invs])=>({
              vendorId,
              vendorName:invs[0]?.vendors?.name||"Unknown vendor",
              invoices:invs.sort((a,b)=>{
                const ad=a.invoice_date?new Date(a.invoice_date):new Date(a.created_at);
                const bd=b.invoice_date?new Date(b.invoice_date):new Date(b.created_at);
                return bd-ad;
              }),
            }))
            .sort((a,b)=>a.vendorName.localeCompare(b.vendorName));

          const visibleInvoiceGroups=recordsVendorFilter?vendorInvoiceGroups.filter(g=>g.vendorId===recordsVendorFilter):vendorInvoiceGroups;
          const filteredVendorName=recordsVendorFilter?(vendors.find(v=>v.id===recordsVendorFilter)?.name||"Unknown vendor"):null;

          return (
          <div>
            <div style={{fontSize:11,fontWeight:700,color:"rgba(255,255,255,0.65)",letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:8}}>
              {recordsVendorFilter?`Showing ${filteredVendorName} — tap to change`:"Select a vendor, or view all below"}
            </div>
            <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:18}}>
              {vendors.map(v=>{
                const vc=vendorColors.get(v.id)||PALETTE[0];
                const isSelected=recordsVendorFilter===v.id;
                return (
                  <button key={v.id} onClick={()=>setRecordsVendorFilter(isSelected?null:v.id)}
                    style={{fontSize:12,fontWeight:700,padding:"6px 14px",borderRadius:20,cursor:"pointer",
                      background:isSelected?vc.accent:vc.bg,color:isSelected?"white":vc.accent,
                      border:isSelected?`2px solid ${vc.accent}`:"2px solid transparent"}}>
                    {v.name}
                  </button>
                );
              })}
              {recordsVendorFilter&&(
                <button onClick={()=>setRecordsVendorFilter(null)}
                  style={{fontSize:12,fontWeight:700,padding:"6px 14px",borderRadius:20,cursor:"pointer",background:"none",color:"rgba(255,255,255,0.6)",border:"2px solid rgba(255,255,255,0.3)"}}>
                  ✕ Clear
                </button>
              )}
            </div>

            {(flaggedInvoiceLines.length>0||invoiceDerivedItems.length>0)&&(
              <div style={{marginBottom:22}}>
                <Section title="Invoice lines needing review" count={flaggedInvoiceLines.length} emptyText="Nothing flagged — every invoice line matched cleanly.">
                  {flaggedInvoiceLines.map(l=>(
                    <div key={l.lineId} onClick={()=>{if(l.vendorId){setVendorDetailId(l.vendorId);setTab("vendorDetail");}}}
                      style={{background:"white",borderRadius:8,padding:"10px 12px",marginBottom:6,cursor:l.vendorId?"pointer":"default",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
                      <div style={{display:"flex",justifyContent:"space-between"}}>
                        <div style={{fontWeight:600,fontSize:13}}>{l.description}</div>
                        <div style={{fontSize:11,fontWeight:700,color:l.noMatch?"#C62828":"#B26A00"}}>{l.noMatch?"⚠ no match":`🔍 ${l.confidence}% match`}</div>
                      </div>
                      <div style={{fontSize:11,color:"#999",marginTop:2}}>{l.vendorName} — {formatDateMDY(l.invoiceDate)} — {formatMoney(l.price)}</div>
                    </div>
                  ))}
                </Section>
                <Section title="Found on an invoice, not yet confirmed by a price sheet" count={invoiceDerivedItems.length} emptyText="Nothing flagged — every item on file came from a confirmed price sheet.">
                  {invoiceDerivedItems.map(i=>(
                    <div key={i.id} onClick={()=>{setVendorDetailId(i.vendorId);setTab("vendorDetail");}}
                      style={{background:"white",borderRadius:8,padding:"10px 12px",marginBottom:6,cursor:"pointer",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
                      <div style={{display:"flex",justifyContent:"space-between"}}>
                        <div style={{fontWeight:600,fontSize:13}}>{i.description}</div>
                        <div style={{fontSize:11,fontWeight:700,color:"#0288D1"}}>{formatMoney(i.price)}</div>
                      </div>
                      <div style={{fontSize:11,color:"#999",marginTop:2}}>{i.vendorName} — this price came from an invoice; will be replaced once a real price sheet confirms it</div>
                    </div>
                  ))}
                </Section>
              </div>
            )}

            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",margin:"0 0 12px"}}>
              <h3 style={{margin:0,fontSize:16,color:"white"}}>Invoice History</h3>
              <button onClick={()=>downloadTextFile(`price-variance-report-${new Date().toISOString().split("T")[0]}.csv`,buildVarianceReportCSV(invoices,vendors),"text/csv")}
                style={{...btn("#2E7D32","white",{fontSize:11,padding:"7px 12px"})}}>📄 Export Variance Report (CSV)</button>
            </div>
            {visibleInvoiceGroups.length===0?(
              <div style={{background:"white",borderRadius:10,padding:32,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
                <div style={{fontSize:32,marginBottom:8}}>🧾</div>
                <p style={{color:"#888",margin:"0 0 4px"}}>{recordsVendorFilter?`No invoices recorded for ${filteredVendorName} yet`:"No invoices recorded yet"}</p>
                <button onClick={()=>{setSelectedVendorId(recordsVendorFilter);setImportMode("invoice");setShowPaste(true);}} style={{...btn("#003584","white",{fontSize:12,padding:"8px 16px",marginTop:8})}}>🧾 Import Invoice</button>
              </div>
            ):visibleInvoiceGroups.map(vendorGroup=>{
              const vc=vendorColors.get(vendorGroup.vendorId)||PALETTE[0];
              return (
                <div key={vendorGroup.vendorId} style={{marginBottom:16}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
                    <div style={{fontWeight:800,fontSize:13,color:vc.accent,paddingLeft:2}}>{vendorGroup.vendorName}</div>
                    <button onClick={()=>{setSelectedVendorId(vendorGroup.vendorId);setImportMode("invoice");setShowPaste(true);}}
                      style={{...btn("white",vc.accent,{fontSize:10,padding:"4px 9px",border:`1px solid ${vc.accent}`})}}>🧾 Import</button>
                  </div>
                  {vendorGroup.invoices.map(inv=>{
              const isOpen=expandedInvoiceId===inv.id;
              const hasVariance=inv._flagged.length>0;
              return (
              <div key={inv.id} style={{background:"white",borderRadius:8,marginBottom:8,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",overflow:"hidden"}}>
                <div style={{padding:14,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div style={{cursor:inv.file_path?"pointer":inv._lines.length?"pointer":"default"}}
                    onClick={()=>{
                      // Pulling up an invoice means seeing what was actually
                      // submitted, not a reformatted table - if the original
                      // document is on file, that's what opens. The extracted
                      // "Price verification" breakdown is still available, just
                      // as its own explicit action below, not the default one.
                      if(inv.file_path) viewStoredFile(inv.file_path);
                      else if(inv._lines.length) setExpandedInvoiceId(isOpen?null:inv.id);
                    }}>
                    <div style={{fontWeight:700,fontSize:13}}>{formatDateMDY(inv.invoice_date)||new Date(inv.created_at).toLocaleDateString()} · {inv.status}{inv.invoice_number?` · #${inv.invoice_number}`:""}</div>
                    {hasVariance&&(
                      <div style={{fontSize:11,fontWeight:700,color:inv._totalVariance>0?"#E65100":"#0A8A4B",marginTop:2}}>
                        ⚠️ {inv._totalVariance>0?"Paid":"Paid"} {formatMoney(Math.abs(inv._totalVariance))} {inv._totalVariance>0?"more":"less"} than quoted
                      </div>
                    )}
                    {inv.file_path&&<div style={{fontSize:10,color:"#0288D1",marginTop:2}}>📄 Tap to view the original invoice</div>}
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <div style={{fontWeight:800,fontSize:16}}>{formatMoney(inv.total_amount)}</div>
                    {inv._lines.length>0&&(
                      <button onClick={()=>setExpandedInvoiceId(isOpen?null:inv.id)}
                        title="Show the extracted, line-by-line price verification instead of the original document"
                        style={{...btn("#EEE","#555",{fontSize:11,padding:"5px 10px"})}}>{isOpen?"Hide":"Verify"} {isOpen?"▲":"▾"}</button>
                    )}
                    <button onClick={()=>setEditingInvoice(inv)} style={{background:"none",border:"none",cursor:"pointer",color:"#888",fontSize:14,padding:0}} title="Edit">✎</button>
                    <button onClick={()=>deleteInvoice(inv)} style={{background:"none",border:"none",cursor:"pointer",color:"#E65100",fontSize:16,padding:0}} title="Delete">×</button>
                  </div>
                </div>
                {isOpen&&inv._lines.length>0&&(
                  <div style={{borderTop:"1px solid #F0F0F0",padding:"10px 14px",background:"#FAFAFA"}}>
                    <div style={{fontSize:10,color:"#BBB",fontWeight:700,letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:6}}>Price verification</div>
                    {inv._lines.map(line=>(
                      <div key={line.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:12,padding:"5px 0",borderBottom:"1px solid #F0F0F0"}}>
                        <div>
                          {line.description}
                          {!line.vendor_item_id?(
                            <span title="No vendor item on file matched this line at all - nothing to compare its price against" style={{marginLeft:6,fontSize:10,fontWeight:700,color:"#C62828",background:"#FFEBEE",padding:"1px 5px",borderRadius:4}}>⚠ no match</span>
                          ):line.match_method==="fuzzy"?(
                            <span title="Matched by wording similarity, not an exact code or description - worth double-checking" style={{marginLeft:6,fontSize:10,fontWeight:700,color:"#B26A00",background:"#FFF3E0",padding:"1px 5px",borderRadius:4}}>🔍 {line.match_confidence}% match</span>
                          ):null}
                        </div>
                        <div style={{textAlign:"right"}}>
                          <div style={{fontWeight:700}}>{formatMoney(line.unit_price)}</div>
                          {line.price_variance!=null&&Math.abs(line.price_variance)>0.009?(
                            <div style={{fontSize:10,color:line.price_variance>0?"#E65100":"#0A8A4B"}}>
                              quoted {formatMoney(line.unit_price-line.price_variance)} ({line.price_variance>0?"+":""}{formatMoney(line.price_variance)})
                            </div>
                          ):line.price_variance!=null?(
                            <div style={{fontSize:10,color:"#0A8A4B"}}>matches quote ✓</div>
                          ):(
                            <div style={{fontSize:10,color:"#CCC"}}>no quote on file</div>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
              );
                  })}
                </div>
              );
            })}
          </div>
          );
        })()}

        {/* PRICE SHEETS TAB */}
        {tab==="priceSheets"&&(()=>{
          const vendorItemMap=new Map(vendorItems.map(vi=>[vi.id,vi]));
          const periodGroups=new Map();
          (priceHistory||[]).forEach(ph=>{
            const vi=vendorItemMap.get(ph.vendor_item_id);
            if(!vi) return;
            const key=`${vi.vendor_id}__${ph.effective_date}`;
            if(!periodGroups.has(key)) periodGroups.set(key,{vendorId:vi.vendor_id,date:ph.effective_date,entries:[]});
            periodGroups.get(key).entries.push({...ph,description:vi.description});
          });
          const byVendor=new Map();
          for(const period of periodGroups.values()){
            if(!byVendor.has(period.vendorId)) byVendor.set(period.vendorId,[]);
            byVendor.get(period.vendorId).push(period);
          }
          const vendorPriceGroups=[...byVendor.entries()]
            .map(([vendorId,periods])=>({
              vendorId,
              vendorName:vendors.find(v=>v.id===vendorId)?.name||"Unknown vendor",
              periods:periods.sort((a,b)=>new Date(b.date)-new Date(a.date)),
            }))
            .sort((a,b)=>a.vendorName.localeCompare(b.vendorName));

          const visiblePriceGroups=priceSheetVendorFilter?vendorPriceGroups.filter(g=>g.vendorId===priceSheetVendorFilter):vendorPriceGroups;
          const filteredVendorName=priceSheetVendorFilter?(vendors.find(v=>v.id===priceSheetVendorFilter)?.name||"Unknown vendor"):null;

          return (
          <div>
            <div style={{fontSize:11,fontWeight:700,color:"rgba(255,255,255,0.65)",letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:8}}>
              {priceSheetVendorFilter?`Showing ${filteredVendorName} — tap to change`:"Select a vendor, or view all below"}
            </div>
            <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:18}}>
              {vendors.map(v=>{
                const vc=vendorColors.get(v.id)||PALETTE[0];
                const isSelected=priceSheetVendorFilter===v.id;
                return (
                  <button key={v.id} onClick={()=>setPriceSheetVendorFilter(isSelected?null:v.id)}
                    style={{fontSize:12,fontWeight:700,padding:"6px 14px",borderRadius:20,cursor:"pointer",
                      background:isSelected?vc.accent:vc.bg,color:isSelected?"white":vc.accent,
                      border:isSelected?`2px solid ${vc.accent}`:"2px solid transparent"}}>
                    {v.name}
                  </button>
                );
              })}
              {priceSheetVendorFilter&&(
                <button onClick={()=>setPriceSheetVendorFilter(null)}
                  style={{fontSize:12,fontWeight:700,padding:"6px 14px",borderRadius:20,cursor:"pointer",background:"none",color:"rgba(255,255,255,0.6)",border:"2px solid rgba(255,255,255,0.3)"}}>
                  ✕ Clear
                </button>
              )}
            </div>

            {(priceUnavailableItems.length>0||expiredItems.length>0)&&(
              <div style={{marginBottom:22}}>
                <Section title="Prices marked unavailable" count={priceUnavailableItems.length} emptyText="Nothing flagged — no vendor has an open 'no price given' item.">
                  {priceUnavailableItems.map(i=>(
                    <div key={i.id} onClick={()=>{setVendorDetailId(i.vendorId);setTab("vendorDetail");}}
                      style={{background:"white",borderRadius:8,padding:"10px 12px",marginBottom:6,cursor:"pointer",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
                      <div style={{fontWeight:600,fontSize:13}}>{i.description}</div>
                      <div style={{fontSize:11,color:"#999",marginTop:2}}>{i.vendorName} — last confirmed {daysAgo(i.lastUpdated)}, using last known price</div>
                    </div>
                  ))}
                </Section>
                <Section title="Stale prices (past refresh window)" count={expiredItems.length} emptyText={org?.settings?.price_refresh_days?"Nothing flagged — every price is within the refresh window.":"Price refresh period is off — set one in Admin to enable this check."}>
                  {expiredItems.map(i=>(
                    <div key={i.id} onClick={()=>{setVendorDetailId(i.vendorId);setTab("vendorDetail");}}
                      style={{background:"white",borderRadius:8,padding:"10px 12px",marginBottom:6,cursor:"pointer",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
                      <div style={{fontWeight:600,fontSize:13}}>{i.description}</div>
                      <div style={{fontSize:11,color:"#999",marginTop:2}}>{i.vendorName} — last confirmed {daysAgo(i.lastUpdated)}</div>
                    </div>
                  ))}
                </Section>
              </div>
            )}

            <h3 style={{margin:"0 0 12px",fontSize:16,color:"white"}}>Price Sheet History</h3>
            {visiblePriceGroups.length===0?(
              <div style={{background:"white",borderRadius:10,padding:32,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
                <div style={{fontSize:32,marginBottom:8}}>📋</div>
                <p style={{color:"#888",margin:"0 0 4px"}}>{priceSheetVendorFilter?`No price sheets imported for ${filteredVendorName} yet`:"No price sheets imported yet"}</p>
                {org.role!=="employee"&&(
                  <button onClick={()=>{setSelectedVendorId(priceSheetVendorFilter);setImportMode("pricelist");setShowPaste(true);}} style={{...btn("#003584","white",{fontSize:12,padding:"8px 16px",marginTop:8})}}>📋 Import Price Sheet</button>
                )}
              </div>
            ):visiblePriceGroups.map(vendorGroup=>{
              const vc=vendorColors.get(vendorGroup.vendorId)||PALETTE[0];
              return (
                <div key={vendorGroup.vendorId} style={{marginBottom:16}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
                    <div style={{fontWeight:800,fontSize:13,color:vc.accent,paddingLeft:2}}>{vendorGroup.vendorName}</div>
                    {org.role!=="employee"&&(
                      <button onClick={()=>{setSelectedVendorId(vendorGroup.vendorId);setImportMode("pricelist");setShowPaste(true);}}
                        style={{...btn(vc.accent,"white",{fontSize:10,padding:"4px 9px"})}}>📋 Import</button>
                    )}
                  </div>
                  {vendorGroup.periods.map(period=>{
                    const periodKey=`${period.vendorId}__${period.date}`;
                    const isOpen=expandedPricePeriod===periodKey;
                    return (
                      <div key={periodKey} style={{background:"white",borderRadius:8,marginBottom:8,boxShadow:"0 1px 3px rgba(0,0,0,0.06)",overflow:"hidden"}}>
                        <button onClick={()=>setExpandedPricePeriod(isOpen?null:periodKey)}
                          style={{width:"100%",background:"none",border:"none",cursor:"pointer",padding:14,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                          <div style={{textAlign:"left"}}>
                            <div style={{fontWeight:700,fontSize:13}}>Week of {new Date(period.date).toLocaleDateString()}</div>
                            <div style={{fontSize:11,color:"#888"}}>{period.entries.length} item{period.entries.length===1?"":"s"} in this sheet</div>
                          </div>
                          <span style={{color:"#CCC"}}>{isOpen?"▲":"▼"}</span>
                        </button>
                        {isOpen&&(
                          <div style={{borderTop:"1px solid #F0F0F0",padding:"10px 14px",background:"#FAFAFA"}}>
                            {period.entries.map(entry=>(
                              <div key={entry.id} style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"5px 0",borderBottom:"1px solid #F0F0F0"}}>
                                <div>{entry.description}</div>
                                <div style={{fontWeight:700}}>{formatMoney(entry.price)}</div>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              );
            })}
          </div>
          );
        })()}

        {/* VENDOR DETAIL TAB */}
        {tab==="vendorDetail"&&vendorDetailId&&(()=>{
          const v=vendors.find(x=>x.id===vendorDetailId);
          if(!v) return <p>Vendor not found.</p>;
          const vc=vendorColors.get(v.id)||PALETTE[0];
          return (
            <VendorDetail
              vendor={v} vc={vc} vendorItems={vendorItems} invoices={invoices} purchaseOrders={purchaseOrders} priceHistory={priceHistory}
              mappings={mappings} catalogItems={catalogItems}
              orgId={org.id} myRole={org.role}
              onBack={()=>setTab("order")}
              onOpenImport={()=>{setSelectedVendorId(v.id);setImportMode("pricelist");setShowPaste(true);}}
              onOpenRecordInvoice={()=>{setSelectedVendorId(v.id);setImportMode("invoice");setShowPaste(true);}}
              onUpdated={loadData}
              onEditInvoice={setEditingInvoice}
              onDeleteInvoice={deleteInvoice}
            />
          );
        })()}

        {/* ITEM CATALOG TAB */}
        {tab==="catalog"&&(
          <>
            {org.role!=="employee"&&unmappedCount>0&&(
              <div style={{background:"#FFF3E0",borderRadius:10,padding:16,marginBottom:14,textAlign:"center"}}>
                <div style={{fontWeight:700,color:"#E65100",marginBottom:4}}>{unmappedCount} imported item{unmappedCount===1?"":"s"} not showing up for ordering</div>
                <p style={{color:"#8A5A00",fontSize:12,margin:"0 0 10px"}}>These were imported before catalog linking existed, so they're invisible on the Order Guide page. One-time fix, safe to run anytime.</p>
                <button onClick={backfillMappings} disabled={backfilling} style={{...btn("#E65100")}}>
                  {backfilling?"Linking...":`Link ${unmappedCount} item${unmappedCount===1?"":"s"} to your catalog`}
                </button>
              </div>
            )}
            <ItemCatalogPanel orgId={org.id} productList={productList} vendors={vendors} catalogItems={catalogItems} mappings={mappings}
              vendorItems={vendorItems} categories={categories}
              onOpenVendor={(vendorId)=>{setVendorDetailId(vendorId);setTab("vendorDetail");}} onUpdated={loadData} />
          </>
        )}

        {/* TEAM TAB */}
        {tab==="team"&&(org.role==="owner"||org.role==="manager")&&(
          <>
            <TeamPanel orgId={org.id} orgName={org.name} orgIndustry={org.industry} orgSettings={org.settings} categories={categories} myRole={org.role} currentUserId={session.user.id} currentUserEmail={session.user.email} onOrgUpdated={loadData}
              logoUrl={logoUrl} onLogoUpload={handleLogoUpload} logoUploading={logoUploading} />
            <div style={{height:28}} />
            <div style={{background:"white",borderRadius:10,padding:16,boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                <h4 style={{margin:0,fontSize:14}}>Your Vendors</h4>
                <button onClick={()=>setShowAddVendor(true)} style={{...btn("#003584","white",{fontSize:12,padding:"6px 12px"})}}>+ Add Vendor</button>
              </div>
              {vendors.map(v=>{
                const vc=vendorColors.get(v.id)||PALETTE[0];
                const count=vendorItems.filter(vi=>vi.vendor_id===v.id).length;
                const invCount=invoices.filter(inv=>inv.vendor_id===v.id).length;
                return (
                  <div key={v.id} style={{padding:"10px 12px",borderRadius:8,marginBottom:6,background:vc.bg}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                      <div style={{cursor:"pointer"}} onClick={()=>{setVendorDetailId(v.id);setTab("vendorDetail");}}>
                        <div style={{fontWeight:700,color:vc.accent}}>{v.name}</div>
                        <div style={{fontSize:11,color:"#888"}}>{count} items · {invCount} invoice{invCount===1?"":"s"} · tap for full history</div>
                      </div>
                      <div style={{display:"flex",gap:6}}>
                        <button onClick={()=>{setSelectedVendorId(v.id);setImportMode("pricelist");setShowPaste(true);}} style={{...btn(vc.accent,"white",{fontSize:11,padding:"6px 10px"})}}>📋 Prices</button>
                        <button onClick={()=>{setSelectedVendorId(v.id);setImportMode("invoice");setShowPaste(true);}} style={{...btn("white",vc.accent,{fontSize:11,padding:"6px 10px",border:`1px solid ${vc.accent}`})}}>🧾 Invoice</button>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
            <div style={{height:28}} />
            <CatalogPanel orgId={org.id} orgIndustry={org.industry} categories={categories} catalogItems={catalogItems} onUpdated={loadData} />
          </>
        )}
      </div>

      {showPaste&&<PasteModal vendors={vendors} orgId={org.id} catalogItems={catalogItems} categories={categories} onClose={()=>setShowPaste(false)} onDone={loadData} initialVendorId={selectedVendorId} initialMode={importMode} />}
      {showAddVendor&&<AddVendorModal orgId={org.id} onClose={()=>setShowAddVendor(false)} onDone={loadData} />}
      {editingInvoice&&<InvoiceEditModal invoice={editingInvoice} vendors={vendors} onClose={()=>setEditingInvoice(null)} onDone={loadData} />}
    </div>
  );
}
