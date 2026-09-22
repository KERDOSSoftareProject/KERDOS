import { useState, useEffect, useMemo } from "react";
import { data as supabase, backend, sessionController, backendInfo } from "./data.js";
import { parseDocument, findDate, findInvoiceNumber } from "./ingestion.js";
import { fileToText } from "./document-reader.js";
import { createDocumentService } from "./services/documents.js";
import { loadSnapshot, saveSnapshot } from "./offline-store.js";
import { configureVocabulary, classifyCategory, nextCategoryRange, eachPrice, pricePerUnit, unitsForDimension, parsePackSize, packsEquivalent, brandsMatch, bestCatalogMatch, bestInvoiceMatch, compareProductIdentity, quoteStatus, MATCH_POLICY } from "./procurement.js";

const documents=createDocumentService(backend);


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

// ── FILE STORAGE ────────────────────────────────────────────────────
// Uploads the original, untouched file through KERDOS document storage so it can
// be reopened later exactly as it was received — separate from
// whatever data got extracted from it.
async function uploadOriginalFile(orgId, vendorId, file) {
  return documents.uploadOriginal(orgId,vendorId,file);
}

async function viewSourceDocument(documentId){
  const {data,error}=await supabase.from("import_documents").select("original_text,file_path,file_name")
    .eq("id",documentId).single();
  if(error){alert(`Cannot load original document: ${error.message}`);return;}
  if(data.file_path){await viewStoredFile(data.file_path);return;}
  const url=URL.createObjectURL(new Blob([data.original_text||""],{type:"text/plain;charset=utf-8"}));
  window.open(url,"_blank");
  setTimeout(()=>URL.revokeObjectURL(url),60000);
}

async function viewStoredFile(path) {
  try{window.open(await documents.signedUrl(path),"_blank");}
  catch(error){alert("Couldn't open file: "+error.message);}
}

// A business's own logo, so their home screen feels like theirs, not a
// generic KERDOS screen. Stored in the same private bucket as documents,
// under a dedicated "logo" folder per organization.
async function uploadOrgLogo(orgId, file) {
  return documents.uploadLogo(orgId,file);
}

async function getSignedUrl(path) {
  try{return await documents.signedUrl(path);}
  catch{return null;}
}

// ── UTILITIES ────────────────────────────────────────────────────────
function r2(n) { return Math.round(n * 100) / 100; }

// The single rule for whether a vendor option can be ordered right now.
// Two things block an option: a stale price (past the org's refresh
// window) and a brand lock (the client locked this item to one brand and
// this vendor's item is a different or unknown brand). Every screen and
// the solver ask this one question instead of each checking its own list.
function orderable(o) { return !o.expired && !o.brandMismatch && !o.priceUnavailable && !o.invoiceOnly && !o.unverified; }

// Every database write goes through this, so a rejected save (a
// permissions rule, a constraint, a dropped connection) becomes a plain
// error the screen shows, never a silent "saved" with nothing saved.
async function write(promise, what) {
  const { data, error } = await promise;
  if (error) throw new Error(`${what}: ${error.message}`);
  return data;
}
function blockReason(o) {
  if (o.expired) return "Quote expired — refresh needed";
  if (o.priceUnavailable) return "No current quoted price";
  if (o.invoiceOnly) return "Invoice charge only — quotation required";
  if (o.unverified) return "Product mapping needs review";
  if (o.brandMismatch) return o.brand ? `${o.brand} — not the locked brand` : "Brand not listed";
  return null;
}

// ── CATALOG MATCHING ────────────────────────────────────────────────
// This is the actual "master catalog" engine: every vendor uses its own
// item numbers and wording, so the only way to let someone compare "the
// same product" across vendors is to link each vendor's item to one
// shared catalog entry. Deliberately rule-based (no AI at runtime, same
// as the rest of ingestion) — uses the size-aware safeProductScore
// engine imported from procurement.js (shared with invoice matching,
// and free of any industry vocabulary): plain word overlap cannot tell
// two near-identical products, or two pack sizes, apart. A strong match reuses an existing
// catalog item (comparison_track "exact"); a partial match still reuses
// it but is marked "similar" for lower-confidence display; no reasonable
// match creates a brand-new catalog item so the product is at least
// visible and orderable, ready to pick up a second vendor later.

// The holding pen: the one category flagged is_holding_pen, where an item
// lands when it matches no category's keywords and needs a person to
// allocate it. It is identified by the flag, never by its name, so an
// org can call it whatever reads naturally to them. A "miscellaneous"
// category with its own keywords is an ordinary category, not this.
function holdingPen(categories) { return categories.find(c => c.is_holding_pen) || null; }

async function ensureHoldingPen(orgId, categories) {
  const existing = holdingPen(categories);
  if (existing) return existing;
  const { range_start, range_end } = nextCategoryRange(categories);
  const data = await write(supabase.from("catalog_categories").insert({
    organization_id: orgId, name: "Uncategorized", is_holding_pen: true, range_start, range_end, keywords: [],
  }).select().single(), "Could not create the holding category");
  categories.push(data);
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

// Copies an industry's starter pack - categories with keywords, and
// vocabulary (units, packaging words, stopwords, synonyms) - into this
// org's own editable rows. Shared by TeamPanel (auto-load right after
// picking an industry, with an opt-out) and CatalogPanel's manual button,
// so there's exactly one place that does this. Skips anything the org
// already has, so running it twice never creates duplicates.
async function loadStarterPackForIndustry(orgId, industry, categories) {
  const [templates, vocabRows] = await Promise.all([
    loadIndustryTemplates(industry),
    supabase.from("industry_vocabulary").select("kind,term,canonical").ilike("industry", industry.trim()).then(r => r.data || []),
  ]);
  if (!templates.length && !vocabRows.length) return { added: 0, addedVocabulary: 0, found: false };

  const existingNames = new Set(categories.map(c => c.name.toLowerCase()));
  const workingCategories = [...categories];
  const toInsert = templates.filter(t => !existingNames.has(t.category_name.toLowerCase())).map(t => {
    const { range_start, range_end } = nextCategoryRange(workingCategories);
    workingCategories.push({ range_start, range_end }); // reserve this block before allocating the next
    return { organization_id: orgId, name: t.category_name, keywords: t.keywords, range_start, range_end };
  });
  if (toInsert.length) await write(supabase.from("catalog_categories").insert(toInsert), "Could not add starter categories");

  const existingVocab = await write(supabase.from("org_vocabulary").select("kind,term").eq("organization_id", orgId), "Could not read vocabulary");
  const have = new Set(existingVocab.map(v => `${v.kind}:${String(v.term).toLowerCase()}`));
  const vocabToInsert = vocabRows.filter(v => !have.has(`${v.kind}:${String(v.term).toLowerCase()}`))
    .map(v => ({ organization_id: orgId, kind: v.kind, term: String(v.term).toLowerCase(), canonical: v.canonical || null }));
  if (vocabToInsert.length) await write(supabase.from("org_vocabulary").insert(vocabToInsert), "Could not add starter vocabulary");

  return { added: toInsert.length, addedVocabulary: vocabToInsert.length, found: true };
}

// Finds the best existing catalog item to attach a newly-imported vendor
// item to, or creates a new one when nothing reasonably matches. Mutates
// `workingCatalogItems` in place so multiple rows in the same import batch
// correctly match against catalog items created earlier in that same batch.
// `categories` is this org's own catalog_categories rows (with keywords);
// mutated in place the same way when the Uncategorized fallback gets
// created on first use.
async function matchOrCreateCatalogItem(orgId, description, workingCatalogItems, categories) {
  const match = bestCatalogMatch(description, workingCatalogItems);
  // A 'similar' candidate is NOT an established identity. Keep the new
  // item independent so an unrelated vendor price cannot enter the basket.
  if (match?.track === "exact") {
    return { catalogItemId: match.catalogItem.id, track: "exact", score: match.score };
  }
  const category = classifyCategory(description, categories) || await ensureHoldingPen(orgId, categories);
  // Numbered within the category's own block, not one global counter -
  // this is what actually makes the numbers read as a series per
  // category (e.g. everything in the 3000s is one category) instead of
  // one running count across the whole catalog.
  const itemsInCategory = workingCatalogItems.filter(ci => ci.category_id === category?.id);
  const nextNumber = itemsInCategory.length
    ? Math.max(...itemsInCategory.map(ci => ci.master_item_number || 0)) + 1
    : (category?.range_start || 1);
  const created = await write(supabase.from("catalog_items").insert({
    organization_id: orgId, category_id: category?.id || null,
    master_item_number: nextNumber, name: description.slice(0, 120),
    matching_behavior: "flexible", canonical_unit: null, brand_locked: false,
  }).select().single(), "Could not create the catalog item");
  workingCatalogItems.push(created);
  return { catalogItemId: created.id, track: match ? "review" : "new", score: match?.score ?? null };
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
  const header = ["Master #","Item","Category","Best Price","Best Vendor","Best Per Unit","Status","Confidence",
    ...vendors.map(v => v.name)];
  const rows = [header];
  for (const item of productList) {
    const usable = item.options.filter(orderable);
    const cheapest = usable[0] || item.options[0] || null;
    let status = "No price on file";
    let confidence = "";
    if (cheapest) {
      if (!orderable(cheapest)) status = blockReason(cheapest);
      else if (cheapest.matchTrack === "similar") { status = "Needs review"; confidence = `${cheapest.matchConfidence}%`; }
      else status = "100% matched";
    }
    const vendorCells = vendors.map(v => {
      const opt = item.options.find(o => o.vendorId === v.id && orderable(o));
      return opt ? formatMoney(opt.casePrice) : "";
    });
    const best = cheapest && orderable(cheapest) ? cheapest : null;
    rows.push([
      item.masterItemNumber, item.name, item.category,
      best ? formatMoney(best.casePrice) : "",
      best ? best.vendorName : "",
      best?.perUnit ? `${formatMoney(best.perUnit.price)}/${best.perUnit.unit}` : "",
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

// ── LOCALE ──────────────────────────────────────────────────────────
// Money and dates are shown the way this organization reads them, from
// organizations.settings (locale, currency). Nothing else in the app
// knows a currency symbol or a date order; every amount and date on
// screen, in messages, and in exports comes through these two functions.
const DEFAULT_LOCALE = (typeof navigator !== "undefined" && navigator.language) || "en-US";
let LOCALE = { locale: DEFAULT_LOCALE, currency: "USD" };
let MONEY = null, DATE = null;

function configureLocale(settings) {
  const locale = String(settings?.locale || DEFAULT_LOCALE);
  const currency = String(settings?.currency || "USD").toUpperCase();
  try { MONEY = new Intl.NumberFormat(locale, { style: "currency", currency }); }
  catch { MONEY = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }); }
  try { DATE = new Intl.DateTimeFormat(locale, { year: "numeric", month: "numeric", day: "numeric" }); }
  catch { DATE = new Intl.DateTimeFormat("en-US", { year: "numeric", month: "numeric", day: "numeric" }); }
  LOCALE = { locale, currency };
}
configureLocale();

function currencyCode() { return LOCALE.currency; }

function formatMoney(n) {
  return MONEY.format(parseFloat(n || 0));
}

// Accepts a date-only string (kept as a calendar date, no timezone shift)
// or a full timestamp; anything unreadable is shown as received.
function formatDate(value) {
  if (!value) return "";
  const str = String(value);
  const d = /^\d{4}-\d{2}-\d{2}$/.test(str) ? new Date(`${str}T00:00:00`) : new Date(str);
  return isNaN(d.getTime()) ? str : DATE.format(d);
}

// ── SOLVER ───────────────────────────────────────────────────────────
function solve(cartItems, vendors) {
  if (!cartItems.length) return [];
  let assignments = cartItems.map(item => {
    // Expired (stale-price) options are shown to the person so they know a
    // vendor's price needs refreshing, but must never be picked as the
    // "cheapest" option — their last quoted amount remains visible, but cannot be selected.
    const validOptions = item.options.filter(orderable);
    const sorted = [...validOptions].sort((a,b)=>a.price-b.price);
    const cheapestOption = sorted[0] || null;
    if (!cheapestOption) {
      // Nothing orderable: every vendor is stale, or none carries the
      // locked brand. Never fall back to a blocked option's price.
      return {...item, assignedVendorId:null, assignedVendorName:null, vendorItemId:null,
        price:0, packSize:null, orderUnit:item.orderUnit, lineTotal:0,
        cheapestPrice:0, premiumPaid:0, locked:false, unorderable:true};
    }
    // A negotiated price belongs to a specific vendor. If that vendor's
    // quote expires, NEVER transfer its negotiated amount to a fallback.
    const forced=item.forcedVendorId?validOptions.find(o=>o.vendorId===item.forcedVendorId):null;
    if(item.forcedVendorId&&!forced){
      return {...item,assignedVendorId:null,assignedVendorName:null,vendorItemId:null,
        price:null,packSize:null,lineTotal:0,cheapestPrice:cheapestOption.price,
        premiumPaid:0,locked:true,unorderable:true};
    }
    const best=forced||cheapestOption;
    const effectivePrice=item.forcedPrice!=null ? item.forcedPrice : best.price;
    const lineTotal = r2(effectivePrice*item.quantity);
    return {...item, assignedVendorId:best.vendorId, assignedVendorName:best.vendorName,
      vendorItemId:best.vendorItemId, price:effectivePrice, packSize:best.packSize,
      orderUnit:best.orderUnit, lineTotal,
      cheapestPrice:cheapestOption.price, premiumPaid:r2(Math.max(0,lineTotal-cheapestOption.price*item.quantity)),
      locked:!!item.forcedVendorId || item.forcedPrice!=null, unorderable:false};
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
          const opt=a.options.find(o=>o.vendorId===req.id&&orderable(o));
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
        const alt=[...a.options].filter(o=>o.vendorId!==req.id&&orderable(o)).sort((x,y)=>x.price-y.price)[0];
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
  const [notice,setNotice]=useState("");

  async function submit(e) {
    e.preventDefault(); setLoading(true); setError(""); setNotice("");
    try{
      if(mode==="login") await backend.session.signIn({email,password});
      else await backend.session.signUp({email,password});
      if(mode==="signup") setNotice("Check your email to confirm, then sign in.");
    }catch(err){setError(err.message||String(err));}
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
      value:"Built for audit-ready record keeping."},
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
          <button onClick={()=>{setMode("login");setError("");setNotice("");}}
            style={{flex:1,padding:"9px",borderRadius:6,border:"none",cursor:"pointer",fontWeight:700,fontSize:13,
              background:mode==="login"?"white":"transparent",color:mode==="login"?"#003584":"#888",
              boxShadow:mode==="login"?"0 1px 3px rgba(0,0,0,0.1)":"none"}}>
            Sign In
          </button>
          <button onClick={()=>{setMode("signup");setError("");setNotice("");}}
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
          {notice&&<div style={{background:"#E8F5E9",color:"#2E7D32",padding:"10px 12px",borderRadius:8,fontSize:13,marginBottom:12}}>{notice}</div>}
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
      await write(supabase.from("organization_members").insert({organization_id:org.id,user_id:user.id,role:"owner"}),"Could not add you as owner");
      const vrows=vendors.filter(v=>v.name.trim()).map(v=>({
        organization_id:org.id, name:v.name.trim(),
        delivery_minimum_dollar:v.minDollar?parseFloat(v.minDollar):null,
        delivery_minimum_units:v.minUnits?parseInt(v.minUnits):null,
      }));
      if(vrows.length) await write(supabase.from("vendors").insert(vrows),"Could not save vendors");
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
            <input style={inp} value={industry} onChange={e=>setIndustry(e.target.value)} placeholder="Your industry" />
          </div>
          <button onClick={()=>setStep(2)} disabled={!orgName.trim()} style={{...btn("#003584"),width:"100%"}}>Next →</button>
        </>}

        {step===2&&<>
          <p style={{fontSize:13,color:"#555",margin:"0 0 14px"}}>Add your vendors — you can add more later.</p>
          {vendors.map((v,i)=>(
            <div key={i} style={{background:"#F8F9FA",borderRadius:8,padding:12,marginBottom:10}}>
              <div style={{marginBottom:8}}>
                <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Vendor name</div>
                <input style={inp} value={v.name} onChange={e=>{const vv=[...vendors];vv[i].name=e.target.value;setVendors(vv);}} placeholder="Vendor name" />
              </div>
              <div style={{display:"flex",gap:8}}>
                <div style={{flex:1}}>
                  <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Min order ({currencyCode()})</div>
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
      await backend.team.acceptInvite(cleanCode,user.id);
      onComplete();
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

// Industries that have a starter category template - read from
// industry_templates, so adding an industry is a data insert, never a
// code change. Shown as quick-pick chips in Admin; the industry field
// itself stays free text for anything not templated yet.
async function loadTemplatedIndustries() {
  const { data } = await supabase.from("industry_templates").select("industry");
  return [...new Set((data || []).map(r => r.industry).filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

// Common price-refresh cadences shown as quick-pick chips in Admin - a
// shortcut for the most common choices, with a custom number field
// always available alongside for anything else (or "Never" to turn the
// whole feature off). Cadences are not tied to any industry.
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
  const [editingLocale,setEditingLocale]=useState(false);
  const [localeInput,setLocaleInput]=useState(orgSettings?.locale||DEFAULT_LOCALE);
  const [currencyInput,setCurrencyInput]=useState(orgSettings?.currency||"USD");
  const [savingLocale,setSavingLocale]=useState(false);
  const [editingRefresh,setEditingRefresh]=useState(false);
  const [refreshInput,setRefreshInput]=useState(orgSettings?.price_refresh_days!=null?String(orgSettings.price_refresh_days):"");
  const [refreshMode,setRefreshMode]=useState(orgSettings?.price_refresh_mode||(orgSettings?.price_refresh_days?"automatic":"manual"));
  const [savingRefresh,setSavingRefresh]=useState(false);
  const [templatedIndustries,setTemplatedIndustries]=useState([]);
  const [members,setMembers]=useState([]);
  const [codes,setCodes]=useState([]);
  const [loading,setLoading]=useState(true);
  const [showInvite,setShowInvite]=useState(false);
  const [inviteRole,setInviteRole]=useState("employee");
  const [newCode,setNewCode]=useState(null);
  const [error,setError]=useState("");

  async function load() {
    setLoading(true);
    const [mr,cr,industries]=await Promise.all([
      supabase.from("organization_members").select("*").eq("organization_id",orgId),
      supabase.from("invite_codes").select("*").eq("organization_id",orgId).order("created_at",{ascending:false}),
      loadTemplatedIndustries(),
    ]);
    setTemplatedIndustries(industries);
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
    setSavingOrgName(true); setError("");
    try{
      await write(supabase.from("organizations").update({name:orgNameInput.trim()}).eq("id",orgId),"Could not rename the organization");
      setEditingOrgName(false);
      onOrgUpdated();
    }catch(err){ setError(err.message); }
    setSavingOrgName(false);
  }

  async function saveIndustry(){
    setSavingIndustry(true); setError("");
    setIndustryMsg("");
    try{
      await write(supabase.from("organizations").update({industry:industryInput.trim()||null}).eq("id",orgId),"Could not save the industry");
      if(autoLoadCategories&&industryInput.trim()){
        const {added,addedVocabulary,found}=await loadStarterPackForIndustry(orgId,industryInput.trim(),categories);
        setIndustryMsg(!found?`Saved. No starter pack found for "${industryInput.trim()}" yet — add categories and vocabulary manually in Admin below.`:
          (added||addedVocabulary)?`Saved — added ${added} starter categor${added===1?"y":"ies"} and ${addedVocabulary} vocabulary term${addedVocabulary===1?"":"s"} for ${industryInput.trim()}.`:
          "Saved. This industry's starter pack was already all present.");
      }
      setEditingIndustry(false);
      onOrgUpdated();
    }catch(err){ setError(err.message); }
    setSavingIndustry(false);
  }

  // Every org sets its OWN refresh cadence - weekly price sheets and
  // monthly ones want very different windows. Nothing here assumes any
  // particular schedule; it's just a number stored per org. Merging into
  // existing settings so other keys (if any get added later) aren't wiped.
  async function saveRefreshDays(){
    setSavingRefresh(true); setError("");
    const days=refreshMode==="manual"?null:(refreshInput.trim()===""?null:Math.max(1,parseInt(refreshInput,10)||0)||null);
    if(refreshMode==="automatic"&&!days){setError("Select a validity period in days.");setSavingRefresh(false);return;}
    try{
      await write(supabase.from("organizations").update({settings:{...(orgSettings||{}),price_refresh_mode:refreshMode,price_refresh_days:days}}).eq("id",orgId),"Could not save the refresh period");
      setEditingRefresh(false);
      onOrgUpdated();
    }catch(err){ setError(err.message); }
    setSavingRefresh(false);
  }

  // Money and date format for this org. The currency list is the
  // browser's own ISO 4217 data; the locale is any BCP 47 tag (en-US,
  // en-GB, de-DE...). Merged into settings like every other org setting.
  async function saveLocale(){
    setSavingLocale(true); setError("");
    const locale=localeInput.trim()||DEFAULT_LOCALE;
    const currency=currencyInput.trim().toUpperCase()||"USD";
    try{
      new Intl.NumberFormat(locale,{style:"currency",currency});
    }catch{
      setError(`"${locale}" / "${currency}" is not a locale and currency the browser recognises.`);
      setSavingLocale(false); return;
    }
    try{
      await write(supabase.from("organizations").update({settings:{...(orgSettings||{}),locale,currency}}).eq("id",orgId),"Could not save the format settings");
      setEditingLocale(false);
      onOrgUpdated();
    }catch(err){ setError(err.message); }
    setSavingLocale(false);
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
    setError("");
    try{ await write(supabase.from("invite_codes").delete().eq("id",c.id),"Could not revoke the code"); load(); }
    catch(err){ setError(err.message); }
  }

  async function changeRole(m,newRole){
    if(m.role==="owner"&&newRole!=="owner"&&roleCounts.owner<=1){
      alert("You can't change the only owner's role — make someone else an owner first.");
      return;
    }
    setError("");
    try{ await write(supabase.from("organization_members").update({role:newRole}).eq("organization_id",orgId).eq("user_id",m.user_id),"Could not change the role"); load(); }
    catch(err){ setError(err.message); }
  }

  async function removeMember(m){
    if(m.user_id===currentUserId){ alert("You can't remove yourself from the team."); return; }
    if(m.role==="owner"&&roleCounts.owner<=1){ alert("You can't remove the only owner."); return; }
    if(!window.confirm("Remove this person from your team? They'll lose access immediately.")) return;
    setError("");
    try{ await write(supabase.from("organization_members").delete().eq("organization_id",orgId).eq("user_id",m.user_id),"Could not remove the member"); load(); }
    catch(err){ setError(err.message); }
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
      {error&&!showInvite&&<div style={{background:"#FFF3E0",color:"#E65100",padding:"10px 12px",borderRadius:8,fontSize:13,marginBottom:14}}>{error}</div>}
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
                {templatedIndustries.map(p=>(
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
                Also load this industry's starter pack (categories and vocabulary) when I save
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

          <div style={{fontSize:10,fontWeight:700,color:"#AAA",letterSpacing:"0.06em",textTransform:"uppercase",margin:"14px 0 4px"}}>Money and dates</div>
          {!editingLocale?(
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div style={{fontWeight:700,fontSize:15}}>{currencyCode()} · {LOCALE.locale} <span style={{fontWeight:400,color:"#AAA",fontSize:12}}>e.g. {formatMoney(1234.5)}, {formatDate(new Date().toISOString())}</span></div>
              <button onClick={()=>{setLocaleInput(orgSettings?.locale||DEFAULT_LOCALE);setCurrencyInput(orgSettings?.currency||"USD");setEditingLocale(true);}} style={{background:"none",border:"none",cursor:"pointer",color:"#888",fontSize:12,padding:0}}>✎ Edit</button>
            </div>
          ):(
            <div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                {(()=>{
                  const codes=(typeof Intl.supportedValuesOf==="function")?Intl.supportedValuesOf("currency"):null;
                  return codes?(
                    <select style={{...inp,width:"auto"}} value={currencyInput} onChange={e=>setCurrencyInput(e.target.value)}>
                      {codes.map(c=><option key={c} value={c}>{c}</option>)}
                    </select>
                  ):(
                    <input style={{...inp,width:110}} value={currencyInput} onChange={e=>setCurrencyInput(e.target.value)} placeholder="USD" />
                  );
                })()}
                <input style={{...inp,flex:1,minWidth:140}} value={localeInput} onChange={e=>setLocaleInput(e.target.value)} placeholder={DEFAULT_LOCALE} />
                <button onClick={()=>setEditingLocale(false)} style={{...btn("#EEE","#555",{padding:"10px 14px"})}}>Cancel</button>
                <button onClick={saveLocale} disabled={savingLocale} style={{...btn("#003584",undefined,{padding:"10px 14px"})}}>{savingLocale?"...":"Save"}</button>
              </div>
              <div style={{fontSize:11,color:"#999",marginTop:6}}>
                Currency is the ISO code. Locale sets number and date order — your browser's is {DEFAULT_LOCALE}.
              </div>
            </div>
          )}

          <div style={{fontSize:10,fontWeight:700,color:"#AAA",letterSpacing:"0.06em",textTransform:"uppercase",margin:"14px 0 4px"}}>Price refresh period</div>
          {!editingRefresh?(
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div style={{fontWeight:700,fontSize:15,color:orgSettings?.price_refresh_days?"#111":"#BBB"}}>
                {(orgSettings?.price_refresh_mode||(orgSettings?.price_refresh_days?"automatic":"manual"))==="manual"?"Manual expiration — prices remain current until explicitly expired":`${orgSettings?.price_refresh_days||"?"} days — automatic expiration`}
              </div>
              <button onClick={()=>{setRefreshInput(orgSettings?.price_refresh_days!=null?String(orgSettings.price_refresh_days):"");setRefreshMode(orgSettings?.price_refresh_mode||(orgSettings?.price_refresh_days?"automatic":"manual"));setEditingRefresh(true);}} style={{background:"none",border:"none",cursor:"pointer",color:"#888",fontSize:12,padding:0}}>✎ Edit</button>
            </div>
          ):(
            <div>
              <div style={{display:"flex",gap:8,marginBottom:10}}>
                <button onClick={()=>setRefreshMode("manual")} style={{...btn(refreshMode==="manual"?"#003584":"#EEE",refreshMode==="manual"?"white":"#555"),flex:1}}>Manual expiration</button>
                <button onClick={()=>setRefreshMode("automatic")} style={{...btn(refreshMode==="automatic"?"#003584":"#EEE",refreshMode==="automatic"?"white":"#555"),flex:1}}>Automatic expiration</button>
              </div>
              {refreshMode==="automatic"&&<>
              <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:8}}>
                <button onClick={()=>setRefreshInput("")}
                  style={{fontSize:11,fontWeight:700,padding:"5px 11px",borderRadius:14,cursor:"pointer",
                    background:refreshInput===""?"#003584":"#EEF2F8",color:refreshInput===""?"white":"#003584",
                    border:refreshInput===""?"2px solid #003584":"2px solid transparent"}}>
                  Choose a period
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
              </div></>}
              <div style={{display:"flex",gap:8,marginTop:8}}>
                <button onClick={()=>setEditingRefresh(false)} style={{...btn("#EEE","#555",{padding:"10px 14px"})}}>Cancel</button>
                <button onClick={saveRefreshDays} disabled={savingRefresh} style={{...btn("#003584",undefined,{padding:"10px 14px"})}}>{savingRefresh?"...":"Save"}</button>
              </div>
              <div style={{fontSize:11,color:"#999",marginTop:6}}>
                Automatic mode expires quotations after the selected period; manual mode expires only when you choose Expire in Price Sheets. Vendor-specified end dates always apply. Historical quotations and recorded invoices remain unchanged.
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
// client: the same panel runs for every org; the only difference is
// which rows exist in THEIR catalog_categories.

// Confirms a fuzzy-matched mapping is correct - bumps it to a full,
// exact match so it stops showing up for review.
async function confirmMapping(mappingId) {
  return write(supabase.from("item_mappings").update({comparison_track:"exact",confidence_score:100}).eq("id",mappingId),"Could not confirm the match");
}

// Re-points a mapping at a different, existing catalog item - for when
// the auto-match picked the wrong one.
async function remapToExistingItem(mappingId, newCatalogItemId) {
  return write(supabase.from("item_mappings").update({catalog_item_id:newCatalogItemId,comparison_track:"exact",confidence_score:100}).eq("id",mappingId),"Could not remap the item");
}

// Splits a mapping out into its own brand-new catalog item - for when
// the auto-match merged it into something that isn't actually the same
// product at all.
async function remapToNewItem(orgId, mappingId, description, catalogItems, categories) {
  const category=classifyCategory(description,categories)||await ensureHoldingPen(orgId,categories);
  const itemsInCategory=catalogItems.filter(ci=>ci.category_id===category?.id);
  const nextNumber=itemsInCategory.length
    ? Math.max(...itemsInCategory.map(ci=>ci.master_item_number||0))+1
    : (category?.range_start||1);
  const created=await write(supabase.from("catalog_items").insert({
    organization_id:orgId, category_id:category?.id||null, master_item_number:nextNumber,
    name:description.slice(0,120), matching_behavior:"flexible", canonical_unit:null, brand_locked:false,
  }).select().single(),"Could not create the new item");
  return write(supabase.from("item_mappings").update({catalog_item_id:created.id,comparison_track:"exact",confidence_score:100}).eq("id",mappingId),"Could not point the vendor item at it");
}

// Merges two catalog items that a person has manually identified as the
// same product (via drag-and-drop in the Item Catalog) - every vendor
// mapping pointing at the dragged (source) item gets re-pointed at the
// drop target, marked exact/100% since a human just confirmed it
// directly, and the now-empty source item is removed. The target item's
// identity (name, master number, category) survives; the source does not.
async function mergeCatalogItems(sourceCatalogItemId, targetCatalogItemId) {
  if(sourceCatalogItemId===targetCatalogItemId) return;
  const sourceMappings=await write(supabase.from("item_mappings").select("id").eq("catalog_item_id",sourceCatalogItemId),"Could not read the item's vendor links");
  for(const m of (sourceMappings||[])){
    await write(supabase.from("item_mappings").update({catalog_item_id:targetCatalogItemId,comparison_track:"exact",confidence_score:100}).eq("id",m.id),"Could not move a vendor link");
  }
  await write(supabase.from("catalog_items").delete().eq("id",sourceCatalogItemId),"Could not remove the merged item");
}

// The catalog item's name is the CLIENT's identity for that product, not
// whichever vendor's raw wording happened to trigger its creation first -
// this is what lets the client actually own that naming instead of being
// stuck with it.
async function renameCatalogItem(catalogItemId, newName) {
  return write(supabase.from("catalog_items").update({name:newName.slice(0,120)}).eq("id",catalogItemId),"Could not rename the item");
}

// The client's per-item rules: which brand (if any) this item is locked
// to, whether vendor matching must be strict, and which unit its
// per-unit price reads in. One write path for all three so the schema
// fields are set from exactly one place.
async function updateCatalogItemSettings(catalogItemId, patch) {
  return write(supabase.from("catalog_items").update(patch).eq("id", catalogItemId), "Could not save that setting");
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
  return write(supabase.from("catalog_items").insert({
    organization_id: orgId, category_id: categoryId || null, master_item_number: nextNumber,
    name: name.slice(0, 120), matching_behavior: "flexible", canonical_unit: null, brand_locked: false,
  }).select().single(), "Could not create the item");
}

// Directly links one specific vendor's item to a specific client item
// number - not limited to items the matching engine happened to flag.
// Works whether this vendor item has never been mapped at all, or
// already has a mapping that needs correcting. Marked exact/100% since
// a person just typed the exact number themselves - the most certain
// kind of match there is.
async function assignVendorItemMapping(orgId, vendorItemId, targetCatalogItemId, existingMappingId) {
  if (existingMappingId) {
    return write(supabase.from("item_mappings").update({
      catalog_item_id: targetCatalogItemId, comparison_track: "exact", confidence_score: 100, match_method: "manual",
    }).eq("id", existingMappingId), "Could not update the link");
  }
  return write(supabase.from("item_mappings").insert({
    organization_id: orgId, catalog_item_id: targetCatalogItemId, vendor_item_id: vendorItemId,
    comparison_track: "exact", confidence_score: 100, match_method: "manual",
  }), "Could not create the link");
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

// Search finds an item by the client's own name for it, by any vendor's
// wording for it, or by any vendor's item code - whichever the person
// has in front of them.
function itemMatchesSearch(item, query){
  const q=String(query||"").trim().toLowerCase();
  if(!q) return true;
  if(item.name.toLowerCase().includes(q)) return true;
  if(String(item.masterItemNumber||"")===q) return true;
  return item.options.some(o=>String(o.description||"").toLowerCase().includes(q)||String(o.vendorItemCode||"").toLowerCase()===q);
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
  const [error,setError]=useState("");

  const vMap=useMemo(()=>new Map(vendors.map(v=>[v.id,v])),[vendors]);
  const viMap=useMemo(()=>new Map(vendorItems.map(vi=>[vi.id,vi])),[vendorItems]);
  const ciMap=useMemo(()=>new Map(catalogItems.map(ci=>[ci.id,ci])),[catalogItems]);

  // Categories are ordered alphabetically - same as Order Guide, so both screens'
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
    new Set(productList.filter(item=>item.options.some(o=>["similar","review"].includes(o.matchTrack))).map(item=>item.catalogItemId)),
  [productList]);

  // Grouped by category (whatever this org's own categories are) - these headers ARE the
  // browsing structure, alphabetical top to bottom. Ordered within each
  // by whichever sortMode is active (chronological by default).
  const groupedItems=useMemo(()=>{
    const items=productList.filter(item=>{
      if(categoryFilter===REVIEW_FILTER){ if(!needsReviewItems.has(item.catalogItemId)) return false; }
      else if(categoryFilter&&item.category!==categoryFilter) return false;
      return itemMatchesSearch(item,search);
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
    const usable=item.options.filter(orderable);
    const cheapest=usable[0]||item.options[0]||null;
    if(!cheapest) return {label:"No price on file",color:"#999",bg:"#F5F5F5"};
    if(!orderable(cheapest)) return {label:blockReason(cheapest),color:"#B26A00",bg:"#FFF3E0"};
    if(["similar","review"].includes(cheapest.matchTrack)) return {label:`Needs review — ${cheapest.matchConfidence}%`,color:"#B26A00",bg:"#FFF3E0"};
    return {label:"✓ 100% matched",color:"#2E7D32",bg:"#E8F5E9"};
  }

  // One wrapper for every catalog action: show the failure where the
  // person is looking, and only refresh when the write actually landed.
  async function act(fn){
    setError("");
    try{ await fn(); onUpdated(); return true; }
    catch(err){ setError(err.message||String(err)); return false; }
  }
  async function handleConfirm(mappingId){
    setBusyMappingId(mappingId);
    await act(()=>confirmMapping(mappingId));
    setBusyMappingId(null);
  }
  async function handleRemapExisting(mappingId,newCatalogItemId){
    setBusyMappingId(mappingId);
    if(await act(()=>remapToExistingItem(mappingId,newCatalogItemId))){ setRemapOpenFor(null); setRemapSearch(""); }
    setBusyMappingId(null);
  }
  async function handleRemapNew(mappingId,description){
    setBusyMappingId(mappingId);
    if(await act(()=>remapToNewItem(orgId,mappingId,description,catalogItems,categories))){ setRemapOpenFor(null); setRemapSearch(""); }
    setBusyMappingId(null);
  }

  function handleDrop(targetItem){
    const sourceId=draggedItemId;
    setDraggedItemId(null); setDragOverItemId(null);
    if(!sourceId||sourceId===targetItem.catalogItemId) return;
    const sourceItem=productList.find(p=>p.catalogItemId===sourceId);
    if(!sourceItem) return;
    if(!window.confirm(`Merge "${sourceItem.name}" into "${targetItem.name}"?\n\nAll of "${sourceItem.name}"'s vendor prices will move under "${targetItem.name}", and "${sourceItem.name}" will be removed as its own item. This can't be undone automatically.`)) return;
    setMerging(true);
    act(()=>mergeCatalogItems(sourceId,targetItem.catalogItemId)).finally(()=>setMerging(false));
  }

  async function saveRename(catalogItemId){
    const trimmed=renameValue.trim();
    setRenamingId(null);
    if(!trimmed) return;
    await act(()=>renameCatalogItem(catalogItemId,trimmed));
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
        if(i>=0) working[i]={...working[i],category_id:newCategoryId,master_item_number:n};
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

  // Per-item rules (brand lock, strict matching, per-unit display unit).
  // A small write, then the same reload every other edit uses.
  async function handleItemSetting(catalogItemId,patch){
    await act(()=>updateCatalogItemSettings(catalogItemId,patch));
  }

  async function handleAssignCategory(catalogItemId,newCategoryId){
    setCategoryEditId(null);
    if(!newCategoryId) return;
    setCategoryEditBusy(true);
    await act(()=>assignItemCategory(catalogItemId,newCategoryId,catalogItems,categories));
    setCategoryEditBusy(false);
  }

  async function handleAddItem(){
    if(!newItemName.trim()) return;
    setAddingBusy(true);
    if(await act(()=>createClientCatalogItem(orgId,newItemName.trim(),newItemCategoryId||null,catalogItems,categories))){
      setNewItemName(""); setNewItemCategoryId(""); setAddingItem(false);
    }
    setAddingBusy(false);
  }

  const lowConfidenceMatches=useMemo(()=>
    mappings.filter(m=>["similar","review"].includes(m.comparison_track)).map(m=>{
      const vi=viMap.get(m.vendor_item_id); const ci=ciMap.get(m.catalog_item_id);
      const v=vi?vMap.get(vi.vendor_id):null;
      const others=catalogItems.filter(c=>c.id!==m.catalog_item_id);
      const suggestion=vi?bestCatalogMatch(vi.description,others):null;
      return {mappingId:m.id, catalogItemId:m.catalog_item_id, catalogName:ci?.name||"(deleted item)", vendorName:v?.name||"—",
        vendorDescription:vi?.description||"—", confidence:m.confidence_score, track:m.comparison_track,
        suggestion:suggestion?{catalogItemId:suggestion.catalogItem.id,name:suggestion.catalogItem.name,number:suggestion.catalogItem.master_item_number,score:Math.round(suggestion.score*100)}:null};
    }).sort((a,b)=>(a.confidence??0)-(b.confidence??0)),
  [mappings,viMap,ciMap,vMap,catalogItems]);

  async function handleMergeSuggested(mapping){
    if(!mapping.suggestion) return;
    if(!window.confirm(`Merge "${mapping.catalogName}" into "#${mapping.suggestion.number} ${mapping.suggestion.name}"?\n\nIts vendor price moves under that item and "${mapping.catalogName}" is removed as its own item.`)) return;
    setBusyMappingId(mapping.mappingId);
    try{
      await act(()=>mergeCatalogItems(mapping.catalogItemId,mapping.suggestion.catalogItemId));
    }finally{
      setBusyMappingId(null);
    }
  }

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

      {error&&<div style={{background:"#FFF3E0",color:"#E65100",padding:"10px 12px",borderRadius:8,fontSize:13,marginBottom:14}}>{error}</div>}

      {addingItem&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div style={{background:"white",borderRadius:12,padding:24,width:"100%",maxWidth:380}}>
            <h3 style={{margin:"0 0 4px",fontSize:16}}>Add a catalog item</h3>
            <p style={{margin:"0 0 14px",fontSize:12,color:"#888"}}>This is your item — name it however makes sense to you. Vendor prices get matched into it as they come in.</p>
            <input style={{...inp,width:"100%",marginBottom:10,boxSizing:"border-box"}} placeholder="Item name" value={newItemName} onChange={e=>setNewItemName(e.target.value)} autoFocus />
            <select style={{...inp,width:"100%",marginBottom:14,boxSizing:"border-box"}} value={newItemCategoryId} onChange={e=>setNewItemCategoryId(e.target.value)}>
              <option value="">No category (Uncategorized)</option>
              {categories.filter(c=>!c.is_holding_pen).map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
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
                <div style={{fontSize:11,color:"#999",marginTop:2,marginBottom:m.suggestion?4:8}}>{m.vendorName} — "{m.vendorDescription}"</div>
                {m.suggestion&&(
                  <div style={{fontSize:11,color:"#555",marginBottom:8}}>
                    Looks like <b>#{m.suggestion.number} {m.suggestion.name}</b> <span style={{color:"#999"}}>({m.suggestion.score}%)</span>
                  </div>
                )}
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
                      style={{...btn("#2E7D32","white",{fontSize:11,padding:"6px 12px"})}}>{busyMappingId===m.mappingId?"...":m.track==="review"?"✓ Keep as its own item":"✓ Confirm match"}</button>
                    {m.suggestion&&(
                      <button disabled={busyMappingId===m.mappingId} onClick={()=>handleMergeSuggested(m)}
                        style={{...btn("#1565C0","white",{fontSize:11,padding:"6px 12px"})}}>⇢ Merge into #{m.suggestion.number}</button>
                    )}
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

      <input style={{...inp,marginBottom:10}} placeholder="Search by name, vendor wording, or item code..." value={search} onChange={e=>setSearch(e.target.value)} />

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
                      <span key={o.vendorId} title={orderable(o)?undefined:blockReason(o)} style={{fontSize:11,background:orderable(o)?"#F5F7FA":"#FFF3E0",borderRadius:5,padding:"3px 8px"}}>
                        {o.vendorName}: <b>{formatMoney(o.casePrice)}{!orderable(o)?" · "+blockReason(o):""}</b>
                        {orderable(o)&&o.perUnit&&<span style={{color:"#888",marginLeft:4}}>{formatMoney(o.perUnit.price)}/{o.perUnit.unit}</span>}
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
                    {(()=>{
                      // Brands come from this item's own vendor options - a lock can
                      // only ever name a brand some vendor actually lists.
                      const brands=[...new Set(item.options.map(o=>String(o.brand||"").trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
                      const units=item.unitDimension?unitsForDimension(item.unitDimension):[];
                      const sel={...inp,fontSize:11,padding:"4px 6px",width:"auto"};
                      const lbl={fontSize:10,fontWeight:700,color:"#AAA",letterSpacing:"0.06em",textTransform:"uppercase"};
                      return (
                        <div style={{display:"flex",flexWrap:"wrap",gap:14,alignItems:"flex-end",background:"white",border:"1px solid #EEE",borderRadius:6,padding:"8px 10px",marginBottom:8}}>
                          <div>
                            <div style={lbl}>Brand lock</div>
                            {brands.length?(
                              <select style={sel} value={item.lockedBrand||""}
                                onChange={e=>handleItemSetting(item.catalogItemId,{brand_locked:!!e.target.value,locked_brand:e.target.value||null})}>
                                <option value="">Any brand</option>
                                {brands.map(b=><option key={b} value={b}>{b}</option>)}
                              </select>
                            ):(
                              <div style={{fontSize:11,color:"#AAA",paddingTop:4}}>No vendor lists a brand for this item</div>
                            )}
                          </div>
                          <div>
                            <div style={lbl}>Vendor matching</div>
                            <select style={sel} value={item.matchingBehavior}
                              onChange={e=>handleItemSetting(item.catalogItemId,{matching_behavior:e.target.value})}>
                              <option value="flexible">Flexible — close matches link, flagged for review</option>
                              <option value="strict">Strict — only exact matches link</option>
                            </select>
                          </div>
                          <div>
                            <div style={lbl}>Compare per</div>
                            {units.length?(
                              <select style={sel} value={item.canonicalUnit||""}
                                onChange={e=>handleItemSetting(item.catalogItemId,{canonical_unit:e.target.value||null})}>
                                <option value="">Auto ({item.displayUnit||"—"})</option>
                                {units.map(u=><option key={u} value={u}>{u}</option>)}
                              </select>
                            ):(
                              <div style={{fontSize:11,color:"#AAA",paddingTop:4}}>{item.displayUnit?`per ${item.displayUnit}`:"No readable pack size yet"}</div>
                            )}
                          </div>
                        </div>
                      );
                    })()}
                    {item.options.length===0&&<div style={{fontSize:11,color:"#AAA",marginBottom:6}}>No vendor is currently linked to this item — it'll pick one up automatically the next time a price sheet mentions it, or wait for a manual link once a vendor item exists to point at.</div>}
                    {item.options.map(o=>(
                      <div key={o.vendorItemId} style={{background:"white",borderRadius:6,padding:8,marginBottom:6,border:"1px solid #EEE"}}>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                          <div style={{fontSize:12,fontWeight:600}}>{o.vendorName}</div>
                          <div style={{fontSize:11,fontWeight:700,color:["similar","review"].includes(o.matchTrack)?"#B26A00":"#2E7D32"}}>
                            {["similar","review"].includes(o.matchTrack)?`🔍 ${o.matchConfidence}%`:"✓ exact"}
                          </div>
                        </div>
                        <div style={{fontSize:11,color:"#999",margin:"2px 0 6px"}}>"{o.description}"{o.brand?` · ${o.brand}`:""} — {formatMoney(o.casePrice)}{!orderable(o)?` · ${blockReason(o)}`:""}{orderable(o)&&o.perUnit?` (${formatMoney(o.perUnit.price)}/${o.perUnit.unit})`:""}</div>
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
  const uncategorized = holdingPen(categories);
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
  await write(supabase.from("catalog_items").update({
    category_id: target.id, master_item_number: nextNumber,
  }).eq("id", catalogItemId), "Could not move the item");
  return nextNumber;
}

function CatalogPanel({orgId,orgIndustry,categories,catalogItems,vocabulary,onUpdated}) {
  const [adding,setAdding]=useState(false);
  const [vocabKind,setVocabKind]=useState("synonym");
  const [vocabTerm,setVocabTerm]=useState("");
  const [vocabCanonical,setVocabCanonical]=useState("");
  const [vocabBusy,setVocabBusy]=useState(false);
  const [newName,setNewName]=useState("");
  const [newKeywords,setNewKeywords]=useState("");
  const [editingId,setEditingId]=useState(null);
  const [editKeywords,setEditKeywords]=useState("");
  const [loadingTemplate,setLoadingTemplate]=useState(false);
  const [templateMsg,setTemplateMsg]=useState("");
  const [reclassifying,setReclassifying]=useState(false);
  const [reclassifyMsg,setReclassifyMsg]=useState("");
  const [error,setError]=useState("");

  const uncategorizedCategory=useMemo(()=>holdingPen(categories),[categories]);

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
    setError("");
    try{
      await write(supabase.from("catalog_categories").update({keywords}).eq("id",cat.id),"Could not save keywords");
      setEditingId(null);
      onUpdated();
    }catch(err){ setError(err.message); }
  }

  async function deleteCategory(cat){
    if(itemCounts[cat.id]>0){
      alert(`Can't delete "${cat.name}" — ${itemCounts[cat.id]} catalog item(s) are still assigned to it. Reassign them first.`);
      return;
    }
    if(!window.confirm(`Delete category "${cat.name}"?`)) return;
    setError("");
    try{
      await write(supabase.from("catalog_categories").delete().eq("id",cat.id),"Could not delete the category");
      onUpdated();
    }catch(err){ setError(err.message); }
  }

  // The org's vocabulary is how the matching engine learns this trade's
  // language: a unit, a packaging word, a stopword, or a synonym. Every
  // row is data the org owns; the engine itself never changes.
  async function addVocabulary(){
    const term=vocabTerm.trim().toLowerCase();
    const canonical=vocabCanonical.trim();
    const needsCanonical=vocabKind==="unit"||vocabKind==="synonym";
    if(!term||(needsCanonical&&!canonical)) return;
    setVocabBusy(true); setError("");
    try{
      await write(supabase.from("org_vocabulary").insert({organization_id:orgId,kind:vocabKind,term,canonical:needsCanonical?canonical:null}),"Could not add the term");
      setVocabTerm(""); setVocabCanonical("");
      onUpdated();
    }catch(err){ setError(err.message); }
    setVocabBusy(false);
  }
  async function removeVocabulary(row){
    setError("");
    try{
      await write(supabase.from("org_vocabulary").delete().eq("id",row.id),"Could not remove the term");
      onUpdated();
    }catch(err){ setError(err.message); }
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
      const {added,addedVocabulary,found}=await withTimeout(loadStarterPackForIndustry(orgId,orgIndustry,categories),20000,"Load starter pack");
      setTemplateMsg(!found?`No starter pack found for "${orgIndustry}" yet — add categories and vocabulary manually below, or ask for that industry to be added.`:
        (added||addedVocabulary)?`Added ${added} starter categor${added===1?"y":"ies"} and ${addedVocabulary} vocabulary term${addedVocabulary===1?"":"s"}.`:"This industry's starter pack is already all present.");
      onUpdated();
    }catch(err){
      setError(`Load starter pack failed: ${err.message||String(err)}`);
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
          Pull the starter pack for this org's industry{orgIndustry?<> (<b>{orgIndustry}</b>)</>:""}: categories with keywords, and the vocabulary the matching engine reads. Fully editable after — nothing here is locked to any one client.
        </div>
        <button onClick={loadStarterTemplate} disabled={loadingTemplate} style={{...btn("#2E7D32","white",{fontSize:12,padding:"8px 14px"})}}>
          {loadingTemplate?"Loading...":"Load starter pack"}
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
          <p style={{color:"#888",fontSize:13,margin:0}}>No categories yet — everything imported so far lands in the holding category automatically. Add categories or load a starter pack above.</p>
        </div>
      ):categories.map(cat=>{
        const isEditing=editingId===cat.id;
        return (
          <div key={cat.id} style={{background:"white",borderRadius:8,padding:"12px 14px",marginBottom:8,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:isEditing?8:0}}>
              <div style={{fontWeight:700,fontSize:14}}>{cat.name} <span style={{fontWeight:400,fontSize:11,color:"#AAA"}}>({itemCounts[cat.id]||0} item{itemCounts[cat.id]===1?"":"s"})</span></div>
              <div style={{display:"flex",gap:10,alignItems:"center"}}>
                {!isEditing&&<button onClick={()=>{setEditingId(cat.id);setEditKeywords((cat.keywords||[]).join(", "));}} style={{background:"none",border:"none",cursor:"pointer",color:"#888",fontSize:12,padding:0}}>✎ Keywords</button>}
                {!cat.is_holding_pen&&<button onClick={()=>deleteCategory(cat)} style={{background:"none",border:"none",cursor:"pointer",color:"#E65100",fontSize:16,padding:0}} title="Delete">×</button>}
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

      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",margin:"24px 0 6px"}}>
        <h3 style={{margin:0,fontSize:16,color:"white"}}>Vocabulary</h3>
      </div>
      <div style={{fontSize:12,color:"rgba(255,255,255,0.7)",marginBottom:10}}>
        How the matching engine reads this trade's language. Units and packaging words are kept out of product names; synonyms make two spellings the same word; stopwords carry no meaning.
      </div>
      <div style={{background:"white",borderRadius:10,padding:14,marginBottom:12,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
        {(()=>{
          const KINDS=[["synonym","Synonym","chix","chicken"],["unit","Unit","sheets","SHEET"],["packaging","Packaging word","bundle",""],["stopword","Stopword","premium",""]];
          const needsCanonical=vocabKind==="unit"||vocabKind==="synonym";
          const [,,exTerm,exCanon]=KINDS.find(k=>k[0]===vocabKind)||KINDS[0];
          return (
            <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
              <select style={{...inp,width:"auto"}} value={vocabKind} onChange={e=>setVocabKind(e.target.value)}>
                {KINDS.map(([k,label])=><option key={k} value={k}>{label}</option>)}
              </select>
              <input style={{...inp,flex:1,minWidth:120}} placeholder={`Term, e.g. ${exTerm}`} value={vocabTerm} onChange={e=>setVocabTerm(e.target.value)}
                onKeyDown={e=>{if(e.key==="Enter") addVocabulary();}} />
              {needsCanonical&&<span style={{color:"#AAA"}}>→</span>}
              {needsCanonical&&<input style={{...inp,flex:1,minWidth:120}} placeholder={vocabKind==="unit"?`Unit code, e.g. ${exCanon}`:`Means, e.g. ${exCanon}`} value={vocabCanonical} onChange={e=>setVocabCanonical(e.target.value)}
                onKeyDown={e=>{if(e.key==="Enter") addVocabulary();}} />}
              <button onClick={addVocabulary} disabled={vocabBusy||!vocabTerm.trim()||(needsCanonical&&!vocabCanonical.trim())} style={{...btn("#003584","white",{fontSize:12,padding:"9px 14px"})}}>{vocabBusy?"...":"Add"}</button>
            </div>
          );
        })()}
      </div>
      {(vocabulary||[]).length===0?(
        <div style={{background:"white",borderRadius:10,padding:16,textAlign:"center",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
          <p style={{color:"#888",fontSize:13,margin:0}}>No vocabulary yet — load the starter pack above, or add terms as you meet them on price sheets.</p>
        </div>
      ):["synonym","unit","packaging","stopword"].map(kind=>{
        const rows=(vocabulary||[]).filter(v=>v.kind===kind).sort((a,b)=>a.term.localeCompare(b.term));
        if(!rows.length) return null;
        const title={synonym:"Synonyms",unit:"Units",packaging:"Packaging words",stopword:"Stopwords"}[kind];
        return (
          <div key={kind} style={{background:"white",borderRadius:8,padding:"10px 14px",marginBottom:8,boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
            <div style={{fontSize:10,fontWeight:700,color:"#AAA",letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:6}}>{title} ({rows.length})</div>
            <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
              {rows.map(v=>(
                <span key={v.id} style={{fontSize:12,background:"#F5F7FA",borderRadius:14,padding:"4px 10px",display:"inline-flex",alignItems:"center",gap:6}}>
                  {v.term}{v.canonical?<span style={{color:"#888"}}>→ {v.canonical}</span>:null}
                  <button onClick={()=>removeVocabulary(v)} title="Remove" style={{background:"none",border:"none",cursor:"pointer",color:"#E65100",fontSize:13,padding:0,lineHeight:1}}>×</button>
                </span>
              ))}
            </div>
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
    setError("");
    try{
      await write(supabase.from("vendors").update({is_active:false}).eq("id",vendor.id),"Could not remove the vendor");
      onUpdated();
      onBack();
    }catch(err){ setError(err.message); }
  }

  async function savePriceEdit(item){
    const newPrice=parseFloat(priceEditValue);
    if(isNaN(newPrice)||newPrice<=0){ setPriceEditId(null); return; }
    setError("");
    try{
      if(Math.abs((item.price||0)-newPrice)>0.001){
        await write(supabase.from("price_history").insert({vendor_item_id:item.id,organization_id:orgId,price:newPrice,source:"manual_edit",effective_date:new Date().toISOString(),source_description:item.description,source_line:`Manual price confirmation: ${item.description} — ${newPrice}`}),"Could not record the price change");
        await write(supabase.from("vendor_items").update({price:newPrice,last_updated:new Date().toISOString(),price_source:"manual_edit",price_expired_at:null,price_quote_valid_until:null,price_unavailable:!item.pack_size}).eq("id",item.id),"Could not save the price");
      }
      setPriceEditId(null);
      onUpdated();
    }catch(err){ setError(err.message); }
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
    try{
      await assignVendorItemMapping(orgId,vendorItem.id,target.id,mapping?.id||null);
      setMapEditId(null); setMapEditValue("");
      onUpdated();
    }catch(err){ setMapError(err.message); }
    setMapBusy(false);
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
      {error&&!editing&&<div style={{background:"#FFF3E0",color:"#E65100",padding:"10px 12px",borderRadius:8,fontSize:13,marginBottom:14}}>{error}</div>}
      <div style={{background:"white",borderRadius:12,padding:20,marginBottom:14,boxShadow:"0 1px 3px rgba(0,0,0,0.08)"}}>
        {!editing?(<>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}>
            <div>
              <div style={{fontWeight:900,fontSize:20,color:vc.accent}}>{vendor.name}</div>
              <div style={{fontSize:12,color:"#888",marginTop:2}}>{vendorItems.filter(vi=>vi.vendor_id===vendor.id).length} items · Min {formatMoney(vendor.delivery_minimum_dollar)} · {vendor.delivery_minimum_units||0} units</div>
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
              <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Min order ({currencyCode()})</div>
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
                        <span style={{fontWeight:700,fontSize:13}}>{formatMoney(item.price)}</span>
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
          <div style={{fontSize:12,color:"#888"}}>{formatDate(inv.invoice_date)||formatDate(inv.created_at)}{inv.invoice_number?` · #${inv.invoice_number}`:""}</div>
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
                <div style={{fontSize:12,color:"#888"}}>{formatDate(po.created_at)} · {lines.length} item{lines.length===1?"":"s"}</div>
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
                      <div style={{fontWeight:700}}>{formatMoney(line.line_total)}</div>
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
                <div style={{fontWeight:700,fontSize:13}}>Quoted {formatDate(period.date)}</div>
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
                      <div style={{textAlign:"right"}}><div style={{fontWeight:700}}>{formatMoney(entry.price)}</div>
                                  {entry.quote_valid_until&&<div style={{fontSize:10,color:"#999"}}>Vendor valid through {formatDate(entry.quote_valid_until)}</div>}
                                  {entry.source_file_name&&<div style={{fontSize:10,color:"#999"}}>{entry.source_file_name}</div>}
                                  {entry.source_document_id&&<button onClick={()=>viewSourceDocument(entry.source_document_id)} style={{background:"none",border:"none",fontSize:10,color:"#003584",cursor:"pointer"}}>Original source ↗</button>}
                                  {!entry.source_document_id&&entry.source_file_path&&<button onClick={()=>viewStoredFile(entry.source_file_path)} style={{background:"none",border:"none",fontSize:10,color:"#003584",cursor:"pointer"}}>Original ↗</button>}
                                </div>
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
          <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Total amount ({currencyCode()})</div>
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
          <input style={inp} value={name} onChange={e=>setName(e.target.value)} placeholder="Vendor name" />
        </div>
        <div style={{marginBottom:10}}>
          <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Email — where to send orders</div>
          <input style={inp} type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="orders@vendor.com" />
        </div>
        <div style={{display:"flex",gap:8,marginBottom:16}}>
          <div style={{flex:1}}>
            <div style={{fontSize:12,fontWeight:600,color:"#666",marginBottom:4}}>Min order ({currencyCode()})</div>
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
function PasteModal({vendors,orgId,orgSettings,catalogItems,categories,onClose,onDone,initialVendorId,initialMode}) {
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
  const [acceptedIssues,setAcceptedIssues]=useState(new Set());
  const [saveReview,setSaveReview]=useState("");

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
      const parsed=parseDocument(d.text);
      return {...d,rows:parsed.rows,skipped:parsed.skipped,documentKind:parsed.documentKind,quoteValidUntil:parsed.quoteValidUntil,invoiceDate:findDate(d.text)||""};
    });
    setParsedGroups(groups);
    setStep(2);
  }

  const allRows=parsedGroups.flatMap(g=>g.rows.map(row=>({...row,_source:g.name})));
  const allSkipped=parsedGroups.flatMap(g=>g.skipped.map(s=>({...s,_source:g.name})));
  const allIncomplete=allRows.filter(r=>r.priceUnavailable);
  const showSourceTags=parsedGroups.length>1;
  const unsafeDocuments=parsedGroups.filter(g=>g.documentKind==="mixed"||(g.documentKind!=="unknown"&&g.documentKind!==mode));
  const missingInvoiceDates=mode==="invoice"?parsedGroups.filter(g=>g.rows.length&&!/^\d{4}-\d{2}-\d{2}$/.test(g.invoiceDate||"")):[];
  const needsReview=parsedGroups.flatMap(g=>g.rows.map((row,index)=>({groupId:g.id,index,row})))
    .filter(x=>x.row.issues?.length&&!acceptedIssues.has(`${x.groupId}:${x.index}`));
  function updateParsedRow(groupId,index,patch){
    // Editing an ambiguous row invalidates a prior approval of its old value.
    setAcceptedIssues(prev=>{const next=new Set(prev);next.delete(`${groupId}:${index}`);return next;});
    setParsedGroups(groups=>groups.map(g=>g.id!==groupId?g:{...g,rows:g.rows.map((r,i)=>i===index?{...r,...patch}:r)}));
  }
  function approveIssue(groupId,index){setAcceptedIssues(prev=>new Set([...prev,`${groupId}:${index}`]));}

  async function doSave(){
    if(missingInvoiceDates.length){setSaveReview("Confirm the invoice date for each document before recording invoice charges.");return;}
    if(unsafeDocuments.length){setSaveReview("This file identifies itself as an invoice and/or a price update inconsistent with the selected import mode. Split mixed documents, or select the correct mode before saving.");return;}
    if(needsReview.length){setSaveReview(`${needsReview.length} ambiguous row(s) still need a deliberate correction or approval. No unreviewed amount will change a current vendor quote.`);return;}
    setSaveReview("");setLoading(true);
    const vendor=vendors.find(v=>v.id===vendorId);
    let updated=0,created=0,invoiceTotal=0,invoicesCreated=0,mapped=0;
    let saveError=null;

    if(mode==="pricelist"){
      const workingCatalogItems=[...catalogItems];
      const workingCategories=[...categories];
      const importBatchTime=new Date().toISOString();

      let failedRows=0;
      for(const group of parsedGroups){
        if(!group.rows.length)continue;
        let sourceFilePath=null;
        // Stable SHA-256 of the original extracted text: re-importing the
        // same vendor quotation cannot silently append duplicate history.
        const digest=await crypto.subtle.digest("SHA-256",new TextEncoder().encode(group.text));
        const fingerprint=Array.from(new Uint8Array(digest)).map(n=>n.toString(16).padStart(2,"0")).join("");
        const existingDoc=await supabase.from("import_documents").select("id,status")
          .eq("organization_id",orgId).eq("vendor_id",vendorId).eq("document_kind","pricelist")
          .eq("fingerprint",fingerprint).maybeSingle();
        if(existingDoc.error){failedRows+=group.rows.length;saveError=(saveError?saveError+" ":"")+`Could not verify whether ${group.name} was already imported: ${existingDoc.error.message}`;continue;}
        if(existingDoc.data){saveError=(saveError?saveError+" ":"")+`${group.name}: previously imported (${existingDoc.data.status}); repeated upload blocked to avoid duplicate pricing. Review existing import before retrying.`;continue;}
        if(group.file){
          const uploaded=await uploadOriginalFile(orgId,vendorId,group.file);
          if(uploaded.error){failedRows+=group.rows.length;saveError=(saveError?saveError+" ":"")+`Could not preserve the original ${group.name}: ${uploaded.error.message}. Prices from that file were not applied.`;continue;}
          sourceFilePath=uploaded.path;
        }
        const storedDoc=await supabase.from("import_documents").insert({
          organization_id:orgId,vendor_id:vendorId,document_kind:"pricelist",fingerprint,
          original_text:group.text,file_name:group.name,file_path:sourceFilePath,status:"processing",
        }).select("id").single();
        if(storedDoc.error){failedRows+=group.rows.length;saveError=(saveError?saveError+" ":"")+`Could not preserve ${group.name} before accepting prices: ${storedDoc.error.message}`;continue;}
        const sourceDocumentId=storedDoc.data.id;
        const failuresBeforeGroup=failedRows;
      for(const row of group.rows){
        try{
        if(!row.priceUnavailable&&(!Number.isFinite(Number(row.price))||Number(row.price)<=0)) throw new Error("No confirmed positive unit price");
        const q=row.code
          ?supabase.from("vendor_items").select("*").eq("organization_id",orgId).eq("vendor_id",vendorId).eq("vendor_item_code",row.code).maybeSingle()
          :supabase.from("vendor_items").select("*").eq("organization_id",orgId).eq("vendor_id",vendorId).eq("description",row.description).maybeSingle();
        let ex=await write(q,"Could not look up the item");
        if(!ex&&!row.code){
          const candidates=await write(supabase.from("vendor_items").select("*")
            .eq("organization_id",orgId).eq("vendor_id",vendorId),"Could not inspect existing vendor products");
          const established=candidates.filter(vi=>compareProductIdentity(row.description,vi.description).status==="same" &&
            (!row.packSize||!vi.pack_size||packsEquivalent(row.packSize,vi.pack_size)));
          if(established.length===1) ex=established[0];
          else if(established.length>1) throw new Error("Multiple existing vendor products have this identity. Verify the vendor item code before updating a quote.");
        }
        let vendorItemId;
        if(ex){
          const identity=compareProductIdentity(row.description,ex.description);
          if(identity.status!=="same") throw new Error(`Vendor item code or description points to an unverified product (${identity.reason}). Existing price and mapping were preserved.`);
          if(row.packSize&&ex.pack_size&&!packsEquivalent(row.packSize,ex.pack_size))
            throw new Error("The pack size changed for the same vendor item; verify the package and purchasing price before updating the quote.");
        }
        // The provider persists current quote and history in one transaction.
        // It is impossible to update one and lose the other midway through.
        vendorItemId=await backend.pricing.applyQuote({
          vendorItemId:ex?.id||null,organizationId:orgId,vendorId,
          vendorItemCode:row.code,description:row.description,
          packSize:row.packSize||ex?.pack_size||null,price:row.price,
          priceUnavailable:!!row.priceUnavailable,effectiveDate:importBatchTime,
          quoteValidUntil:group.quoteValidUntil,sourceFilePath,
          sourceFileName:group.name,sourceLine:row.sourceLine||null,sourceDocumentId,
        });
        if(ex) updated++; else created++;

        // Whether this vendor item is brand new or was just updated, it
        // must be linked to a catalog item: an unlinked price is invisible
        // to ordering and to cross-vendor comparison.
        if(vendorItemId){
          const existingMapping=await write(supabase.from("item_mappings").select("id").eq("organization_id",orgId).eq("vendor_item_id",vendorItemId).maybeSingle(),"Could not check the catalog link");
          if(!existingMapping){
            const match=await matchOrCreateCatalogItem(orgId,row.description,workingCatalogItems,workingCategories);
            if(match){
              await write(supabase.from("item_mappings").insert({
                organization_id:orgId, catalog_item_id:match.catalogItemId, vendor_item_id:vendorItemId,
                confidence_score:Math.round((match.score??0)*100),
                match_method:"rule_based", comparison_track:match.track,
              }),"Could not link the item to your catalog");
              mapped++;
            }
          }
        }
        }catch(err){
          // Report the first failure verbatim and count the rest; a row
          // that failed is never counted as updated or created.
          failedRows++;
          if(!saveError) saveError=`"${row.description}" — ${err.message||String(err)}`;
        }
      }
        const finalStatus=failedRows===failuresBeforeGroup?"complete":"partial";
        const finished=await supabase.from("import_documents").update({status:finalStatus}).eq("id",sourceDocumentId);
        if(finished.error)saveError=(saveError?saveError+" ":"")+`Could not finalize source record for ${group.name}: ${finished.error.message}`;
      }
      if(failedRows>1) saveError+=` (and ${failedRows-1} more row${failedRows===2?"":"s"} failed the same way)`;
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
      const {data:vendorItemsForMatch}=await supabase.from("vendor_items").select("*").eq("organization_id",orgId).eq("vendor_id",vendorId);
      const viList=vendorItemsForMatch||[];
      // Same working-copy pattern as price-sheet import, so an item that
      // genuinely matches nothing on file still ends up in the catalog
      // and orderable, instead of vanishing into a permanent "no match".
      const workingCatalogItems=[...catalogItems];
      const workingCategories=[...categories];

      for(const group of parsedGroups){
        if(!group.rows.length) continue;
        const groupTotal=group.rows.reduce((s,row)=>s+(row.amount!=null?row.amount:row.price),0);
        const invoiceNo=findInvoiceNumber(group.text);
        const duplicate=invoiceNo
          ?await supabase.from("invoices").select("id").eq("organization_id",orgId).eq("vendor_id",vendorId).eq("invoice_number",invoiceNo).limit(1)
          :await supabase.from("invoices").select("id").eq("organization_id",orgId).eq("vendor_id",vendorId).eq("raw_text",group.text).limit(1);
        if(duplicate.error){saveError=(saveError?saveError+" ":"")+`Could not check existing invoice: ${duplicate.error.message}`;continue;}
        if(duplicate.data?.length){saveError=(saveError?saveError+" ":"")+`Invoice ${invoiceNo||group.name} is already recorded; duplicate skipped.`;continue;}

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
          let matched=null, confidence=null, method=null,codeConflict=false,pendingCatalog=null;
          if(row.code){
            matched=viList.find(vi=>vi.vendor_item_code===row.code)||null;
            if(matched&&compareProductIdentity(row.description,matched.description).status==="same"){
              confidence=100; method="code";
            }else if(matched){matched=null;codeConflict=true;method="identity_conflict";}
          }
          if(!matched&&!codeConflict&&row.description){
            matched=viList.find(vi=>vi.description===row.description)||null;
            if(matched){ confidence=100; method="exact_description"; }
          }
          if(!matched&&!codeConflict&&row.description&&viList.length){
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
          if(!matched&&!codeConflict&&row.description){
            try{
              pendingCatalog=await matchOrCreateCatalogItem(orgId,row.description,workingCatalogItems,workingCategories);
              method="created_from_invoice"; confidence=null;
            }catch(err){
              method="unmatched";
              saveError=(saveError?saveError+" ":"")+`"${row.description}": ${err.message||String(err)}`;
            }
          }
          if(codeConflict)saveError=(saveError?saveError+" ":"")+`Invoice line ${row.description}: item code conflicts with the previously identified product, so the original invoice line was recorded without a product link.`;
          let quotedPrice=null;
          if(matched?.id&&method!=="created_from_invoice"){
            const invoiceDate=group.invoiceDate;
            const {data:historical,error:histErr}=await supabase.from("price_history")
              .select("price,effective_date,quote_valid_until,source")
              .eq("organization_id",orgId).eq("vendor_item_id",matched.id)
              .lte("effective_date",`${invoiceDate}T23:59:59.999Z`)
              .order("effective_date",{ascending:false}).limit(1);
            if(histErr) saveError=(saveError?saveError+" ":"")+`Unable to verify historic quote for ${row.description}: ${histErr.message}`;
            const quote=historical?.[0];
            const withinVendorTerm=!quote?.quote_valid_until||quote.quote_valid_until>=invoiceDate;
            const refreshDays=Number(orgSettings?.price_refresh_days);
            const mode=orgSettings?.price_refresh_mode||(refreshDays>0?"automatic":"manual");
            const withinClientWindow=mode!=="automatic"||!refreshDays||
              ((new Date(`${invoiceDate}T12:00:00Z`)-new Date(quote?.effective_date||0))/86400000)<=refreshDays;
            if(quote&&withinVendorTerm&&withinClientWindow&&["price_list","manual_edit"].includes(quote.source))quotedPrice=Number(quote.price);
          }
          const variance=quotedPrice!=null?r2(row.price-quotedPrice):null;
          linesToInsert.push({
            vendor_item_id:matched?.id||null, vendor_item_code:row.code, description:row.description,
            unit_price:row.price, line_total:(row.amount!=null?row.amount:row.price),
            price_variance:variance, match_confidence:confidence, match_method:method,
            create_vendor_item:!matched&&!codeConflict&&method==="created_from_invoice",
            pack_size:row.packSize||null,catalog_item_id:pendingCatalog?.catalogItemId||null,
            catalog_confidence:pendingCatalog?Math.round((pendingCatalog.score??0)*100):null,
            catalog_track:pendingCatalog?.track||null,
          });
        }
        try{
          await backend.invoices.record({
            organization_id:orgId,vendor_id:vendorId,total_amount:r2(groupTotal),
            raw_text:group.text,invoice_date:group.invoiceDate,invoice_number:invoiceNo,
            file_path:filePath,file_name:fileName,
          },linesToInsert);
          invoicesCreated++;
          invoiceTotal+=groupTotal;
        }catch(error){
          if(filePath){try{await documents.remove([filePath]);}catch{}}
          saveError=(saveError?saveError+" ":"")+`"${group.name}" couldn't be recorded: ${error.message||String(error)}`;
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
              <input type="file" multiple accept=".csv,.txt,.tsv,.xlsx,.xls,.pdf,.eml,.html,.htm"
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
          {mode==="invoice"&&parsedGroups.map(g=><label key={g.id} style={{display:"block",fontSize:12,marginBottom:7}}>
            Invoice date for {g.name}: <input type="date" value={g.invoiceDate||""} onChange={e=>setParsedGroups(groups=>groups.map(x=>x.id===g.id?{...x,invoiceDate:e.target.value}:x))} />
            {!g.invoiceDate&&<span style={{color:"#B26A00"}}> Required</span>}
          </label>)}
          {unsafeDocuments.length>0&&<div style={{background:"#FFEBEE",color:"#B71C1C",padding:10,marginBottom:10,fontSize:12}}>Document type mismatch or mixed invoice and price update: {unsafeDocuments.map(g=>g.name).join(", ")}. Separate the sections, or switch the import type. Nothing from these documents will be saved until resolved.</div>}
          <div style={{maxHeight:340,overflowY:"auto",marginBottom:14}}>
            {parsedGroups.flatMap(g=>g.rows.map((row,i)=>(
              <div key={`${g.id}:${i}`} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:"1px solid #F0F0F0",fontSize:13}}>
                <div>
                  {row.code&&<span style={{color:"#888",marginRight:8,fontFamily:"monospace",fontSize:11}}>{row.code}</span>}
                  <span>{row.description}</span>
                  <input aria-label="Verify or correct product description" value={row.description} onChange={e=>updateParsedRow(g.id,i,{description:e.target.value})}
                    style={{...inp,fontSize:11,padding:"4px 6px",marginTop:5}} />
                  {row.packSize&&<span style={{color:"#AAA",marginLeft:6,fontSize:11}}>{row.packSize}</span>}
                  {showSourceTags&&<span style={{color:"#BBB",marginLeft:6,fontSize:10}}>· {g.name}</span>}
                  {!!row.issues?.length&&<div style={{color:"#B26A00",fontSize:11,marginTop:3}}>
                    ⚠ {row.issues.join("; ")}
                    <button onClick={()=>approveIssue(g.id,i)} disabled={acceptedIssues.has(`${g.id}:${i}`)} style={{marginLeft:8,border:"none",background:"#FFF3E0",color:"#9A5700",cursor:"pointer",fontWeight:700}}>{acceptedIssues.has(`${g.id}:${i}`)?"Acknowledged":"Approve as entered"}</button>
                  </div>}
                  {mode==="pricelist"&&!row.packSize&&<input aria-label="Confirm the purchasing pack" placeholder="Confirm pack: e.g. 40 LB" value={row.packSize||""} onChange={e=>updateParsedRow(g.id,i,{packSize:e.target.value})} style={{...inp,fontSize:11,padding:"4px 6px",marginTop:5}} />}
                </div>
                {row.priceUnavailable?(
                  <span style={{fontWeight:700,flexShrink:0,marginLeft:8,color:"#B26A00",fontSize:11}}>no price listed</span>
                ):(
                  <span style={{fontWeight:700,flexShrink:0,marginLeft:8}}>{formatMoney(row.price)}</span>
                )}
              </div>
            )))}
          </div>
          {allIncomplete.length>0&&(
            <div style={{background:"#FFF3E0",padding:"10px 14px",borderRadius:8,marginBottom:14,fontSize:12}}>
              <div style={{fontWeight:700,color:"#B26A00"}}>{allIncomplete.length} product{allIncomplete.length>1?"s":""} listed with no price — a confirmed quote is still needed</div>
              <div style={{color:"#996600",marginTop:2}}>The last-known amount remains in history, but an unavailable quote cannot be selected for ordering.</div>
            </div>
          )}
          {allSkipped.length>0&&(
            <div style={{background:"#FFF3E0",padding:"10px 14px",borderRadius:8,marginBottom:14,fontSize:12}}>
              <div style={{fontWeight:700,color:"#E65100",marginBottom:6}}>{allSkipped.length} line{allSkipped.length>1?"s":""} weren't interpreted as product lines — retained in the original source</div>
              <div style={{maxHeight:120,overflowY:"auto"}}>
                {allSkipped.map((s,i)=>(
                  <div key={i} style={{color:"#999",marginBottom:3,fontFamily:"monospace",fontSize:11}}>
                    "{s.line.slice(0,60)}" — {s.reason}{showSourceTags?` (${s._source})`:""}
                  </div>
                ))}
              </div>
            </div>
          )}
          {saveReview&&<div style={{color:"#B71C1C",fontSize:12,marginBottom:8}}>{saveReview}</div>}
          <div style={{display:"flex",gap:8}}>
            <button onClick={()=>setStep(1)} style={{...btn("#EEE","#555"),flex:1}}>← Back</button>
            <button onClick={doSave} disabled={loading||!allRows.length||unsafeDocuments.length>0||needsReview.length>0||missingInvoiceDates.length>0} style={{...btn("#003584"),flex:2}}>
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
  const [organizations,setOrganizations]=useState([]);
  const [vendors,setVendors]=useState([]);
  const [catalogItems,setCatalogItems]=useState([]);
  const [categories,setCategories]=useState([]);
  const [vendorItems,setVendorItems]=useState([]);
  const [mappings,setMappings]=useState([]);
  const [invoices,setInvoices]=useState([]);
  const [purchaseOrders,setPurchaseOrders]=useState([]);
  const [priceHistory,setPriceHistory]=useState([]);
  const [priceHistoryHasMore,setPriceHistoryHasMore]=useState(false);
  const [loadingOlderPrices,setLoadingOlderPrices]=useState(false);
  const [vocabulary,setVocabulary]=useState([]);
  const [quantities,setQuantities]=useState({}); // {catalogItemId_case: n, catalogItemId_each: n}
  const [offline,setOffline]=useState(null);
  const [restoredBasketKey,setRestoredBasketKey]=useState(null);
  const [dragItem,setDragItem]=useState(null);
  const [dropTarget,setDropTarget]=useState(null);
  const [dropNotice,setDropNotice]=useState("");
  const [tab,setTab]=useState(()=>{
    try { return sessionStorage.getItem("kerdos.activeTab") || "order"; }
    catch { return "order"; }
  });
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
  const [clockTick,setClockTick]=useState(0);
  useEffect(()=>{const timer=setInterval(()=>setClockTick(n=>n+1),60000);return ()=>clearInterval(timer);},[]);
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

  // Preserve the current workspace on a normal browser refresh. A genuine
  // sign-in still intentionally starts on Order Guide via sessionController.
  useEffect(()=>{
    try { sessionStorage.setItem("kerdos.activeTab", tab); } catch {}
  },[tab]);

  useEffect(()=>{
    if(session===undefined) return;
    if(!session){setLoading(false);return;}
    supabase.from("profiles").upsert({id:session.user.id,email:session.user.email,updated_at:new Date().toISOString()})
      .then(({error})=>{ if(error) console.error("Profile sync failed:",error.message); });
    loadData();
  },[session]);

  useEffect(()=>{
    const reconnect=()=>{ if(session) loadData(); };
    window.addEventListener("online",reconnect);
    return ()=>window.removeEventListener("online",reconnect);
  },[session,org?.id]);

  async function loadData(requestedOrganizationId=null){
    setLoading(true);
    const membershipKey=`memberships:${session.user.id}`;
    let mem, usingLocal=null;
    try{
      mem=await backend.workspace.memberships(session.user.id);
      void saveSnapshot(membershipKey,mem);
    }catch(err){
      const kept=await loadSnapshot(membershipKey);
      if(!kept){
        setLoading(false);
        alert(err.message||String(err));
        return;
      }
      mem=kept.value;
      usingLocal={since:kept.savedAt};
    }
    if(!mem?.length){setOrganizations([]);setOrg(null);setLoading(false);return;}
    const available=mem.map(m=>({...m.organizations,role:m.role}));
    setOrganizations(available);
    let savedId=null;
    try{savedId=sessionStorage.getItem("kerdos.organizationId");}catch{}
    const selectedId=requestedOrganizationId||org?.id||savedId;
    const o=available.find(candidate=>candidate.id===selectedId)||available[0];
    try{sessionStorage.setItem("kerdos.organizationId",o.id);}catch{}
    configureLocale(o.settings);
    setOrg(o);
    if(o.logo_url) getSignedUrl(o.logo_url).then(setLogoUrl); else setLogoUrl(null);
    const id=o.id;
    const snapshotKey=`snapshot:${id}`;
    let snapshot;
    try{
      snapshot=await backend.workspace.snapshot(id);
      void saveSnapshot(snapshotKey,snapshot);
      usingLocal=null;
    }
    catch(error){
      const kept=await loadSnapshot(snapshotKey);
      if(!kept){
        setLoading(false);
        alert(`KERDOS kept the existing workspace because it could not refresh data: ${error.message}`);
        return;
      }
      snapshot=kept.value;
      usingLocal={since:kept.savedAt};
    }
    setOffline(usingLocal);
    // The engine reads this org's vocabulary from here on - before any
    // matching, parsing, or per-unit pricing in this session runs.
    configureVocabulary(snapshot.vocabulary);
    setVocabulary(snapshot.vocabulary);
    setVendors(snapshot.vendors);
    setCatalogItems(snapshot.catalogItems);
    setCategories(snapshot.categories);
    setVendorItems(snapshot.vendorItems);
    setMappings(snapshot.mappings);
    setInvoices(snapshot.invoices);
    setPurchaseOrders(snapshot.purchaseOrders);
    setPriceHistory(snapshot.priceHistory);
    setPriceHistoryHasMore(snapshot.priceHistory.length===2000);
    setLoading(false);
  }

  async function switchOrganization(organizationId){
    if(!organizationId||organizationId===org?.id)return;
    setQuantities({});setVendorOverride({});setPriceOverride({});setUnitSelection({});
    await loadData(organizationId);
  }

  useEffect(()=>{
    if(!org) return;
    return backend.realtime.subscribeToOrganization(org.id,loadData);
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
    let linked=0, failed=0, firstError=null;
    for(const vi of unmapped){
      try{
        const match=await matchOrCreateCatalogItem(org.id,vi.description,workingCatalogItems,workingCategories);
        if(!match) continue;
        await write(supabase.from("item_mappings").insert({
          organization_id:org.id, catalog_item_id:match.catalogItemId, vendor_item_id:vi.id,
          confidence_score:Math.round((match.score??0)*100),
          match_method:"rule_based", comparison_track:match.track,
        }),"Could not link the item");
        linked++;
      }catch(err){ failed++; if(!firstError) firstError=err.message||String(err); }
    }
    if(failed) alert(`Linked ${linked} item${linked===1?"":"s"}; ${failed} could not be linked. First error: ${firstError}`);
    await loadData();
    setBackfilling(false);
  }

  const productList=useMemo(()=>{
    if(!catalogItems.length) return [];
    const viMap=new Map(vendorItems.map(vi=>[vi.id,vi]));
    const vMap=new Map(vendors.map(v=>[v.id,v]));
    const penName=holdingPen(categories)?.name||"Uncategorized";

    return catalogItems.map(ci=>{
      const ciMappings=mappings.filter(m=>m.catalog_item_id===ci.id);
      const lockedBrand=ci.brand_locked?(ci.locked_brand||null):null;
      const options=ciMappings.map(m=>{
        const vi=viMap.get(m.vendor_item_id);
        const v=vi?vMap.get(vi.vendor_id):null;
        if(!vi||!v||!vi.price) return null;
        // Preserve the actual quoted price even when expired. Block it from
        // ordering and comparison; never turn a historical quote into $0.
        const quote=quoteStatus(vi,org?.settings||{});
        const expired=quote==="expired";
        const invoiceOnly=quote==="invoice_only";
        const price=parseFloat(vi.price);
        const pack=vi.pack_size;
        const each=quote==="current"?eachPrice(price,pack):null;
        // A locked brand blocks every vendor item that is a different
        // brand - and one with no brand listed, since "unknown" cannot
        // be verified as the locked brand.
        const brandMismatch=!!lockedBrand&&!brandsMatch(vi.brand,lockedBrand);
        return {
          vendorId:v.id, vendorName:v.name,
          vendorItemId:vi.id, vendorItemCode:vi.vendor_item_code,
          brand:vi.brand, packSize:pack, description:vi.description,
          casePrice:price,
          eachPrice:each?.price||null, eachSize:each?.size||null,
          mappingId:m.id,
          matchConfidence:m.confidence_score, matchTrack:m.comparison_track,
          expired,
          priceUnavailable:quote==="unavailable",
          invoiceOnly,
          unverified:["similar","review"].includes(m.comparison_track),
          brandMismatch,
        };
      }).filter(Boolean).sort((a,b)=>{
        // Blocked options (stale price, wrong brand) always sink to the
        // bottom regardless of price, so they can never look "cheapest".
        const ab=!orderable(a), bb=!orderable(b);
        if(ab!==bb) return ab?1:-1;
        return a.casePrice-b.casePrice;
      });

      // Per-unit price for the honest cross-vendor comparison. Shown in
      // the unit the client chose for this item (canonical_unit), or,
      // until they choose one, the unit of the first readable pack.
      const firstPack=options.map(o=>parsePackSize(o.packSize)).find(p=>p?.parsed);
      const displayUnit=ci.canonical_unit||firstPack?.unit||null;
      for(const o of options){
        o.perUnit=(orderable(o)&&displayUnit)?pricePerUnit(o.casePrice,o.packSize,displayUnit):null;
        // Never represent unlike dimensions as competing per-unit offers.
        if(o.perUnit && o.perUnit.unit!==displayUnit) o.perUnit=null;
      }

      return {
        catalogItemId:ci.id,
        masterItemNumber:ci.master_item_number,
        name:ci.name,
        category:ci.catalog_categories?.name||penName,
        createdAt:ci.created_at,
        brandLocked:ci.brand_locked||false,
        lockedBrand,
        matchingBehavior:ci.matching_behavior||"flexible",
        canonicalUnit:ci.canonical_unit||null,
        displayUnit,
        unitDimension:firstPack?.dimension||null,
        options,
      };
    });
  },[catalogItems,categories,vendorItems,mappings,vendors,vendorColors,vocabulary,org?.settings,clockTick]);

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
    return vendorItems.filter(vi=>quoteStatus(vi,org?.settings||{})==="expired").map(vi=>{
      const v=vMap.get(vi.vendor_id);
      return {id:vi.id, vendorName:v?.name||"—", vendorId:vi.vendor_id, description:vi.description, lastUpdated:vi.last_updated};
    });
  },[vendorItems,vMap,org?.settings,clockTick]);

  // Item Catalog's nav badge is scoped to catalog MAPPING issues only
  // (fuzzy vendor-item matches) - invoice-line issues get their own
  // badge on Invoices, price-sheet health (unavailable/stale) gets its
  // own badge on Price Sheets. Each tab's badge reflects only what's
  // actually reviewable on that tab.
  const needsAttentionCount=useMemo(()=>
    mappings.filter(m=>["similar","review"].includes(m.comparison_track)).length,
  [mappings]);
  const invoiceReviewCount=flaggedInvoiceLines.length;
  const priceSheetReviewCount=priceUnavailableItems.length+expiredItems.length;

  const setQty=(key,val)=>setQuantities(p=>({...p,[key]:Math.max(0,val)}));

  const basketKey=org?`kerdos.basket.${org.id}`:null;
  useEffect(()=>{
    if(!basketKey) return;
    setRestoredBasketKey(null);
    let saved=null;
    try{ saved=JSON.parse(localStorage.getItem(basketKey)||"null"); }catch{}
    setQuantities(saved?.quantities||{});
    setUnitSelection(saved?.unitSelection||{});
    setVendorOverride(saved?.vendorOverride||{});
    setPriceOverride(saved?.priceOverride||{});
    setRestoredBasketKey(basketKey);
  },[basketKey]);

  useEffect(()=>{
    if(!basketKey||restoredBasketKey!==basketKey) return;
    try{
      localStorage.setItem(basketKey,JSON.stringify({quantities,unitSelection,vendorOverride,priceOverride}));
    }catch{}
  },[basketKey,restoredBasketKey,quantities,unitSelection,vendorOverride,priceOverride]);

  function clearBasket(){
    setQuantities({});
    setVendorOverride({});
    setPriceOverride({});
  }

  function dropOnBasket(vendorId){
    const dragged=dragItem;
    setDragItem(null);
    setDropTarget(null);
    if(!dragged) return;
    const product=productList.find(item=>item.catalogItemId===dragged.catalogItemId);
    if(!product) return;
    if(vendorId){
      const wantsEach=dragged.key.endsWith("_each");
      const option=product.options.find(candidate=>candidate.vendorId===vendorId&&orderable(candidate)&&(!wantsEach||candidate.eachPrice));
      if(!option){
        const vendorName=vendors.find(vendor=>vendor.id===vendorId)?.name||"That vendor";
        setDropNotice(`${vendorName} has no current quote for ${product.name}.`);
        window.setTimeout(()=>setDropNotice(""),4000);
        return;
      }
      setVendorOverride(previous=>({...previous,[dragged.key]:vendorId}));
      setPriceOverride(previous=>{const next={...previous};delete next[dragged.key];return next;});
    }
    setQty(dragged.key,(quantities[dragged.key]||0)+1);
  }

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
      if(a.unorderable||!a.assignedVendorId) continue;
      const b=map.get(a.assignedVendorId)||{vendorId:a.assignedVendorId,vendorName:a.assignedVendorName,items:[],dollar:0,units:0};
      b.items.push(a);b.dollar=r2(b.dollar+a.lineTotal);b.units+=a.quantity;
      map.set(a.assignedVendorId,b);
    }
    return Array.from(map.values());
  },[assignments]);

  const baselineSpend=useMemo(()=>cartItems.reduce((s,i)=>{
    const cheapest=i.options.filter(orderable).sort((a,b)=>a.price-b.price)[0];
    return s+(cheapest?.price||0)*i.quantity;
  },0),[cartItems]);
  const blockedCartItems=assignments.filter(a=>a.unorderable||!a.assignedVendorId);
  const totalSpend=baskets.reduce((s,b)=>s+b.dollar,0);
  const totalUnits=baskets.reduce((s,b)=>s+b.units,0);

  const orderCategoryList=useMemo(()=>{
    // Order Guide only ever lists items with a vendor price (see the
    // options.length>0 filter below) - the category chip list is scoped
    // to that same orderable set, not the full catalog, so a category
    // that's entirely unpriced doesn't show an empty chip here. Order
    // Guide's categories are alphabetical - not the number-range order
    // Item Catalog uses, since this screen is for FINDING something to
    // order, not for working the numbering itself.
    const set=new Set(productList.filter(i=>i.options.length>0).map(p=>p.category));
    return [...set].sort((a,b)=>a.localeCompare(b));
  },[productList]);

  const filtered=useMemo(()=>{
    // Order Guide is for ORDERING - a client-created item with no vendor
    // price mapped to it yet has nothing to order, so it's excluded here
    // even though it's fully visible in Item Catalog.
    const matches=productList.filter(i=>{
      if(i.options.length===0) return false;
      if(orderCategoryFilter&&i.category!==orderCategoryFilter) return false;
      return itemMatchesSearch(i,search);
    });
    // Category alphabetical first, then ordered within each category by
    // whichever sort mode is active (defaults to alphabetical too, so
    // "alphabetize category, then alphabetize in the category" is the
    // out-of-the-box behavior).
    const groups=new Map();
    matches.forEach(item=>{
      const cat=item.category;
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
      try {
        await write(supabase.from("organizations").update({ logo_url: result.path }).eq("id", org.id), "Could not save the logo");
        const url = await getSignedUrl(result.path);
        setLogoUrl(url);
        setOrg(o => ({ ...o, logo_url: result.path }));
      } catch (err) { alert(err.message); }
    }
    setLogoUploading(false);
  }

  async function deleteInvoice(inv){
    if(!window.confirm(`Delete this ${inv.vendors?.name||""} invoice? This can't be undone.`)) return;
    try{
      await write(supabase.from("invoice_lines").delete().eq("invoice_id",inv.id),"Could not delete the invoice lines");
      await write(supabase.from("invoices").delete().eq("id",inv.id),"Could not delete the invoice");
      if(inv.file_path){
        await documents.remove([inv.file_path]);
      }
    }catch(err){ alert(err.message); }
    loadData();
  }

  async function loadOlderPriceHistory(){
    if(!org?.id||loadingOlderPrices)return;
    setLoadingOlderPrices(true);
    try{
      const page=await write(supabase.from("price_history").select("*").eq("organization_id",org.id)
        .order("effective_date",{ascending:false}).order("id",{ascending:false})
        .range(priceHistory.length,priceHistory.length+1999),"Could not load earlier price sheets");
      setPriceHistory(current=>{
        const existing=new Set(current.map(row=>row.id));
        return [...current,...page.filter(row=>!existing.has(row.id))];
      });
      setPriceHistoryHasMore(page.length===2000);
    }catch(err){alert(err.message);}
    setLoadingOlderPrices(false);
  }

  async function expireVendorQuotes(vendorId){
    const vendor=vendors.find(v=>v.id===vendorId);
    const active=vendorItems.filter(vi=>vi.vendor_id===vendorId && quoteStatus(vi,org?.settings||{})==="current");
    if(!active.length){alert("No current vendor prices to expire.");return;}
    if(!window.confirm(`Expire ${active.length} current price(s) for ${vendor?.name||"this vendor"}? All quoted amounts remain in Price Sheet History. Nothing is deleted.`)) return;
    try{
      await write(supabase.from("vendor_items").update({price_expired_at:new Date().toISOString()})
        .eq("organization_id",org.id).eq("vendor_id",vendorId).in("id",active.map(vi=>vi.id)),"Could not expire vendor quotations");
      await loadData();
    }catch(err){alert(err.message);}
  }

  async function expireOneQuote(item){
    if(!window.confirm(`Expire ${item.description} for ${vendors.find(v=>v.id===item.vendor_id)?.name||"this vendor"}? Its price stays in history.`))return;
    try{
      await write(supabase.from("vendor_items").update({price_expired_at:new Date().toISOString()})
        .eq("organization_id",org.id).eq("id",item.id),"Could not expire the quotation");
      await loadData();
    }catch(err){alert(err.message);}
  }

  async function submitOrders(){
    if(blockedCartItems.length){alert(`${blockedCartItems.length} item(s) need a confirmed current quote. Remove them or refresh their prices before submitting.`);return;}
    if(!baskets.length){alert("No orderable items in the basket.");return;}
    let placed=0;
    try{
      for(const basket of baskets){
        const order=await write(supabase.from("purchase_orders").insert({
          organization_id:org.id,vendor_id:basket.vendorId,
          created_by:session.user.id,status:"submitted",total_amount:basket.dollar,
        }).select().single(),`Could not submit the ${basket.vendorName} order`);
        await write(supabase.from("purchase_order_lines").insert(basket.items.map(item=>({
          purchase_order_id:order.id,catalog_item_id:item.catalogItemId.replace("_each",""),
          vendor_item_id:item.vendorItemId,quantity:item.quantity,
          unit_price:item.price,line_total:item.lineTotal,
        }))),`Could not save the ${basket.vendorName} order lines`);
        placed++;
      }
      alert(`${placed} order${placed===1?"":"s"} recorded in KERDOS. These records do NOT send orders to vendors. Use the Email/Text draft buttons and send from your messaging app.`);
      clearBasket();
    }catch(err){
      alert(`${err.message}${placed?` (${placed} order${placed===1?"":"s"} before it did go through.)`:""}`);
    }
    // Recent Orders reads purchase_orders, which the live channel does not
    // watch, so refresh explicitly rather than waiting for the next reload.
    loadData();
  }

  // Send one vendor basket through the user's normal email or text app
  // without changing the optimized basket or recording it as submitted.
  function basketOrderMessage(vendor,basket){
    const lines=[
      "KERDOS Purchase Order",
      `Customer: ${org.name}`,
      `Vendor: ${vendor.name}`,
      "",
      ...basket.items.map(item=>{
        const vi=vendorItems.find(v=>v.id===item.vendorItemId);
        const code=vi?.vendor_item_code?`[${vi.vendor_item_code}] `:"";
        const pack=item.packSize?` (${item.packSize})`:"";
        return `${item.quantity} ${item.orderUnit} × ${code}${vi?.description||item.name}${pack} — ${formatMoney(item.lineTotal)}`;
      }),
      "",
      `Order total: ${formatMoney(basket.dollar)}`,
    ];
    return lines.join("\n");
  }

  function emailBasket(vendor,basket){
    const destination=(vendor.email||window.prompt(`Email address for ${vendor.name}:`,"")||"").trim();
    if(!destination) return;
    const body=basketOrderMessage(vendor,basket);
    if(!window.confirm(`Open an email for ${vendor.name} with this ${basket.items.length}-item order totaling ${formatMoney(basket.dollar)}?`)) return;
    const subject=`${org.name} purchase order — ${vendor.name}`;
    window.location.href=`mailto:${destination}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  }

  function textBasket(vendor,basket){
    const phone=(window.prompt(`Mobile number for ${vendor.name}:`,"")||"").trim();
    if(!phone) return;
    const cleanPhone=phone.replace(/[^+\d]/g,"");
    if(!cleanPhone) return;
    const body=basketOrderMessage(vendor,basket);
    if(!window.confirm(`Open a text message for ${vendor.name} with this ${basket.items.length}-item order totaling ${formatMoney(basket.dollar)}?`)) return;
    window.location.href=`sms:${cleanPhone}?body=${encodeURIComponent(body)}`;
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
              <div style={{fontWeight:800,fontSize:17}}>{formatMoney(totalSpend)}</div>
              {totalSpend>baselineSpend+0.01&&<div style={{fontSize:10,color:"#FF9800"}}>+{formatMoney(totalSpend-baselineSpend)} vs cheapest</div>}
            </div>
          )}
          <div style={{fontSize:11,opacity:0.6,textAlign:"right"}}>
              {organizations.length>1
                ?<select aria-label="Active organization" value={org.id} onChange={e=>switchOrganization(e.target.value)}
                  style={{maxWidth:190,fontSize:11,color:"white",background:"#00204F",border:"1px solid #4A90D9",borderRadius:4,padding:"2px 4px"}}>
                  {organizations.map(o=><option key={o.id} value={o.id}>{o.name}</option>)}
                </select>
                :<div>{org.name}</div>}
            <button onClick={()=>{try{sessionStorage.removeItem("kerdos.activeTab");}catch{} backend.session.signOut().catch(err=>alert(err.message));}} style={{background:"none",border:"none",color:"#4A90D9",cursor:"pointer",fontSize:11,padding:0}}>Sign out</button>
          </div>
        </div>
      </header>

      {offline&&(
        <div style={{background:"#FFF3E0",color:"#8A4F00",padding:"8px 16px",fontSize:12,fontWeight:600,display:"flex",justifyContent:"space-between",alignItems:"center",gap:12}}>
          <span>Working offline — showing your workspace as of {formatDate(offline.since)}. Your basket remains saved on this device.</span>
          <button onClick={()=>loadData()} style={{...btn("#8A4F00","white",{fontSize:11,padding:"5px 10px"})}}>Retry</button>
        </div>
      )}

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
              .order-aside-inner { position:sticky; top:90px; }
              @media (max-width:720px) {
                .order-layout { flex-direction:column; }
                .order-aside { width:100%; }
                .order-aside-inner { position:static; }
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
                placeholder={`🔍 Search ${productList.length} items by name, vendor wording, or code...`}
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
                        const cheapestUnitPrice=unitOptions.find(orderable)?.unitPrice;
                        const bestPerUnit=unitOptions.filter(o=>orderable(o)&&o.perUnit).sort((a,b)=>a.perUnit.price-b.perUnit.price)[0]||null;
                        const menuOpen=openPriceMenu===activeKey;
                        const activeBlocked=assignment?assignment.unorderable:!!(cheapest&&!orderable(cheapest));
                        const activeBlockReason=activeBlocked?(cheapest?blockReason(cheapest):"No price on file"):null;
                        const activeOption=item.options.find(o=>o.vendorItemId===(assignment?.vendorItemId))||cheapest;
                        const activeMatchTrack=activeOption?.matchTrack;
                        const activeMatchConfidence=activeOption?.matchConfidence;

                        const cells=[
                          <div key={item.catalogItemId+"_name"} draggable={!activeBlocked}
                            onDragStart={event=>{event.dataTransfer.effectAllowed="copy";event.dataTransfer.setData("text/plain",activeKey);setDragItem({catalogItemId:item.catalogItemId,key:activeKey,name:item.name});}}
                            onDragEnd={()=>{setDragItem(null);setDropTarget(null);}}
                            title={activeBlocked?undefined:"Drag onto a vendor basket to order from that vendor"}
                            style={{minWidth:0,padding:"8px 8px 8px 2px",borderTop:"1px solid #F2F2F2",cursor:activeBlocked?"default":"grab",opacity:dragItem?.key===activeKey?0.5:1}}>
                            <div style={{fontWeight:600,fontSize:12.5,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                              {item.name}
                              {item.lockedBrand&&<span title={`Locked to ${item.lockedBrand} - other brands are never ordered for this item`} style={{marginLeft:4,fontSize:9,background:"#E3F2FD",color:"#1565C0",padding:"1px 4px",borderRadius:4,fontWeight:700}}>🔒 {item.lockedBrand}</span>}
                              {["similar","review"].includes(activeMatchTrack)&&(
                                <span title="Auto-matched to this product below full confidence - worth double-checking it's really the same item"
                                  style={{marginLeft:4,fontSize:9,background:"#FFF3E0",color:"#B26A00",padding:"1px 4px",borderRadius:4,fontWeight:700}}>
                                  🔍 {activeMatchConfidence}%
                                </span>
                              )}
                            </div>
                            <div style={{fontSize:10,color:"#AAA"}}>
                              {activePackSize||""}
                              {!activeBlocked&&activeOption?.perUnit&&<span style={{marginLeft:6,color:"#666",fontWeight:600}}>{formatMoney(activeOption.perUnit.price)}/{activeOption.perUnit.unit}</span>}
                            </div>
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
                            <button onClick={()=>{if(!activeBlocked) setQty(activeKey,activeQty+1);}} disabled={activeBlocked}
                              title={activeBlocked?activeBlockReason:undefined}
                              style={{width:26,height:26,borderRadius:7,border:"none",background:activeBlocked?"#DDD":"#2E7D32",cursor:activeBlocked?"not-allowed":"pointer",fontSize:15,fontWeight:800,color:"white",flexShrink:0}}>+</button>
                          </div>,
                          <div key={item.catalogItemId+"_price"} style={{textAlign:"right",padding:"6px",borderTop:"1px solid #F2F2F2",borderLeft:"1px solid #EEE",background:vc.bg}}>
                            <button onClick={()=>{
                                setOpenPriceMenu(menuOpen?null:activeKey);
                                setCustomPriceInput(activePrice!=null?String(activePrice):"");
                              }}
                              style={{background:"white",border:`1px solid ${activeBlocked?"#FFCC80":vc.light}`,borderRadius:8,cursor:unitOptions.length?"pointer":"default",padding:"5px 8px",textAlign:"right",width:"100%"}}>
                              {activeBlocked?(
                                <div style={{fontWeight:700,fontSize:11,color:"#E65100"}}>⚠ {activeBlockReason}{activeOption?` (last ${formatMoney(activeOption.casePrice)})`:""}</div>
                              ):(
                                <div style={{fontWeight:800,fontSize:13,color:vc.accent}}>{formatMoney(activePrice)}{isCustomPrice&&<span title="Custom price" style={{marginLeft:2,fontSize:9}}>✎</span>}</div>
                              )}
                              <div style={{fontSize:9,fontWeight:700,color:activeBlocked?"#E65100":vc.accent}}>{activeBlocked?"":activeVendorName} {unitOptions.length>1&&(menuOpen?"▲":"▾")}</div>
                            </button>
                          </div>,
                        ];

                        if(menuOpen){
                          cells.push(
                            <div key={item.catalogItemId+"_menu"} style={{gridColumn:"1 / -1",background:"#FAFAFA",borderTop:"1px solid #F0F0F0",borderRadius:6,padding:"8px 10px",marginBottom:4}}>
                              <div style={{fontSize:10,color:"#BBB",fontWeight:700,letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:6}}>Power ranked — tap to select</div>
                              {unitOptions.map(opt=>{
                                const isSelected=assignment?.vendorItemId===opt.vendorItemId&&!isCustomPrice;
                                const blocked=!orderable(opt);
                                return (
                                  <button key={opt.vendorItemId} onClick={()=>{
                                      if(blocked) return;
                                      setVendorOverride(prev=>({...prev,[activeKey]:opt.vendorId}));
                                      setPriceOverride(prev=>{const n={...prev};delete n[activeKey];return n;});
                                      setOpenPriceMenu(null);
                                    }}
                                    disabled={blocked}
                                    title={blocked?blockReason(opt):undefined}
                                    style={{width:"100%",display:"flex",justifyContent:"space-between",alignItems:"center",
                                      background:blocked?"#FAFAFA":isSelected?"#E8F5E9":"white",border:`1px solid ${blocked?"#F0F0F0":isSelected?"#A5D6A7":"#EEE"}`,
                                      borderRadius:6,padding:"7px 10px",marginBottom:4,cursor:blocked?"not-allowed":"pointer",textAlign:"left",opacity:blocked?0.6:1}}>
                                    <span style={{fontSize:12,fontWeight:isSelected?700:500}}>{isSelected&&"✓ "}{opt.vendorName}
                                      {opt.brand&&<span style={{marginLeft:5,fontSize:10,color:"#999"}}>{opt.brand}</span>}
                                      {!blocked&&opt.priceUnavailable&&(
                                        <span title="Vendor's latest price sheet listed no price for this item - showing the last known price instead" style={{marginLeft:6,fontSize:10,fontWeight:700,color:"#B26A00",background:"#FFF3E0",padding:"1px 5px",borderRadius:4}}>
                                          ⓘ last known price
                                        </span>
                                      )}
                                      {!blocked&&["similar","review"].includes(opt.matchTrack)&&(
                                        <span title="Auto-matched to this product below full confidence - worth double-checking" style={{marginLeft:6,fontSize:10,fontWeight:700,color:"#B26A00",background:"#FFF3E0",padding:"1px 5px",borderRadius:4}}>
                                          🔍 {opt.matchConfidence}% match
                                        </span>
                                      )}
                                    </span>
                                    {blocked?(
                                      <span style={{fontSize:11,fontWeight:700,color:"#E65100"}}>{opt.expired?"⚠":"🔒"} {blockReason(opt)}</span>
                                    ):(
                                      <span style={{textAlign:"right"}}>
                                        <span style={{fontSize:12,fontWeight:700}}>{formatMoney(opt.unitPrice)} {cheapestUnitPrice!=null&&opt.unitPrice>cheapestUnitPrice&&<span style={{color:"#E65100",fontWeight:600}}>(+{formatMoney(opt.unitPrice-cheapestUnitPrice)})</span>}</span>
                                        {opt.perUnit&&(
                                          <span style={{display:"block",fontSize:10,color:bestPerUnit?.vendorItemId===opt.vendorItemId?"#2E7D32":"#888",fontWeight:bestPerUnit?.vendorItemId===opt.vendorItemId?700:500}}>
                                            {formatMoney(opt.perUnit.price)}/{opt.perUnit.unit}{bestPerUnit?.vendorItemId===opt.vendorItemId&&unitOptions.filter(orderable).length>1?" · best per unit":""}
                                          </span>
                                        )}
                                      </span>
                                    )}
                                  </button>
                                );
                              })}
                              <div style={{display:"flex",gap:6,marginTop:8,alignItems:"center"}}>
                                <span style={{fontSize:11,color:"#888",flexShrink:0}}>Negotiated price:</span>
                                <input type="number" step="0.01" value={customPriceInput} onChange={e=>setCustomPriceInput(e.target.value)}
                                  style={{...inp,padding:"5px 8px",fontSize:12,flex:1}} placeholder={formatMoney(0)} />
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
                            <div style={{fontSize:12,color:"#888"}}>{formatDate(po.created_at)} · {lines.length} item{lines.length===1?"":"s"}</div>
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
              <div className="order-aside-inner"
                onDragOver={event=>{if(dragItem){event.preventDefault();event.dataTransfer.dropEffect="copy";if(dropTarget==null)setDropTarget("auto");}}}
                onDragLeave={event=>{if(dropTarget==="auto"&&!event.currentTarget.contains(event.relatedTarget))setDropTarget(null);}}
                onDrop={event=>{event.preventDefault();event.stopPropagation();dropOnBasket(dropTarget&&dropTarget!=="auto"?dropTarget:null);}}
                style={dragItem?{outline:"2px dashed rgba(255,255,255,0.6)",outlineOffset:6,borderRadius:10}:undefined}>
                {dragItem&&(
                  <div style={{background:"rgba(255,255,255,0.12)",border:"1px solid rgba(255,255,255,0.35)",borderRadius:9,padding:10,marginBottom:10}}>
                    <div style={{fontSize:11,fontWeight:700,color:"white",marginBottom:6}}>Drop <b>{dragItem.name}</b> on a vendor, or anywhere here for the best price</div>
                    <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                      {vendors.map(vendor=>{
                        const color=vendorColors.get(vendor.id)||PALETTE[0];
                        const selected=dropTarget===vendor.id;
                        return <div key={vendor.id}
                          onDragEnter={event=>{event.preventDefault();event.stopPropagation();setDropTarget(vendor.id);}}
                          onDragOver={event=>{event.preventDefault();event.stopPropagation();event.dataTransfer.dropEffect="copy";if(dropTarget!==vendor.id)setDropTarget(vendor.id);}}
                          onDrop={event=>{event.preventDefault();event.stopPropagation();dropOnBasket(vendor.id);}}
                          style={{fontSize:12,fontWeight:700,padding:"8px 14px",borderRadius:20,background:selected?"white":color.bg,color:color.accent,outline:selected?`3px solid ${color.accent}`:"none"}}>{vendor.name}</div>;
                      })}
                    </div>
                  </div>
                )}
                {dropNotice&&<div style={{background:"#FFF3E0",color:"#8A4F00",borderRadius:9,padding:"8px 12px",marginBottom:10,fontSize:12,fontWeight:600}}>{dropNotice}</div>}
                {blockedCartItems.length>0&&<div style={{background:"#FFF3E0",color:"#8A4F00",borderRadius:9,padding:12,marginBottom:10,fontSize:12}}>
                  <b>{blockedCartItems.length} item(s) need attention — not in vendor baskets.</b>
                  {blockedCartItems.map(a=><div key={a.catalogItemId} style={{marginTop:5}}>{a.quantity} × {a.name}: confirm a current eligible vendor quote or remove from order.</div>)}
                </div>}
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                  <div style={{fontSize:11,fontWeight:800,color:"rgba(255,255,255,0.7)",letterSpacing:"0.08em",textTransform:"uppercase"}}>Order Baskets</div>
                  {baskets.length>0&&(
                    <button onClick={clearBasket}
                      style={{background:"rgba(255,255,255,0.12)",border:"1px solid rgba(255,255,255,0.35)",color:"white",borderRadius:6,padding:"5px 8px",fontSize:10,fontWeight:700,cursor:"pointer"}}>
                      Clear All
                    </button>
                  )}
                </div>

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
                      <div key={vendor.id}
                        onDragEnter={event=>{if(dragItem){event.preventDefault();event.stopPropagation();setDropTarget(vendor.id);}}}
                        onDragOver={event=>{if(dragItem){event.preventDefault();event.stopPropagation();event.dataTransfer.dropEffect="copy";if(dropTarget!==vendor.id)setDropTarget(vendor.id);}}}
                        onDrop={event=>{event.preventDefault();event.stopPropagation();dropOnBasket(vendor.id);}}
                        style={{background:"white",borderRadius:10,marginBottom:10,
                        border:`2px solid ${dropTarget===vendor.id?"#1565C0":meetsAll?vc.accent:"#FFB74D"}`,overflow:"hidden",
                        boxShadow:dropTarget===vendor.id?"0 0 0 4px rgba(21,101,192,0.25)":"none"}}>
                        <div style={{background:meetsAll?vc.bg:"#FFF8E1",padding:"9px 12px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                          <div>
                            <div style={{fontWeight:800,fontSize:13,color:vc.accent}}>{vendor.name}</div>
                            <div style={{fontSize:10,color:"#999"}}>{basket.units} items</div>
                          </div>
                          <div style={{textAlign:"right"}}>
                            <div style={{fontWeight:800,fontSize:15}}>{formatMoney(basket.dollar)}</div>
                            {meetsAll&&<div style={{fontSize:9,color:vc.accent,fontWeight:700}}>✓ READY</div>}
                          </div>
                        </div>
                        {vendor.delivery_minimum_dollar&&(
                          <div style={{padding:"6px 12px 0"}}>
                            <div style={{display:"flex",justifyContent:"space-between",fontSize:9,color:"#AAA",marginBottom:2}}>
                              <span>Min {formatMoney(vendor.delivery_minimum_dollar)}</span>
                              <span style={{color:meetsDollar?vc.accent:"#FF9800",fontWeight:600}}>
                                {meetsDollar?"✓":`${formatMoney(vendor.delivery_minimum_dollar-basket.dollar)} to go`}
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
                              <span style={{fontWeight:700,flexShrink:0,marginLeft:6,color:item.premiumPaid>0.005?"#FF9800":"#333"}}>{formatMoney(item.lineTotal)}</span>
                            </div>
                          ))}
                        </div>
                        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,padding:"0 12px 10px"}}>
                          <button onClick={()=>emailBasket(vendor,basket)}
                            style={{border:"1px solid #D7E0EA",background:"#F7FAFD",color:"#24445F",borderRadius:6,padding:"6px 7px",fontSize:10,fontWeight:800,cursor:"pointer"}}>
                            ✉️ Email Order
                          </button>
                          <button onClick={()=>textBasket(vendor,basket)}
                            style={{border:"1px solid #D7E0EA",background:"#F7FAFD",color:"#24445F",borderRadius:6,padding:"6px 7px",fontSize:10,fontWeight:800,cursor:"pointer"}}>
                            💬 Text Order
                          </button>
                        </div>
                        {vendor.email&&(
                          <div style={{fontSize:9,color:"#999",padding:"0 12px 9px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>Email: {vendor.email}</div>
                        )}
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
                      <span style={{fontWeight:600}}>{formatMoney(baselineSpend)}</span>
                    </div>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}>
                      <span style={{fontSize:11,opacity:0.6}}>Optimized total</span>
                      <span style={{fontWeight:800,fontSize:16,color:totalSpend>baselineSpend+0.01?"#FF9800":"#69F0AE"}}>{formatMoney(totalSpend)}</span>
                    </div>
                    <button onClick={submitOrders} disabled={blockedCartItems.length>0}
                      title={blockedCartItems.length?"Resolve the items needing a price first":undefined}
                      style={{...btn(blockedCartItems.length?"#777":"#4A90D9"),width:"100%"}}>
                      Record {baskets.length} Order{baskets.length>1?"s":""} (not sent)
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
                      <div style={{fontSize:11,color:"#999",marginTop:2}}>{l.vendorName} — {formatDate(l.invoiceDate)} — {formatMoney(l.price)}</div>
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
                    <div style={{fontWeight:700,fontSize:13}}>{formatDate(inv.invoice_date)||formatDate(inv.created_at)} · {inv.status}{inv.invoice_number?` · #${inv.invoice_number}`:""}</div>
                    {hasVariance&&(
                      <div style={{fontSize:11,fontWeight:700,color:inv._totalVariance>0?"#E65100":"#0A8A4B",marginTop:2}}>
                        ⚠️ Paid {formatMoney(Math.abs(inv._totalVariance))} {inv._totalVariance>0?"more":"less"} than quoted
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
            periodGroups.get(key).entries.push({...ph,description:ph.source_description||vi.description});
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
                      <div style={{fontSize:11,color:"#999",marginTop:2}}>{i.vendorName} — last confirmed {daysAgo(i.lastUpdated)}; ordering blocked until a valid quote arrives</div>
                    </div>
                  ))}
                </Section>
                <Section title="Expired quotations (manual, client period, or vendor term)" count={expiredItems.length} emptyText={org?.settings?.price_refresh_days?"Nothing flagged — every price is within the refresh window.":"Price refresh period is off — set one in Admin to enable this check."}>
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

            <h3 style={{margin:"0 0 12px",fontSize:16,color:"white"}}>Current vendor quotations</h3>
            {vendors.filter(v=>!priceSheetVendorFilter||v.id===priceSheetVendorFilter).map(v=>(
              <div key={v.id} style={{background:"white",borderRadius:8,padding:12,marginBottom:10}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
                  <b style={{fontSize:13}}>{v.name}</b>
                  {org.role!=="employee"&&<button onClick={()=>expireVendorQuotes(v.id)}
                    style={{...btn("#FFF3E0","#A14C00",{fontSize:11,padding:"6px 9px"})}}>Expire this vendor's current prices</button>}
                </div>
                {vendorItems.filter(vi=>vi.vendor_id===v.id).map(vi=>{
                  const status=quoteStatus(vi,org?.settings||{});
                  return <div key={vi.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,borderTop:"1px solid #EEE",padding:"5px 0",fontSize:12}}>
                    <span style={{minWidth:0,flex:1}}>{vi.description} · {vi.pack_size||"pack not confirmed"}</span>
                    <span style={{fontWeight:700}}>{vi.price!=null?formatMoney(vi.price):"—"}</span>
                    <span style={{color:status==="current"?"#2E7D32":"#B26A00"}}>{status==="current"?"Current":status==="expired"?"Expired":status==="invoice_only"?"Invoice only — quote needed":"Price needed"}</span>
                    {org.role!=="employee"&&status==="current"&&<button onClick={()=>expireOneQuote(vi)} style={{...btn("#EEE","#555",{fontSize:10,padding:"4px 7px"})}}>Expire</button>}
                  </div>;
                })}
              </div>
            ))}
            <h3 style={{margin:"20px 0 12px",fontSize:16,color:"white"}}>Price Sheet History — every original quote retained</h3>
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
                            <div style={{fontWeight:700,fontSize:13}}>Quoted {formatDate(period.date)}</div>
                            <div style={{fontSize:11,color:"#888"}}>{period.entries.length} item{period.entries.length===1?"":"s"} in this sheet</div>
                          </div>
                          <span style={{color:"#CCC"}}>{isOpen?"▲":"▼"}</span>
                        </button>
                        {isOpen&&(
                          <div style={{borderTop:"1px solid #F0F0F0",padding:"10px 14px",background:"#FAFAFA"}}>
                            {period.entries.map(entry=>(
                              <div key={entry.id} style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"5px 0",borderBottom:"1px solid #F0F0F0"}}>
                                <div>{entry.description}</div>
                                <div style={{textAlign:"right"}}><div style={{fontWeight:700}}>{formatMoney(entry.price)}</div>
                                  {entry.quote_valid_until&&<div style={{fontSize:10,color:"#999"}}>Vendor valid through {formatDate(entry.quote_valid_until)}</div>}
                                  {entry.source_file_name&&<div style={{fontSize:10,color:"#999"}}>{entry.source_file_name}</div>}
                                  {entry.source_document_id&&<button onClick={()=>viewSourceDocument(entry.source_document_id)} style={{background:"none",border:"none",fontSize:10,color:"#003584",cursor:"pointer"}}>Original source ↗</button>}
                                  {!entry.source_document_id&&entry.source_file_path&&<button onClick={()=>viewStoredFile(entry.source_file_path)} style={{background:"none",border:"none",fontSize:10,color:"#003584",cursor:"pointer"}}>Original ↗</button>}
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
            {priceHistoryHasMore&&<button onClick={loadOlderPriceHistory} disabled={loadingOlderPrices}
              style={{...btn("white","#003584",{marginTop:10})}}>{loadingOlderPrices?"Loading earlier price sheets...":"Load earlier price sheet history"}</button>}
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
            <CatalogPanel orgId={org.id} orgIndustry={org.industry} categories={categories} catalogItems={catalogItems} vocabulary={vocabulary} onUpdated={loadData} />
          </>
        )}
      </div>

      {showPaste&&<PasteModal vendors={vendors} orgId={org.id} orgSettings={org.settings} catalogItems={catalogItems} categories={categories} onClose={()=>setShowPaste(false)} onDone={loadData} initialVendorId={selectedVendorId} initialMode={importMode} />}
      {showAddVendor&&<AddVendorModal orgId={org.id} onClose={()=>setShowAddVendor(false)} onDone={loadData} />}
      {editingInvoice&&<InvoiceEditModal invoice={editingInvoice} vendors={vendors} onClose={()=>setEditingInvoice(null)} onDone={loadData} />}
    </div>
  );
}
