// KERDOS deterministic procurement primitives.
// Pure functions only: no UI, database, client, vendor, or industry assumptions.

// One named place for the auto-link cutoff instead of the same magic
// number repeated at each call site. At or above this score two
// descriptions are treated as the same product and linked automatically;
// below it, the pair is kept for human review rather than silently
// merged.
const MATCH_POLICY = Object.freeze({
  autoLink: 0.85,   // at/above: same product, linked automatically
  reviewFloor: 0.5, // at/above (but below autoLink): linked, flagged for review
});                 // below reviewFloor: not the same product

// ── VOCABULARY ────────────────────────────────────────────────────
// The engine ships a small base vocabulary and knows nothing beyond it.
// Everything industry- or language-specific arrives as data rows and is
// merged in by configureVocabulary(): the org's own vocabulary, seeded
// from its industry's pack and taught by its own corrections. Four kinds
// of row:
//   unit      term -> canonical unit code (an existing code, or a new one
//                     such as SHEET that compares only against itself)
//   packaging term is a container word, never a product word
//   stopword  term carries no meaning for matching
//   synonym   term -> canonical word, so "chix" reads as "chicken"
const BASE_STOPWORDS = ["the","a","an","of","and","or","with","in","new"];

const UNIT_DEFINITIONS = {
  // mass -> grams
  G:{dimension:"mass",base:"G",factor:1,aliases:["g","gram","grams"]},
  KG:{dimension:"mass",base:"G",factor:1000,aliases:["kg","kgs","kilogram","kilograms"]},
  OZ:{dimension:"mass",base:"G",factor:28.349523125,aliases:["oz","ounce","ounces"]},
  LB:{dimension:"mass",base:"G",factor:453.59237,aliases:["lb","lbs","pound","pounds"]},
  // volume -> milliliters
  ML:{dimension:"volume",base:"ML",factor:1,aliases:["ml","milliliter","milliliters"]},
  L:{dimension:"volume",base:"ML",factor:1000,aliases:["l","liter","liters","litre","litres"]},
  FLOZ:{dimension:"volume",base:"ML",factor:29.5735295625,aliases:["fl oz","floz","fluid ounce","fluid ounces"]},
  PT:{dimension:"volume",base:"ML",factor:473.176473,aliases:["pt","pint","pints"]},
  QT:{dimension:"volume",base:"ML",factor:946.352946,aliases:["qt","quart","quarts"]},
  GAL:{dimension:"volume",base:"ML",factor:3785.411784,aliases:["gal","gallon","gallons","gl","ga"]},
  // length -> millimeters
  MM:{dimension:"length",base:"MM",factor:1,aliases:["mm","millimeter","millimeters"]},
  CM:{dimension:"length",base:"MM",factor:10,aliases:["cm","centimeter","centimeters"]},
  M:{dimension:"length",base:"MM",factor:1000,aliases:["m","meter","meters","metre","metres"]},
  IN:{dimension:"length",base:"MM",factor:25.4,aliases:["in","inch","inches"]},
  FT:{dimension:"length",base:"MM",factor:304.8,aliases:["ft","foot","feet"]},
  YD:{dimension:"length",base:"MM",factor:914.4,aliases:["yd","yard","yards"]},
  // count -> each
  EA:{dimension:"count",base:"EA",factor:1,aliases:["ea","each","piece","pieces","pc","pcs"]},
  CT:{dimension:"count",base:"EA",factor:1,aliases:["ct","count"]},
  DOZ:{dimension:"count",base:"EA",factor:12,aliases:["doz","dozen","dz"]},
};

const BASE_PACKAGING = [
  "case","cases","cs","carton","cartons","box","boxes","bx","bag","bags",
  "pack","packs","pk","pallet","pallets","roll","rolls","container","containers",
  "bottle","bottles","can","cans","jar","jars","tub","tubs","tray","trays",
];

// The live vocabulary: base plus whatever configureVocabulary() merged in.
// One organization is active per session, so this is process state rather
// than a parameter threaded through every call.
let STOPWORDS, PACKAGING_ALIASES, UNIT_LOOKUP, SYNONYMS, UNIT_ALTERNATION, MEASURE_RE;

function escapeRegExp(text) { return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }

function configureVocabulary(rows=[]) {
  STOPWORDS = new Set(BASE_STOPWORDS);
  PACKAGING_ALIASES = new Set(BASE_PACKAGING);
  UNIT_LOOKUP = new Map();
  SYNONYMS = new Map();
  for (const [code, def] of Object.entries(UNIT_DEFINITIONS)) {
    UNIT_LOOKUP.set(code.toLowerCase(), code);
    for (const alias of def.aliases) UNIT_LOOKUP.set(alias.toLowerCase(), code);
  }
  for (const r of rows) {
    const term = String(r?.term || "").trim().toLowerCase();
    const canonical = String(r?.canonical || "").trim();
    if (!term) continue;
    if (r.kind === "unit" && canonical) UNIT_LOOKUP.set(term, canonical.toUpperCase());
    else if (r.kind === "packaging") PACKAGING_ALIASES.add(term);
    else if (r.kind === "stopword") STOPWORDS.add(canonicalWord(term));
    else if (r.kind === "synonym" && canonical) SYNONYMS.set(canonicalWord(term), canonicalWord(canonical));
  }
  // Measurements in free text are recognised by the unit vocabulary too,
  // so a unit taught to the org is found in descriptions as well as in
  // pack sizes. Single letters are excluded: "3 g" is too easily a code.
  const words = [...UNIT_LOOKUP.keys()].filter(k => k.length > 1)
    .sort((a, b) => b.length - a.length).map(k => escapeRegExp(k).replace(/\s+/g, "\\s*"));
  UNIT_ALTERNATION = words.join("|");
  MEASURE_RE = new RegExp(`\\b(\\d+(?:\\.\\d+)?)\\s*(${UNIT_ALTERNATION})\\b`, "g");
}

function round(n, places=6) {
  const p = 10 ** places;
  return Math.round((n + Number.EPSILON) * p) / p;
}

function canonicalWord(word) {
  let w = String(word || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  if (w.length > 4 && w.endsWith("ies")) w = w.slice(0,-3) + "y";
  else if (w.length > 4 && /(?:oes|xes|zes|ches|shes)$/.test(w)) w = w.slice(0,-2);
  else if (w.length > 3 && w.endsWith("s") && !w.endsWith("ss")) w = w.slice(0,-1);
  return w;
}

function normalizeForMatch(value) {
  return String(value || "").toLowerCase().replace(/[^a-z0-9\s]/g," ").split(/\s+/)
    .map(canonicalWord).map(w => SYNONYMS.get(w) || w).filter(w => w.length > 1 && !STOPWORDS.has(w));
}

configureVocabulary();

function wordsMatch(a,b) { return canonicalWord(a) === canonicalWord(b); }

function classifyCategory(description, categories=[]) {
  const descWords = normalizeForMatch(description);
  if (!descWords.length) return null;
  const scored=[];
  for (const category of categories) {
    let score=0;
    let longestPhrase=0;
    for (const keyword of (Array.isArray(category.keywords) ? category.keywords : [])) {
      const kwWords=normalizeForMatch(keyword);
      if (kwWords.length && kwWords.every(k => descWords.some(d => wordsMatch(k,d)))) {
        // Specific phrases outrank generic one-word hits without embedding
        // product-specific exceptions in the engine.
        score += kwWords.length * kwWords.length;
        longestPhrase=Math.max(longestPhrase,kwWords.length);
      }
    }
    if (score>0) scored.push({category,score,longestPhrase});
  }
  scored.sort((a,b)=>b.longestPhrase-a.longestPhrase || b.score-a.score);
  if (!scored.length) return null;
  // An exact tie is genuinely ambiguous. Do not let array order silently
  // decide the category; send it to Uncategorized/review instead.
  if (scored[1] && scored[0].longestPhrase===scored[1].longestPhrase && scored[0].score===scored[1].score) return null;
  return scored[0].category;
}

function nextCategoryRange(categories=[], blockSize=10000) {
  const maxEnd=categories.reduce((max,c)=>Math.max(max,Number(c.range_end)||0),0);
  const start=Math.ceil((maxEnd+1)/blockSize)*blockSize || blockSize;
  return {range_start:start,range_end:start+blockSize-1};
}

function normalizeUnit(raw) {
  if (!raw) return null;
  const key=String(raw).trim().toLowerCase().replace(/\./g,"").replace(/\s+/g," ");
  return UNIT_LOOKUP.get(key) || String(raw).trim().toUpperCase();
}

function measurement(quantity, unitRaw) {
  const quantityNumber=Number(quantity);
  if (!Number.isFinite(quantityNumber) || quantityNumber <= 0) return null;
  const unit=normalizeUnit(unitRaw);
  const def=UNIT_DEFINITIONS[unit];
  return def
    ? {quantity:quantityNumber,unit,dimension:def.dimension,baseUnit:def.base,baseQuantity:round(quantityNumber*def.factor)}
    : {quantity:quantityNumber,unit,dimension:"unknown",baseUnit:unit,baseQuantity:quantityNumber};
}

// Parses common packaging expressions without restricting KERDOS to a fixed industry.
// Unknown units remain valid and comparable only to the same unknown unit.
function parsePackSize(raw) {
  if (!raw || !String(raw).trim()) return null;
  const source=String(raw).trim();
  const clean=source.toLowerCase().replace(/[×x]/g,"x").replace(/\s+/g," ").trim();
  const number="(\\d+(?:\\.\\d+)?)";
  // A known unit (any number of words, longest first) is preferred; an
  // unknown single word is still accepted as a unit of its own.
  const unit=`(${UNIT_ALTERNATION}|[a-z]+)`;
  let m=clean.match(new RegExp(`^${number}\\s*(?:/|x)\\s*${number}\\s*${unit}\\b`));
  let outerQty=1, innerQty, unitRaw;
  if (m) { outerQty=Number(m[1]); innerQty=Number(m[2]); unitRaw=m[3]; }
  else {
    m=clean.match(new RegExp(`^${number}\\s*${unit}\\b`));
    if (!m) return {raw:source,parsed:false,levels:[],total:null,unit:null,dimension:"unknown"};
    innerQty=Number(m[1]); unitRaw=m[2];
  }
  const measure=measurement(innerQty,unitRaw);
  if (!measure) return null;
  const total=outerQty*innerQty;
  return {
    raw:source,parsed:true,caseQty:outerQty,unitQty:innerQty,unit:measure.unit,total,
    dimension:measure.dimension,baseUnit:measure.baseUnit,baseTotal:round(outerQty*measure.baseQuantity),
    levels: outerQty>1 ? [{quantity:outerQty,type:"PACKAGE"},{quantity:innerQty,type:measure.unit}] : [{quantity:innerQty,type:measure.unit}],
    eachStr:`${innerQty} ${measure.unit}`,caseStr:outerQty>1?`${outerQty}/${innerQty} ${measure.unit}`:`${innerQty} ${measure.unit}`,
  };
}

function normalizedPrice(price, pack) {
  const p=parsePackSize(pack), n=Number(price);
  if (!p?.parsed || !Number.isFinite(n) || !p.baseTotal) return null;
  return {price:round(n/p.baseTotal),unit:p.baseUnit,dimension:p.dimension};
}

// Numeric totals alone do not identify a pack: 1 gram, 1 milliliter,
// 1 sheet, and 1 each are different purchasing configurations. Unknown
// units intentionally compare only with the same normalized custom unit.
function packsEquivalent(a, b) {
  const x=parsePackSize(a), y=parsePackSize(b);
  if(!x?.parsed || !y?.parsed) return false;
  return x.dimension===y.dimension && x.baseUnit===y.baseUnit && x.baseTotal===y.baseTotal;
}

// Brand equality is case/space-insensitive; an empty brand never matches
// a locked one, since "unknown brand" is not the same as "this brand".
function brandsMatch(a, b) {
  const norm = v => String(v || "").trim().toLowerCase().replace(/\s+/g, " ");
  const x = norm(a), y = norm(b);
  return !!x && !!y && x === y;
}

function eachPrice(casePrice, pack) {
  const p=parsePackSize(pack), n=Number(casePrice);
  if (!p?.parsed || !Number.isFinite(n) || p.caseQty<=1) return null;
  return {price:round(n/p.caseQty,2),size:p.eachStr};
}

// Price per one unit of measure, for honest comparison across vendors
// whose packs are expressed differently (4/1 GAL against 6/32 FL OZ).
// `targetUnit` is the unit the client wants to read the price in; when
// it is absent, or belongs to a different dimension than this pack, the
// price is given per the pack's own unit instead - never silently
// converted across dimensions. Unknown units stay comparable only to
// the same unknown unit, same rule as parsePackSize.
function pricePerUnit(price, pack, targetUnit) {
  const p=parsePackSize(pack), n=Number(price);
  if (!p?.parsed || !Number.isFinite(n) || !p.total) return null;
  const target=targetUnit ? normalizeUnit(targetUnit) : p.unit;
  const def=UNIT_DEFINITIONS[target];
  if (def && def.dimension===p.dimension && p.baseTotal) {
    return {price:round(n/(p.baseTotal/def.factor),4),unit:target,dimension:p.dimension};
  }
  return {price:round(n/p.total,4),unit:p.unit,dimension:p.dimension};
}

// The units a client may choose to read a price in, for one dimension.
// Driven by UNIT_DEFINITIONS so adding a unit there is enough.
function unitsForDimension(dimension) {
  return Object.entries(UNIT_DEFINITIONS).filter(([,d])=>d.dimension===dimension).map(([code])=>code);
}

function extractComparableMeasurements(value) {
  const text=String(value||"").toLowerCase();
  const out=[];
  const re=new RegExp(MEASURE_RE.source, "g");
  let m;
  while ((m=re.exec(text))) {
    const x=measurement(Number(m[1]),m[2]);
    if (x) out.push(`${x.dimension}:${x.baseUnit}:${round(x.baseQuantity,3)}`);
  }
  const ratios=[...text.matchAll(/\b(\d+)\s*[\/x]\s*(\d+)\b/g)].map(m=>`ratio:${m[1]}x${m[2]}`);
  return new Set([...out,...ratios]);
}

function productCoreWords(value) {
  const unitWords=new Set([...UNIT_LOOKUP.keys(),...PACKAGING_ALIASES]);
  return normalizeForMatch(value).filter(w => !/^\d+(?:\.\d+)?$/.test(w) && !unitWords.has(w));
}

// Product identity is evaluated BEFORE overlap scores. This is intentionally
// industry-neutral: a shared generic term never proves two products identical.
// Organization vocabulary supplies contextual aliases; unrecognized defining
// words stay significant instead of being silently discarded.
function productIdentity(description) {
  const text=String(description||"");
  const measurements=extractComparableMeasurements(text);
  const clean=text.replace(new RegExp(MEASURE_RE.source,"gi")," ").replace(/\b\d+\s*[\/x]\s*\d+\b/gi," ");
  return { core:new Set(productCoreWords(clean)), measurements };
}
function compareProductIdentity(a,b) {
  const x=productIdentity(a), y=productIdentity(b);
  const shared=[...x.core].filter(w=>y.core.has(w));
  if(!shared.length) return {status:"different",reason:"No shared defining product terms"};
  const onlyA=[...x.core].filter(w=>!y.core.has(w));
  const onlyB=[...y.core].filter(w=>!x.core.has(w));
  // Different defining terms on BOTH sides cannot be explained by an
  // omitted field: "chicken base" vs "chicken breast", "breast" vs
  // "thigh", "tomato sauce" vs "tomato paste" remain separate.
  if(onlyA.length&&onlyB.length) return {status:"different",reason:`Conflicting product terms: ${onlyA.join(", ")} vs ${onlyB.join(", ")}`};
  if(onlyA.length||onlyB.length) return {status:"review",reason:`Unverified defining terms: ${[...onlyA,...onlyB].join(", ")}`};
  const aSpecs=x.measurements,bSpecs=y.measurements;
  if(aSpecs.size&&bSpecs.size){
    if(aSpecs.size!==bSpecs.size||[...aSpecs].some(z=>!bSpecs.has(z)))
      return {status:"review",reason:"Different or incomplete measurements / purchasing configurations"};
  }else if(aSpecs.size||bSpecs.size) return {status:"review",reason:"A measurement is missing from one description"};
  return {status:"same",reason:"All recognized defining product terms and measurements agree"};
}

function safeProductScore(a,b) {
  const aWords=new Set(productCoreWords(a)), bWords=new Set(productCoreWords(b));
  if (!aWords.size || !bWords.size) return 0;
  let shared=0;
  for (const w of aWords) if (bWords.has(w)) shared++;
  if (!shared) return 0;

  const specsA=extractComparableMeasurements(a), specsB=extractComparableMeasurements(b);
  if (specsA.size && specsB.size && ![...specsA].some(x=>specsB.has(x))) return 0;

  const subset=shared/Math.min(aWords.size,bWords.size);
  const union=shared/(aWords.size+bWords.size-shared);
  // Keeps legitimate abbreviated/subset descriptions viable while penalizing
  // near-neighbors such as "chicken breast" vs "chicken thigh".
  return round(0.65*subset + 0.35*union,4);
}

// Picks which existing catalog item a vendor description belongs to.
// "strict" items only accept auto-link-grade matches: a partial
// resemblance must never attach to them, because that is exactly how two
// genuinely different products get merged. "flexible" (the default) also
// accepts review-grade matches, flagged as "similar" for a person to
// confirm. Below the review floor nothing matches and the caller creates
// a new item.
function bestCatalogMatch(description, catalogItems=[], policy=MATCH_POLICY) {
  let best=null,bestScore=0,bestTrack=null;
  for(const ci of catalogItems){
    if(!ci?.name) continue;
    const identity=compareProductIdentity(description,ci.name);
    if(identity.status==="different") continue;
    const score=safeProductScore(description,ci.name);
    if(identity.status==="same" && score>=policy.autoLink && score>bestScore){
      best=ci; bestScore=score; bestTrack="exact";
    }
  }
  // Prefer an established identity over any fuzzy candidate.
  if(best) return {catalogItem:best,score:bestScore,track:bestTrack};
  for(const ci of catalogItems){
    if(!ci?.name||ci.matching_behavior==="strict") continue;
    const identity=compareProductIdentity(description,ci.name);
    if(identity.status!=="review") continue;
    const score=safeProductScore(description,ci.name);
    if(score>=policy.reviewFloor&&score>bestScore){best=ci;bestScore=score;bestTrack="similar";}
  }
  // A similar candidate is only a suggestion: CALLERS MUST NOT link it.
  return best?{catalogItem:best,score:bestScore,track:bestTrack}:null;
}

function bestInvoiceMatch(description,candidates=[],threshold=MATCH_POLICY.autoLink) {
  let best=null,bestScore=0;
  for(const candidate of candidates){
    if(!candidate?.description||compareProductIdentity(description,candidate.description).status!=="same") continue;
    const score=safeProductScore(description,candidate.description);
    if(score>=threshold&&score>bestScore){best=candidate;bestScore=score;}
  }
  return best?{vendorItem:best,score:bestScore}:null;
}

// A current quotation is NOT a past invoice charge. Expiry can be client-
// configured or manual, but an explicit vendor validity date always wins.
function quoteStatus(item,settings={},now=new Date()){
  if(!item||item.price_source==="invoice") return "invoice_only";
  if(item.price_unavailable||item.price==null||!Number.isFinite(Number(item.price))||Number(item.price)<=0) return "unavailable";
  if(item.price_expired_at) return "expired";
  const date=now instanceof Date?now:new Date(now);
  if(!Number.isFinite(date.getTime())) return "expired";
  const localDate=`${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,"0")}-${String(date.getDate()).padStart(2,"0")}`;
  if(item.price_quote_valid_until && String(item.price_quote_valid_until).slice(0,10)<localDate) return "expired";
  const days=Number(settings?.price_refresh_days);
  const mode=settings?.price_refresh_mode||(days>0?"automatic":"manual");
  if(mode==="automatic"&&days>0){
    const updated=new Date(item.last_updated);
    if(!item.last_updated||!Number.isFinite(updated.getTime())) return "expired";
    if((date.getTime()-updated.getTime())/(86400000)>days) return "expired";
  }
  return "current";
}

export {
  MATCH_POLICY, UNIT_DEFINITIONS, configureVocabulary, normalizeUnit, measurement, parsePackSize, packsEquivalent, normalizedPrice, eachPrice,
  unitsForDimension, pricePerUnit, brandsMatch,
  normalizeForMatch, wordsMatch, classifyCategory, nextCategoryRange,
  safeProductScore, productIdentity, compareProductIdentity, quoteStatus, bestCatalogMatch, bestInvoiceMatch,
};
