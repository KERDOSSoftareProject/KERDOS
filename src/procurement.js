// KERDOS deterministic procurement primitives.
// Pure functions only: no UI, database, client, vendor, or industry assumptions.

// One named place for the auto-link cutoff instead of the same magic
// number repeated at each call site. At or above this score two
// descriptions are treated as the same product and linked automatically;
// below it, the pair is kept for human review rather than silently
// merged.
const MATCH_POLICY = Object.freeze({ autoLink: 0.85 });

const STOPWORDS = new Set(["the","a","an","of","and","or","with","in","new"]);

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

const PACKAGING_ALIASES = new Set([
  "case","cases","cs","carton","cartons","box","boxes","bx","bag","bags",
  "pack","packs","pk","pallet","pallets","roll","rolls","container","containers",
  "bottle","bottles","can","cans","jar","jars","tub","tubs","tray","trays",
]);

const UNIT_LOOKUP = new Map();
for (const [code, def] of Object.entries(UNIT_DEFINITIONS)) {
  UNIT_LOOKUP.set(code.toLowerCase(), code);
  for (const alias of def.aliases) UNIT_LOOKUP.set(alias.toLowerCase(), code);
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
    .map(canonicalWord).filter(w => w.length > 1 && !STOPWORDS.has(w));
}

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
  const unit="([a-z]+(?:\\s+oz)?)";
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

function eachPrice(casePrice, pack) {
  const p=parsePackSize(pack), n=Number(casePrice);
  if (!p?.parsed || !Number.isFinite(n) || p.caseQty<=1) return null;
  return {price:round(n/p.caseQty,2),size:p.eachStr};
}

function extractComparableMeasurements(value) {
  const text=String(value||"").toLowerCase();
  const out=[];
  const re=/\b(\d+(?:\.\d+)?)\s*(fl\s*oz|floz|lbs?|pounds?|oz|ounces?|kgs?|kilograms?|grams?|gals?|gallons?|qts?|quarts?|pts?|pints?|liters?|litres?|ml|mm|cm|meters?|metres?|inches?|inch|in|feet|foot|ft|yards?|yd|dozen|doz|dz|each|ea|ct|count)\b/g;
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

function bestInvoiceMatch(description,candidates=[],threshold=MATCH_POLICY.autoLink) {
  let best=null,bestScore=0;
  for (const candidate of candidates) {
    if (!candidate?.description) continue;
    const score=safeProductScore(description,candidate.description);
    if (score>=threshold && score>bestScore) { best=candidate; bestScore=score; }
  }
  return best ? {vendorItem:best,score:bestScore} : null;
}

export {
  MATCH_POLICY, UNIT_DEFINITIONS, normalizeUnit, measurement, parsePackSize, normalizedPrice, eachPrice,
  normalizeForMatch, wordsMatch, classifyCategory, nextCategoryRange,
  safeProductScore, bestInvoiceMatch,
};
