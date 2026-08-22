// ════════════════════════════════════════════════════════════════════
// KERDOS INGESTION MODULE
// ════════════════════════════════════════════════════════════════════
//
// PURPOSE
// Turns any pasted or dropped document (price list or invoice, in any
// vendor's format) into clean structured rows: description, code,
// quantity, unit price, line total.
//
// DESIGN
// Unlike the old line-by-line guessing, this works in four stages:
//   1. SNIFF STRUCTURE  — figure out how the document is delimited
//                          (tabs, commas, or fixed-width columns)
//   2. FIND HEADER       — locate the real header row, even if it's
//                          mid-document or oddly labeled
//   3. MAP COLUMNS       — match header text to meaning (description,
//                          qty, price, total, etc). If no header is
//                          found, fall back to profiling each column's
//                          data to infer its role.
//   4. EXTRACT ROWS      — read each data row using that structure,
//                          skipping page furniture (totals, addresses,
//                          disclaimers) rather than misreading it.
//
// Nothing here calls any AI or external service. It is plain,
// deterministic JavaScript — same input always produces same output.
//
// This module does NOT touch the database or the UI. It only turns
// text into structured data. The caller decides what to do with the
// result (show a review screen, save it, flag it, etc).
// ════════════════════════════════════════════════════════════════════


// ── Column roles we understand ─────────────────────────────────────
// Every column in a document gets matched to one of these, or "unknown".
const COLUMN_ROLES = {
  description: ["item", "description", "desc", "product", "name"],
  code:        ["item no", "item#", "sku", "code", "cust item no", "item number"],
  qty:         ["qty", "quantity", "count"],
  packSize:    ["unit", "size", "pack", "uom", "pack size"],
  price:       ["price", "unit price", "cost", "unit cost", "rate"],
  amount:      ["amount", "total", "ext", "extended", "ext price", "ext. price", "line total"],
};

// Rows that are clearly not item data — invoice/pricelist "furniture".
const JUNK_ROW_PATTERNS = [
  /^page \d+ of \d+/i,
  /^total\b/i,
  /^subtotal\b/i,
  /^grand total/i,
  /^cash total/i,
  /^card total/i,
  /^signature/i,
  /^note:/i,
  /^terms:/i,
  /^ship to/i,
  /^sold to/i,
  /^tel:|^fax:/i,
  /^order\b/i,
  /^[A-Z]{1,4}:\s*$/,
  /perishable agricultural/i,
  /paca trust/i,
  /return policy/i,
];

// A narrower set — these genuinely mark "the item table has ended"
// (totals, signatures, legal disclaimers). Unlike the broader junk
// list above, these are only used to find where the table stops; a
// stray order-reference line or a bare unit-code label can appear in
// the middle of a table without meaning it's over.
const TABLE_END_PATTERNS = [
  /^total\b/i,
  /^subtotal\b/i,
  /^grand total/i,
  /^cash total/i,
  /^card total/i,
  /^signature/i,
  /^note:/i,
  /perishable agricultural/i,
  /paca trust/i,
  /return policy/i,
];

function isTableEndMarker(cells) {
  const joined = cells.join(" ").trim();
  return TABLE_END_PATTERNS.some(pat => pat.test(joined));
}


// ── Small utilities ─────────────────────────────────────────────────

function parseMoney(str) {
  if (str == null) return null;
  const cleaned = String(str).replace(/[$,]/g, "").trim();
  if (!/^-?\d+(\.\d+)?$/.test(cleaned)) return null;
  const n = parseFloat(cleaned);
  if (isNaN(n)) return null;
  return n;
}

function isJunkRow(cells) {
  const joined = cells.join(" ").trim();
  if (joined.length < 3) return true;
  return JUNK_ROW_PATTERNS.some(pat => pat.test(joined));
}

function median(numbers) {
  if (!numbers.length) return 0;
  const sorted = [...numbers].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}


// ── Stage 1: Sniff document structure ───────────────────────────────
// Tries tab-split, comma-split, and multi-space-split on every line,
// and picks whichever delimiter gives the most CONSISTENT cell count
// across lines (a real table has the same number of columns row to
// row; noise doesn't).

function sniffDelimiter(lines) {
  const candidates = [
    { name: "tab", regex: /\t/ },
    { name: "comma", regex: /,(?=(?:[^"]*"[^"]*")*[^"]*$)/ },
    { name: "multispace", regex: /\s{2,}/ },
  ];

  let best = { name: null, consistency: -1 };

  for (const cand of candidates) {
    const linesWithDelim = lines.filter(l => cand.regex.test(l));
    // A real table uses this delimiter on most lines, not just a
    // scattered few (a handful of junk lines with incidental double
    // spaces shouldn't be mistaken for a whole tabular document).
    const coverage = linesWithDelim.length / lines.length;
    if (coverage < 0.6) continue;

    const counts = linesWithDelim.map(l => l.split(cand.regex).length);
    const modeCount = mostCommon(counts);
    const matchingRows = counts.filter(c => c === modeCount).length;
    const consistency = matchingRows / counts.length;

    if (modeCount >= 2 && consistency > best.consistency) {
      best = { name: cand.name, consistency, modeCount };
    }
  }

  // No reliable tabular delimiter — this is typical of text pulled out
  // of a PDF, where columns are just single spaces apart with no
  // consistent structure. We deliberately return null rather than
  // forcing a bad split; the caller switches to "freeform" mode instead
  // of pretending this is a table.
  if (best.consistency < 0.5) return null;
  return best.name;
}

function mostCommon(arr) {
  const freq = new Map();
  for (const v of arr) freq.set(v, (freq.get(v) || 0) + 1);
  let best = arr[0], bestCount = 0;
  for (const [v, c] of freq) if (c > bestCount) { best = v; bestCount = c; }
  return best;
}

function splitRow(line, delimiterName) {
  if (delimiterName === "tab") return line.split(/\t/);
  if (delimiterName === "comma") return line.split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/);
  return line.split(/\s{2,}/);
}

function cleanCell(cell) {
  return cell.trim().replace(/^"|"$/g, "");
}


// ── Stage 2: Find the header row ────────────────────────────────────
// Scores each of the first several rows on how "header-like" it is:
// mostly text, contains recognized header words, and its column count
// matches the most common column count in the document (meaning it's
// structurally part of the table, not a stray title line).

function scoreHeaderRow(cells, modeColumnCount) {
  if (cells.length < 2) return 0;

  let score = 0;

  // Structural fit: does this row have the same number of columns
  // as most other rows in the document?
  if (cells.length === modeColumnCount) score += 3;

  // Keyword match: how many cells look like known header words?
  const allKeywords = Object.values(COLUMN_ROLES).flat();
  const keywordHits = cells.filter(cell => {
    const lower = cell.toLowerCase().trim();
    return allKeywords.some(kw => lower === kw || lower.includes(kw));
  }).length;
  score += keywordHits * 2;

  // Text-heaviness: headers are mostly words, not numbers.
  const numericCells = cells.filter(c => parseMoney(c) !== null).length;
  if (numericCells === 0) score += 1;

  return score;
}

function findHeaderRow(grid) {
  const columnCounts = grid.map(r => r.cells.length);
  const modeColumnCount = mostCommon(columnCounts);

  const searchLimit = Math.min(grid.length, 15);
  let best = { index: -1, score: 0 };

  for (let i = 0; i < searchLimit; i++) {
    const score = scoreHeaderRow(grid[i].cells, modeColumnCount);
    if (score > best.score) best = { index: i, score };
  }

  // Require a minimum score so we don't mistake a random line for
  // a header when the document genuinely has none.
  return best.score >= 4 ? best.index : -1;
}


// ── Stage 3: Map columns to roles ───────────────────────────────────

function mapColumnsFromHeader(headerCells) {
  const map = {}; // role -> column index
  headerCells.forEach((cell, idx) => {
    const lower = cell.toLowerCase().trim();
    for (const [role, keywords] of Object.entries(COLUMN_ROLES)) {
      if (map[role] !== undefined) continue; // first match wins
      if (keywords.some(kw => lower === kw || lower.includes(kw))) {
        map[role] = idx;
      }
    }
  });
  return map;
}

// Profiles each column's DATA (rather than its header text) to guess
// its role, based on shape: how numeric it is, how long the text is,
// etc. Used in two situations:
//   - no header row exists at all
//   - a header row exists, but its label doesn't match any known word
//     (e.g. Excel's generic "Column2" instead of "Price") — in that
//     case we only use this to fill in whatever roles are still
//     missing, not to override roles we already matched by name.
function profileColumns(grid, dataStartIndex) {
  const dataRows = grid.slice(dataStartIndex);
  const columnCount = mostCommon(dataRows.map(r => r.cells.length));
  const usableRows = dataRows.filter(r => r.cells.length === columnCount);
  if (!usableRows.length) return {};

  const columnStats = [];
  for (let col = 0; col < columnCount; col++) {
    const values = usableRows.map(r => r.cells[col]).filter(v => v !== "");
    if (!values.length) continue;
    const numericValues = values.map(parseMoney).filter(v => v !== null);
    columnStats.push({
      col,
      numericRatio: numericValues.length / values.length,
      avgLength: values.reduce((s, v) => s + v.length, 0) / values.length,
    });
  }

  const map = {};
  const numericCols = columnStats.filter(c => c.numericRatio > 0.6);
  if (numericCols.length >= 2) {
    map.price = numericCols[numericCols.length - 2].col;
    map.amount = numericCols[numericCols.length - 1].col;
  } else if (numericCols.length === 1) {
    map.price = numericCols[0].col;
  }

  const textCols = columnStats.filter(c => c.numericRatio < 0.3);
  if (textCols.length) {
    const widest = textCols.reduce((a, b) => (b.avgLength > a.avgLength ? b : a));
    map.description = widest.col;
  }

  return map;
}

// Fills in any role the header-text pass didn't find, using the data
// profile — without overwriting anything already matched by name.
function fillMissingRolesFromData(columnMap, grid, dataStartIndex) {
  const usedColumns = new Set(Object.values(columnMap));
  const dataProfile = profileColumns(grid, dataStartIndex);

  const filled = { ...columnMap };
  for (const [role, col] of Object.entries(dataProfile)) {
    if (filled[role] !== undefined) continue;   // already matched by header text
    if (usedColumns.has(col)) continue;          // column already claimed by another role
    filled[role] = col;
    usedColumns.add(col);
  }
  return filled;
}


// ── Stage 4: Extract structured rows ────────────────────────────────

function extractRow(cells, columnMap) {
  const get = role => (columnMap[role] !== undefined ? cells[columnMap[role]] : undefined);

  const priceRaw = get("price");
  const amountRaw = get("amount");
  const price = parseMoney(priceRaw);
  const amount = parseMoney(amountRaw);

  // A row needs at least one valid money value to be usable.
  if (price === null && amount === null) return null;

  let description = (get("description") || "").trim();
  // If we don't know which column is the description, fall back to
  // the longest non-numeric cell in the row.
  if (!description) {
    description = cells
      .filter(c => parseMoney(c) === null)
      .sort((a, b) => b.length - a.length)[0] || "";
  }
  description = description.trim();
  if (description.length < 2) return null;

  const code = (get("code") || "").trim() || null;
  const packSize = (get("packSize") || "").trim() || null;
  const qty = parseMoney(get("qty"));

  return {
    code,
    description: description.slice(0, 120),
    packSize,
    qty: qty !== null ? qty : null,
    price: price !== null ? price : amount,
    amount: amount !== null ? amount : price,
  };
}


// ── Freeform extraction (no reliable column delimiter) ──────────────
// PDF text commonly comes out as single-space-separated words with no
// consistent column structure — "6 GREENLEAF LETTUCE GREENLEAF 24CT
// (USA) 23.00 138.00" — a real table, but not one a delimiter can
// split cleanly. Instead of forcing it into fake columns, we read each
// line by its actual shape: numbers at the end are price/amount, an
// optional short code near the start, everything else is the
// description.

function findMoneyTokens(line) {
  const moneyRe = /\$?\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?\b/g;
  return [...line.matchAll(moneyRe)].map(m => ({
    value: parseMoney(m[0]),
    index: m.index,
    length: m[0].length,
  })).filter(t => t.value !== null);
}

// Some invoices split a line's true extended price onto its own
// separate line just below the item's main details — a column that
// got pushed onto its own row by the document's layout. We stitch
// these back together before extraction, so the "last two numbers on
// the line = price, then amount" logic sees both values together.
function mergeBareDollarContinuations(lines) {
  const bareDollar = /^\$?\d{1,3}(?:,\d{3})*\.\d{2}$/;
  const merged = [];
  for (const line of lines) {
    const trimmed = line.trim();
    const prev = merged[merged.length - 1];
    if (
      bareDollar.test(trimmed) &&
      prev !== undefined &&
      !isJunkRow([prev]) &&
      findMoneyTokens(prev).length > 0
    ) {
      merged[merged.length - 1] = prev + " " + trimmed;
    } else {
      merged.push(line);
    }
  }
  return merged;
}


function extractRowFreeform(line) {
  const tokens = findMoneyTokens(line);
  if (tokens.length === 0) return null;

  // The rightmost money-looking number is the line total; the one
  // before it (if present) is the unit price. A lone number is
  // treated as the price.
  const amountTok = tokens[tokens.length - 1];
  const priceTok = tokens.length >= 2 ? tokens[tokens.length - 2] : amountTok;

  if (amountTok.value <= 0 || amountTok.value > 1000000) return null;

  // Everything before the price token is candidate description text.
  let head = line.slice(0, priceTok.index).trim();

  // Strip up to two leading quantity numbers — some formats show an
  // ordered quantity and a separate shipped/adjustment quantity back
  // to back. Only the first is kept as the real quantity.
  let qty = null;
  for (let pass = 0; pass < 2; pass++) {
    const qtyMatch = head.match(/^(\d+(?:\.\d+)?)\s+/);
    if (!qtyMatch) break;
    if (pass === 0) qty = parseFloat(qtyMatch[1]);
    head = head.slice(qtyMatch[0].length);
  }

  // Strip a short unit-type token (CS, EA, LB, BX...) when it comes
  // right before the real item code, so it isn't mistaken for the
  // code itself.
  const unitTypeMatch = head.match(/^([A-Z]{2,4})\s+(?=[A-Z0-9]{4,})/);
  if (unitTypeMatch) {
    head = head.slice(unitTypeMatch[0].length);
  }

  // Strip a leading item code, e.g. "GREENLEAF LETTUCE GREENLEAF..."
  // A code looks like a short all-caps/alphanumeric token immediately
  // followed by more descriptive text.
  let code = null;
  const codeMatch = head.match(/^([A-Z0-9\-]{3,14})\s+/);
  if (codeMatch) {
    code = codeMatch[1];
    head = head.slice(codeMatch[0].length);
  }

  const description = head.trim();
  if (description.length < 2) return null;

  return {
    code,
    description: description.slice(0, 120),
    packSize: null,
    qty,
    price: priceTok.value,
    amount: amountTok.value,
  };
}


// Some PDFs place an item's code, or the tail end of a wrapped
// description, on its own line — visually part of the same row, but
// positioned slightly differently so our line-grouping sees it as
// separate. Before extracting rows, we merge any short, price-less
// line into the row it visually belongs to: any such "orphan" line
// appearing between one priced line and the next belongs to the row
// that came before it.
function mergeOrphanLines(lines) {
  const merged = [];
  let pendingBase = null;
  let pendingExtras = [];

  const flush = () => {
    if (pendingBase !== null) {
      merged.push([pendingBase, ...pendingExtras].join(" "));
      pendingBase = null;
      pendingExtras = [];
    }
  };

  for (const line of lines) {
    const hasMoney = findMoneyTokens(line).length > 0;
    const junk = isJunkRow([line]);

    if (junk) {
      flush();
      merged.push(line); // pass through untouched; filtered downstream
    } else if (hasMoney) {
      flush();
      pendingBase = line;
    } else if (pendingBase !== null) {
      pendingExtras.push(line); // orphan text belongs to the pending row
    } else {
      merged.push(line); // no row to attach to; left as-is
    }
  }
  flush();
  return merged;
}


// Even free-flowing PDF text usually has a real header line somewhere
// ("Qty Item no Description Price Amount") and a clear end ("Total",
// "Signature", disclaimers). We find both boundaries first, so we only
// ever extract from inside the actual item table — never from invoice
// numbers, phone numbers, zip codes, or footer text that happen to
// contain digits.
function findFreeformHeaderIndex(lines) {
  const allKeywords = Object.values(COLUMN_ROLES).flat();
  const searchLimit = Math.min(lines.length, 20);
  for (let i = 0; i < searchLimit; i++) {
    // A line that itself parses as a labeled data row (e.g. "Item: X
    // Qty: 5 Price: $10") is data, not a header — even though it
    // contains header-like keywords, it also has real values attached.
    if (extractLabeledFields(lines[i])) continue;
    const lower = lines[i].toLowerCase();
    const hits = allKeywords.filter(kw => lower.includes(kw)).length;
    if (hits >= 3) return i;
  }
  return -1;
}


// ── Date detection ───────────────────────────────────────────────────
// Finds the invoice/document date, handling several common formats,
// by scanning near the top of the document (where dates typically
// appear) and normalizing to YYYY-MM-DD. Prefers dates that sit near
// a relevant word ("date", "invoice", "delivered", "ship") over any
// other date-looking number nearby (e.g. a phone number fragment).

const MONTH_NAMES = {
  jan:1, january:1, feb:2, february:2, mar:3, march:3, apr:4, april:4,
  may:5, jun:6, june:6, jul:7, july:7, aug:8, august:8,
  sep:9, sept:9, september:9, oct:10, october:10, nov:11, november:11,
  dec:12, december:12,
};

function pad2(n) { return String(n).padStart(2, "0"); }

function isValidDate(y, m, d) {
  return m >= 1 && m <= 12 && d >= 1 && d <= 31 && y >= 2000 && y <= 2100;
}

function findDate(text) {
  const lines = text.split("\n").slice(0, 30);
  const monthPattern = "jan|january|feb|february|mar|march|apr|april|may|jun|june|jul|july|aug|august|sep|sept|september|oct|october|nov|november|dec|december";

  const patterns = [
    { re: /\b(20\d{2})[\/\-](\d{1,2})[\/\-](\d{1,2})\b/g,
      toISO: m => isValidDate(+m[1], +m[2], +m[3]) ? `${m[1]}-${pad2(m[2])}-${pad2(m[3])}` : null },
    { re: /\b(\d{1,2})[\/\-](\d{1,2})[\/\-](20\d{2})\b/g,
      toISO: m => isValidDate(+m[3], +m[1], +m[2]) ? `${m[3]}-${pad2(m[1])}-${pad2(m[2])}` : null },
    { re: new RegExp(`\\b(${monthPattern})\\.?\\s+(\\d{1,2}),?\\s+(20\\d{2})\\b`, "gi"),
      toISO: m => { const mo = MONTH_NAMES[m[1].toLowerCase()]; return mo && isValidDate(+m[3], mo, +m[2]) ? `${m[3]}-${pad2(mo)}-${pad2(m[2])}` : null; } },
    { re: new RegExp(`\\b(\\d{1,2})\\s+(${monthPattern})\\.?,?\\s+(20\\d{2})\\b`, "gi"),
      toISO: m => { const mo = MONTH_NAMES[m[2].toLowerCase()]; return mo && isValidDate(+m[3], mo, +m[1]) ? `${m[3]}-${pad2(mo)}-${pad2(m[1])}` : null; } },
  ];

  const candidates = [];
  lines.forEach((line, idx) => {
    for (const pat of patterns) {
      pat.re.lastIndex = 0;
      let m;
      while ((m = pat.re.exec(line))) {
        const iso = pat.toISO(m);
        if (iso) {
          const nearKeyword = /date|invoice|delivered|ship/i.test(line);
          candidates.push({ iso, lineIndex: idx, nearKeyword });
        }
      }
    }
  });

  if (!candidates.length) return null;

  candidates.sort((a, b) => {
    if (a.nearKeyword !== b.nearKeyword) return a.nearKeyword ? -1 : 1;
    return a.lineIndex - b.lineIndex;
  });

  return candidates[0].iso;
}


// ── Invoice number detection ────────────────────────────────────────
// Looks near the top of the document for something labeled as the
// invoice number ("Invoice #4821", "INVOICE 330228", "Inv No: X-99").
// Works off the label itself, not any particular vendor's format.

function findInvoiceNumber(text) {
  const lines = text.split("\n").slice(0, 20);
  const patterns = [
    /\binvoice\s*(?:no\.?|number|#)?\s*:?\s*([A-Z0-9][A-Z0-9\-]{2,14})\b/i,
    /\binv(?!oice)\.?\s*(?:no\.?|number|#)?\s*:?\s*([A-Z0-9][A-Z0-9\-]{2,14})\b/i,
  ];
  const nonAnswers = /^(date|no|number|for|the|price|amount|list)$/i;

  for (const line of lines) {
    for (const pat of patterns) {
      const m = line.match(pat);
      if (m && !nonAnswers.test(m[1])) {
        return m[1];
      }
    }
  }
  return null;
}


// ── Inline-labeled-field extraction ─────────────────────────────────
// Some documents have no real header row at all, because whoever made
// them typed field labels directly into each line instead of using
// columns — e.g. "Item: Paper Towels Qty: 12 Price: $3.50". This reads
// a line by its explicit labels rather than guessing from position,
// which is more reliable than plain number-guessing whenever labels
// are actually present.
function extractLabeledFields(line) {
  const fieldPatterns = {
    code: /(?:sku|item\s*#|item\s*no\.?)\s*[:#]\s*([A-Z0-9\-]{2,15})/i,
    qty: /(?:qty|quantity)\s*[:#]\s*(\d+(?:\.\d+)?)/i,
    price: /(?:unit\s*price|price|cost)\s*[:#]\s*\$?(\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?)/i,
    amount: /(?:total|amount|ext(?:ended)?(?:\s*price)?)\s*[:#]\s*\$?(\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?)/i,
    description: /(?:item|description|desc|product)\s*[:#]\s*(.+)/i,
  };

  const matches = {};
  let hits = 0;
  for (const [key, re] of Object.entries(fieldPatterns)) {
    const m = line.match(re);
    if (m) {
      const valueStart = m.index + m[0].indexOf(m[1]);
      matches[key] = { value: m[1].trim(), labelStart: m.index, valueStart };
      hits++;
    }
  }

  if (hits < 2) return null;

  const price = matches.price ? parseMoney(matches.price.value) : null;
  const amount = matches.amount ? parseMoney(matches.amount.value) : null;
  if (price === null && amount === null) return null;

  let description = matches.description ? matches.description.value : "";
  if (description && matches.description) {
    const laterLabelStarts = Object.entries(matches)
      .filter(([k]) => k !== "description")
      .map(([, v]) => v.labelStart)
      .filter(s => s > matches.description.valueStart);
    if (laterLabelStarts.length) {
      const cutoff = Math.min(...laterLabelStarts) - matches.description.valueStart;
      description = description.slice(0, cutoff);
    }
  }
  description = description.trim().replace(/[,;:]+$/, "");
  if (description.length < 2) return null;

  return {
    code: matches.code ? matches.code.value : null,
    description: description.slice(0, 120),
    packSize: null,
    qty: matches.qty ? parseFloat(matches.qty.value) : null,
    price: price !== null ? price : amount,
    amount: amount !== null ? amount : price,
  };
}


// ── Main entry point ─────────────────────────────────────────────────
//
// parseDocument(text) -> {
//   headerFound: boolean,
//   columnMap: { role: columnIndex, ... },
//   rows: [ { code, description, packSize, qty, price, amount }, ... ],
//   skipped: [ { line, reason }, ... ]   // for transparency in the review UI
// }

function parseDocument(text) {
  const rawLines = text.split("\n").map(l => l.trimEnd()).filter(l => l.trim().length > 0);

  const delimiter = sniffDelimiter(rawLines);

  // ── FREEFORM MODE ──
  // No reliable column delimiter was found (typical of PDF-extracted
  // text). Read each line by its shape instead of forcing fake columns.
  if (delimiter === null) {
    const stitchedLines = mergeBareDollarContinuations(rawLines);
    const headerIdx = findFreeformHeaderIndex(stitchedLines);
    let dataStart = headerIdx !== -1 ? headerIdx + 1 : 0;

    // Junk can appear both before the real item table (addresses,
    // order-reference lines) and after it (totals, disclaimers). A
    // junk line only marks "the end" once real item data has actually
    // started — so first find where a genuine, parseable row begins.
    for (let i = dataStart; i < stitchedLines.length; i++) {
      if (isJunkRow([stitchedLines[i]])) continue;
      if (extractLabeledFields(stitchedLines[i]) || extractRowFreeform(stitchedLines[i])) {
        dataStart = i;
        break;
      }
    }

    // Now find where the item table ends: the first genuine end-marker
    // line after real data has begun (Total, Signature, disclaimers).
    // Ordinary noise (a stray order line, a bare unit label) can occur
    // mid-table without meaning the table has ended.
    let dataEnd = stitchedLines.length;
    for (let i = dataStart; i < stitchedLines.length; i++) {
      if (isTableEndMarker([stitchedLines[i]])) { dataEnd = i; break; }
    }

    const beforeCount = dataStart;
    const afterCount = stitchedLines.length - dataEnd;
    const itemLines = stitchedLines.slice(dataStart, dataEnd);
    const mergedLines = mergeOrphanLines(itemLines);

    const rows = [];
    const skipped = [];
    for (const line of mergedLines) {
      if (isJunkRow([line])) {
        skipped.push({ line, reason: "looks like a total/note/address line" });
        continue;
      }
      const row = extractLabeledFields(line) || extractRowFreeform(line);
      if (!row) {
        skipped.push({ line, reason: "couldn't find a valid price or description" });
        continue;
      }
      rows.push(row);
    }
    if (beforeCount > 0) {
      skipped.unshift({ line: `(${beforeCount} header/address lines before the item table)`, reason: "outside the item table" });
    }
    if (afterCount > 0) {
      skipped.push({ line: `(${afterCount} footer/note lines after the item table)`, reason: "outside the item table" });
    }
    return { mode: "freeform", headerFound: headerIdx !== -1, columnMap: {}, rows, skipped };
  }

  // ── TABULAR MODE ──
  // A real, consistent delimiter was found — treat this as a proper
  // table with columns.
  const grid = rawLines.map(line => ({
    line,
    cells: splitRow(line, delimiter).map(cleanCell),
  }));

  const headerIndex = findHeaderRow(grid);
  const headerFound = headerIndex !== -1;
  const dataStart = headerFound ? headerIndex + 1 : 0;

  // Start with whatever the header text tells us (most reliable when
  // it's clearly labeled), then use the actual column data to fill in
  // any role the header didn't clearly name — covers both "no header
  // at all" and "header exists but is mislabeled" (e.g. "Column2").
  const columnMap = fillMissingRolesFromData(
    headerFound ? mapColumnsFromHeader(grid[headerIndex].cells) : {},
    grid,
    dataStart
  );
  const rows = [];
  const skipped = [];

  for (let i = dataStart; i < grid.length; i++) {
    const { line, cells } = grid[i];

    if (isJunkRow(cells)) {
      skipped.push({ line, reason: "looks like a total/note/address line" });
      continue;
    }

    const row = extractRow(cells, columnMap);
    if (!row) {
      skipped.push({ line, reason: "couldn't find a valid price or description" });
      continue;
    }

    rows.push(row);
  }

  return { mode: "tabular", headerFound, columnMap, rows, skipped };
}


// Exported for use elsewhere in the app.
export { parseDocument, findDate, findInvoiceNumber };
