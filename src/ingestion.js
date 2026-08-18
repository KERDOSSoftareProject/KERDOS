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
  /^\*{2,}/,
  /perishable agricultural/i,
  /paca trust/i,
  /return policy/i,
];


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
    const counts = lines
      .filter(l => cand.regex.test(l))
      .map(l => l.split(cand.regex).length);
    if (counts.length < 2) continue;

    const modeCount = mostCommon(counts);
    const matchingRows = counts.filter(c => c === modeCount).length;
    const consistency = matchingRows / counts.length;

    if (modeCount >= 2 && consistency > best.consistency) {
      best = { name: cand.name, consistency, modeCount };
    }
  }

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

function scoreHeaderRow(cells, modeColumnCount) {
  if (cells.length < 2) return 0;

  let score = 0;

  if (cells.length === modeColumnCount) score += 3;

  const allKeywords = Object.values(COLUMN_ROLES).flat();
  const keywordHits = cells.filter(cell => {
    const lower = cell.toLowerCase().trim();
    return allKeywords.some(kw => lower === kw || lower.includes(kw));
  }).length;
  score += keywordHits * 2;

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

  return best.score >= 4 ? best.index : -1;
}


// ── Stage 3: Map columns to roles ───────────────────────────────────

function mapColumnsFromHeader(headerCells) {
  const map = {};
  headerCells.forEach((cell, idx) => {
    const lower = cell.toLowerCase().trim();
    for (const [role, keywords] of Object.entries(COLUMN_ROLES)) {
      if (map[role] !== undefined) continue;
      if (keywords.some(kw => lower === kw || lower.includes(kw))) {
        map[role] = idx;
      }
    }
  });
  return map;
}

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

function fillMissingRolesFromData(columnMap, grid, dataStartIndex) {
  const usedColumns = new Set(Object.values(columnMap));
  const dataProfile = profileColumns(grid, dataStartIndex);

  const filled = { ...columnMap };
  for (const [role, col] of Object.entries(dataProfile)) {
    if (filled[role] !== undefined) continue;
    if (usedColumns.has(col)) continue;
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

  if (price === null && amount === null) return null;

  let description = (get("description") || "").trim();
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

function findMoneyTokens(line) {
  const moneyRe = /\$?\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?\b/g;
  return [...line.matchAll(moneyRe)].map(m => ({
    value: parseMoney(m[0]),
    index: m.index,
    length: m[0].length,
  })).filter(t => t.value !== null);
}

function extractRowFreeform(line) {
  const tokens = findMoneyTokens(line);
  if (tokens.length === 0) return null;

  const amountTok = tokens[tokens.length - 1];
  const priceTok = tokens.length >= 2 ? tokens[tokens.length - 2] : amountTok;

  if (amountTok.value <= 0 || amountTok.value > 1000000) return null;

  let head = line.slice(0, priceTok.index).trim();

  let qty = null;
  const qtyMatch = head.match(/^(\d+(?:\.\d+)?)\s+/);
  if (qtyMatch) {
    qty = parseFloat(qtyMatch[1]);
    head = head.slice(qtyMatch[0].length);
  }

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


// ── Main entry point ─────────────────────────────────────────────────
//
// parseDocument(text) -> {
//   mode: "tabular" | "freeform",
//   headerFound: boolean,
//   columnMap: { role: columnIndex, ... },
//   rows: [ { code, description, packSize, qty, price, amount }, ... ],
//   skipped: [ { line, reason }, ... ]
// }

function parseDocument(text) {
  const rawLines = text.split("\n").map(l => l.trimEnd()).filter(l => l.trim().length > 0);

  const delimiter = sniffDelimiter(rawLines);

  if (delimiter === null) {
    const rows = [];
    const skipped = [];
    for (const line of rawLines) {
      const cellsForJunkCheck = [line];
      if (isJunkRow(cellsForJunkCheck)) {
        skipped.push({ line, reason: "looks like a total/note/address line" });
        continue;
      }
      const row = extractRowFreeform(line);
      if (!row) {
        skipped.push({ line, reason: "couldn't find a valid price or description" });
        continue;
      }
      rows.push(row);
    }
    return { mode: "freeform", headerFound: false, columnMap: {}, rows, skipped };
  }

  const grid = rawLines.map(line => ({
    line,
    cells: splitRow(line, delimiter).map(cleanCell),
  }));

  const headerIndex = findHeaderRow(grid);
  const headerFound = headerIndex !== -1;
  const dataStart = headerFound ? headerIndex + 1 : 0;

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
export { parseDocument };
