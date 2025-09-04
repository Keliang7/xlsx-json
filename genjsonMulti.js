// Usage: npm run xlsx2json <folderName>
// Reads projectRoot/_xlsx/<folderName>/ and writes JSON files to
// projectRoot/_json/<folderName>/
// Each Excel file under _xlsx/<folderName>/ contains exactly one language column
// (besides the "key" column). We will generate <fileName>.json where <fileName>
// is the base name (without extension) of the Excel file.
//
// Example:
// _xlsx/i18n/zh-CN.xlsx  -> _json/i18n/zh-CN.json
// _xlsx/i18n/en-US.xlsx  -> _json/i18n/en-US.json

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const norm = (s) => String(s || '').trim().toLowerCase();
const isIndex = (p) => /^\d+$/.test(p);

/**
 * Turn a flat object with dot paths / numeric indices into nested objects/arrays.
 * - Supports keys like "a.b.c", "list.0.name", "a.0.b.1.c"
 * - Creates arrays when the *next* path segment is a numeric index.
 * - Avoids the previous "array-like object" mutation pitfalls by always assigning
 *   into the parent container explicitly when we need to switch container types.
 */
function unflatten(flatObj) {
  const root = {};

  const setDeep = (parts, value) => {
    let cur = root;
    /** @type {{container:any,key:string|number|null}} */
    let parentInfo = { container: null, key: null };

    for (let i = 0; i < parts.length; i++) {
      const p = parts[i];
      const last = i === parts.length - 1;
      const nextIsIndex = !last && isIndex(parts[i + 1]);

      if (last) {
        if (isIndex(p)) {
          const idx = Number(p);
          if (!Array.isArray(cur)) {
            // If we expected an array here but cur isn't, convert the parent slot to an array.
            if (parentInfo.container) {
              const arr = Array.isArray(parentInfo.container[parentInfo.key])
                ? parentInfo.container[parentInfo.key]
                : [];
              parentInfo.container[parentInfo.key] = arr;
              cur = arr;
            } // else at root: we allow numeric keys on object as fallback
          }
          cur[idx] = value;
        } else {
          cur[p] = value;
        }
        return;
      }

      // Not last segment, ensure container exists and has correct type
      if (isIndex(p)) {
        const idx = Number(p);
        // Ensure current is an array when stepping through an index
        if (!Array.isArray(cur)) {
          if (parentInfo.container) {
            const arr = Array.isArray(parentInfo.container[parentInfo.key])
              ? parentInfo.container[parentInfo.key]
              : [];
            parentInfo.container[parentInfo.key] = arr;
            cur = arr;
          } // else at root: allow object with numeric keys (rare in i18n)
        }
        if (cur[idx] == null || typeof cur[idx] !== 'object') {
          cur[idx] = nextIsIndex ? [] : {};
        }
        parentInfo = { container: cur, key: idx };
        cur = cur[idx];
      } else {
        if (cur[p] == null || typeof cur[p] !== 'object') {
          cur[p] = nextIsIndex ? [] : {};
        }
        parentInfo = { container: cur, key: p };
        cur = cur[p];
      }
    }
  };

  for (const [flatKey, value] of Object.entries(flatObj || {})) {
    if (!flatKey) continue;
    const parts = String(flatKey).split('.');
    setDeep(parts, value);
  }

  return root;
}

function readSheetRows(ws) {
  // Keep empty cells as '' so we don't drop keys
  return XLSX.utils.sheet_to_json(ws, { defval: '', blankrows: false });
}

/**
 * Given sheet rows, find the "key" column (case-insensitive) and the ONLY value column.
 * Returns { keyCol, valueCol }
 */
function detectColumns(rows) {
  const headerNames = Object.keys(rows[0] || {});
  const keyCol = headerNames.find((h) => norm(h) === 'key');
  if (!keyCol) {
    throw new Error('Cannot find a "key" column in the sheet header.');
  }
  const valueCols = headerNames.filter((h) => norm(h) !== 'key');
  if (!valueCols.length) {
    throw new Error('No value column found (need one column besides "key").');
  }
  if (valueCols.length > 1) {
    // We allow multiple but warn: we will use the first non-key column.
    console.warn(
      `[xlsx2json] Detected multiple non-key columns (${valueCols.join(
        ', '
      )}); using the first: ${valueCols[0]}`
    );
  }
  return { keyCol, valueCol: valueCols[0] };
}

function processWorkbookToJSON(wb) {
  const sheetName = wb.SheetNames[0];
  if (!sheetName) throw new Error('Workbook has no sheets.');
  const ws = wb.Sheets[sheetName];
  const rows = readSheetRows(ws);
  if (!rows.length) return { keys: 0, data: {} };

  const { keyCol, valueCol } = detectColumns(rows);

  // Build flat map: { key: value }
  const flatMap = {};
  for (const r of rows) {
    const k = String(r[keyCol] ?? '').trim();
    if (!k) continue;
    flatMap[k] = r[valueCol];
  }
  const nested = unflatten(flatMap);
  return { keys: Object.keys(flatMap).length, data: nested };
}

function main() {
  const folderName = process.argv[2];
  if (!folderName) {
    console.error('[xlsx2json] Missing folder name. Example: npm run xlsx2json i18n');
    process.exit(1);
  }

  const projectRoot = process.cwd();
  const inputDir = path.join(projectRoot, '_xlsx', folderName);
  const outputDir = path.join(projectRoot, '_json', folderName);

  if (!fs.existsSync(inputDir) || !fs.statSync(inputDir).isDirectory()) {
    console.error(`[xlsx2json] Input directory not found: ${inputDir}`);
    process.exit(1);
  }

  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

  const entries = fs.readdirSync(inputDir, { withFileTypes: true });
  const excelFiles = entries
    .filter((e) => e.isFile())
    .map((e) => e.name)
    .filter((name) => /\.xlsx$/i.test(name));

  if (!excelFiles.length) {
    console.warn(`[xlsx2json] No .xlsx files found under ${path.relative(projectRoot, inputDir)}`);
    return;
  }

  let written = 0;
  for (const fileName of excelFiles) {
    const filePath = path.join(inputDir, fileName);
    let wb;
    try {
      wb = XLSX.readFile(filePath);
    } catch (e) {
      console.error(`[xlsx2json] Failed to read "${fileName}": ${e.message}`);
      continue;
    }

    let result;
    try {
      result = processWorkbookToJSON(wb);
    } catch (e) {
      console.error(`[xlsx2json] Skipped "${fileName}": ${e.message}`);
      continue;
    }

    const base = path.parse(fileName).name;
    const safeBase = String(base).replace(/[\\/:*?"<>|]/g, '_');
    const outPath = path.join(outputDir, `${safeBase}.json`);
    fs.writeFileSync(outPath, JSON.stringify(result.data, null, 2), 'utf8');
    console.log(
      `[xlsx2json] Wrote ${path.relative(projectRoot, outPath)} (${result.keys} keys)`
    );
    written++;
  }

  console.log(
    `[xlsx2json] Done: ${written} file(s) generated under ${path.relative(projectRoot, outputDir)}`
  );
}

if (require.main === module) {
  main();
}
