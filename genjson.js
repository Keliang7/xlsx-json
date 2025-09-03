// Usage: npm run xlsx2json <folderName>
// Reads projectRoot/_xlsx/<folderName>.xlsx and writes JSON files to
// projectRoot/_json/<folderName>/: one file per non-`key` column (e.g. zh-CN.json)

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const norm = (s) => String(s || '').trim().toLowerCase();

// Expand dot-path keys like "a.b.c" or with array indices like "a.b.0.c" into nested objects/arrays
function unflatten(flatObj) {
  const root = {};
  const isIndex = (p) => /^\d+$/.test(p);

  for (const [flatKey, value] of Object.entries(flatObj || {})) {
    if (!flatKey) continue;
    const parts = String(flatKey).split('.');

    let cur = root;
    for (let i = 0; i < parts.length; i++) {
      const p = parts[i];
      const last = i === parts.length - 1;

      if (last) {
        if (isIndex(p)) {
          if (!Array.isArray(cur)) {
            // Convert object to array if needed
            const arr = [];
            Object.assign(arr, cur);
            cur = arr;
          }
          cur[Number(p)] = value;
        } else {
          cur[p] = value;
        }
      } else {
        const next = parts[i + 1];
        const shouldBeArray = isIndex(next);

        if (isIndex(p)) {
          // Current segment is array index under current container
          if (!Array.isArray(cur)) {
            const arr = [];
            Object.assign(arr, cur);
            // We cannot reassign parent reference here directly, so handle via temporary holder
          }
          // When an index appears but cur isn't an array, we coerce by direct set on numeric prop
        }

        // Ensure the child container exists
        if (isIndex(p)) {
          // index under an array-like container
          const idx = Number(p);
          if (!Array.isArray(cur)) {
            // create array if needed
            const arr = [];
            // copy existing numeric keys if any
            for (const k of Object.keys(cur)) {
              if (/^\d+$/.test(k)) arr[Number(k)] = cur[k];
            }
            // attach array-like back to object (works since JS arrays are objects)
            Object.setPrototypeOf(arr, Array.prototype);
            // replace properties of cur with arr (mutate cur)
            for (const k of Object.keys(cur)) delete cur[k];
            cur.push = Array.prototype.push; // ensure it's array-like
          }
          if (cur[idx] == null) cur[idx] = shouldBeArray ? [] : {};
          if (typeof cur[idx] !== 'object') cur[idx] = shouldBeArray ? [] : {};
          cur = cur[idx];
        } else {
          if (cur[p] == null) cur[p] = shouldBeArray ? [] : {};
          if (typeof cur[p] !== 'object') cur[p] = shouldBeArray ? [] : {};
          cur = cur[p];
        }
      }
    }
  }
  return root;
}

function readSheetRows(ws) {
  // Keep empty cells as '' so we don't drop keys
  return XLSX.utils.sheet_to_json(ws, { defval: '', blankrows: false });
}

function main() {
  const folderName = process.argv[2];
  if (!folderName) {
    console.error('[xlsx2json] Missing folder name. Example: npm run xlsx2json i18n');
    process.exit(1);
  }

  const projectRoot = process.cwd();
  const inputPath = path.join(projectRoot, '_xlsx', `${folderName}.xlsx`);
  const outputDir = path.join(projectRoot, '_json', folderName);

  if (!fs.existsSync(inputPath)) {
    console.error(`[xlsx2json] Input file not found: ${inputPath}`);
    process.exit(1);
  }

  let wb;
  try {
    wb = XLSX.readFile(inputPath);
  } catch (e) {
    console.error(`[xlsx2json] Failed to read workbook: ${e.message}`);
    process.exit(1);
  }

  const sheetName = wb.SheetNames[0];
  if (!sheetName) {
    console.error('[xlsx2json] Workbook has no sheets.');
    process.exit(1);
  }

  const ws = wb.Sheets[sheetName];
  const rows = readSheetRows(ws);
  if (!rows.length) {
    console.warn('[xlsx2json] Sheet is empty, nothing to convert.');
    return;
  }

  const headerNames = Object.keys(rows[0]);
  const keyCol = headerNames.find((h) => norm(h) === 'key');
  if (!keyCol) {
    console.error('[xlsx2json] Cannot find a "key" column in the first sheet header.');
    process.exit(1);
  }

  const valueCols = headerNames.filter((h) => norm(h) !== 'key');
  if (!valueCols.length) {
    console.error('[xlsx2json] No value columns found (need at least one column besides "key").');
    process.exit(1);
  }

  // Build flat maps per column: { colName: { key: value } }
  const flatMaps = {};
  for (const col of valueCols) flatMaps[col] = {};

  for (const r of rows) {
    const k = String(r[keyCol] ?? '').trim();
    if (!k) continue;
    for (const col of valueCols) {
      flatMaps[col][k] = r[col];
    }
  }

  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  let fileCount = 0;
  for (const col of valueCols) {
    const nested = unflatten(flatMaps[col]);
    // Sanitize file name (avoid slashes or special chars)
    const safeCol = String(col).replace(/[\\/:*?"<>|]/g, '_');
    const outPath = path.join(outputDir, `${safeCol}.json`);
    fs.writeFileSync(outPath, JSON.stringify(nested, null, 2), 'utf8');
    console.log(`[xlsx2json] Wrote ${path.relative(projectRoot, outPath)} (${Object.keys(flatMaps[col]).length} keys)`);
    fileCount++;
  }

  console.log(`[xlsx2json] Done: ${fileCount} file(s) generated under ${path.relative(projectRoot, outputDir)}`);
}

if (require.main === module) {
  main();
}
