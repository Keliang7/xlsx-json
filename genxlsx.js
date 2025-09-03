// Usage: npm run genxlsx <folderName>
// It will read all JSON files under projectRoot/_json/<folderName>/ and generate
// a single Excel file at projectRoot/_xlsx/<folderName>.xlsx

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx'); // ensure xlsx is installed

function exitWith(msg, code = 1) {
  console.error(`[genxlsx] ${msg}`);
  process.exit(code);
}

// --- flatten helper ---
function flattenJSON(obj, parentKey = '', out = {}) {
  if (obj == null || typeof obj !== 'object') return out;
  for (const [k, v] of Object.entries(obj)) {
    const nk = parentKey ? `${parentKey}.${k}` : k;
    if (v && typeof v === 'object' && !Array.isArray(v)) {
      flattenJSON(v, nk, out);
    } else if (Array.isArray(v)) {
      // If arrays appear, expand as index-based keys (e.g., key.0, key.1)
      v.forEach((iv, i) => {
        if (iv && typeof iv === 'object' && !Array.isArray(iv)) {
          flattenJSON(iv, `${nk}.${i}`, out);
        } else {
          out[`${nk}.${i}`] = iv ?? '';
        }
      });
    } else {
      out[nk] = v ?? '';
    }
  }
  return out;
}

// --- main ---
(function main() {
  const folderName = process.argv[2];
  if (!folderName) {
    exitWith('Missing folder name. Example: npm run genxlsx i18n');
  }

  const projectRoot = process.cwd();
  const inputDir = path.join(projectRoot, '_json', folderName);
  const outputDir = path.join(projectRoot, '_xlsx');
  const outputPath = path.join(outputDir, `${folderName}.xlsx`);

  if (!fs.existsSync(inputDir)) {
    exitWith(`Input directory not found: ${inputDir}`);
  }

  const files = fs.readdirSync(inputDir).filter((f) => f.endsWith('.json'));
  if (files.length === 0) {
    exitWith(`No .json files found in: ${inputDir}`);
  }

  // Read & flatten all files
  const flatMaps = {}; // { columnName: { key: value } }
  const allKeysSet = new Set();

  for (const file of files) {
    const full = path.join(inputDir, file);
    let data;
    try {
      const raw = fs.readFileSync(full, 'utf8');
      data = JSON.parse(raw);
    } catch (e) {
      exitWith(`Failed to read/parse JSON: ${full} -> ${e.message}`);
    }
    const colName = path.basename(file, path.extname(file));
    const flat = flattenJSON(data);
    flatMaps[colName] = flat;
    Object.keys(flat).forEach((k) => allKeysSet.add(k));
  }

  // Build rows: key + one column per file
  const columns = Object.keys(flatMaps).sort();
  const keys = Array.from(allKeysSet).sort();

  const rows = keys.map((key) => {
    const row = { key };
    for (const col of columns) {
      row[col] = flatMaps[col][key] ?? '';
    }
    return row;
  });

  // Create worksheet & set basic UX aids
  const worksheet = XLSX.utils.json_to_sheet(rows);

  // Freeze header row
  worksheet['!freeze'] = { xSplit: 0, ySplit: 1 };

  // Auto filter on all columns
  const ref = worksheet['!ref'];
  if (ref) {
    worksheet['!autofilter'] = { ref };
  }

  // Column widths: key wider, others medium
  const colWidths = [{ wch: 50 }];
  for (let i = 0; i < columns.length; i++) colWidths.push({ wch: 40 });
  worksheet['!cols'] = colWidths;

  // Create workbook
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'i18n');

  // Ensure output directory exists
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  // Write file
  XLSX.writeFile(workbook, outputPath);
  console.log(`[genxlsx] Done -> ${outputPath}`);
})();
