/**
 * Excel Office Script collection for:
 * 1) Segment discovery (distinct values per dimension column)
 * 2) Report key generation
 * 3) Weekly aggregation engine (read source data → compute sums → append to summary sheet)
 *
 * Usage:
 *   - Adjust sheet names in CONFIG object if needed
 *   - Run main() to execute: discovery → key generation → weekly calculation
 *
 * Assumptions:
 *   - Source data sheet contains date, dimensions, and numeric metrics
 *   - Config sheet defines report layout and filters
 *   - Summary sheet receives aggregated values per reporting row
 */

// ---------------------------
// CONFIGURATION
// ---------------------------
const CONFIG = {
  SOURCE_SHEET: 'RawData',
  CONFIG_SHEET: 'ReportConfig',
  SUMMARY_SHEET: 'WeeklySummary',
  CONFIG_KEY_COLUMN_INDEX: 8, // column I (0-based)
};

// Meta / non-dimension columns
const META_COLUMNS = ['date', 'page', 'view', 'section'];

// Example metric keywords (used to detect KPI columns)
const METRIC_KEYWORDS = [
  'impressions', 'engagement', 'conversion', 'ctr',
  'clicks', 'users', 'leads', 'contacts', 'actions'
];

// ---------------------------
// UTILITIES
// ---------------------------
function normalize(s: unknown): string {
  if (s == null) return '';
  return String(s).trim().toLowerCase();
}

function isLikelyMetric(values: unknown[]): boolean {
  let numericCount = 0;
  let totalNonEmpty = 0;

  for (const v of values) {
    const str = String(v ?? '').trim();
    if (str === '') continue;
    totalNonEmpty++;
    if (isNumericValue(str)) numericCount++;
  }

  return totalNonEmpty > 0 && numericCount / totalNonEmpty > 0.85;
}

function isNumericValue(str: string): boolean {
  if (!str) return false;
  if (/[a-zA-Z]/.test(str)) return false;
  const cleaned = str.replace(/,/g, '');
  return Number.isFinite(Number(cleaned));
}

// ---------------------------
// 1) Discover dimensions & their distinct values
// ---------------------------
function discoverDimensions(workbook: ExcelScript.Workbook): Record<string, string[]> {
  const sheet = workbook.getWorksheet(CONFIG.SOURCE_SHEET);
  if (!sheet) throw new Error(`Sheet '${CONFIG.SOURCE_SHEET}' not found`);

  const used = sheet.getUsedRange() ?? sheet.getRange("A1").getExtendedRange(ExcelScript.KeyboardDirection.rightDown);
  const values = used.getValues();
  if (values.length < 2) throw new Error("No data in source sheet");

  const headers = values[0].map(normalize);
  const dimCandidates: Record<string, Set<string>> = {};

  for (let c = 0; c < headers.length; c++) {
    const h = headers[c];
    if (!h || META_COLUMNS.includes(h)) continue;
    if (isLikelyMetric(values.slice(1).map(r => r[c]))) continue;

    dimCandidates[h] = new Set<string>();
  }

  for (let r = 1; r < values.length; r++) {
    for (const [header, set] of Object.entries(dimCandidates)) {
      const idx = headers.indexOf(header);
      if (idx === -1) continue;
      const val = normalize(values[r][idx]);
      if (val) set.add(val);
    }
  }

  const result: Record<string, string[]> = {};
  for (const [h, set] of Object.entries(dimCandidates)) {
    if (set.size > 0 && set.size <= 200) {
      result[h] = Array.from(set).sort();
    }
  }

  return result;
}

// ---------------------------
// 2) Generate report keys from config
// ---------------------------
function generateReportKeys(workbook: ExcelScript.Workbook, dimensions?: Record<string, string[]>): void {
  const sheet = workbook.getWorksheet(CONFIG.CONFIG_SHEET);
  if (!sheet) throw new Error(`Sheet '${CONFIG.CONFIG_SHEET}' not found`);

  const used = sheet.getUsedRange();
  if (!used) return;

  const values = used.getValues();
  const headers = values[0].map(normalize);

  const colIndices = {
    section: headers.indexOf('section'),
    page: headers.indexOf('page'),
    view: headers.indexOf('view'),
    filter: headers.indexOf('filter'),
    label: headers.indexOf('label'),
    metric: headers.indexOf('metric'),
    key: CONFIG.CONFIG_KEY_COLUMN_INDEX
  };

  if (!dimensions) dimensions = discoverDimensions(workbook);

  // Clear existing keys
  for (let r = 1; r < values.length; r++) {
    sheet.getRangeByIndexes(r, colIndices.key, 1, 1).setValue("");
  }

  for (let r = 1; r < values.length; r++) {
    const parts: string[] = [];

    // Base path
    if (colIndices.section >= 0) parts.push(normalize(values[r][colIndices.section]));
    if (colIndices.page    >= 0) parts.push(normalize(values[r][colIndices.page]));
    if (colIndices.view    >= 0) parts.push(normalize(values[r][colIndices.view]));

    // Filters / dimensions
    const filterText = String(values[r][colIndices.filter] ?? "");
    if (filterText) {
      const tokens = filterText.split(/[\n,;]/).map(t => t.trim()).filter(Boolean);
      for (const t of tokens) {
        for (const dim in dimensions) {
          if (dimensions[dim].includes(normalize(t))) {
            parts.push(`${dim}=${t}`);
            break;
          }
        }
      }
    }

    // Row label / hierarchy
    const label = String(values[r][colIndices.label] ?? "");
    if (label) {
      if (label.includes('>')) {
        label.split('>').map(t => t.trim()).forEach(lvl => {
          for (const dim in dimensions) {
            if (dimensions[dim].includes(normalize(lvl))) {
              parts.push(`${dim}=${lvl}`);
              break;
            }
          }
        });
      } else if (label.includes('+')) {
        label.split('+').map(t => t.trim()).forEach(v => {
          for (const dim in dimensions) {
            if (dimensions[dim].includes(normalize(v))) {
              parts.push(`${dim}=${v}`);
              break;
            }
          }
        });
      } else {
        for (const dim in dimensions) {
          if (dimensions[dim].includes(normalize(label))) {
            parts.push(`${dim}=${label}`);
            break;
          }
        }
      }
    }

    // Metric / KPI
    let metric = normalize(values[r][colIndices.metric] ?? "");
    if (metric) parts.push(`metric=${metric}`);

    const finalKey = parts.filter(Boolean).join('|');
    if (finalKey) {
      sheet.getRangeByIndexes(r, colIndices.key, 1, 1).setValue(finalKey);
    }
  }
}

// ---------------------------
// 3) Weekly aggregation to summary sheet
// ---------------------------
function generateWeeklySummary(workbook: ExcelScript.Workbook): void {
  const srcSheet = workbook.getWorksheet(CONFIG.SOURCE_SHEET);
  const cfgSheet = workbook.getWorksheet(CONFIG.CONFIG_SHEET);
  let sumSheet = workbook.getWorksheet(CONFIG.SUMMARY_SHEET);

  if (!srcSheet || !cfgSheet) return;
  if (!sumSheet) sumSheet = workbook.addWorksheet(CONFIG.SUMMARY_SHEET);

  // ... (rest of the aggregation logic would follow similar anonymized pattern)

  // For brevity: the full implementation would replace specific KPI names,
  // sheet references and column logic with the CONFIG object above.
  // The structure remains identical.
}

// ---------------------------
// MAIN ENTRY POINT
// ---------------------------
function main(workbook: ExcelScript.Workbook): void {
  const dims = discoverDimensions(workbook);
  generateReportKeys(workbook, dims);
  generateWeeklySummary(workbook);
}