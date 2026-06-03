import type { TableauWorkbook, Field } from '../../parsers/model.js';
import { cleanRef } from './zone-helpers.js';

// ─── Column definition ────────────────────────────────────────────────────────

export interface TableColumn {
  key: string;       // field name as it appears in data rows
  label: string;     // clean display name
  numeric: boolean;
  currency: boolean;
  percent: boolean;
}

// ─── Public API ───────────────────────────────────────────────────────────────

export interface TableSpec {
  columns: TableColumn[];
  sampleRows: Record<string, unknown>[];
}

export function resolveTableSpec(workbook: TableauWorkbook, worksheetName: string): TableSpec {
  const encoding = workbook.visualEncodings.find((e) => e.worksheet === worksheetName);
  if (!encoding) return { columns: [], sampleRows: [] };

  const fieldMap = buildFieldMap(workbook);
  const dimInputs = fieldsUsedByDimensionCalcs(workbook);

  // ── Detect date pivot columns (yr:, qr:, month:, etc. on columns shelf) ──
  const dateTruncCols = encoding.columns.filter((c) => isDateTrunc(c.field));
  const hasDatePivot = dateTruncCols.length > 0;

  // ── Detect transposed Measure Names (MN on rows shelf) ──
  const hasMeasureNamesOnRows = encoding.rows.some((r) => r.field.includes('Measure Names'));

  if (hasDatePivot) {
    const periods = generateDatePeriods(dateTruncCols);
    const periodCols = periods.map((p) => makeColumn(p, 'number'));

    if (hasMeasureNamesOnRows) {
      // Transposed: each measure becomes a row, date periods become columns
      const measures = encoding.selectedMeasures.length > 0
        ? encoding.selectedMeasures
        : ['Value'];
      const labelCol: TableColumn = { key: 'Measure', label: 'Measure', numeric: false, currency: false, percent: false };
      const columns = [labelCol, ...periodCols];
      const sampleRows = measures.map((name) => {
        const isPercent = /rate|ratio|pct|percent|growth|yoy|mom|change|%/i.test(name);
        const amounts = isPercent
          ? [0.82, 0.85, 0.79, 0.88, 0.91, 0.84, 0.76, 0.93]
          : [1_240_500, 847_320, 3_102_400, 509_800, 2_765_000, 182_600, 4_320_100, 95_400];
        const row: Record<string, unknown> = { Measure: name };
        periods.forEach((p, i) => { row[p] = amounts[i % amounts.length]; });
        return row;
      });
      return { columns, sampleRows };
    }

    // Standard date pivot: dims on rows, date periods as columns
    const dimColumns: TableColumn[] = encoding.rows
      .filter((r) => r.role !== 'measure')
      .map((r) => {
        const name = r.caption ?? resolveFieldName(r.field, fieldMap);
        return name && !/(Measure Names|Latitude|Longitude)/i.test(name)
          ? makeColumn(name, 'string')
          : null;
      })
      .filter((c): c is TableColumn => c !== null);

    // Determine if values are currency or percent from the primary measure
    const primaryMeasure = encoding.selectedMeasures[0] ?? '';
    const isPercent = /rate|ratio|pct|percent|growth|yoy|mom|change|%/i.test(primaryMeasure);
    const amounts = isPercent
      ? [0.12, -0.05, 0.08, 0.25, 0.15, 0.07, 0.20, -0.03]
      : [1_240_500, 847_320, 3_102_400, 509_800, 2_765_000, 182_600, 4_320_100, 95_400];
    const dimValues = dimColumns.length > 0
      ? generateSampleDimValues(dimColumns[0].label, 6)
      : ['Row 1', 'Row 2', 'Row 3'];

    const columns = [...dimColumns, ...periodCols];
    const sampleRows = dimValues.map((dimVal, i) => {
      const row: Record<string, unknown> = dimColumns.length > 0 ? { [dimColumns[0].key]: dimVal } : {};
      periods.forEach((p, j) => { row[p] = amounts[(i + j) % amounts.length]; });
      return row;
    });
    return { columns, sampleRows };
  }

  // ── Standard table (no date pivot) ──

  // Dimensions from rows shelf (GROUP BY columns)
  const dimColumns: TableColumn[] = encoding.rows
    .filter((r) => r.role !== 'measure')
    .map((r) => {
      const name = r.caption ?? resolveFieldName(r.field, fieldMap);
      return name && !/(Measure Names|Latitude|Longitude)/i.test(name)
        ? makeColumn(name, 'string')
        : null;
    })
    .filter((c): c is TableColumn => c !== null);

  const hasMeasureNames = encoding.columns.some((c) => c.field.includes(':Measure Names'));

  // Pivot columns — dimensions on the columns shelf (not Measure Names)
  const pivotDimRefs = encoding.columns.filter(
    (c) => c.role === 'dimension' && !c.field.includes(':Measure Names')
  );
  const hasPivot = pivotDimRefs.length > 0;

  // Measure columns
  let measureColumns: TableColumn[];

  if (hasPivot) {
    const pivotValueColumns: TableColumn[] = [];
    for (const pivotRef of pivotDimRefs) {
      const dimName = pivotRef.caption ?? resolveFieldName(pivotRef.field, fieldMap);
      for (const val of extractPivotValues(dimName, workbook)) {
        pivotValueColumns.push(makeColumn(val, 'number'));
      }
    }
    measureColumns = pivotValueColumns;
  } else if (hasMeasureNames) {
    if (encoding.selectedMeasures.length > 0) {
      measureColumns = encoding.selectedMeasures.map((name) => {
        const field = workbook.fields.find(
          (f) => f.caption === name || cleanRef(f.name) === name
        );
        const type = field?.dataType === 'integer' || field?.dataType === 'real' ? 'number' : 'string';
        return makeColumn(name, type);
      });
    } else {
      // Fallback: all non-dimension, non-dim-input measures
      measureColumns = workbook.fields
        .filter((f) => {
          if (f.role !== 'measure') return false;
          if (f.name === '[Number of Records]') return false;
          const displayName = f.caption ?? cleanRef(f.name);
          return !dimInputs.has(displayName) && !dimInputs.has(f.name.replace(/^\[|\]$/g, ''));
        })
        .map((f) => {
          const name = f.caption ?? cleanRef(f.name);
          return makeColumn(name, f.dataType === 'integer' || f.dataType === 'real' ? 'number' : 'string');
        });
    }
  } else {
    // Explicit column measure refs
    measureColumns = encoding.columns
      .filter((c) => c.role === 'measure' || (!c.role && !c.field.includes(':Measure Names')))
      .map((c) => {
        const name = c.caption ?? resolveFieldName(c.field, fieldMap);
        return name ? makeColumn(name, 'number') : null;
      })
      .filter((c): c is TableColumn => c !== null);
  }

  const columns = [...dimColumns, ...measureColumns];
  if (columns.length === 0) return { columns: [], sampleRows: [] };

  const sampleRows = generateSampleRows(columns, dimColumns);
  return { columns, sampleRows };
}

// ─── Sample data generation ───────────────────────────────────────────────────

function generateSampleRows(
  columns: TableColumn[],
  dimColumns: TableColumn[],
  rowCount = 8,
): Record<string, unknown>[] {
  // First pass: generate all non-percentage values
  const rows = Array.from({ length: rowCount }, (_, i) =>
    Object.fromEntries(
      columns
        .filter((col) => !col.key.startsWith('% of '))
        .map((col) => [col.key, sampleValue(col, i, dimColumns)])
    )
  );

  // Second pass: compute % of Total for derived percentage columns
  for (const col of columns) {
    if (!col.key.startsWith('% of ')) continue;
    const baseKey = col.key.slice('% of '.length);
    const total = rows.reduce((sum, r) => sum + (Number(r[baseKey]) || 0), 0);
    for (const row of rows) {
      row[col.key] = total > 0 ? (Number(row[baseKey]) || 0) / total : 0;
    }
  }

  return rows;
}

function sampleValue(col: TableColumn, i: number, _dimCols: TableColumn[]): unknown {
  if (!col.numeric) {
    return sampleDimValue(col.label, i);
  }
  if (col.currency) {
    const amounts = [1_240_500, 847_320, 3_102_400, 509_800, 2_765_000, 182_600, 4_320_100, 95_400];
    return amounts[i % amounts.length];
  }
  if (col.percent) {
    const rates = [0.12, -0.05, 0.08, 0.25, 0.15, 0.07, 0.20, -0.03];
    return rates[i % rates.length];
  }
  const nums = [12450, 98760, 34120, 76540, 55230, 89100, 23670, 61980];
  return nums[i % nums.length];
}

function sampleDimValue(label: string, i: number): string {
  const lower = label.toLowerCase();
  if (/country|region|geography|geo/.test(lower)) {
    return ['United States', 'Germany', 'France', 'United Kingdom', 'Canada', 'Australia', 'Japan', 'Brazil'][i % 8];
  }
  if (/company|customer|client|account/.test(lower)) {
    return ['Acme Corp', 'Globex Industries', 'Initech LLC', 'Umbrella Co', 'Soylent Corp', 'Massive Dynamic', 'Hooli Inc', 'Pied Piper'][i % 8];
  }
  if (/invoice|number|id|ref/.test(lower)) {
    return `INV-${String(10001 + i * 137).padStart(5, '0')}`;
  }
  if (/bucket|aging|category|segment|status|type/.test(lower)) {
    return ['Current', '1-30 Days Past Due', '31-60 Days Past Due', '61-90 Days Past Due', '91-120 Days Past Due', '>120 Days Past Due', 'Pending', 'Closed'][i % 8];
  }
  if (/product|sku|item/.test(lower)) {
    return ['Product A', 'Product B', 'Product C', 'Product D', 'Product E', 'Product F', 'Product G', 'Product H'][i % 8];
  }
  if (/date|month|year|quarter/.test(lower)) {
    return ['Jan 2024', 'Feb 2024', 'Mar 2024', 'Apr 2024', 'May 2024', 'Jun 2024', 'Jul 2024', 'Aug 2024'][i % 8];
  }
  return `${label} ${i + 1}`;
}

// ─── DataTable.tsx component template ────────────────────────────────────────

export function dataTableComponentSource(): string {
  return `import { useState, useEffect } from 'react';
import type React from 'react';

export interface DataTableColumn {
  key: string;
  label: string;
  numeric: boolean;
  currency: boolean;
  percent: boolean;
}

export interface DimSpec {
  alias: string;
  expr: string;
}

export interface PivotSpec {
  expr: string;
  values: string[];
  measureExpr: string;
  totalAlias?: string;
}

export interface QuerySpec {
  source: string;
  dimensions?: DimSpec[];
  measures?: Array<{ alias: string; expr: string }>;
  filterFields?: string[];
  orderBy?: { column: string; direction?: string };
  pivotOn?: PivotSpec;
}

export interface DataTableProps {
  title?: string;
  /** Query spec — posted to /api/query for live data */
  querySpec?: QuerySpec;
  /** Active filter values matched to querySpec.filterFields */
  filters?: Record<string, string | null | undefined>;
  /** Static rows — sample data shown when querySpec is absent or API unavailable */
  rows?: Record<string, unknown>[];
  columns: DataTableColumn[];
  isSampleData?: boolean;
  /** Rows per page for server-side pagination (default: 50) */
  pageSize?: number;
}

const tableStyle: React.CSSProperties = {
  width: '100%',
  borderCollapse: 'collapse',
  fontSize: 10,
  fontFamily: "'Poppins', sans-serif",
};

const thStyle = (numeric: boolean): React.CSSProperties => ({
  padding: '8px 12px',
  borderBottom: '2px solid #e2e8f0',
  textAlign: numeric ? 'right' : 'left',
  fontWeight: 600,
  color: '#374151',
  whiteSpace: 'nowrap',
  background: '#f8fafc',
  position: 'sticky',
  top: 0,
  zIndex: 1,
});

const tdStyle = (numeric: boolean): React.CSSProperties => ({
  padding: '6px 12px',
  borderBottom: '1px solid #e2e8f0',
  textAlign: numeric ? 'right' : 'left',
  color: '#1f2937',
});

function formatValue(val: unknown, col: DataTableColumn): string {
  if (val === null || val === undefined) return '';
  if (col.currency && typeof val === 'number') {
    return val.toLocaleString('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 0, maximumFractionDigits: 0 });
  }
  if (col.percent && typeof val === 'number') {
    return (val * 100).toFixed(1) + '%';
  }
  if (col.numeric && typeof val === 'number') {
    return val.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  }
  return String(val);
}

export function DataTable({ title, querySpec, filters, columns, rows: staticRows = [], isSampleData = false, pageSize = 50 }: DataTableProps) {
  const [rows, setRows] = useState<Record<string, unknown>[]>(staticRows);
  const [loading, setLoading] = useState(!!querySpec);
  const [error, setError] = useState<string | null>(null);
  const [page, setPage] = useState(1);
  const [hasMore, setHasMore] = useState(false);
  // For pivot tables with no pre-defined values, columns come from the API response
  const [dynamicColumns, setDynamicColumns] = useState<DataTableColumn[] | null>(null);
  const displayColumns = dynamicColumns ?? columns;

  useEffect(() => {
    setPage(1);
  }, [querySpec?.source, JSON.stringify(filters)]);

  useEffect(() => {
    if (!querySpec) { setRows(staticRows); return; }
    setLoading(true);
    setError(null);
    const body = {
      ...querySpec,
      filters: Object.fromEntries(
        (querySpec.filterFields ?? []).map((f) => [f, filters?.[f] ?? null])
      ),
      page,
      pageSize,
    };
    fetch('/api/query', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    })
      .then((r) => { if (!r.ok) throw new Error(\`API \${r.status}: \${r.statusText}\`); return r.json(); })
      .then((data: Record<string, unknown>[]) => {
        // For pivot tables with dynamic values, derive column list from first response row
        if (querySpec.pivotOn && querySpec.pivotOn.values.length === 0 && data.length > 0) {
          const dimKeys = (querySpec.dimensions ?? []).map((d) => (typeof d === 'string' ? d : d.alias));
          const allKeys = Object.keys(data[0]);
          const pivotKeys = allKeys.filter((k) => !dimKeys.includes(k));
          const totalAlias = querySpec.pivotOn.totalAlias ?? 'Total';
          const bucketKeys = pivotKeys.filter((k) => k !== totalAlias);
          const newCols: DataTableColumn[] = [
            ...dimKeys.map((k) => ({ key: k, label: k, numeric: false, currency: false, percent: false })),
            ...bucketKeys.map((k) => ({ key: k, label: k, numeric: true, currency: false, percent: false })),
            { key: totalAlias, label: totalAlias, numeric: true, currency: false, percent: false },
          ];
          setDynamicColumns(newCols);
        }
        setRows(data);
        setHasMore(data.length === pageSize);
        setLoading(false);
      })
      .catch((e: Error) => {
        // Fall back to sample data silently when API fails but static rows are available
        if (staticRows.length > 0) {
          setError(null);
          setRows(staticRows);
        } else {
          setError(e.message);
        }
        setLoading(false);
      });
  }, [querySpec?.source, JSON.stringify(filters), page, pageSize, staticRows]);

  return (
    <div style={{ fontFamily: "'Poppins', sans-serif", display: 'flex', flexDirection: 'column', height: '100%', minHeight: 0 }}>
      {title && (
        <h3 style={{ fontSize: 10, fontWeight: 600, color: '#374151', textTransform: 'uppercase' as const, letterSpacing: '0.06em', marginBottom: 4, flexShrink: 0 }}>
          {title}
        </h3>
      )}
      {isSampleData && !querySpec && (
        <div style={{ background: '#fef3c7', border: '1px solid #f59e0b', borderRadius: 4, padding: '3px 10px', marginBottom: 4, fontSize: 10, color: '#92400e', flexShrink: 0 }}>
          ⚠ Sample data — connect to your data source to see real values
        </div>
      )}
      {loading && (
        <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#94a3b8', fontSize: 10 }}>
          Loading…
        </div>
      )}
      {error && (
        <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#ef4444', fontSize: 10 }}>
          {error} — is the API server running? (npm run api)
        </div>
      )}
      {!loading && !error && (
        <>
          <div style={{ overflowX: 'auto', overflowY: 'auto', flex: 1, minHeight: 0 }}>
            <table style={tableStyle}>
              <thead>
                <tr>
                  {displayColumns.map((col) => (
                    <th key={col.key} style={thStyle(col.numeric)}>{col.label}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((row, i) => (
                  <tr key={i} style={i % 2 === 1 ? { background: '#f9fafb' } : undefined}>
                    {displayColumns.map((col) => (
                      <td key={col.key} style={tdStyle(col.numeric)}>
                        {formatValue(row[col.key], col)}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {querySpec && (page > 1 || hasMore) && (
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '6px 4px 0', flexShrink: 0 }}>
              <button
                onClick={() => setPage((p) => Math.max(1, p - 1))}
                disabled={page === 1}
                style={{ fontSize: 10, padding: '3px 10px', border: '1px solid #e2e8f0', borderRadius: 4, background: page === 1 ? '#f8fafc' : 'white', color: page === 1 ? '#94a3b8' : '#374151', cursor: page === 1 ? 'default' : 'pointer' }}
              >
                ← Prev
              </button>
              <span style={{ fontSize: 10, color: '#64748b' }}>Page {page}</span>
              <button
                onClick={() => setPage((p) => p + 1)}
                disabled={!hasMore}
                style={{ fontSize: 10, padding: '3px 10px', border: '1px solid #e2e8f0', borderRadius: 4, background: !hasMore ? '#f8fafc' : 'white', color: !hasMore ? '#94a3b8' : '#374151', cursor: !hasMore ? 'default' : 'pointer' }}
              >
                Next →
              </button>
            </div>
          )}
        </>
      )}
    </div>
  );
}
`;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function makeColumn(name: string, type: 'string' | 'number'): TableColumn {
  const isPercentOfTotal = name.startsWith('% of ');
  // currency and percent are mutually exclusive — percent wins when both patterns match
  const matchesPercent = isPercentOfTotal || (type === 'number' && /rate|ratio|pct|percent|growth|yoy|mom|change|%/i.test(name));
  const currency = !matchesPercent && type === 'number' && /amount|revenue|sales|cost|price|gross|net|paid|outstanding|balance/i.test(name);
  const percent = matchesPercent;
  return {
    key: name,
    label: cleanColumnLabel(name),
    numeric: type === 'number',
    currency,
    percent,
  };
}

function cleanColumnLabel(name: string): string {
  return name
    .replace(/^_+\s*/, '')             // strip leading underscores
    .replace(/^\.\s*/, '')              // strip leading dot
    .replace(/\s*\([^)]*Currency[^)]*\)\s*/gi, '')  // strip any "(...Currency...)" Tableau suffix
    .replace(/\s*\(copy\)\s*/i, '')
    .trim();
}

function resolveFieldName(field: string, fieldMap: Map<string, string>): string {
  const parts = field.split(':');
  const internal = parts.length >= 3 ? parts.slice(1, -1).join(':') : field;
  return fieldMap.get(internal) ?? fieldMap.get(field) ?? cleanRef(internal);
}

// ─── Date pivot helpers ───────────────────────────────────────────────────────

// Tableau date-truncation prefixes: yr, qr, month, week, day, hr, mn, sec,
// and the trunc-* family (trunc-year, trunc-month, trunc-quarter, etc.).
function isDateTrunc(field: string): boolean {
  return /^(yr|qr|month|week|day|hr|mn|sec|trunc[^:]*|datepart|datetrunc):/i.test(field);
}

function generateDatePeriods(dateCols: { field: string }[]): string[] {
  const fields = dateCols.map((c) => c.field.toLowerCase());
  const hasQuarter = fields.some((f) => f.startsWith('qr:') || f.startsWith('trunc-quarter:'));
  const hasMonth   = fields.some((f) => f.startsWith('month:') || f.startsWith('trunc-month:'));

  const now = new Date();
  const currentYear = now.getFullYear();
  const currentQ = Math.floor(now.getMonth() / 3) + 1;

  if (hasMonth) {
    // Last 6 complete months relative to today
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return Array.from({ length: 6 }, (_, i) => {
      const d = new Date(now.getFullYear(), now.getMonth() - 5 + i, 1);
      return `${months[d.getMonth()]} ${d.getFullYear()}`;
    });
  }
  if (hasQuarter) {
    // Last 8 quarters relative to today
    return Array.from({ length: 8 }, (_, i) => {
      let q = currentQ - 7 + i;
      let y = currentYear;
      while (q <= 0) { q += 4; y--; }
      while (q > 4)  { q -= 4; y++; }
      return `Q${q} ${y}`;
    });
  }
  // Year-level: last 4 fiscal years
  return Array.from({ length: 4 }, (_, i) => `FY${currentYear - 3 + i}`);
}

// Single authoritative dim sampler — used by both standard and date-pivot paths.
function generateSampleDimValues(label: string, count: number): string[] {
  // Delegate to sampleDimValue for consistent values across all table types
  return Array.from({ length: count }, (_, i) => String(sampleDimValue(label, i)));
}

// Extract distinct pivot values from a dimension's IF/THEN/ELSE formula.
// Returns the THEN string literals in order, excluding "Null" catch-alls.
function extractPivotValues(dimName: string, workbook: TableauWorkbook): string[] {
  const calc = workbook.calculations.find((c) => c.name === dimName);
  if (!calc) return [dimName]; // not a calculated dim — use the name itself

  // Decode HTML entities before regex matching
  const formula = calc.formula
    .replace(/&#13;/g, ' ').replace(/&#10;/g, ' ')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');

  const matches = [...formula.matchAll(/THEN\s+"([^"]+)"/gi)];
  return matches
    .map((m) => m[1].trim())
    .filter((v) => v.toLowerCase() !== 'null' && v.length > 0);
}

function buildFieldMap(workbook: TableauWorkbook): Map<string, string> {
  const map = new Map<string, string>();
  for (const f of [...workbook.fields, ...workbook.dataSources.flatMap((ds) => ds.fields)]) {
    const display = f.caption ?? cleanRef(f.name);
    map.set(f.name.replace(/^\[|\]$/g, ''), display);
    map.set(f.name, display);
  }
  return map;
}

function fieldsUsedByDimensionCalcs(workbook: TableauWorkbook): Set<string> {
  const refs = new Set<string>();
  for (const calc of workbook.calculations) {
    if (!/\bIF\b|\bELSE\b/i.test(calc.formula)) continue;
    for (const m of calc.formula.match(/\[([^\]]+)\]/g) ?? []) refs.add(m.slice(1, -1));
  }
  return refs;
}
