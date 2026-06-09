import type { TableauWorkbook } from '../../parsers/model.js';
import type { ParquetTable } from '../../extractors/hyper-extractor.js';

export interface DimSpec {
  alias: string;
  /** SQL expression — quoted column name for raw fields, CASE WHEN for calcs */
  expr: string;
}

export interface QuerySpec {
  /** Parquet file slug (filename without .parquet) */
  source: string;
  /** GROUP BY dimensions — raw columns use `"name"`, calcs use their SQL expression */
  dimensions: DimSpec[];
  /** Computed measure columns (non-pivot tables) */
  measures: Array<{ alias: string; expr: string }>;
  /** Filter fields used by interactive wiring */
  filterFields: string[];
  /** Pivot config — when present, generates conditional aggregation per bucket */
  pivotOn?: {
    /** SQL expression for the pivot dimension (e.g. CASE WHEN formula) */
    expr: string;
    /** Distinct values to pivot on — each becomes a column */
    values: string[];
    /** The measure expression to aggregate per bucket (plain, no SUM wrapper) */
    measureExpr: string;
    /** Alias prefix for the total column */
    totalAlias?: string;
  };
}

export interface ApiServerResult {
  serverTs: string;
  /** Query spec per worksheet — embedded in Dashboard.tsx */
  specs: Record<string, QuerySpec>;
}

// ─── Public entry point ───────────────────────────────────────────────────────

export function generateApiServer(
  workbook: TableauWorkbook,
  parquetTables: ParquetTable[],
  workbookSlug: string,
): ApiServerResult {
  const specs: Record<string, QuerySpec> = {};
  const defaultSlug = parquetTables[0]?.slug ?? workbookSlug;
  const actionFilters = extractActionFilters(workbook);

  for (const enc of workbook.visualEncodings) {
    if (enc.effectiveMarkType !== 'automatic') continue;

    const dims: DimSpec[] = enc.rows
      .filter((r) => r.role !== 'measure')
      .map((r) => {
        const alias = r.caption ?? cleanFieldRef(r.field);
        if (!alias || /(Measure Names|Latitude|Longitude)/i.test(alias)) return null;
        const expr = resolveDimExpr(alias, workbook);
        if (expr === null) return null; // Tableau Group/Set — no SQL equivalent
        return { alias, expr };
      })
      .filter((d): d is DimSpec => d !== null);

    const filterFields = actionFilters
      .filter((f) => f.appliedTo.includes(enc.worksheet))
      .map((f) => f.field);

    // Detect pivot columns — dimensions on the columns shelf (not Measure Names)
    const pivotDimRefs = enc.columns.filter(
      (c) => c.role === 'dimension' && !c.field.includes(':Measure Names')
    );

    let pivotOn: QuerySpec['pivotOn'] | undefined;
    let measures: Array<{ alias: string; expr: string }>;

    if (pivotDimRefs.length > 0 && enc.selectedMeasures.length > 0) {
      // Pivot table — build pivotOn config.
      // Prefer the finest date granularity ref (qr: > yr: > plain).
      const pivotRef = pivotDimRefs.find((r) => r.field.startsWith('qr:'))
        ?? pivotDimRefs.find((r) => r.field.startsWith('month:'))
        ?? pivotDimRefs.find((r) => r.field.startsWith('yr:'))
        ?? pivotDimRefs[0];
      const pivotDimName = pivotRef.caption ?? cleanFieldRef(pivotRef.field);
      const pivotExpr = resolveDatePivotExpr(pivotRef.field, pivotDimName, workbook);
      const pivotValues = extractPivotValues(pivotDimName, workbook);
      const baseMeasureName = enc.selectedMeasures[0];
      // measureExpr must be plain (no outer SUM) — the server wraps it in SUM(CASE WHEN ...)
      const measureExpr = resolveFormulaExpr(baseMeasureName, workbook, false);

      if (!isUntranslatable(measureExpr)) {
        pivotOn = { expr: pivotExpr, values: pivotValues, measureExpr, totalAlias: baseMeasureName };
      } else {
        // Measure references Tableau Parameters/federated — fall back to first raw numeric field
        const fallbackExpr = findFallbackMeasureExpr(workbook);
        if (fallbackExpr) {
          pivotOn = { expr: pivotExpr, values: pivotValues, measureExpr: fallbackExpr, totalAlias: baseMeasureName };
        }
      }
      measures = []; // measures are expressed as pivot columns
    } else {
      // Filter out measures whose expressions can't run in DuckDB
      measures = resolveQueryMeasures(enc.selectedMeasures, workbook)
        .filter((m) => !isUntranslatable(m.expr));
      // If all measures were untranslatable, try a raw numeric fallback
      if (measures.length === 0) {
        const fallbackExpr = findFallbackMeasureExpr(workbook);
        if (fallbackExpr) {
          const fallbackAlias = enc.selectedMeasures[0] ?? 'Total';
          measures = [{ alias: fallbackAlias, expr: `SUM(${fallbackExpr})` }];
        }
      }
    }

    // Skip specs with no pivot and no measures — they'd return raw rows that
    // don't match any column definition, so sample data is better.
    if (!pivotOn && measures.length === 0) continue;

    specs[enc.worksheet] = {
      source: defaultSlug,
      dimensions: dims,
      measures,
      filterFields,
      ...(pivotOn ? { pivotOn } : {}),
    };
  }

  const knownSources = parquetTables.map((t) => t.slug);
  const serverTs = buildServerFile(knownSources);

  return { serverTs, specs };
}

/** True when an expression contains Tableau-specific constructs DuckDB can't execute. */
function isUntranslatable(expr: string): boolean {
  return /\bParameters\b|\bfederated\.[a-z0-9]+\b/i.test(expr);
}

/**
 * Find a raw (non-calculated) numeric field from the workbook to use as a
 * fallback measure when the real measure is untranslatable. Returns the
 * quoted SQL column reference, e.g. `"Item Amount Total"`.
 */
/**
 * Find a raw (non-calculated) numeric field from the workbook to use as a
 * fallback measure when the real measure is untranslatable.
 * Skips fields with captions (caption ≠ column name in the parquet).
 * Returns the quoted SQL column reference, e.g. `"Item Amount Total"`.
 */
function findFallbackMeasureExpr(workbook: TableauWorkbook): string | null {
  const candidates = workbook.fields.filter(
    (f) =>
      f.role === 'measure' &&
      !f.calculation?.formula &&
      f.dataType === 'real' &&
      !f.caption && // skip aliased fields — their caption ≠ parquet column name
      !/Number of Records|Latitude|Longitude/i.test(f.name),
  );
  // Prefer fields with "Total" in the name (more likely to be the summary amount column)
  const field = candidates.sort((a, b) =>
    (/total/i.test(b.name) ? 1 : 0) - (/total/i.test(a.name) ? 1 : 0)
  )[0];
  if (!field) return null;
  return `"${field.name.replace(/^\[|\]$/g, '')}"`;
}

// ─── Measure resolver ─────────────────────────────────────────────────────────

function resolveQueryMeasures(
  selectedMeasures: string[],
  workbook: TableauWorkbook,
): Array<{ alias: string; expr: string }> {
  return selectedMeasures.map((name) => {
    if (name.startsWith('% of ')) {
      const base = name.slice('% of '.length);
      const baseExpr = resolveFormulaExpr(base, workbook);
      return {
        alias: name,
        expr: `(${baseExpr}) / NULLIF(SUM(${baseExpr}) OVER (), 0)`,
      };
    }
    return { alias: name, expr: resolveFormulaExpr(name, workbook) };
  });
}

/**
 * Translate a Tableau column-shelf date ref (e.g. "qr:Date Closed:ok") to a DuckDB
 * expression that produces orderable string labels like "Q3 2024" or "2024".
 * Falls back to resolveDimExpr for non-date refs.
 */
function resolveDatePivotExpr(fieldRef: string, dimName: string, workbook: TableauWorkbook): string {
  if (fieldRef.startsWith('qr:')) {
    const col = cleanFieldRef(fieldRef);
    // ISO-sortable format "YYYY-QN" so alphabetical order == chronological order
    return `CAST(DATEPART('year', "${col}"::DATE) AS VARCHAR) || '-Q' || CAST(DATEPART('quarter', "${col}"::DATE) AS VARCHAR)`;
  }
  if (fieldRef.startsWith('month:')) {
    const col = cleanFieldRef(fieldRef);
    return `STRFTIME("${col}"::DATE, '%Y-%m')`;
  }
  if (fieldRef.startsWith('yr:') || fieldRef.startsWith('year:')) {
    const col = cleanFieldRef(fieldRef);
    return `CAST(DATEPART('year', "${col}"::DATE) AS VARCHAR)`;
  }
  return resolveDimExpr(dimName, workbook) ?? `"${dimName}"`;
}

/**
 * Returns the SQL expression for a dimension.
 * Returns null when the field is a Tableau Group/Set (present in fields but no formula)
 * — these have no SQL equivalent and must be skipped.
 */
function resolveDimExpr(name: string, workbook: TableauWorkbook): string | null {
  // Named calculation (IF/THEN/CASE WHEN style)
  const calc = workbook.calculations.find((c) => c.name === name);
  if (calc) return translateFormula(calc.formula, workbook, false);

  // Field-level calculation (caption lookup)
  const field = workbook.fields.find((f) => f.caption === name);
  if (field) {
    if (field.calculation?.formula) return translateFormula(field.calculation.formula, workbook, false);
    // Field exists but has no formula → Tableau Group/Set, untranslatable
    return null;
  }

  // Assume raw parquet column
  return `"${name}"`;
}

/** Extract THEN literal values from an IF/THEN/ELSE calculation formula. */
function extractPivotValues(dimName: string, workbook: TableauWorkbook): string[] {
  const calc = workbook.calculations.find((c) => c.name === dimName);
  if (!calc) return [];
  const formula = calc.formula
    .replace(/&#13;/g, ' ').replace(/&#10;/g, ' ')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
  return [...formula.matchAll(/THEN\s+"([^"]+)"/gi)]
    .map((m) => m[1].trim())
    .filter((v) => v.toLowerCase() !== 'null' && v.length > 0);
}

function resolveFormulaExpr(name: string, workbook: TableauWorkbook, aggContext = true): string {
  const calc = workbook.calculations.find((c) => c.name === name);
  if (calc) return translateFormula(calc.formula, workbook, aggContext);

  // Try matching by field internal name (Calculation_XXXXXX → caption)
  const field = workbook.fields.find((f) => f.caption === name);
  if (field?.calculation?.formula) {
    return translateFormula(field.calculation.formula, workbook, true);
  }

  // Plain field — default to SUM
  return `SUM("${name}")`;
}

/**
 * Translate a Tableau formula to DuckDB SQL.
 * When `aggContext` is true (GROUP BY query), leaf field refs that have no
 * explicit aggregation in the formula are automatically wrapped with SUM().
 * Formulas that already contain SUM/AVG/MIN/MAX manage their own aggregation.
 */
function translateFormula(formula: string, workbook: TableauWorkbook, aggContext = false): string {
  let f = formula
    .replace(/&#13;/g, ' ').replace(/&#10;/g, ' ')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');

  // Convert Tableau string literals "..." to SQL single-quoted strings.
  // In Tableau formulas, double quotes are string delimiters (not column refs).
  // Column refs use [brackets]. Do this BEFORE field ref substitution.
  // Also trim leading/trailing whitespace from string values so comparisons
  // match the trimmed values produced by extractPivotValues.
  f = f.replace(/"([^"]+)"/g, (_m, s: string) => `'${s.trim()}'`);

  // If the formula already has explicit aggregation, translate refs as plain fields
  const hasExplicitAgg = /\b(SUM|AVG|MIN|MAX|COUNT|COUNTD)\s*\(/i.test(f);
  const wrapWithSum = aggContext && !hasExplicitAgg;

  // Resolve [FieldName] references
  f = f.replace(/\[([^\]]+)\]/g, (_match, ref: string) => {
    // Calculation by caption name
    const calc = workbook.calculations.find((c) => c.name === ref);
    if (calc) return `(${translateFormula(calc.formula, workbook, aggContext)})`;

    // Calculation by internal ID
    const field = workbook.fields.find(
      (fld) => fld.name === `[${ref}]` || fld.name.replace(/^\[|\]$/g, '') === ref
    );
    if (field?.calculation?.formula) {
      return `(${translateFormula(field.calculation.formula, workbook, aggContext)})`;
    }

    // Leaf field — wrap with SUM when in an aggregated context
    return wrapWithSum ? `SUM("${ref}")` : `"${ref}"`;
  });

  // IF/THEN → CASE WHEN
  if (/\bIF\b/i.test(f)) {
    f = f
      .replace(/\bIF\b\s*/gi, 'CASE WHEN ')
      .replace(/\bELSE IF\b\s*/gi, 'WHEN ')
      .replace(/\bTHEN\b\s*/gi, 'THEN ')
      .replace(/\bELSE\b\s*/gi, 'ELSE ')
      .replace(/\bEND\b/gi, 'END');
  }

  return f.replace(/\s+/g, ' ').trim();
}

// ─── Generic server template ──────────────────────────────────────────────────

function buildServerFile(knownSources: string[]): string {
  const sourcesJson = JSON.stringify(knownSources);

  return `import express from 'express';
import cors from 'cors';
import duckdb from 'duckdb';
import { existsSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const DATA_DIR = join(__dirname, 'data');
const PORT = process.env.API_PORT ? Number(process.env.API_PORT) : 3001;
const KNOWN_SOURCES: string[] = ${sourcesJson};

const db = new duckdb.Database(':memory:');

// One connection per request — safe for concurrent queries.
// DuckDB connections on an in-memory DB share the same engine but
// have independent transaction state, so concurrent calls never collide.
function dbQuery(
  sql: string,
  params: (string | null)[] = [],
): Promise<Record<string, unknown>[]> {
  return new Promise((resolve, reject) => {
    const conn = db.connect();
    conn.all(sql, ...params, (err: Error | null, rows: Record<string, unknown>[]) => {
      conn.close();
      if (err) reject(err);
      else resolve(rows ?? []);
    });
  });
}

interface MeasureSpec { alias: string; expr: string; }
interface DimSpec { alias: string; expr: string; }
interface PivotSpec { expr: string; values: string[]; measureExpr: string; totalAlias?: string; }

interface QueryRequest {
  source: string;
  dimensions?: (DimSpec | string)[];
  measures?: MeasureSpec[];
  pivotOn?: PivotSpec;
  filters?: Record<string, string | null | undefined>;
  orderBy?: { column: string; direction?: 'ASC' | 'DESC' };
  /** Pagination — page is 1-indexed, pageSize defaults to 50 */
  page?: number;
  pageSize?: number;
}

// Normalise a dimension entry — new callers send DimSpec; legacy callers send plain strings.
function dimAlias(d: DimSpec | string): string { return typeof d === 'string' ? d : d.alias; }
function dimExpr(d: DimSpec | string): string { return typeof d === 'string' ? \`"\${d}"\` : d.expr; }
// Only emit AS alias when the expression isn't already a plain quoted column ref of the same name,
// which avoids DuckDB's "alias cannot be used in GROUP BY" error.
function dimSelect(d: DimSpec | string): string {
  const alias = dimAlias(d);
  const expr = dimExpr(d);
  return expr === \`"\${alias}"\` ? \`  \${expr}\` : \`  \${expr} AS "\${alias.replace(/"/g, '\\\\"')}"\`;
}

function buildSQL(parquetPath: string, req: QueryRequest): { sql: string; params: (string | null)[] } {
  const params: (string | null)[] = [];

  // WHERE — use ? placeholders. For calculated dimensions (CASE WHEN etc.) use their
  // SQL expression rather than the alias, since the alias is not a real column.
  const dimExprMap = new Map<string, string>(
    (req.dimensions ?? []).map((d) => [dimAlias(d), dimExpr(d)])
  );
  const whereParts: string[] = [];
  if (req.filters) {
    for (const [field, value] of Object.entries(req.filters)) {
      if (value !== null && value !== undefined) {
        const expr = dimExprMap.get(field) ?? \`"\${field}"\`;
        whereParts.push(\`  (\${expr}) = ?\`);
        params.push(value);
      }
    }
  }

  // Ordinal GROUP BY — avoids alias/column-name conflicts in DuckDB.
  // Dimensions are always emitted first in SELECT, so positions are 1..N.
  const groupBy = req.dimensions?.length
    ? \`GROUP BY \${req.dimensions.map((_, i) => i + 1).join(', ')}\`
    : '';

  let selectParts: string[];
  let orderBy: string;

  if (req.pivotOn) {
    const { expr: pivotExpr, values, measureExpr, totalAlias } = req.pivotOn;
    // Conditional aggregation: one SUM(CASE WHEN ...) per bucket value
    const bucketCols = values.map((v) =>
      \`  SUM(CASE WHEN (\${pivotExpr}) = '\${v.replace(/'/g, "''")}' THEN (\${measureExpr}) ELSE 0 END) AS "\${v.replace(/"/g, '\\\\"')}"\`
    );
    const totalExpr = \`SUM(\${measureExpr})\`;
    selectParts = [
      ...(req.dimensions ?? []).map((d) => dimSelect(d)),
      ...bucketCols,
      \`  \${totalExpr} AS "\${(totalAlias ?? 'Total').replace(/"/g, '\\\\"')}"\`,
    ];
    orderBy = \`ORDER BY \${totalExpr} DESC NULLS LAST\`;
  } else {
    selectParts = [
      ...(req.dimensions ?? []).map((d) => dimSelect(d)),
      ...(req.measures ?? []).map((m) => \`  \${m.expr} AS "\${m.alias.replace(/"/g, '\\\\"')}"\`),
    ];
    if (!selectParts.length) selectParts.push('  *');
    orderBy = req.orderBy
      ? \`ORDER BY "\${req.orderBy.column}" \${req.orderBy.direction ?? 'DESC'} NULLS LAST\`
      : '';
  }

  const pageSize = req.pageSize ?? 50;
  const page = Math.max(1, req.page ?? 1);
  const offset = (page - 1) * pageSize;

  const sql = [
    'SELECT',
    selectParts.join(',\\n'),
    \`FROM read_parquet('\${parquetPath}')\`,
    whereParts.length ? \`WHERE\\n\${whereParts.join('\\n  AND\\n')}\` : '',
    groupBy,
    orderBy,
    \`LIMIT \${pageSize} OFFSET \${offset}\`,
  ].filter(Boolean).join('\\n');

  return { sql, params };
}

const app = express();
app.use(cors());
app.use(express.json());

app.get('/api/health', (_req, res) => {
  res.json({ status: 'ok', sources: KNOWN_SOURCES });
});

app.get('/api/sources', (_req, res) => {
  const sources = KNOWN_SOURCES.map((slug) => ({
    slug,
    available: existsSync(join(DATA_DIR, \`\${slug}.parquet\`)),
  }));
  res.json(sources);
});

app.post('/api/query', async (req, res) => {
  const body = req.body as QueryRequest;
  const { source } = body;

  if (!source || !KNOWN_SOURCES.includes(source)) {
    return res.status(400).json({
      error: \`Unknown source "\${source}". Available: \${KNOWN_SOURCES.join(', ')}\`,
    });
  }

  const parquetPath = join(DATA_DIR, \`\${source}.parquet\`);
  if (!existsSync(parquetPath)) {
    return res.status(503).json({
      error: \`Data not extracted yet. Run: drexo migrate <workbook.twbx>\`,
    });
  }

  try {
    // Auto-discover pivot values when none are supplied (e.g. date-based pivots)
    if (body.pivotOn && body.pivotOn.values.length === 0) {
      const distSql = \`SELECT DISTINCT (\${body.pivotOn.expr}) AS v FROM read_parquet('\${parquetPath}') WHERE (\${body.pivotOn.expr}) IS NOT NULL ORDER BY v\`;
      const distRows = await dbQuery(distSql);
      body.pivotOn.values = distRows.map((r) => String(r['v']));
    }
    const { sql, params } = buildSQL(parquetPath, body);
    const rows = await dbQuery(sql, params);
    res.json(rows);
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

app.listen(PORT, () => {
  console.log(\`[api] drexo API → http://localhost:\${PORT}\`);
  console.log(\`[api] Sources: \${KNOWN_SOURCES.join(', ')}\`);
});
`;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

interface ActionFilter { field: string; appliedTo: string[]; }

function extractActionFilters(workbook: TableauWorkbook): ActionFilter[] {
  const result: ActionFilter[] = [];
  for (const f of workbook.filters) {
    // f.name is now cleaned — use raw f.field to detect action filters
    const match = (f.field ?? f.name).match(/Action \(([^)]+)\)/);
    if (!match) continue;
    for (const field of match[1].split(',').map((s) => s.trim())) {
      const existing = result.find((af) => af.field === field);
      if (existing) {
        for (const ws of f.appliedTo) {
          if (!existing.appliedTo.includes(ws)) existing.appliedTo.push(ws);
        }
      } else {
        result.push({ field, appliedTo: [...f.appliedTo] });
      }
    }
  }
  return result;
}

function cleanFieldRef(field: string): string {
  const parts = field.split(':');
  const mid = parts.length >= 3 ? parts.slice(1, -1).join(':') : field;
  return mid.replace(/^\[|\]$/g, '');
}
