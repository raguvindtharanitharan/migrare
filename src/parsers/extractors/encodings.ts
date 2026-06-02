import type { FieldRef, FieldRole, MarkType, VisualEncoding } from '../model.js';
import { toArray, extractFormattedTitle } from './_helpers.js';
import { classifyMarks, resolveEffectiveMarkType } from './marks.js';

/**
 * Extract per-worksheet visual encodings (shelves + mark + color/size/...).
 *
 * Pill grammar: Tableau emits shelf fields as either `[field]` or the
 * qualified form `[datasource].[field]`. The pill parser handles both
 * and keeps the datasource separate on the resulting `FieldRef`.
 *
 * Field captions are resolved from a workbook-wide lookup built from
 * every `<column>` node (datasource definitions live at workbook
 * root, not under each worksheet). Resolution is name-only — the
 * lookup stores multiple name variants per field so raw expressions
 * like `[ds].[avg:Sales:qk]` find the caption defined on the
 * `<column name="[avg:Sales:qk]">` node.
 *
 * Shelf entries are de-duplicated by `${datasource}|${field}` after
 * resolution. Multi-pane worksheets repeat the same `<rows>` /
 * `<color>` block per pane, and we collapse those copies here.
 */
export function extractVisualEncodings(workbook: any): VisualEncoding[] {
  const worksheetsRaw = workbook?.worksheets?.worksheet;
  if (!worksheetsRaw) return [];

  const lookup = buildWorkbookFieldLookup(workbook);
  return toArray(worksheetsRaw).map((ws: any) => mapEncoding(ws, lookup));
}

type FieldLookup = Map<string, { caption?: string; role?: FieldRole }>;

function mapEncoding(ws: any, lookup: FieldLookup): VisualEncoding {
  const worksheet: string = ws['@_name'] ?? 'Unnamed';
  const title = extractFormattedTitle(ws);

  const markTypes = classifyMarks(ws);
  const markType: MarkType = markTypes[0] ?? 'automatic';

  const rows = dedupe(extractShelf(ws, 'rows', lookup));
  const columns = dedupe(extractShelf(ws, 'cols', lookup));
  const effectiveMarkType = resolveEffectiveMarkType(markType, rows, columns);

  // Always extract which measures are instantiated in the view.
  // For Measure Names tables: driven by groupfilter (display order preserved).
  // For pivot/standard tables: driven by column-instances (ground truth).
  const selectedMeasures = extractSelectedMeasures(ws, lookup);

  const encodings = collectEncodingNodes(ws);
  const color = dedupe(encodingsForAttr(encodings, 'color', lookup));
  const size = dedupe(encodingsForAttr(encodings, 'size', lookup));
  const shape = dedupe(encodingsForAttr(encodings, 'shape', lookup));
  const detail = dedupe(encodingsForAttr(encodings, 'detail', lookup));
  const tooltip = dedupe(encodingsForAttr(encodings, 'tooltip', lookup));
  const label = dedupe(encodingsForAttr(encodings, 'label', lookup));

  return {
    worksheet,
    title,
    markType,
    effectiveMarkType,
    isDualAxis: false,
    rows,
    columns,
    selectedMeasures,
    color: toEncodingValue(color),
    size: toEncodingValue(size),
    shape: toEncodingValue(shape),
    detail,
    tooltip,
    label,
    sorting: [],
    referenceLines: [],
    tableCalculations: [],
  };
}

// ---------------------------------------------------------------------------
// Measure Names filter — extract which measures Tableau selected to display
// ---------------------------------------------------------------------------

function extractSelectedMeasures(ws: any, lookup: FieldLookup): string[] {
  const view = ws?.table?.view;
  if (!view) return [];

  // Primary: groupfilter preserves Tableau's exact display order.
  const fromFilter = extractMeasuresFromGroupFilter(view, lookup);
  if (fromFilter.length > 0) return fromFilter;

  // Fallback: column-instances — works even when no explicit filter is set.
  return extractMeasuresFromColumnInstances(view, lookup);
}

function extractMeasuresFromColumnInstances(view: any, lookup: FieldLookup): string[] {
  const resolved: string[] = [];
  const deps = toArray(view['datasource-dependencies'] ?? []);

  for (const dep of deps) {
    const instances = toArray(dep['column-instance'] ?? []);
    for (const inst of instances) {
      const colRef = String(inst['@_name'] ?? '').replace(/^\[|\]$/g, '');
      if (!colRef) continue;

      // Only measure aggregations — skip dimension instances (none:, attr: for dims, etc.)
      if (/^(none|trunc|datepart|datetrunc):/i.test(colRef)) continue;

      // pcto: = percent of total — emit as "% of <BaseName>"
      if (/^pcto:/i.test(colRef)) {
        const baseRef = colRef.replace(/^pcto:/i, '').replace(/:qk(:\d+)?$/, '').replace(/^sum:/, '');
        const baseCaption = resolveFieldName(baseRef, lookup)?.caption;
        if (baseCaption) {
          const label = `% of ${baseCaption}`;
          if (!resolved.includes(label)) resolved.push(label);
        }
        continue;
      }

      // Skip pure table calcs that require window functions
      if (/^(running_sum|lookup|window|rank|size|index|first|last|previous_value):/i.test(colRef)) continue;

      // Must start with a numeric measure aggregation.
      // attr: is Tableau's Attribute function — it is a dimension aggregation
      // (returns a value only when all rows agree), NOT a numeric measure.
      if (!/^(sum|avg|min|max|cnt|count|median|stdev|var):/i.test(colRef)) continue;

      const caption = resolveFieldName(colRef, lookup)?.caption;
      if (caption && !resolved.includes(caption)) resolved.push(caption);
    }
  }

  return resolved;
}

function extractMeasuresFromGroupFilter(view: any, lookup: FieldLookup): string[] {
  const filters = toArray(view.filter ?? []);
  const mnFilter = filters.find(
    (f: any) => typeof f['@_column'] === 'string' && f['@_column'].includes('Measure Names')
  );
  if (!mnFilter?.groupfilter) return [];

  const innerGf = mnFilter.groupfilter.groupfilter ?? mnFilter.groupfilter;
  const resolved: string[] = [];

  for (const m of toArray(innerGf)) {
    if (m['@_function'] !== 'member' || !m['@_member']) continue;
    const raw = String(m['@_member']).replace(/^"|"$/g, '');
    const match = raw.match(/\.\[([^\]]+)\]$/);
    const fieldRef = match ? match[1] : raw.replace(/^\[|\]$/g, '');

    if (/^pcto:/i.test(fieldRef)) {
      const baseRef = fieldRef.replace(/^pcto:/i, '').replace(/:qk(:\d+)?$/, '').replace(/^sum:/, '');
      const baseCaption = resolveFieldName(baseRef, lookup)?.caption;
      if (baseCaption) {
        const label = `% of ${baseCaption}`;
        if (!resolved.includes(label)) resolved.push(label);
      }
      continue;
    }

    if (/^(running_sum|lookup|window|rank|size|index|first|last|previous_value):/i.test(fieldRef)) continue;

    const caption = resolveFieldName(fieldRef, lookup)?.caption;
    if (caption && !resolved.includes(caption)) resolved.push(caption);
  }

  return resolved;
}

function toEncodingValue(refs: FieldRef[]): FieldRef | FieldRef[] | null {
  if (refs.length === 0) return null;
  if (refs.length === 1) return refs[0]!;
  return refs;
}

// ---------------------------------------------------------------------------
// Pill parser
// ---------------------------------------------------------------------------

interface Pill {
  datasource?: string;
  field: string;
}

/**
 * Parse a shelf expression into pills. Each pill is one field
 * reference, optionally qualified by a datasource.
 *
 *   "[foo]"               → [{ field: "foo" }]
 *   "[ds].[foo]"          → [{ datasource: "ds", field: "foo" }]
 *   "[a]+[b]"             → [{ field: "a" }, { field: "b" }]
 *   "[ds].[a]+[ds].[b]"   → [{ datasource: "ds", field: "a" },
 *                             { datasource: "ds", field: "b" }]
 */
export function parsePills(expr: string): Pill[] {
  const re = /\[([^\]]+)\](?:\.\[([^\]]+)\])?/g;
  const pills: Pill[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(expr)) !== null) {
    if (m[2] !== undefined) {
      pills.push({ datasource: m[1]!, field: m[2] });
    } else {
      pills.push({ field: m[1]! });
    }
  }
  return pills;
}

function pillToFieldRef(pill: Pill, lookup: FieldLookup): FieldRef {
  const resolved = resolveFieldName(pill.field, lookup);
  const ref: FieldRef = { field: pill.field };
  if (pill.datasource !== undefined) ref.datasource = pill.datasource;
  if (resolved?.caption !== undefined) ref.caption = resolved.caption;
  if (resolved?.role !== undefined) ref.role = resolved.role;
  return ref;
}

// ---------------------------------------------------------------------------
// Field lookup
// ---------------------------------------------------------------------------

function buildWorkbookFieldLookup(workbook: any): FieldLookup {
  const lookup: FieldLookup = new Map();
  walk(workbook, (node) => {
    if (looksLikeColumnNode(node)) addColumnVariants(node, lookup);
  });
  return lookup;
}

function looksLikeColumnNode(node: any): boolean {
  if (!node || typeof node !== 'object') return false;
  if (!node['@_name']) return false;
  return !!(node['@_datatype'] || node.calculation || node['@_role'] || node['@_caption']);
}

function addColumnVariants(col: any, lookup: FieldLookup): void {
  const name: string = col['@_name'];
  const caption: string | undefined = col['@_caption'];
  const role: FieldRole | undefined = col['@_role'] as FieldRole | undefined;
  const entry = { caption, role };

  const variants = new Set<string>();
  variants.add(name);
  variants.add(name.toLowerCase());
  const unbracketed = name.replace(/^\[|\]$/g, '');
  variants.add(unbracketed);
  variants.add(unbracketed.toLowerCase());

  const lastDot = unbracketed.split('.').pop() ?? unbracketed;
  variants.add(lastDot);
  variants.add(lastDot.toLowerCase());

  if (unbracketed.includes(':')) {
    const parts = unbracketed.split(':');
    const middle = parts.slice(1, -1).join(':');
    if (middle) {
      variants.add(middle);
      variants.add(middle.toLowerCase());
    }
    if (parts.length >= 2 && parts[1]) {
      variants.add(parts[1]);
      variants.add(parts[1]!.toLowerCase());
    }
  }

  for (const v of variants) {
    if (v && v.length > 1) {
      const existing = lookup.get(v);
      if (!existing || (!existing.caption && caption)) {
        lookup.set(v, entry);
      }
    }
  }
}

function resolveFieldName(
  raw: string,
  lookup: FieldLookup
): { caption?: string; role?: FieldRole } | undefined {
  let m = lookup.get(raw);
  if (m) return m;

  const unbracketed = raw.replace(/^\[|\]$/g, '');
  m = lookup.get(unbracketed);
  if (m) return m;

  // Tableau field refs use a colon-delimited grammar like
  // `agg:basename:typeClass` (`avg:Sales:qk`) or, with table calcs and
  // virtual-table variants, deeper forms like
  // `pcto:usr:Calculation_299...:qk:2`. The base field name can sit at
  // any middle position. Try every segment against the lookup; the
  // first that matches a known column wins.
  if (unbracketed.includes(':')) {
    const parts = unbracketed.split(':');
    for (const part of parts) {
      if (!part) continue;
      m = lookup.get(part) ?? lookup.get(part.toLowerCase());
      if (m) return m;
    }
  }

  const lastDot = unbracketed.split('.').pop() ?? unbracketed;
  m = lookup.get(lastDot);
  if (m) return m;

  return undefined;
}

// ---------------------------------------------------------------------------
// Shelves (rows / cols)
// ---------------------------------------------------------------------------

function extractShelf(ws: any, shelfName: 'rows' | 'cols', lookup: FieldLookup): FieldRef[] {
  const expressions: string[] = [];
  walk(ws, (node) => {
    if (typeof node !== 'object' || !node) return;
    if (shelfName in node) {
      expressions.push(...readShelfValue(node[shelfName]));
    }
  });
  return expressions.flatMap((expr) =>
    parsePills(expr).map((pill) => pillToFieldRef(pill, lookup))
  );
}

function readShelfValue(node: any): string[] {
  if (node === undefined || node === null) return [];
  if (typeof node === 'string') return [node];
  if (Array.isArray(node)) {
    return node
      .map((item) => (typeof item === 'string' ? item : item?.['#text'] ?? ''))
      .filter(Boolean);
  }
  if (typeof node === 'object') {
    const text = node['#text'] ?? node[''] ?? '';
    return text ? [text] : [];
  }
  return [];
}

// ---------------------------------------------------------------------------
// <encoding attr="color" field="..."> nodes
// ---------------------------------------------------------------------------

function collectEncodingNodes(ws: any): any[] {
  const out: any[] = [];
  walk(ws, (node) => {
    if (node && typeof node === 'object' && node['@_attr']) {
      out.push(node);
    }
  });
  return out;
}

function encodingsForAttr(nodes: any[], attr: string, lookup: FieldLookup): FieldRef[] {
  return nodes
    .filter((n) => n['@_attr'] === attr)
    .flatMap((n) => toFieldRefs(n, lookup));
}

function toFieldRefs(enc: any, lookup: FieldLookup): FieldRef[] {
  const fieldRaw: string | undefined = enc['@_field'];
  if (!fieldRaw) return [];
  const pills = parsePills(fieldRaw);
  if (pills.length === 0) return [];
  return pills.map((pill) => pillToFieldRef(pill, lookup));
}

// ---------------------------------------------------------------------------
// Dedupe
// ---------------------------------------------------------------------------

function dedupe(refs: FieldRef[]): FieldRef[] {
  const seen = new Set<string>();
  const out: FieldRef[] = [];
  for (const ref of refs) {
    const key = `${ref.datasource ?? ''}|${ref.field}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(ref);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Generic tree walk
// ---------------------------------------------------------------------------

function walk(node: any, visit: (node: any) => void): void {
  if (!node || typeof node !== 'object') return;
  visit(node);
  for (const child of Object.values(node)) {
    if (child && typeof child === 'object') walk(child, visit);
  }
}
