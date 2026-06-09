import type { Filter, FieldRole } from '../model.js';
import { toArray } from './_helpers.js';

/**
 * Extract filters from each worksheet.
 *
 * v0.1 reports filters at worksheet scope only; v0.2+ will detect
 * dashboard-scoped filters via Tableau's `<filter-classes>` and
 * dashboard `<action>` graph.
 */
export function extractFilters(workbook: any): Filter[] {
  const worksheetsRaw = workbook?.worksheets?.worksheet;
  if (!worksheetsRaw) return [];

  const collected: Filter[] = [];

  for (const ws of toArray(worksheetsRaw)) {
    const worksheetName: string = ws['@_name'] ?? 'Unnamed';
    for (const f of collectFilters(ws)) {
      collected.push({ ...f, appliedTo: [worksheetName], scope: 'worksheet' });
    }
  }

  const merged = mergeAcrossWorksheets(collected);

  // Build calc-id → caption lookup from datasource columns
  const calcCaption = new Map<string, string>();
  for (const ds of toArray(workbook?.datasources?.datasource ?? [])) {
    for (const col of toArray(ds?.column ?? [])) {
      const n: string | undefined = col['@_name'];
      const cap: string | undefined = col['@_caption'];
      if (n && cap) calcCaption.set(n.replace(/^\[|\]$/g, ''), cap);
    }
  }

  // Collect fields shown as interactive controls in any dashboard (type="filter" zones).
  // visibleNames stores the human-readable field name (caption or cleaned field name).
  const visibleNames = new Set<string>();
  for (const db of toArray(workbook?.dashboards?.dashboard ?? [])) {
    collectFilterCtrlFields(db?.zones?.zone, visibleNames, calcCaption);
  }

  for (const f of merged) {
    if (visibleNames.has(f.name)) f.visible = true;
  }

  return merged;
}

function collectFilterCtrlFields(zones: any, out: Set<string>, calcCaption: Map<string, string>): void {
  for (const z of toArray(zones ?? [])) {
    if (z['@_type'] === 'filter' && z['@_param']) {
      // Extract core id from param: "[ds].[none:Calculation_xxx:nk]" → "Calculation_xxx"
      const raw = String(z['@_param']).replace(/^\[|\]$/g, '');
      const afterDot = raw.includes('].[') ? raw.split('].[').pop()! : raw;
      const stripped = afterDot.replace(/^\[/, '').replace(/\]$/, '');
      const core = stripped.replace(/^:?none:/i, '').replace(/:[a-z]+$/i, '');
      // Resolve to human-readable name via caption map, then fall back to cleanFilterName
      const caption = calcCaption.get(core) ?? cleanFilterName(raw);
      out.add(caption);
    }
    if (z.zone) collectFilterCtrlFields(z.zone, out, calcCaption);
  }
}

function collectFilters(ws: any): Omit<Filter, 'appliedTo' | 'scope'>[] {
  const result: Omit<Filter, 'appliedTo' | 'scope'>[] = [];

  function walk(node: any): void {
    if (!node || typeof node !== 'object') return;
    if (node.filter) {
      for (const f of toArray(node.filter)) {
        const item = mapFilter(f);
        if (item) result.push(item);
      }
    }
    for (const child of Object.values(node)) {
      if (child && typeof child === 'object') walk(child);
    }
  }
  walk(ws);
  return result;
}

function mapFilter(f: any): Omit<Filter, 'appliedTo' | 'scope'> | undefined {
  const fieldRaw: string | undefined = f['@_column'] ?? f['@_field'];
  if (!fieldRaw) return undefined;

  const field = fieldRaw.replace(/^\[|\]$/g, '');
  const filterClass: string | undefined = f['@_class'];
  const type: FieldRole = filterClass === 'quantitative' ? 'measure' : 'dimension';
  const name = cleanFilterName(field);

  // Extract selected/included member values for categorical filters.
  // Tableau uses two formats: <member value="..." included="true"/> and
  // <groupfilter function="member" member="&quot;Value&quot;"/>
  const members = toArray(f?.members?.member ?? f?.member ?? []);
  const memberValues: string[] = members
    .filter((m: any) => m['@_included'] !== 'false')
    .map((m: any) => String(m['@_value'] ?? ''))
    .filter((v: string) => v !== '');

  const groupFilters = toArray(f?.groupfilter ?? []);
  const groupValues: string[] = groupFilters
    .filter((g: any) => g['@_function'] === 'member' && g['@_member'])
    .map((g: any) => String(g['@_member']).replace(/^"|"$/g, '').trim())
    .filter((v: string) => v !== '');

  const values = [...memberValues, ...groupValues];

  return { name, field, type, ...(values.length > 0 ? { values } : {}) };
}

/**
 * Convert a raw Tableau field ref to a human-readable filter name.
 *
 * Examples:
 *   "excel.41723].[none:TRAN_STATUS:nk"  → "TRAN_STATUS"
 *   "excel.41723].[Action (Ship-To Country)" → "Ship-To Country"
 *   "excel.41723].[:Measure Names" → "Measure Names"
 *   "[none:Company:nk]" → "Company"
 */
function cleanFilterName(raw: string): string {
  // Strip datasource prefix "datasource.xxx].[" leaving just the field part
  const afterDot = raw.includes('].[') ? raw.split('].[').pop()! : raw;
  // Remove trailing bracket
  const stripped = afterDot.replace(/\]$/, '').trim();

  // "Action (Field Name)" → "Field Name"
  const actionMatch = stripped.match(/^Action\s+\((.+)\)$/i);
  if (actionMatch) return actionMatch[1].trim();

  // "none:FieldName:nk" or ":FieldName:nk" → "FieldName"
  const noneMatch = stripped.match(/^:?none:(.+?)(?::nk)?$/i) ?? stripped.match(/^:(.+?)(?::nk)?$/);
  if (noneMatch) return noneMatch[1].trim();

  return stripped;
}

function mergeAcrossWorksheets(filters: Filter[]): Filter[] {
  const byKey = new Map<string, Filter>();
  for (const f of filters) {
    const key = f.field ?? f.name;
    const existing = byKey.get(key);
    if (existing) {
      existing.appliedTo = Array.from(new Set([...existing.appliedTo, ...f.appliedTo]));
    } else {
      byKey.set(key, { ...f });
    }
  }
  return Array.from(byKey.values());
}
