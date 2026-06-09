import type { TableauWorkbook, MarkType } from '../../parsers/model.js';
import {
  collectLeafZones, buildFiltersByWorksheet, pickDashboard,
  cleanRef, markTypeLabel, zoneColors, pct,
  type LeafKind,
} from './zone-helpers.js';
import { resolveTableSpec, type TableColumn } from './table-generator.js';
import type { QuerySpec } from './api-server.js';

// ─── Zone view model ──────────────────────────────────────────────────────────

interface ZoneViewModel {
  label: string;
  internalName: string;
  kind: LeafKind;
  markType: MarkType;
  isControl: boolean;
  isTextButton: boolean;
  filterFields: string[];
  rowFields: string[];
  colFields: string[];
  hasMeasureNames: boolean;
  controlMode?: string;
  paramRef?: string;
  // Table-specific
  isTable: boolean;
  tableColumns: TableColumn[];
  dataFile?: string;        // sample data fallback (relative path)
  querySpec?: QuerySpec;    // live API query spec
  left: string; top: string; width: string; height: string;
}

export interface DashboardGeneratorResult {
  dashboardTsx: string;
  dataFiles: Array<{ relativePath: string; content: string }>;
}

export function generateDashboardComponent(
  workbook: TableauWorkbook,
  slug: string,
  querySpecs: Record<string, QuerySpec> = {},
): DashboardGeneratorResult {
  const dashboard = pickDashboard(workbook);

  if (!dashboard) return { dashboardTsx: fallbackComponent(workbook.metadata.name), dataFiles: [] };

  const leaves = collectLeafZones(dashboard.zones);
  if (leaves.length === 0) return { dashboardTsx: fallbackComponent(workbook.metadata.name), dataFiles: [] };

  const minX = Math.min(...leaves.map((z) => z.x));
  const minY = Math.min(...leaves.map((z) => z.y));
  const maxX = Math.max(...leaves.map((z) => z.x + z.w));
  const maxY = Math.max(...leaves.map((z) => z.y + z.h));
  const totalW = maxX - minX || 1;
  const totalH = maxY - minY || 1;
  const totalArea = totalW * totalH;
  const aspectRatio = Math.round((totalW / totalH) * 1000) / 1000;

  // Build filter field map (non-action filters per worksheet)
  const filtersByWorksheet = buildFiltersByWorksheet(workbook);

  // Map internal calc IDs → captions (e.g. Calculation_316... → "AR Aging Buckets")
  // Used to resolve filterctrl param refs to human-readable field names.
  const calcIdToCaption = new Map<string, string>();
  for (const f of workbook.fields) {
    if (f.name && f.caption) {
      const id = f.name.replace(/^\[|\]$/g, '');
      calcIdToCaption.set(id, f.caption);
    }
  }

  // Compute max area per worksheet name for control detection
  const areaByName = new Map<string, number>();
  for (const z of leaves) {
    const area = z.w * z.h;
    if (!areaByName.has(z.worksheetOrName) || area > areaByName.get(z.worksheetOrName)!) {
      areaByName.set(z.worksheetOrName, area);
    }
  }

  const dataFiles: Array<{ relativePath: string; content: string }> = [];

  const zones: ZoneViewModel[] = leaves.map((z) => {
    const encoding = workbook.visualEncodings.find((e) => e.worksheet === z.worksheetOrName);
    const markType: MarkType = encoding?.effectiveMarkType ?? 'unsupported';
    const area = z.w * z.h;
    const maxArea = areaByName.get(z.worksheetOrName) ?? area;
    const hasEncodings = encoding ? encoding.rows.length > 0 || encoding.columns.length > 0 : false;
    const isTextButton = z.kind === 'worksheet' && !hasEncodings && area < maxArea * 0.5;
    // filterctrl zones are always controls; worksheet zones are never controls (they are tables or unsupported)
    const isControl = z.kind === 'filterctrl' || z.kind === 'paramctrl' || z.kind === 'text' || isTextButton;
    const isTable = !isControl && z.kind === 'worksheet' && markType === 'automatic';

    const rowFields = encoding?.rows
      .map((r) => r.caption ?? cleanRef(r.field))
      .filter((n) => n && !/Latitude|Longitude/i.test(n)) ?? [];
    const colFields = encoding?.columns
      .map((c) => c.caption ?? cleanRef(c.field))
      .filter((n) => n && n !== ':Measure Names') ?? [];
    const hasMeasureNames = encoding?.columns.some((c) => c.field.includes(':Measure Names')) ?? false;
    // filterctrl zones carry the explicit field ref in worksheetOrName (from zone.filterField)
    // — resolve via caption map (for calc IDs) then fall back to the cleaned raw ref.
    const filterFields = z.kind === 'filterctrl'
      ? (() => {
          const raw = cleanFilterRef(z.worksheetOrName);
          const caption = calcIdToCaption.get(raw) ?? raw;
          return [caption];
        })()
      : (filtersByWorksheet.get(z.worksheetOrName) ?? []);

    // Resolve table spec — columns for schema, sample rows as fallback
    let tableColumns: TableColumn[] = [];
    let dataFile: string | undefined;
    let querySpec: QuerySpec | undefined;
    if (isTable) {
      const spec = resolveTableSpec(workbook, z.worksheetOrName);
      tableColumns = spec.columns;
      querySpec = querySpecs[z.worksheetOrName];

      if (spec.columns.length > 0) {
        const wsSlug = z.worksheetOrName.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
        const relPath = `src/workbooks/${slug}/data/${wsSlug}.json`;
        // Always write sample data as fallback for when API is unavailable
        dataFiles.push({ relativePath: relPath, content: JSON.stringify(spec.sampleRows, null, 2) });
        dataFile = `./data/${wsSlug}.json`;
      }
    }

    return {
      label: z.displayLabel || z.worksheetOrName,
      internalName: z.worksheetOrName,
      kind: z.kind,
      markType,
      isControl,
      isTextButton,
      isTable,
      tableColumns,
      dataFile,
      querySpec,
      filterFields,
      rowFields,
      colFields,
      hasMeasureNames,
      controlMode: z.controlMode,
      paramRef: z.paramRef,
      left:   pct((z.x - minX) / totalW),
      top:    pct((z.y - minY) / totalH),
      width:  pct(z.w / totalW),
      height: pct(z.h / totalH),
    };
  });

  // Build a map of filterField → {source, dim} for dropdown population.
  // Collect all candidate sources, then pick the one that has the field as an actual
  // dimension (with its SQL expression). Only fall back to plain column ref when no
  // table has the field as a dimension.
  const filterMeta = new Map<string, { source: string; dim: { alias: string; expr: string } }>();
  // Collect all filter fields mentioned across all table querySpecs
  const allFilterFields = new Set<string>();
  for (const z of zones) {
    if (!z.querySpec) continue;
    for (const f of z.querySpec.filterFields ?? []) allFilterFields.add(f);
    for (const d of z.querySpec.dimensions ?? []) allFilterFields.add(d.alias);
  }
  for (const field of allFilterFields) {
    // Prefer the zone that has this field as a real dimension (carries the SQL expr)
    const zoneWithDim = zones.find((z) => z.querySpec?.dimensions?.some((d) => d.alias === field));
    if (zoneWithDim?.querySpec) {
      const d = zoneWithDim.querySpec.dimensions!.find((d) => d.alias === field)!;
      filterMeta.set(field, { source: zoneWithDim.querySpec.source, dim: d });
      continue;
    }
    // Fall back to any zone that mentions the field in filterFields
    const zoneWithField = zones.find((z) => z.querySpec?.filterFields?.includes(field));
    if (zoneWithField?.querySpec) {
      filterMeta.set(field, {
        source: zoneWithField.querySpec.source,
        dim: { alias: field, expr: `"${field}"` },
      });
    }
  }
  // Resolve filter fields from filterctrl zones not yet in the map (e.g. raw column quick filters).
  const firstTableSource = zones.find((z) => z.querySpec?.source)?.querySpec?.source;
  for (const z of zones) {
    if (!z.isControl) continue;
    for (const field of z.filterFields) {
      if (filterMeta.has(field)) continue;
      const tableWithField = zones.find((t) => t.querySpec?.filterFields?.includes(field));
      const source = tableWithField?.querySpec?.source ?? firstTableSource;
      if (!source) continue;
      const d = tableWithField?.querySpec?.dimensions?.find((d) => d.alias === field);
      filterMeta.set(field, { source, dim: d ?? { alias: field, expr: `"${field}"` } });
    }
  }

  // Collect default values for non-visible (always-on) background filters.
  // Add them to filterMeta so they reach the table filters props, and pre-populate
  // filterState so queries match Tableau's default view on first load.
  const bgDefaults: Record<string, string> = {};
  for (const f of workbook.filters) {
    if (f.visible || !f.values || f.values.length !== 1) continue;
    const field = f.name;
    // Add to filterMeta if not already there (using the first available table source)
    if (!filterMeta.has(field) && firstTableSource) {
      filterMeta.set(field, { source: firstTableSource, dim: { alias: field, expr: `"${field}"` } });
    }
    if (filterMeta.has(field)) bgDefaults[field] = f.values[0];
  }
  const filterStateInit = Object.keys(bgDefaults).length > 0
    ? JSON.stringify(bgDefaults)
    : '{}';

  const hasFilters = filterMeta.size > 0;
  const zonesJsx = zones.map((z) => renderZone(z, filterMeta)).join('\n');

  // Build DataTable import only when at least one table zone exists.
  // Sample data files are imported as fallback; API is used when available.
  const tableZones = zones.filter((z) => z.isTable && z.tableColumns.length > 0);
  const seenVarNames = new Set<string>();
  const dataImports = tableZones
    .filter((z) => z.dataFile)
    .filter((z) => {
      const v = dataVarName(z.internalName);
      if (seenVarNames.has(v)) return false;
      seenVarNames.add(v);
      return true;
    })
    .map((z) => `import ${dataVarName(z.internalName)} from '${z.dataFile}';`)
    .join('\n');
  const hasTable = tableZones.length > 0;

  const dashboardTsx = [
    `/** ${workbook.metadata.name} — layout scaffold generated by drexo */`,
    hasTable ? `import { DataTable } from '../../components/DataTable';` : '',
    hasFilters ? `import { useState, useEffect } from 'react';` : '',
    hasTable && dataImports ? dataImports : '',
    '',
    hasFilters ? filterDropdownComponent() : '',
    `export function Dashboard() {`,
    hasFilters ? `  const [filterState, setFilterState] = useState<Record<string, string | null>>(${filterStateInit});` : '',
    hasFilters ? `  const setFilter = (field: string, value: string | null) => setFilterState((prev) => ({ ...prev, [field]: value }));` : '',
    `  return (`,
    `    <div style={{ position: 'relative', width: '100%', aspectRatio: ${aspectRatio}, fontFamily: "'Poppins', sans-serif", background: 'white', borderRadius: 8, overflow: 'hidden', boxShadow: '0 1px 4px rgba(0,0,0,0.06)' }}>`,
    zonesJsx,
    `    </div>`,
    `  );`,
    `}`,
    '',
  ].filter((l) => l !== null && l !== false).join('\n');

  return { dashboardTsx, dataFiles };
}

// ─── Zone renderer ────────────────────────────────────────────────────────────

function dataVarName(worksheetName: string): string {
  return worksheetName
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_|_$/g, '') + 'Data';
}

function renderZone(z: ZoneViewModel, filterMeta: Map<string, { source: string; dim: { alias: string; expr: string } }>): string {
  // Table zone — render DataTable with API endpoint + sample fallback
  if (z.isTable && z.tableColumns.length > 0) {
    const varName = z.dataFile ? dataVarName(z.internalName) : null;
    const columnsJson = JSON.stringify(z.tableColumns.map((c) => ({
      key: c.key, label: c.label, numeric: c.numeric, currency: c.currency, percent: c.percent,
    })), null, 6).replace(/^/gm, '      ');

    const specJson = z.querySpec
      ? JSON.stringify(z.querySpec, null, 6).replace(/^/gm, '      ')
      : null;

    // Wire filters to shared filterState — null means "no filter applied" (server skips nulls).
    // Include: declared filterFields, dimension aliases (have SQL expr in dimExprMap),
    // AND any control-zone filter from the same source (e.g. TRAN_STATUS raw columns).
    const tableSource = z.querySpec?.source;
    const filterableFields = Array.from(new Set([
      ...(z.querySpec?.filterFields ?? []),
      ...(z.querySpec?.dimensions?.map((d) => d.alias) ?? []),
      ...Array.from(filterMeta.entries())
        .filter(([, meta]) => meta.source === tableSource)
        .map(([field]) => field),
    ])).filter((f) => filterMeta.has(f));
    const filterParamsStr = filterableFields.length
      ? `{{ ${filterableFields.map((f) => `"${esc(f)}": filterState["${esc(f)}"] ?? null`).join(', ')} }}`
      : 'undefined';

    return `      {/* ${esc(z.label)} */}
      <div style={{ position: 'absolute', left: '${z.left}', top: '${z.top}', width: '${z.width}', height: '${z.height}', boxSizing: 'border-box', padding: 8, display: 'flex', flexDirection: 'column' as const, overflow: 'hidden' }}>
        <DataTable
          title="${esc(z.label)}"
          columns={${columnsJson.trim()}}
          ${specJson ? `querySpec={${specJson.trim()}}` : ''}
          ${filterParamsStr !== 'undefined' ? `filters=${filterParamsStr}` : ''}
          ${varName ? `rows={${varName}}` : 'rows={[]}'}
          isSampleData={${!z.querySpec}}
        />
      </div>`;
  }

  const base = `position: 'absolute', left: '${z.left}', top: '${z.top}', width: '${z.width}', height: '${z.height}', boxSizing: 'border-box', padding: 10, overflow: 'hidden', display: 'flex', flexDirection: 'column' as const, alignItems: 'center', justifyContent: 'center', gap: 4`;

  if (z.kind === 'paramctrl') {
    const ctrlType = z.controlMode ?? 'dropdown';
    const icon = ctrlType === 'slider' ? '⟼' : '▾';
    const paramLine = z.paramRef
      ? `\n      <span style={{ fontSize: 8, color: '#92400e', opacity: 0.7, textAlign: 'center' }}>${esc(z.paramRef)}</span>`
      : '';
    return `      {/* Parameter Control: ${esc(z.label)} */}
      <div style={{ ${base}, background: '#fffbeb', border: '2px dashed #f59e0b', borderRadius: 6 }}>
        <span style={{ fontSize: 8, fontWeight: 600, letterSpacing: '0.06em', textTransform: 'uppercase' as const, color: '#b45309' }}>Parameter Control</span>
        <span style={{ fontSize: 8, fontWeight: 600, color: '#78350f', textAlign: 'center' }}>${esc(z.label)}</span>${paramLine}
        <span style={{ fontSize: 8, background: '#fef3c7', color: '#92400e', padding: '1px 6px', borderRadius: 4 }}>${icon} ${ctrlType}</span>
      </div>`;
  }

  if (z.kind === 'text') {
    return `      {/* Text Zone */}
      <div style={{ ${base}, background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 4 }}>
        <span style={{ fontSize: 8, color: '#94a3b8', textTransform: 'uppercase' as const, letterSpacing: '0.04em' }}>Title / Text</span>
      </div>`;
  }

  if (z.isTextButton) {
    return `      {/* Text / Button: ${esc(z.internalName)} */}
      <div style={{ ${base}, background: '#f0fdf4', border: '1px dashed #86efac', borderRadius: 4 }}>
        <span style={{ fontSize: 8, fontWeight: 600, color: '#16a34a', textTransform: 'uppercase' as const, letterSpacing: '0.05em' }}>Text / Button</span>
        <span style={{ fontSize: 8, color: '#15803d' }}>${esc(z.internalName)}</span>
      </div>`;
  }

  if (z.isControl) {
    const activeFields = z.filterFields.filter((f) => filterMeta.has(f));
    // Render live filter dropdowns when we have metadata; fall back to placeholder otherwise
    if (activeFields.length > 0) {
      const dropdowns = activeFields.map((f) => {
        const meta = filterMeta.get(f)!;
        return `        <FilterDropdown label="${esc(f)}" source="${meta.source}" dim={${JSON.stringify(meta.dim)}} value={filterState["${esc(f)}"] ?? null} onChange={(v) => setFilter("${esc(f)}", v)} />`;
      }).join('\n');
      return `      {/* Quick Filter: ${activeFields.map(esc).join(', ')} */}
      <div style={{ position: 'absolute', left: '${z.left}', top: '${z.top}', width: '${z.width}', height: '${z.height}', boxSizing: 'border-box', padding: 6, display: 'flex', flexDirection: 'row' as const, flexWrap: 'wrap' as const, alignItems: 'flex-start', gap: 6, overflow: 'auto', background: 'white', border: '1px solid #e2e8f0', borderRadius: 6 }}>
${dropdowns}
      </div>`;
    }
    const filterLabel = z.filterFields.length > 0 ? z.filterFields : [z.label];
    const ctrlType = z.filterFields.some(f => /status|type/i.test(f)) ? 'single-value list' : 'dropdown';
    const fieldLines = filterLabel.map(f =>
      `\n        <span style={{ fontSize: 8, fontWeight: 500, color: '#78350f', textAlign: 'center' }}>${esc(f)}</span>`
    ).join('');
    return `      {/* Quick Filter: ${filterLabel.join(', ')} */}
      <div style={{ ${base}, background: '#fffbeb', border: '2px dashed #f59e0b', borderRadius: 6 }}>
        <span style={{ fontSize: 8, fontWeight: 600, letterSpacing: '0.06em', textTransform: 'uppercase' as const, color: '#b45309' }}>Quick Filter</span>${fieldLines}
        <span style={{ fontSize: 8, background: '#fef3c7', color: '#92400e', padding: '1px 6px', borderRadius: 4 }}>type: ${ctrlType}</span>
      </div>`;
  }

  // Regular worksheet zone
  const { bg, border, badgeBg, badgeColor } = zoneColors(z.markType);
  const rowLine = z.rowFields.length > 0
    ? `\n        <span style={{ fontSize: 8, color: '#64748b' }}><span style={{ color: '#94a3b8', fontWeight: 500 }}>ROWS </span>${esc(z.rowFields.join(' · '))}</span>`
    : '';
  const colLine = z.hasMeasureNames || z.colFields.length > 0
    ? `\n        <span style={{ fontSize: 8, color: '#64748b' }}><span style={{ color: '#94a3b8', fontWeight: 500 }}>COLS </span>${esc(z.hasMeasureNames ? 'Measure Names' + (z.colFields.length ? ' · ' + z.colFields.join(' · ') : '') : z.colFields.join(' · '))}</span>`
    : '';

  return `      {/* ${esc(z.label)} */}
      <div style={{ ${base}, background: '${bg}', border: '2px solid ${border}', borderRadius: 6 }}>
        <span style={{ fontSize: 13, fontWeight: 600, color: '#1e293b', textAlign: 'center', lineHeight: 1.3 }}>${esc(z.label)}</span>
        <span style={{ fontSize: 8, fontWeight: 500, letterSpacing: '0.05em', textTransform: 'uppercase' as const, background: '${badgeBg}', color: '${badgeColor}', padding: '2px 8px', borderRadius: 999 }}>${markTypeLabel(z.markType)}</span>${rowLine}${colLine}
      </div>`;
}

// ─── FilterDropdown component template ───────────────────────────────────────

function filterDropdownComponent(): string {
  return `
function FilterDropdown({ label, source, dim, value, onChange }: {
  label: string;
  source: string;
  dim: { alias: string; expr: string };
  value: string | null;
  onChange: (v: string | null) => void;
}) {
  const [options, setOptions] = useState<string[]>([]);
  useEffect(() => {
    fetch('/api/query', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ source, dimensions: [dim], measures: [], pageSize: 500 }),
    })
      .then((r) => (r.ok ? r.json() : []))
      .then((rows: Record<string, unknown>[]) => {
        setOptions(
          rows.map((r) => String(r[dim.alias])).filter((v) => v && v !== 'null').sort()
        );
      })
      .catch(() => {});
  }, [source, dim.alias]);
  return (
    <div style={{ flex: '1 1 100px', minWidth: 90 }}>
      <div style={{ fontSize: 8, fontWeight: 600, color: '#6b7280', textTransform: 'uppercase' as const, letterSpacing: '0.05em', marginBottom: 2 }}>{label}</div>
      <select
        value={value ?? ''}
        onChange={(e) => onChange(e.target.value || null)}
        style={{ fontSize: 10, padding: '3px 6px', border: '1px solid #e2e8f0', borderRadius: 4, width: '100%', background: 'white', color: '#374151', cursor: 'pointer' }}
      >
        <option value="">All</option>
        {options.map((o) => <option key={o} value={o}>{o}</option>)}
      </select>
    </div>
  );
}

`;
}

// ─── Local helpers ────────────────────────────────────────────────────────────

/**
 * Clean a raw Tableau field ref from a filterctrl zone's param attribute into a
 * human-readable field name. The workbook parser stores field captions; this
 * extracts the internal calc/column name so the generator can resolve it via
 * filterMeta to the right display name.
 *
 * "[excel.xxx].[none:Calculation_3160325163145427:nk]" → "Calculation_3160325163145427"
 * "[excel.xxx].[none:Company:nk]"                     → "Company"
 */
function cleanFilterRef(raw: string): string {
  const afterDot = raw.includes('].[') ? raw.split('].[').pop()! : raw;
  const stripped = afterDot.replace(/^\[/, '').replace(/\]$/, '').trim();
  const noneMatch = stripped.match(/^:?none:(.+?)(?::nk)?$/i);
  if (noneMatch) return noneMatch[1].trim();
  return stripped;
}

function esc(s: string): string {
  // Also escape */ to prevent it from prematurely closing JSX block comments
  return s.replace(/'/g, "\\'").replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\*\//g, '*\\/');
}

function fallbackComponent(name: string): string {
  return `export function Dashboard() {
  return <div style={{ padding: 24, color: '#ef4444' }}>No layout zones found for ${esc(name)}.</div>;
}
`;
}
