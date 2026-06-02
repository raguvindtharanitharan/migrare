import type { TableauWorkbook, MarkType } from '../../parsers/model.js';
import {
  collectLeafZones, buildFiltersByWorksheet, pickDashboard,
  cleanRef, markTypeLabel, zoneColors, pct,
  type LeafKind,
} from './zone-helpers.js';
import { resolveTableSpec, type TableColumn } from './table-generator.js';

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
  dataFile?: string;   // relative import path for this zone's data
  left: string; top: string; width: string; height: string;
}

export interface DashboardGeneratorResult {
  dashboardTsx: string;
  dataFiles: Array<{ relativePath: string; content: string }>;
}

export function generateDashboardComponent(workbook: TableauWorkbook, slug: string): DashboardGeneratorResult {
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
    const isControl = z.kind !== 'worksheet' || (area < maxArea && hasEncodings) || isTextButton;
    const isTable = !isControl && markType === 'automatic';

    const rowFields = encoding?.rows
      .map((r) => r.caption ?? cleanRef(r.field))
      .filter((n) => n && !/Latitude|Longitude/i.test(n)) ?? [];
    const colFields = encoding?.columns
      .map((c) => c.caption ?? cleanRef(c.field))
      .filter((n) => n && n !== ':Measure Names') ?? [];
    const hasMeasureNames = encoding?.columns.some((c) => c.field.includes(':Measure Names')) ?? false;
    const filterFields = filtersByWorksheet.get(z.worksheetOrName) ?? [];

    // Resolve table spec and generate data file for text tables
    let tableColumns: TableColumn[] = [];
    let dataFile: string | undefined;
    if (isTable) {
      const spec = resolveTableSpec(workbook, z.worksheetOrName);
      tableColumns = spec.columns;
      if (spec.columns.length > 0) {
        const wsSlug = z.worksheetOrName.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
        const relPath = `src/workbooks/${slug}/data/${wsSlug}.json`;
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

  const zonesJsx = zones.map((z) => renderZone(z)).join('\n');

  // Build DataTable import only when at least one table zone exists.
  // Deduplicate by varName — two worksheets producing the same identifier
  // would cause a duplicate import binding compile error.
  const tableZones = zones.filter((z) => z.isTable && z.dataFile);
  const seenVarNames = new Set<string>();
  const dataImports = tableZones
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
    hasTable ? dataImports : '',
    '',
    `export function Dashboard() {`,
    `  return (`,
    `    <div style={{ position: 'relative', width: '100%', aspectRatio: ${aspectRatio}, fontFamily: "'Poppins', sans-serif", background: 'white', borderRadius: 8, overflow: 'hidden', boxShadow: '0 1px 4px rgba(0,0,0,0.06)' }}>`,
    zonesJsx,
    `    </div>`,
    `  );`,
    `}`,
    '',
  ].filter((l) => l !== null).join('\n');

  return { dashboardTsx, dataFiles };
}

// ─── Zone renderer ────────────────────────────────────────────────────────────

function dataVarName(worksheetName: string): string {
  return worksheetName
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_|_$/g, '') + 'Data';
}

function renderZone(z: ZoneViewModel): string {
  // Table zone — render real DataTable component
  if (z.isTable && z.dataFile && z.tableColumns.length > 0) {
    const varName = dataVarName(z.internalName);
    const columnsJson = JSON.stringify(z.tableColumns.map((c) => ({
      key: c.key, label: c.label, numeric: c.numeric, currency: c.currency, percent: c.percent,
    })), null, 6).replace(/^/gm, '      ');
    return `      {/* ${esc(z.label)} */}
      <div style={{ position: 'absolute', left: '${z.left}', top: '${z.top}', width: '${z.width}', height: '${z.height}', boxSizing: 'border-box', padding: 8, display: 'flex', flexDirection: 'column' as const, overflow: 'hidden' }}>
        <DataTable
          title="${esc(z.label)}"
          columns={${columnsJson.trim()}}
          rows={${varName}}
          isSampleData={true}
        />
      </div>`;
  }

  const base = `position: 'absolute', left: '${z.left}', top: '${z.top}', width: '${z.width}', height: '${z.height}', boxSizing: 'border-box', padding: 10, overflow: 'hidden', display: 'flex', flexDirection: 'column' as const, alignItems: 'center', justifyContent: 'center', gap: 4`;

  if (z.kind === 'paramctrl') {
    const ctrlType = z.controlMode ?? 'dropdown';
    const icon = ctrlType === 'slider' ? '⟼' : '▾';
    const paramLine = z.paramRef
      ? `\n      <span style={{ fontSize: 9, color: '#92400e', opacity: 0.7, textAlign: 'center' }}>${esc(z.paramRef)}</span>`
      : '';
    return `      {/* Parameter Control: ${esc(z.label)} */}
      <div style={{ ${base}, background: '#fffbeb', border: '2px dashed #f59e0b', borderRadius: 6 }}>
        <span style={{ fontSize: 9, fontWeight: 600, letterSpacing: '0.06em', textTransform: 'uppercase' as const, color: '#b45309' }}>Parameter Control</span>
        <span style={{ fontSize: 12, fontWeight: 600, color: '#78350f', textAlign: 'center' }}>${esc(z.label)}</span>${paramLine}
        <span style={{ fontSize: 9, background: '#fef3c7', color: '#92400e', padding: '1px 6px', borderRadius: 4 }}>${icon} ${ctrlType}</span>
      </div>`;
  }

  if (z.kind === 'text') {
    return `      {/* Text Zone */}
      <div style={{ ${base}, background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 4 }}>
        <span style={{ fontSize: 10, color: '#94a3b8', textTransform: 'uppercase' as const, letterSpacing: '0.04em' }}>Title / Text</span>
      </div>`;
  }

  if (z.isTextButton) {
    return `      {/* Text / Button: ${esc(z.internalName)} */}
      <div style={{ ${base}, background: '#f0fdf4', border: '1px dashed #86efac', borderRadius: 4 }}>
        <span style={{ fontSize: 9, fontWeight: 600, color: '#16a34a', textTransform: 'uppercase' as const, letterSpacing: '0.05em' }}>Text / Button</span>
        <span style={{ fontSize: 10, color: '#15803d' }}>${esc(z.internalName)}</span>
      </div>`;
  }

  if (z.isControl) {
    const filterLabel = z.filterFields.length > 0 ? z.filterFields : [z.label];
    const ctrlType = z.filterFields.some(f => /status|type/i.test(f)) ? 'single-value list' : 'dropdown';
    const fieldLines = filterLabel.map(f =>
      `\n        <span style={{ fontSize: 11, fontWeight: 500, color: '#78350f', textAlign: 'center' }}>${esc(f)}</span>`
    ).join('');
    return `      {/* Quick Filter: ${filterLabel.join(', ')} */}
      <div style={{ ${base}, background: '#fffbeb', border: '2px dashed #f59e0b', borderRadius: 6 }}>
        <span style={{ fontSize: 9, fontWeight: 600, letterSpacing: '0.06em', textTransform: 'uppercase' as const, color: '#b45309' }}>Quick Filter</span>${fieldLines}
        <span style={{ fontSize: 9, background: '#fef3c7', color: '#92400e', padding: '1px 6px', borderRadius: 4 }}>type: ${ctrlType}</span>
      </div>`;
  }

  // Regular worksheet zone
  const { bg, border, badgeBg, badgeColor } = zoneColors(z.markType);
  const rowLine = z.rowFields.length > 0
    ? `\n        <span style={{ fontSize: 10, color: '#64748b' }}><span style={{ color: '#94a3b8', fontWeight: 500 }}>ROWS </span>${esc(z.rowFields.join(' · '))}</span>`
    : '';
  const colLine = z.hasMeasureNames || z.colFields.length > 0
    ? `\n        <span style={{ fontSize: 10, color: '#64748b' }}><span style={{ color: '#94a3b8', fontWeight: 500 }}>COLS </span>${esc(z.hasMeasureNames ? 'Measure Names' + (z.colFields.length ? ' · ' + z.colFields.join(' · ') : '') : z.colFields.join(' · '))}</span>`
    : '';

  return `      {/* ${esc(z.label)} */}
      <div style={{ ${base}, background: '${bg}', border: '2px solid ${border}', borderRadius: 6 }}>
        <span style={{ fontSize: 14, fontWeight: 600, color: '#1e293b', textAlign: 'center', lineHeight: 1.3 }}>${esc(z.label)}</span>
        <span style={{ fontSize: 10, fontWeight: 500, letterSpacing: '0.05em', textTransform: 'uppercase' as const, background: '${badgeBg}', color: '${badgeColor}', padding: '2px 8px', borderRadius: 999 }}>${markTypeLabel(z.markType)}</span>${rowLine}${colLine}
      </div>`;
}

// ─── Local helpers ────────────────────────────────────────────────────────────

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
