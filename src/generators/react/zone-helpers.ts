import type { TableauWorkbook, Zone, MarkType } from '../../parsers/model.js';

// ─── Shared types ─────────────────────────────────────────────────────────────

export type LeafKind = 'worksheet' | 'paramctrl' | 'filterctrl' | 'text';

export interface LeafZone {
  worksheetOrName: string;
  kind: LeafKind;
  displayLabel?: string;
  controlMode?: string;
  paramRef?: string;
  showTitle?: boolean;
  x: number; y: number; w: number; h: number;
}

export interface ZoneColors {
  bg: string; border: string; badgeBg: string; badgeColor: string;
}

// ─── Leaf zone collection ─────────────────────────────────────────────────────

export function collectLeafZones(zones: Zone[]): LeafZone[] {
  const leaves: LeafZone[] = [];
  function walk(z: Zone) {
    if (z.children.length === 0) {
      if (z.worksheet) {
        leaves.push({
          worksheetOrName: z.worksheet, kind: 'worksheet',
          displayLabel: z.displayLabel, showTitle: z.showTitle,
          x: z.x, y: z.y, w: z.w, h: z.h,
        });
      } else if (z.type === 'filterctrl') {
        leaves.push({
          worksheetOrName: z.filterField ?? z.name ?? '', kind: 'filterctrl',
          displayLabel: z.displayLabel, controlMode: z.controlMode,
          x: z.x, y: z.y, w: z.w, h: z.h,
        });
      } else if (z.type === 'paramctrl') {
        leaves.push({
          worksheetOrName: z.name ?? '', kind: 'paramctrl',
          displayLabel: z.displayLabel, controlMode: z.controlMode, paramRef: z.paramRef,
          x: z.x, y: z.y, w: z.w, h: z.h,
        });
      } else if (z.type === 'text') {
        leaves.push({
          worksheetOrName: z.name ?? 'Title', kind: 'text',
          displayLabel: z.displayLabel,
          x: z.x, y: z.y, w: z.w, h: z.h,
        });
      }
      return;
    }
    for (const c of z.children) walk(c);
  }
  for (const z of zones) walk(z);
  return leaves;
}

// ─── Filter field map ─────────────────────────────────────────────────────────

export function buildFiltersByWorksheet(workbook: TableauWorkbook): Map<string, string[]> {
  const map = new Map<string, string[]>();
  for (const f of workbook.filters) {
    // Use raw field ref to detect action filters — f.name is now cleaned
    if ((f.field ?? f.name).includes('Action (')) continue;
    const raw = f.name.split('].').pop() ?? '';
    const field = raw
      .replace(/^\[/, '').replace(/\]$/, '')
      .replace(/^none:/, '').replace(/^usr:/, '').replace(/^sum:/, '')
      .replace(/:[a-z]+$/i, '');
    if (!field || /Latitude|Longitude|Measure Names/i.test(field)) continue;
    for (const ws of f.appliedTo) {
      const arr = map.get(ws) ?? [];
      if (!arr.includes(field)) arr.push(field);
      map.set(ws, arr);
    }
  }
  return map;
}

// ─── Dashboard selector ───────────────────────────────────────────────────────

// Prefer exact name match → case-insensitive exact → last dashboard.
// Avoids substring match ambiguity for short workbook names like "AR".
export function pickDashboard(workbook: TableauWorkbook) {
  const { dashboards, metadata } = workbook;
  if (dashboards.length === 0) return undefined;
  return (
    dashboards.find((d) => d.name === metadata.name) ??
    dashboards.find((d) => d.name.toLowerCase() === metadata.name.toLowerCase()) ??
    dashboards[dashboards.length - 1]
  );
}

// ─── Formatting helpers ───────────────────────────────────────────────────────

export function cleanRef(field: string): string {
  const parts = field.split(':');
  const mid = parts.length >= 3 ? parts.slice(1, -1).join(':') : field;
  return mid.replace(/^\[|\]$/g, '').replace(/^_+\s*/, '').replace(/^\./, '');
}

export function markTypeLabel(mark: MarkType): string {
  const labels: Partial<Record<MarkType, string>> = {
    automatic: 'Text Table', text: 'Text Table', bar: 'Bar Chart',
    line: 'Line Chart', area: 'Area Chart', pie: 'Pie Chart',
    map: 'Map', heatmap: 'Heat Map', circle: 'Scatter', unsupported: 'Unsupported',
  };
  return labels[mark] ?? mark;
}

export function zoneColors(mark: MarkType): ZoneColors {
  switch (mark) {
    case 'automatic': case 'text':
      return { bg: '#dbeafe', border: '#3b82f6', badgeBg: '#eff6ff', badgeColor: '#1d4ed8' };
    case 'map':
      return { bg: '#f1f5f9', border: '#94a3b8', badgeBg: '#e2e8f0', badgeColor: '#475569' };
    case 'bar': case 'line': case 'area': case 'pie':
      return { bg: '#dcfce7', border: '#22c55e', badgeBg: '#f0fdf4', badgeColor: '#15803d' };
    default:
      return { bg: '#f8fafc', border: '#cbd5e1', badgeBg: '#f1f5f9', badgeColor: '#64748b' };
  }
}

export function pct(n: number): string {
  return `${Math.round(n * 10000) / 100}%`;
}
