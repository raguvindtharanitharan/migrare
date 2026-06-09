import type {
  Dashboard,
  DashboardSize,
  Zone,
  ZoneType,
} from '../model.js';
import { toArray, extractFormattedTitle } from './_helpers.js';

/**
 * Extract dashboards and their nested zone tree.
 *
 * Tableau dashboards are built from a hierarchical layout of zones.
 * Each zone is either a layout container (`layout-basic`,
 * `layout-flow`) or a content zone (worksheet, parameter control,
 * text). We preserve the full tree so the generator can later
 * reconstruct it as a responsive React grid (v0.3 work).
 */
export function extractDashboards(workbook: any): Dashboard[] {
  const dashboardsRaw = workbook?.dashboards?.dashboard;
  if (!dashboardsRaw) return [];

  return toArray(dashboardsRaw).map((db: any) => mapDashboard(db));
}

function mapDashboard(db: any): Dashboard {
  const name: string = db['@_name'] ?? 'Unnamed Dashboard';
  const title = extractFormattedTitle(db);
  const size = extractSize(db);

  const zones = db?.zones?.zone ? toArray(db.zones.zone).map((z: any) => mapZone(z)) : [];

  return {
    name,
    title,
    size,
    layoutType: 'tiled',           // v0.2+: detect floating layouts
    zones,
    worksheetsUsed: collectWorksheetsUsed(zones),
    floatingObjects: [],           // v0.2+
  };
}

function extractSize(db: any): DashboardSize {
  const size = db?.size ?? {};
  return {
    width: parseInt(size['@_maxwidth'] ?? size['@_minwidth'] ?? '0', 10),
    height: parseInt(size['@_maxheight'] ?? size['@_minheight'] ?? '0', 10),
  };
}

function mapZone(zone: any): Zone {
  const typeV2: string | undefined = zone['@_type-v2'] ?? zone['@_type'];
  const name: string | undefined = zone['@_name'];
  const isLayoutZone = typeV2?.startsWith('layout') ?? false;
  const isFilterCtrl = typeV2 === 'filter';
  const isWorksheetZone = name !== undefined && !isLayoutZone && !isFilterCtrl && typeV2 !== 'paramctrl' && typeV2 !== 'text' && typeV2 !== 'title';

  const type: ZoneType = isFilterCtrl ? 'filterctrl' : (isWorksheetZone ? 'worksheet' : normalizeZoneType(typeV2));

  const children = zone.zone ? toArray(zone.zone).map((z: any) => mapZone(z)) : [];

  // Extract the human-readable label from formatted-text (the visible title shown in the dashboard)
  const displayLabel = extractZoneLabel(zone);

  // Extract paramctrl/filterctrl fields
  const controlMode = normalizeControlMode(zone['@_mode']);
  // Strip the datasource namespace prefix generically: '[DataSource].[Name]' → 'Name'
  // Works for any language/locale — does not hardcode 'Parameters'.
  const paramRef = zone['@_param']
    ? String(zone['@_param'])
        .replace(/^\[[^\]]*\]\.\[/, '')  // strip [AnyDataSource].[ prefix
        .replace(/\]$/, '')              // strip trailing ]
    : undefined;

  const showTitle = zone['@_show-title'] !== 'false';

  // filterctrl: extract the field reference from the @_param attribute
  const filterField = isFilterCtrl && zone['@_param']
    ? String(zone['@_param'])
    : undefined;

  return {
    id: String(zone['@_id'] ?? ''),
    name,
    worksheet: isWorksheetZone ? name : undefined,
    type,
    filterField,
    x: parseFloat(zone['@_x'] ?? '0'),
    y: parseFloat(zone['@_y'] ?? '0'),
    w: parseFloat(zone['@_w'] ?? '0'),
    h: parseFloat(zone['@_h'] ?? '0'),
    isFixed: zone['@_is-fixed'] === 'true',
    children,
    displayLabel: displayLabel || undefined,
    controlMode,
    paramRef,
    showTitle,
  };
}

function extractZoneLabel(zone: any): string {
  const ft = zone['formatted-text'];
  if (!ft) return '';
  const runs = toArray(ft.run ?? []);
  return runs
    .map((r: any) => (typeof r === 'string' ? r : (r['#text'] ?? '')))
    // Tableau uses U+00C6 (Æ) and &#10; as internal line/paragraph separators —
    // replace them with spaces rather than dropping the whole run.
    .map((t: string) => t.replace(/Æ|&#10;|&#13;/g, ' '))
    .map((t: string) => t.trim())
    .filter((t: string) => t.length > 0)
    .join(' ')
    .trim();
}

function normalizeControlMode(mode: string | undefined): string | undefined {
  if (!mode) return undefined;
  const map: Record<string, string> = {
    compact:          'dropdown',
    dropdown:         'dropdown',
    slider:           'slider',
    radio:            'radio',
    'multiple-values': 'multi-select',
    'multiple_values': 'multi-select',  // Tableau uses both hyphen and underscore variants
    'single-value':   'dropdown',
    'single_value':   'dropdown',
    checkbox:         'checkbox',
    list:             'list',
  };
  return map[mode] ?? mode;
}

function normalizeZoneType(raw: string | undefined): ZoneType {
  if (!raw) return 'unknown';
  if (raw.startsWith('layout-basic')) return 'layout-basic';
  if (raw.startsWith('layout-flow')) return 'layout-flow';
  if (raw === 'paramctrl') return 'paramctrl';
  if (raw === 'text') return 'text';
  if (raw === 'empty') return 'empty';
  return 'unknown';
}

function collectWorksheetsUsed(zones: Zone[]): string[] {
  const result = new Set<string>();
  function walk(list: Zone[]): void {
    for (const z of list) {
      if (z.worksheet) result.add(z.worksheet);
      if (z.children.length > 0) walk(z.children);
    }
  }
  walk(zones);
  return Array.from(result);
}
