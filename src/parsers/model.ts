/**
 * Canonical typed model representing a parsed Tableau workbook.
 *
 * Source of truth: `docs/schema/workbook-model.md`
 *
 * This model is the contract between parsers (which produce it) and
 * generators (which consume it). The markdown generator (Stage 3)
 * is responsible for converting the camelCase TypeScript field
 * names into the snake_case keys defined in the YAML schema.
 *
 * Conventions:
 * - `undefined` is never emitted to YAML. Optional single-value
 *   fields serialize as `null`; optional collections serialize as
 *   `[]`.
 * - String literal types are used where the schema documents a
 *   closed set of values (e.g. mark types, zone types).
 * - Fields populated only in v0.2+ are explicitly marked.
 */

// ============================================================================
// Top-level workbook
// ============================================================================

export interface TableauWorkbook {
  metadata: WorkbookMetadata;
  executiveSummary?: string;
  dataSources: DataSource[];
  fields: Field[];
  worksheets: Worksheet[];
  visualEncodings: VisualEncoding[];
  dashboards: Dashboard[];
  filters: Filter[];
  parameters: Parameter[];
  actions: Action[];
  calculations: Calculation[];
  notes: string[];
}

// ============================================================================
// Workbook metadata
// ============================================================================

export type WorkbookDomain = 'Finance' | 'Operations' | 'Other';
export type WorkbookComplexity = 'Simple' | 'Medium' | 'Complex';

export interface WorkbookMetadata {
  name: string;
  originalFilename: string;
  tableauVersion?: string;
  sourceUrl?: string;
  domain?: WorkbookDomain;
  complexity?: WorkbookComplexity;
  exportedAt: string; // ISO 8601 timestamp
}

// ============================================================================
// Data sources & fields
// ============================================================================

export type DataSourceType = 'live' | 'extract' | 'parameters' | 'other';

export interface DataSource {
  name: string;
  caption?: string;
  type: DataSourceType;
  connectionType?: string; // e.g. "Excel", "PostgreSQL"
  connectionDetails?: string;
  tables: string[]; // empty in v0.1
  fields: Field[];
}

export type FieldRole = 'dimension' | 'measure';
export type FieldDataType =
  | 'string'
  | 'integer'
  | 'real'
  | 'date'
  | 'datetime'
  | 'boolean';

export interface Field {
  name: string; // internal Tableau name (e.g. "[Calculation_123]")
  caption?: string; // user-facing name
  role?: FieldRole;
  dataType?: FieldDataType;
  folder?: string;
  calculation?: FieldCalculation;
  defaultAggregation?: string;
  description?: string;
}

export interface FieldCalculation {
  isCalculated: boolean;
  formula?: string;
  simplifiedFormula?: string; // v0.2+: human/AI-friendly version
  dependsOn?: string[]; // v0.2+: parsed dependency list
}

// ============================================================================
// Worksheets & visual encodings
// ============================================================================

export type MarkType =
  | 'bar'
  | 'line'
  | 'area'
  | 'pie'
  | 'circle'
  | 'square'
  | 'shape'
  | 'text'
  | 'map'
  | 'gantt'
  | 'heatmap'
  | 'automatic'
  | 'unsupported';

export interface Worksheet {
  name: string;
  title?: string;
  sheetType: 'worksheet';
  markTypes: MarkType[]; // a worksheet may have multiple panes
  description?: string;
}

export interface FieldRef {
  field: string; // raw identifier as it appears in the XML
  datasource?: string; // owning datasource ID when the pill was qualified, e.g. `[ds].[field]`
  caption?: string; // resolved user-facing name
  role?: FieldRole;
  dataType?: FieldDataType;
  isCalculated?: boolean;
  aggregation?: string; // SUM, AVG, MIN, etc.
  datePart?: string; // year, quarter, month, etc.
}

export interface VisualEncoding {
  worksheet: string; // matches Worksheet.name
  title?: string;
  markType: MarkType;
  effectiveMarkType: MarkType; // resolved from shelf inspection when markType is 'automatic'
  isDualAxis: boolean;
  rows: FieldRef[];
  columns: FieldRef[];
  // When Measure Names is on the columns shelf, Tableau applies a categorical
  // filter to select which specific measures to display. This field captures
  // those resolved captions in display order. Empty = all measures shown.
  selectedMeasures: string[];
  color: FieldRef | FieldRef[] | null;
  size: FieldRef | FieldRef[] | null;
  shape: FieldRef | FieldRef[] | null;
  detail: FieldRef[];
  tooltip: FieldRef[];
  label: FieldRef[];
  sorting: SortConfig[]; // empty in v0.1
  referenceLines: ReferenceLineConfig[]; // empty in v0.1
  tableCalculations: TableCalculationConfig[]; // empty in v0.1
}

// Stub types for v0.2+ schema fields. Defined here so the model
// shape is stable; populated by future parser work.

export interface SortConfig {
  field?: string;
  direction?: 'asc' | 'desc';
}

export interface ReferenceLineConfig {
  scope?: string;
  value?: string | number;
}

export interface TableCalculationConfig {
  name?: string;
  expression?: string;
}

// ============================================================================
// Dashboards
// ============================================================================

export type LayoutType = 'tiled' | 'floating';

export type ZoneType =
  | 'layout-basic'
  | 'layout-flow'
  | 'paramctrl'
  | 'filterctrl'   // explicit quick-filter control: type="filter" in the XML
  | 'text'
  | 'empty'
  | 'worksheet'
  | 'unknown';

export interface DashboardSize {
  width: number;
  height: number;
}

export interface Zone {
  id: string;
  name?: string;
  worksheet?: string; // worksheet name if this zone hosts a sheet
  type: ZoneType;
  x: number;
  y: number;
  w: number;
  h: number;
  isFixed: boolean;
  children: Zone[];
  // Display metadata — populated for paramctrl and titled zones
  displayLabel?: string;   // from formatted-text inside the zone (the visible title)
  controlMode?: string;    // paramctrl/filterctrl mode: 'slider' | 'compact' | 'radio' | 'dropdown'
  paramRef?: string;       // paramctrl parameter reference e.g. '[Parameters].[Revenue View]'
  filterField?: string;    // filterctrl: raw param ref of the field being filtered
  showTitle?: boolean;     // whether the zone title is visible
}

export interface Dashboard {
  name: string;
  title?: string;
  size: DashboardSize;
  layoutType: LayoutType;
  zones: Zone[];
  worksheetsUsed: string[];
  floatingObjects: string[]; // empty in v0.1
}

// ============================================================================
// Filters, parameters, actions
// ============================================================================

export type FilterScope = 'worksheet' | 'dashboard';

export interface Filter {
  name: string;
  field?: string;
  type: FieldRole; // reuses 'dimension' | 'measure'
  scope: FilterScope;
  appliedTo: string[]; // worksheet names
  /** Selected/included values for categorical filters (empty = no restriction recorded) */
  values?: string[];
  /** True when this filter is explicitly shown as an interactive UI control in the dashboard */
  visible?: boolean;
}

export interface Parameter {
  name: string;
  dataType?: FieldDataType;
  currentValue?: string;
  allowableValues: string[];
  controlType?: string; // 'slider' | 'dropdown' | 'text-input' — from param-domain-type
}

export type ActionType = 'filter' | 'highlight' | 'url';

export interface Action {
  name: string;
  type: ActionType;
  source?: string;
  target?: string;
}

// ============================================================================
// Calculations
// ============================================================================

export type CalculationComplexity = 'simple' | 'medium' | 'complex';

export interface Calculation {
  name: string;
  formula: string;
  simplifiedFormula?: string; // v0.2+: plain-English restatement from AI enrichment
  dependsOn: string[]; // empty in v0.1; populated when formula parsing lands in v0.2+
  complexity?: CalculationComplexity;
  notes?: string;
}
