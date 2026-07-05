// Shared type contracts for the GridMD reference implementation.
//
// The STABLE model — scalars, cells, sheets, the workbook model, diagnostics
// and the block tree — is fully and strictly typed here (this is the public
// API surface). YAML-derived directive metadata is intentionally open-ended:
// SPEC permits `x-` extension keys and per-directive shapes, so it is carried
// as the permissive `Meta` alias and used through it. See the README section
// "Deliberate divergences" for the rationale (this is the one documented
// divergence from CODING_PRACTICES §1 "no any").

/**
 * Open-ended, YAML-derived data (frontmatter, directive metadata, inline
 * props and `@`-directive bodies). Structurally unconstrained by design; it
 * enters as parsed YAML and is validated by the lint pass, not by the type
 * system. Deliberate, documented divergence from CODING_PRACTICES §1.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type Meta = any;

// ---- diagnostics ----
export interface Diagnostic {
  line: number;
  msg: string;
}

// ---- cell scalars (scalar.ts) ----
export type ScalarKind =
  | 'blank'
  | 'text'
  | 'formula'
  | 'number'
  | 'boolean'
  | 'date'
  | 'time'
  | 'error'
  | 'invalid';

/**
 * One parsed cell scalar. A wide struct that mirrors the runtime object shape
 * exactly (every field optional) so the parser/serializer round-trip without
 * narrowing churn; `kind` is the discriminant every consumer switches on.
 */
export interface Scalar {
  kind: ScalarKind;
  value?: string | number | boolean;
  formula?: string;
  cse?: boolean;
  cached?: Scalar | null;
  problem?: string;
  forced?: boolean;
  quoted?: boolean;
}

// ---- references (refs.ts) ----
export type TargetKind = 'cell' | 'range' | 'cols' | 'rows';

export interface Target {
  kind: TargetKind;
  sheet: string | null;
  c1: number;
  r1: number;
  c2: number;
  r2: number;
  line?: number;
}

export interface CellPos {
  col: number;
  row: number;
}

// ---- minimal XML tree (xml.ts) ----
export interface XmlNode {
  name: string;
  attrs: Record<string, string | undefined>;
  children: XmlNode[];
  text: string;
}

// ---- parse tree (parser.ts) ----
export interface InfoArgs {
  positional: string[];
  flags: Record<string, string | undefined>;
  anchor: string | null;
  size: { w: number; h: number } | null;
}

export interface Row {
  cells: string[];
  line: number;
}

export interface FenceBlock {
  type: 'fence';
  kind: string;
  args: InfoArgs;
  body: string[];
  line: number;
  meta?: Meta;
  rows?: Row[];
  code?: string;
  payload?: string;
}

export interface AtBlock {
  type: 'at';
  targetText: string;
  line: number;
  scalarText: string | null;
  props: Meta | null;
  body: Meta | null;
}

export type Block = FenceBlock | AtBlock;

export interface SheetBlock {
  name: string;
  line: number;
  blocks: Block[];
}

export interface ParseStats {
  defs: number;
  blocks: number;
}

export interface ParsedDocument {
  frontmatter: Meta;
  workbookBlocks: Block[];
  sheets: SheetBlock[];
  errors: Diagnostic[];
  warnings: Diagnostic[];
  mode: string;
  stats?: ParseStats;
}

export interface LintResult {
  doc: ParsedDocument;
  errors: Diagnostic[];
  warnings: Diagnostic[];
  sheets: number;
  cells: number;
  blocks: number;
}

// ---- workbook model (xlsx/model.ts) ----
export interface CellContent {
  formula?: string;
  cse?: boolean;
  cached?: Scalar | null;
  scalar?: Scalar;
  rich?: Meta;
  arrayRef?: string;
  entityFields?: Meta;
  spillCache?: boolean;
}

export interface Cell {
  col: number;
  row: number;
  content: CellContent | null;
  patches: Meta[];
}

export interface TableModel {
  name: string;
  anchor: CellPos;
  columns: string[];
  headerRow: boolean;
  bodyRows: number;
  total: Meta;
  style?: Meta;
  banded: Meta;
  filter: Meta;
  sort: Meta[];
  line: number;
  sheetName?: string;
}

export interface CfBlock {
  sqref: string;
  rules: Meta;
  line: number;
}

export interface ValidationBlock {
  sqref: string;
  meta: Meta;
}

export interface FilterBlock {
  sqref: string;
  meta: Meta;
  line: number;
}

export interface NoteModel {
  col: number;
  row: number;
  text: string;
}

export interface HyperlinkModel {
  col: number;
  row: number;
  target: string;
  tip?: Meta;
}

export interface ChartModel {
  type: string | undefined;
  title: string | null;
  anchor: string | null;
  size: { w: number; h: number } | null;
  meta: Meta;
  line: number;
}

export interface SparklineModel {
  sqref: string;
  meta: Meta;
  line: number;
}

export interface PivotModel {
  name: string;
  anchor: string | null;
  meta: Meta;
  line: number;
}

export interface SlicerModel {
  anchor: string | null;
  size: { w: number; h: number } | null;
  meta: Meta;
  kind: string;
  line: number;
  _name?: string;
  _timeline?: boolean;
}

export interface ImageModel {
  anchor: string | null;
  size: { w: number; h: number } | null;
  src: string;
  alt: Meta;
  line: number;
}

export interface ShapeModel {
  preset: string;
  anchor: string | null;
  size: { w: number; h: number } | null;
  meta: Meta;
  line: number;
}

export interface ThreadModel {
  ref: string;
  comments: Meta;
  line: number;
}

export interface ScenarioModel {
  name: string;
  meta: Meta;
  line: number;
}

export interface OutlineModel {
  rows: Meta[];
  cols: Meta[];
}

export interface Sheet {
  name: string;
  meta: Meta;
  kind: 'worksheet' | 'chart';
  cells: Map<string, Cell>;
  merges: Target[];
  tables: TableModel[];
  cf: CfBlock[];
  validations: ValidationBlock[];
  notes: NoteModel[];
  hyperlinks: HyperlinkModel[];
  outline: OutlineModel;
  page: Meta;
  charts: ChartModel[];
  sparklines: SparklineModel[];
  pivots: PivotModel[];
  slicers: SlicerModel[];
  images: ImageModel[];
  shapes: ShapeModel[];
  threads: ThreadModel[];
  scenarios: ScenarioModel[];
  filters: FilterBlock[];
}

export type ReportAction = 'native' | 'carried' | 'partial' | 'not-emitted';

export interface ReportEntry {
  line?: number;
  feature: string;
  action: ReportAction;
  note?: string;
}

export interface RawPart {
  part: string;
  payload: string | undefined;
  encoding: string | undefined;
  line: number;
}

export interface CarryEntry {
  kind: string;
  line: number;
  feature?: string;
  args?: InfoArgs | null;
  meta?: Meta;
  code?: string | null;
  body?: Meta;
}

export interface TableIndexEntry extends TableModel {
  sheetName: string;
}

export interface WorkbookModel {
  fm: Meta;
  themeColors: Record<string, string>;
  sheets: Sheet[];
  report: ReportEntry[];
  rawParts: RawPart[];
  carry: CarryEntry[];
  tableIndex: Map<string, TableIndexEntry>;
  baseDir: string;
}
