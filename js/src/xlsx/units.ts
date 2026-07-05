// Unit conversions for the XLSX writer (INTEROP.md §2).

const DAY_MS = 86400000;

// ISO date/time (SPEC §6 rule 8) → Excel serial per date-system.
// 1900 system: serial 1 = 1900-01-01; the phantom 1900-02-29 occupies serial
// 60, so real dates before 1900-03-01 sit one lower than the naive diff.
export function isoToSerial(iso: string, dateSystem = 1900): number {
  let datePart: string | null = null;
  let timePart: string | null = null;
  if (/^\d{2}:/.test(iso)) timePart = iso;
  else { const parts = iso.split('T'); datePart = parts[0] ?? null; timePart = parts[1] ?? null; }

  let frac = 0;
  if (timePart) {
    const t = timePart.split(':');
    const hh = Number(t[0] ?? 0), mm = Number(t[1] ?? 0), ss = Number(t[2] ?? 0);
    frac = (hh * 3600 + mm * 60 + (ss || 0)) / 86400;
  }
  if (!datePart) return frac;

  const [y, m, d] = datePart.split('-').map(Number);
  const utc = Date.UTC(y!, m! - 1, d!);
  let days: number;
  if (dateSystem === 1904) {
    days = (utc - Date.UTC(1904, 0, 1)) / DAY_MS;
  } else {
    days = (utc - Date.UTC(1899, 11, 30)) / DAY_MS;
    if (days < 61) days -= 1; // before 1900-03-01 (phantom leap day)
  }
  return days + frac;
}

// px → Excel column width (chars of the max digit width; MDW=7px @ Calibri 11).
export function pxToColWidth(px: number): number {
  return Math.max(0, Math.round(((px - 5) / 7) * 100) / 100);
}

// px → row height in points (96dpi).
export const pxToPt = (px: number): number => Math.round((px * 72 / 96) * 100) / 100;

// cm → inches (page margins).
export const cmToInch = (cm: number): number => Math.round((cm / 2.54) * 10000) / 10000;

// Office default theme palette (Excel 2013+), overridable from frontmatter.
export const DEFAULT_THEME: Record<string, string> = {
  dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6',
  accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000',
  accent5: '5B9BD5', accent6: '70AD47', hlink: '0563C1', folHlink: '954F72',
};

const THEME_REF_RE = /^(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)(?:@(-?\d{1,3}))?$/;

// GridMD color (FORMATTING §3) → ARGB hex, resolving theme slots + tint/shade.
export function resolveColor(value: unknown, themeColors: Record<string, string> = {}): string | null {
  if (typeof value !== 'string' || value === 'auto') return null;
  if (/^#[0-9a-fA-F]{6}$/.test(value)) return `FF${value.slice(1).toUpperCase()}`;
  if (/^#[0-9a-fA-F]{8}$/.test(value)) {
    // GridMD #RRGGBBAA → OOXML AARRGGBB
    return (value.slice(7, 9) + value.slice(1, 7)).toUpperCase();
  }
  const m = THEME_REF_RE.exec(value);
  if (!m) return null;
  const slot = m[1]!;
  const base = themeColors[slot] ?? DEFAULT_THEME[slot]!;
  let rgb = [0, 2, 4].map((i) => parseInt(base.slice(i, i + 2), 16));
  if (m[2] !== undefined) {
    const pct = Number(m[2]) / 100;
    rgb = rgb.map((c) => {
      const out = pct >= 0 ? c + (255 - c) * pct : c * (1 + pct);
      return Math.max(0, Math.min(255, Math.round(out)));
    });
  }
  return `FF${rgb.map((c) => c.toString(16).padStart(2, '0')).join('')}`.toUpperCase();
}

// numfmt aliases (FORMATTING §2).
export const NUMFMT_ALIASES: Record<string, string> = {
  general: 'General',
  number: '0.00',
  comma: '#,##0.00',
  'comma-0': '#,##0',
  currency: '$#,##0.00',
  'currency-0': '$#,##0',
  accounting: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
  percent: '0.00%',
  'percent-0': '0%',
  scientific: '0.00E+00',
  fraction: '# ?/?',
  'short-date': 'm/d/yyyy',
  'long-date': 'dddd, mmmm d, yyyy',
  time: 'h:mm:ss AM/PM',
  text: '@',
};

// Exact built-in numFmt ids (subset; anything else becomes a custom id ≥164).
export const BUILTIN_NUMFMT_IDS = new Map<string, number>(Object.entries({
  'General': 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4,
  '0%': 9, '0.00%': 10, '0.00E+00': 11, '# ?/?': 12, '# ??/??': 13,
  'm/d/yyyy': 14, 'h:mm:ss AM/PM': 19, '@': 49,
}));
