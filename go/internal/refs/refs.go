// Package refs implements A1-reference parsing (SPEC.md §8.2, Appendix A).
package refs

import (
	"regexp"
	"strconv"
	"strings"
)

// MaxCol is column XFD; MaxRow is the Excel row ceiling.
const (
	MaxCol = 16384 // XFD
	MaxRow = 1048576
)

// ColToNum converts column letters (e.g. "AB") to a 1-based index.
func ColToNum(letters string) int {
	n := 0
	for _, ch := range letters {
		n = n*26 + int(ch-'A'+1)
	}
	return n
}

// NumToCol converts a 1-based column index to letters (e.g. 28 -> "AB").
func NumToCol(n int) string {
	var b []byte
	for n > 0 {
		r := (n - 1) % 26
		b = append([]byte{byte('A' + r)}, b...)
		n = (n - 1 - r) / 26
	}
	return string(b)
}

// Cell is a resolved single-cell coordinate (1-based).
type Cell struct {
	Col int
	Row int
}

// Kind discriminates the shape of a parsed target.
type Kind int

// Target kinds.
const (
	KindNone Kind = iota
	KindCell
	KindRange
	KindCols
	KindRows
)

// Target is a parsed @-directive/anchor target.
type Target struct {
	Kind  Kind
	Sheet string // "" when unqualified
	HasSheet bool
	C1, R1 int
	C2, R2 int
}

var (
	cellRe     = regexp.MustCompile(`^(\$?)([A-Z]{1,3})(\$?)([1-9]\d{0,6})$`)
	colRangeRe = regexp.MustCompile(`^\$?([A-Z]{1,3}):\$?([A-Z]{1,3})$`)
	rowRangeRe = regexp.MustCompile(`^\$?([1-9]\d{0,6}):\$?([1-9]\d{0,6})$`)
)

// ParseCell parses a single A1 cell; returns nil when the text is not a cell
// or exceeds the grid bounds.
func ParseCell(text string) *Cell {
	m := cellRe.FindStringSubmatch(text)
	if m == nil {
		return nil
	}
	col := ColToNum(m[2])
	row, _ := strconv.Atoi(m[4])
	if col > MaxCol || row > MaxRow {
		return nil
	}
	return &Cell{Col: col, Row: row}
}

// ParseTarget parses a target: cell | cell:cell | col:col | row:row, with an
// optional leading Sheet! qualifier ('quoted' names supported). Returns nil on
// failure.
func ParseTarget(input string) *Target {
	text := input
	var sheet string
	hasSheet := false
	if bang := strings.LastIndex(text, "!"); bang != -1 {
		sheet = text[:bang]
		hasSheet = true
		if len(sheet) >= 2 && strings.HasPrefix(sheet, "'") && strings.HasSuffix(sheet, "'") {
			sheet = strings.ReplaceAll(sheet[1:len(sheet)-1], "''", "'")
		}
		text = text[bang+1:]
	}
	if cell := ParseCell(text); cell != nil {
		return &Target{Kind: KindCell, Sheet: sheet, HasSheet: hasSheet, C1: cell.Col, R1: cell.Row, C2: cell.Col, R2: cell.Row}
	}
	if strings.Contains(text, ":") {
		parts := strings.Split(text, ":")
		if len(parts) == 2 {
			a := ParseCell(parts[0])
			b := ParseCell(parts[1])
			if a != nil && b != nil {
				return &Target{
					Kind: KindRange, Sheet: sheet, HasSheet: hasSheet,
					C1: min(a.Col, b.Col), R1: min(a.Row, b.Row),
					C2: max(a.Col, b.Col), R2: max(a.Row, b.Row),
				}
			}
			if m := colRangeRe.FindStringSubmatch(text); m != nil {
				c1 := ColToNum(m[1])
				c2 := ColToNum(m[2])
				if c1 <= MaxCol && c2 <= MaxCol {
					return &Target{Kind: KindCols, Sheet: sheet, HasSheet: hasSheet, C1: min(c1, c2), C2: max(c1, c2)}
				}
			}
			if m := rowRangeRe.FindStringSubmatch(text); m != nil {
				r1, _ := strconv.Atoi(m[1])
				r2, _ := strconv.Atoi(m[2])
				if r1 <= MaxRow && r2 <= MaxRow {
					return &Target{Kind: KindRows, Sheet: sheet, HasSheet: hasSheet, R1: min(r1, r2), R2: max(r1, r2)}
				}
			}
		}
	}
	return nil
}

// RefKey is the map key for a (col,row) coordinate.
func RefKey(col, row int) string {
	return strconv.Itoa(col) + "," + strconv.Itoa(row)
}
