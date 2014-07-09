// Package implements creation of XLSX simple spreadsheet files

package xlsx

import (
	"archive/zip"
	"bufio"
	"bytes"
	"fmt"
	"html"
	"io"
	"os"
	"strconv"
	"time"
)

type CellType uint

// Basic spreadsheet cell types
const (
	CellTypeNumber CellType = iota
	CellTypeString
	CellTypeDatetime
	CellTypeInlineString
)

// XLSX Spreadsheet Cell
type Cell struct {
	Type  CellType
	Value string
}

// XLSX Spreadsheet Row
type Row struct {
	Cells []Cell
}

// XLSX Spreadsheet Column
type Column struct {
	Name  string
	Width uint64
}

// XLSX Spreadsheet Document Properties
type DocumentInfo struct {
	CreatedBy  string
	ModifiedBy string
	CreatedAt  time.Time
	ModifiedAt time.Time
}

// XLSX Spreadsheet
type Sheet struct {
	Title           string
	columns         []Column
	rows            []Row
	sharedStringMap map[string]int
	sharedStrings   []string
	DocumentInfo    DocumentInfo
}

// Create a sheet with no dimensions
func NewSheet() Sheet {
	c := make([]Column, 0)
	r := make([]Row, 0)
	ssm := make(map[string]int)
	sst := make([]string, 0)

	s := Sheet{
		Title:           "Data",
		columns:         c,
		rows:            r,
		sharedStringMap: ssm,
		sharedStrings:   sst,
	}

	return s
}

// Create a sheet with dimensions derived from the given columns
func NewSheetWithColumns(c []Column) Sheet {
	r := make([]Row, 0)
	ssm := make(map[string]int)
	sst := make([]string, 0)

	s := Sheet{
		Title:           "Data",
		columns:         c,
		rows:            r,
		sharedStringMap: ssm,
		sharedStrings:   sst,
	}

	s.DocumentInfo.CreatedBy = "xlsx.go"
	s.DocumentInfo.CreatedAt = time.Now()

	s.DocumentInfo.ModifiedBy = s.DocumentInfo.CreatedBy
	s.DocumentInfo.ModifiedAt = s.DocumentInfo.CreatedAt

	return s
}

// Create a new row with a length caculated by the sheets known column count
func (s *Sheet) NewRow() Row {
	c := make([]Cell, len(s.columns))
	r := Row{
		Cells: c,
	}
	return r
}

// Append a row to the sheet
func (s *Sheet) AppendRow(r Row) error {
	if len(r.Cells) != len(s.columns) {
		return fmt.Errorf("the given row has %d cells and %d were expected", len(r.Cells), len(s.columns))
	}

	cells := make([]Cell, len(s.columns))

	for n, c := range r.Cells {
		cells[n].Type = c.Type
		cells[n].Value = c.Value

		if cells[n].Type == CellTypeString {
			// calculate string reference
			cells[n].Value = html.EscapeString(cells[n].Value)
			i, exists := s.sharedStringMap[cells[n].Value]
			if !exists {
				i = len(s.sharedStrings)
				s.sharedStringMap[cells[n].Value] = i
				s.sharedStrings = append(s.sharedStrings, cells[n].Value)
			}
			cells[n].Value = strconv.Itoa(i)
		}
	}

	row := s.NewRow()
	row.Cells = cells

	s.rows = append(s.rows, row)

	return nil
}

// Get the Shared Strings in the order they were added to the map
func (s *Sheet) SharedStrings() []string {
	return s.sharedStrings
}

// Given zero-based array indices output the Excel cell reference. For
// example (0,0) => "A1"; (2,2) => "C3"; (26,45) => "AA46"
func CellIndex(x, y uint64) string {
	return fmt.Sprintf("%s%d", colName(x), y+1)
}

// From a zero-based column number return the Excel column name.
// For example: 0 => "A"; 2 => "C"; 26 => "AA"
func colName(n uint64) string {
	var s string
	n += 1

	for n > 0 {
		n -= 1
		s = fmt.Sprintf("%s%s", string(65+(n%26)), s)
		n /= 26
	}

	return s
}

// Convert time to the OLE Automation format.
func OADate(d time.Time) string {
	epoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	nsPerDay := 24 * time.Hour

	v := -1 * float64(epoch.Sub(d)) / float64(nsPerDay)

	// TODO: deal with dates before epoch
	// e.g. http://stackoverflow.com/questions/15549823/oadate-to-milliseconds-timestamp-in-javascript/15550284#15550284

	if d.Hour() == 0 && d.Minute() == 0 && d.Second() == 0 {
		return fmt.Sprintf("%d", int64(v))
	} else {
		return fmt.Sprintf("%f", v)
	}
}

// Create filename and save the XLSX file
func (s *Sheet) SaveToFile(filename string) error {
	outputfile, err := os.Create(filename)
	if err != nil {
		return err
	}
	w := bufio.NewWriter(outputfile)
	err = s.SaveToWriter(w)
	defer w.Flush()
	return err
}

// Save the XLSX file to the given writer
func (s *Sheet) SaveToWriter(w io.Writer) error {

	ww := NewWorkbookWriter(w)

	sw, err := ww.NewSheetWriter(s)
	if err != nil {
		return err
	}

	err = sw.WriteRows(s.rows)
	if err != nil {
		return err
	}

	err = ww.Close()

	return err
}

// Handles the writing of an XLSX workbook
type WorkbookWriter struct {
	zipWriter     *zip.Writer
	sheetWriter   *SheetWriter
	headerWritten bool
	closed        bool
}

// NewWorkbookWriter creates a new WorkbookWriter, which SheetWriters will
// operate on. It must be closed when all Sheets have been written.
func NewWorkbookWriter(w io.Writer) *WorkbookWriter {
	return &WorkbookWriter{zip.NewWriter(w), nil, false, false}
}

// Write the header files of the workbook
func (ww *WorkbookWriter) WriteHeader(s *Sheet) error {
	if ww.closed {
		panic("Can not write to closed WorkbookWriter")
	}

	if ww.headerWritten {
		panic("Workbook header already written")
	}

	z := ww.zipWriter

	f, err := z.Create("[Content_Types].xml")
	err = TemplateContentTypes.Execute(f, nil)
	if err != nil {
		return err
	}

	f, err = z.Create("docProps/app.xml")
	err = TemplateApp.Execute(f, s)
	if err != nil {
		return err
	}

	f, err = z.Create("docProps/core.xml")
	err = TemplateCore.Execute(f, s.DocumentInfo)
	if err != nil {
		return err
	}

	f, err = z.Create("_rels/.rels")
	err = TemplateRelationships.Execute(f, nil)
	if err != nil {
		return err
	}

	f, err = z.Create("xl/workbook.xml")
	err = TemplateWorkbook.Execute(f, s)
	if err != nil {
		return err
	}

	f, err = z.Create("xl/_rels/workbook.xml.rels")
	err = TemplateWorkbookRelationships.Execute(f, nil)
	if err != nil {
		return err
	}

	f, err = z.Create("xl/styles.xml")
	err = TemplateStyles.Execute(f, nil)
	if err != nil {
		return err
	}

	f, err = z.Create("xl/sharedStrings.xml")
	err = TemplateStringLookups.Execute(f, s.SharedStrings())
	if err != nil {
		return err
	}

	return nil
}

// Closes the WorkbookWriter
func (ww *WorkbookWriter) Close() error {
	if ww.closed {
		panic("WorkbookWriter already closed")
	}

	if ww.sheetWriter != nil {
		err := ww.sheetWriter.Close()
		if err != nil {
			return err
		}
	}

	ww.closed = true

	return ww.zipWriter.Close()
}

// NewSheetWriter creates a new SheetWriter in this workbook using the given sheet.
// It returns a SheetWriter to which rows can be written.
// All rows must be written to the SheetWriter before the next call to NewSheetWriter,
// as this will automatically close the previous SheetWriter.
func (ww *WorkbookWriter) NewSheetWriter(s *Sheet) (*SheetWriter, error) {
	if ww.closed {
		panic("Can not write to closed WorkbookWriter")
	}

	if !ww.headerWritten {
		err := ww.WriteHeader(s)
		if err != nil {
			return nil, err
		}
	}

	f, err := ww.zipWriter.Create("xl/worksheets/" + "sheet1" + ".xml")
	sw := &SheetWriter{f, err, 0, 0, false}

	if ww.sheetWriter != nil {
		err = ww.sheetWriter.Close()
		if err != nil {
			return nil, err
		}
	}

	ww.sheetWriter = sw
	err = sw.WriteHeader(s)

	return sw, err
}

// Handles the writing of a sheet
type SheetWriter struct {
	f            io.Writer
	err          error
	currentIndex uint64
	maxNCols     uint64
	closed       bool
}

// Write the given rows to this SheetWriter
func (sw *SheetWriter) WriteRows(rows []Row) error {
	if sw.closed {
		panic("Can not write to closed SheetWriter")
	}

	var err error

	for i, r := range rows {
		rb := &bytes.Buffer{}

		if sw.maxNCols < uint64(len(r.Cells)) {
			sw.maxNCols = uint64(len(r.Cells))
		}

		for j, c := range r.Cells {

			cell := struct {
				CellIndex string
				Value     string
				Type      CellType
			}{
				CellIndex: CellIndex(uint64(j), uint64(i)+sw.currentIndex),
				Value:     c.Value,
				Type:      c.Type,
			}

			if c.Type == CellTypeDatetime {
				d, err := time.Parse(time.RFC3339, c.Value)
				if err == nil {
					cell.Value = OADate(d)
				}
			} else if c.Type == CellTypeInlineString {
				cell.Value = html.EscapeString(c.Value)
			}

			var cellString string

			switch c.Type {
			case CellTypeString:
				cellString = `<c r="%s" t="s" s="1"><v>%s</v></c>`
			case CellTypeInlineString:
				cellString = `<c r="%s" t="inlineStr"><is><t>%s</t></is></c>`
			case CellTypeNumber:
				cellString = `<c r="%s" t="n" s="1"><v>%s</v></c>`
			case CellTypeDatetime:
				cellString = `<c r="%s" s="2"><v>%s</v></c>`
			}

			io.WriteString(rb, fmt.Sprintf(cellString, cell.CellIndex, cell.Value))

			if err != nil {
				return err
			}
		}

		rowString := fmt.Sprintf(`<row r="%d">%s</row>`, uint64(i)+sw.currentIndex+1, rb.String())

		_, err = io.WriteString(sw.f, rowString)
		if err != nil {
			return err
		}

	}

	sw.currentIndex += uint64(len(rows))

	return nil
}

// Closes the SheetWriter
func (sw *SheetWriter) Close() error {
	if sw.closed {
		panic("SheetWriter already closed")
	}

	sheet := struct {
		Start string
		End   string
	}{
		Start: "A1",
		End:   CellIndex(sw.maxNCols-1, sw.currentIndex-1),
	}

	err := TemplateSheetEnd.Execute(sw.f, sheet)

	sw.closed = true

	return err
}

// Writes the header of a sheet
func (sw *SheetWriter) WriteHeader(s *Sheet) error {
	if sw.closed {
		panic("Can not write to closed SheetWriter")
	}

	sheet := struct {
		Cols []Column
	}{
		Cols: s.columns,
	}

	return TemplateSheetStart.Execute(sw.f, sheet)
}
