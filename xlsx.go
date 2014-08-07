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
	Type    CellType
	Value   string
	Colspan uint64
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

// XLSX Workbook Document Properties
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
func CellIndex(x, y uint64) (string, uint64) {
	return colName(x), (y + 1)
}

// From a zero-based column number return the Excel column name.
// For example: 0 => "A"; 2 => "C"; 26 => "AA"
func colName(n uint64) string {
	var s string
	n += 1

	for n > 0 {
		n -= 1
		s = string(65+(n%26)) + s
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

	ww.SharedStrings = s.sharedStrings

	err = ww.Close()

	return err
}

// Write the header files of the workbook
func (ww *WorkbookWriter) WriteHeader() error {
	if ww.headerWritten {
		panic("Workbook header already written")
	}

	z := ww.zipWriter

	f, err := z.Create("[Content_Types].xml")
	err = TemplateContentTypes.Execute(f, ww.sheetNames)
	if err != nil {
		return err
	}

	f, err = z.Create("docProps/app.xml")
	err = TemplateApp.Execute(f, ww.sheetNames)
	if err != nil {
		return err
	}

	f, err = z.Create("docProps/core.xml")
	err = TemplateCore.Execute(f, ww.documentInfo)
	if err != nil {
		return err
	}

	f, err = z.Create("_rels/.rels")
	err = TemplateRelationships.Execute(f, nil)
	if err != nil {
		return err
	}

	f, err = z.Create("xl/workbook.xml")
	err = TemplateWorkbook.Execute(f, ww.sheetNames)
	if err != nil {
		return err
	}

	f, err = z.Create("xl/_rels/workbook.xml.rels")
	err = TemplateWorkbookRelationships.Execute(f, ww.sheetNames)
	if err != nil {
		return err
	}

	f, err = z.Create("xl/styles.xml")
	err = TemplateStyles.Execute(f, nil)
	if err != nil {
		return err
	}

	f, err = z.Create("xl/sharedStrings.xml")
	err = TemplateStringLookups.Execute(f, ww.SharedStrings)
	if err != nil {
		return err
	}

	return err
}

// Handles the writing of an XLSX workbook
type WorkbookWriter struct {
	zipWriter     *zip.Writer
	sheetWriter   *SheetWriter
	headerWritten bool
	closed        bool
	sheetNames    []string
	SharedStrings []string
	documentInfo  *DocumentInfo
}

// Creates a new WorkbookWriter
func NewWorkbookWriter(w io.Writer) *WorkbookWriter {
	return &WorkbookWriter{zip.NewWriter(w), nil, false, false, []string{}, nil, nil}
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

	if !ww.headerWritten {
		err := ww.WriteHeader()
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

	if ww.sheetWriter != nil {
		err := ww.sheetWriter.Close()
		if err != nil {
			return nil, err
		}
	}

	f, err := ww.zipWriter.Create("xl/worksheets/" + fmt.Sprintf("sheet%s", strconv.Itoa(len(ww.sheetNames)+1)) + ".xml")
	sw := &SheetWriter{f, err, 0, 0, false, "", 0}

	ww.documentInfo = &s.DocumentInfo

	ww.sheetWriter = sw
	err = sw.WriteHeader(s)

	ww.sheetNames = append(ww.sheetNames, s.Title)

	return sw, err
}

// Handles the writing of a sheet
type SheetWriter struct {
	f               io.Writer
	err             error
	currentIndex    uint64
	maxNCols        uint64
	closed          bool
	mergeCells      string
	mergeCellsCount int
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

			cellX, cellY := CellIndex(uint64(j), uint64(i)+sw.currentIndex)

			if c.Type == CellTypeDatetime {
				d, err := time.Parse(time.RFC3339, c.Value)
				if err == nil {
					c.Value = OADate(d)
				} else {
					return err
				}
			} else if c.Type == CellTypeInlineString {
				c.Value = html.EscapeString(c.Value)
			}

			var cellString string

			switch c.Type {
			case CellTypeString:
				cellString = `<c r="%s%d" t="s" s="1"><v>%s</v></c>`
			case CellTypeInlineString:
				cellString = `<c r="%s%d" t="inlineStr"><is><t>%s</t></is></c>`
			case CellTypeNumber:
				cellString = `<c r="%s%d" t="n" s="1"><v>%s</v></c>`
			case CellTypeDatetime:
				cellString = `<c r="%s%d" s="2"><v>%s</v></c>`
			}

			if c.Colspan < 0 {
				panic(fmt.Sprintf("%v is not a valid colspan", c.Colspan))
			} else if c.Colspan > 1 {
				mergeCellX, _ := CellIndex(uint64(j)+c.Colspan-1, uint64(i)+sw.currentIndex)
				sw.mergeCells += fmt.Sprintf(`<mergeCell ref="%[1]s%[2]d:%[3]s%[2]d"/>`, cellX, cellY, mergeCellX)
				sw.mergeCellsCount += 1
			}

			io.WriteString(rb, fmt.Sprintf(cellString, cellX, cellY, c.Value))

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

	cellEndX, cellEndY := CellIndex(sw.maxNCols-1, sw.currentIndex-1)
	sheetEnd := fmt.Sprintf(`<dimension ref="A1:%s%d"/></sheetData>`, cellEndX, cellEndY)
	if sw.mergeCellsCount > 0 {
		sheetEnd += fmt.Sprintf(`<mergeCells count="%v">`, sw.mergeCellsCount)
		sheetEnd += sw.mergeCells
		sheetEnd += `</mergeCells>`
	}
	sheetEnd += `</worksheet>`
	_, err := io.WriteString(sw.f, sheetEnd)

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
