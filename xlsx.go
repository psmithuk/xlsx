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
		} else if cells[n].Type == CellTypeDatetime {
			d, err := time.Parse(time.RFC3339, cells[n].Value)
			if err == nil {
				cells[n].Value = OADate(d)
			}
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

func (sw *SheetWriter) WriteRows(rows []Row) error {

	var err error

	sheet := struct {
		Rows []string
	}{
		Rows: make([]string, len(rows)),
	}

	for i, r := range rows {
		rb := &bytes.Buffer{}
		for j, c := range r.Cells {

			cell := struct {
				CellIndex string
				Value     string
			}{
				CellIndex: CellIndex(uint64(j), uint64(i)),
				Value:     c.Value,
			}

			switch c.Type {
			case CellTypeString:
				err = TemplateCellString.Execute(rb, cell)
			case CellTypeNumber:
				err = TemplateCellNumber.Execute(rb, cell)
			case CellTypeDatetime:
				err = TemplateCellDateTime.Execute(rb, cell)
			}

			if err != nil {
				return err
			}
		}
		sheet.Rows[i] = rb.String()
	}

	err = TemplateSheetRows.Execute(sw.f, sheet)
	if err != nil {
		return err
	}

    return nil
}

// Save the XLSX file to the given writer
func (s *Sheet) SaveToWriter(w io.Writer) error {

	ww := NewWorkbookWriter(w)

	err := ww.WriteHeader(s)
	if err != nil {
		return err
	}

    sw := ww.NewSheetWriter()

    sw.Write(s)
	sw.WriteRows(s.rows)

	err = ww.Close()
	if err != nil {
		return err
	}

	return nil
}

func (ww *WorkbookWriter) WriteHeader(s *Sheet) error {

	z := ww.zipWriter

	f, err := z.Create("[Content_Types].xml")
	err = TemplateContentTypes.Execute(f, nil)
	if err != nil {
		return err
	}

	f, err = z.Create("docProps/app.xml")
	err = TemplateApp.Execute(f, nil)
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
	err = TemplateWorkbook.Execute(f, nil)
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

type WorkbookWriter struct {
	zipWriter *zip.Writer
}

func NewWorkbookWriter(w io.Writer) *WorkbookWriter {
	return &WorkbookWriter{zip.NewWriter(w)}
}

func (ww *WorkbookWriter) Close() error {
	return ww.zipWriter.Close()
}

func (ww *WorkbookWriter) NewSheetWriter() *SheetWriter {
	f, err := ww.zipWriter.Create("xl/worksheets/sheet1.xml")
    return &SheetWriter{f, err}
}

type SheetWriter struct {
	f   io.Writer
	err error
}

func (sw *SheetWriter) Close() error {
    err := TemplateSheetEnd.Execute(sw.f, nil)
    return err
}

func (sw *SheetWriter) Write(s *Sheet) error {
    sheet := struct {
		Cols  []Column
		Start string
		End   string
	}{
		Cols: s.columns,
	}

	sheet.Start = "A1"
	sheet.End = CellIndex(uint64(len(s.columns)-1), uint64(len(s.rows)-1))

    err := TemplateSheetStart.Execute(sw.f, sheet)
    return err
}
