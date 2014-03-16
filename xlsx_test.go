package xlsx

import (
	"bytes"
	"testing"
	"time"
)

type CellIndexTestCase struct {
	x        uint64
	y        uint64
	expected string
}

func TestCellIndex(t *testing.T) {

	tests := []CellIndexTestCase{
		CellIndexTestCase{0, 0, "A1"},
		CellIndexTestCase{2, 2, "C3"},
		CellIndexTestCase{26, 45, "AA46"},
		CellIndexTestCase{2600, 100000, "CVA100001"},
	}

	for _, c := range tests {
		s := CellIndex(c.x, c.y)
		if s != c.expected {
			t.Errorf("expected %s, got %s", c.expected, s)
		}
	}
}

type OADateTestCase struct {
	datetime time.Time
	expected string
}

func TestOADate(t *testing.T) {

	tests := []OADateTestCase{
		OADateTestCase{time.Date(1970, 1, 1, 0, 0, 0, 0, time.UTC), "25569"},
		OADateTestCase{time.Date(1970, 1, 1, 12, 20, 0, 0, time.UTC), "25569.513889"},
		OADateTestCase{time.Date(2014, 12, 20, 0, 0, 0, 0, time.UTC), "41993"},
	}

	for _, d := range tests {
		s := OADate(d.datetime)
		if s != d.expected {
			t.Errorf("expected %s, got %s", d.expected, s)
		}
	}
}

func TestTemplates(t *testing.T) {

	var b bytes.Buffer
	var err error
	var s Sheet

	err = TemplateContentTypes.Execute(&b, nil)
	if err != nil {
		t.Errorf("template TemplateContentTypes failed to Execute returning error %s", err.Error())
	}

	err = TemplateRelationships.Execute(&b, nil)
	if err != nil {
		t.Errorf("template TemplateRelationships failed to Execute returning error %s", err.Error())
	}

	err = TemplateApp.Execute(&b, nil)
	if err != nil {
		t.Errorf("template TemplateApp failed to Execute returning error %s", err.Error())
	}

	err = TemplateCore.Execute(&b, s.DocumentInfo)
	if err != nil {
		t.Errorf("template TemplateCore failed to Execute returning error %s", err.Error())
	}

	err = TemplateWorkbook.Execute(&b, nil)
	if err != nil {
		t.Errorf("template TemplateWorkbook failed to Execute returning error %s", err.Error())
	}

	err = TemplateWorkbookRelationships.Execute(&b, nil)
	if err != nil {
		t.Errorf("template TemplateWorkbookRelationships failed to Execute returning error %s", err.Error())
	}

	err = TemplateStyles.Execute(&b, nil)
	if err != nil {
		t.Errorf("template TemplateStyles failed to Execute returning error %s", err.Error())
	}

	err = TemplateStringLookups.Execute(&b, []string{})
	if err != nil {
		t.Errorf("template TemplateStringLookups failed to Execute returning error %s", err.Error())
	}

	cell := struct {
		CellIndex string
		Value     string
	}{
		CellIndex: "A1",
		Value:     "ABC",
	}

	err = TemplateCellString.Execute(&b, cell)
	if err != nil {
		t.Errorf("template TemplateCellString failed to Execute returning error %s", err.Error())
	}

	err = TemplateCellNumber.Execute(&b, cell)
	if err != nil {
		t.Errorf("template TemplateCellNumber failed to Execute returning error %s", err.Error())
	}

	err = TemplateCellDateTime.Execute(&b, cell)
	if err != nil {
		t.Errorf("template TemplateCellDateTime failed to Execute returning error %s", err.Error())
	}

	sheet := struct {
		Cols  []Column
		Rows  []string
		Start string
		End   string
	}{
		Cols:  []Column{},
		Rows:  []string{},
		Start: "A1",
		End:   "C3",
	}
	err = TemplateSheet.Execute(&b, sheet)
	if err != nil {
		t.Errorf("template TemplateSheet failed to Execute returning error %s", err.Error())
	}
}
