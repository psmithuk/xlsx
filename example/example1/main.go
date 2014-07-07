package main

import (
	"time"

	"github.com/psmithuk/xlsx"
)

func main() {

	c := []xlsx.Column{
		xlsx.Column{Name: "Col1", Width: 10},
		xlsx.Column{Name: "Col2", Width: 10},
		xlsx.Column{Name: "Col3", Width: 10},
	}

	sh := xlsx.NewSheetWithColumns(c)
	r := sh.NewRow()

	r.Cells[0] = xlsx.Cell{
		Type:  xlsx.CellTypeNumber,
		Value: "10",
	}
	r.Cells[1] = xlsx.Cell{
		Type:  xlsx.CellTypeString,
		Value: "Apple",
	}
	r.Cells[2] = xlsx.Cell{
		Type:  xlsx.CellTypeDatetime,
		Value: time.Date(1970, 1, 1, 0, 0, 0, 0, time.UTC).Format(time.RFC3339),
	}

	sh.AppendRow(r)

	r2 := sh.NewRow()

	r2.Cells[0] = xlsx.Cell{
		Type:  xlsx.CellTypeNumber,
		Value: "10",
	}
	r2.Cells[1] = xlsx.Cell{
		Type:  xlsx.CellTypeString,
		Value: "Apple",
	}
	r2.Cells[2] = xlsx.Cell{
		Type:  xlsx.CellTypeDatetime,
		Value: time.Date(1970, 1, 1, 0, 0, 0, 0, time.UTC).Format(time.RFC3339),
	}

	sh.AppendRow(r2)

	err := sh.SaveToFile("test.xlsx")
	if err != nil {
		println(err)
	}
}
