package main

import (
	"github.com/psmithuk/xlsx"
)

func main() {

	c := []xlsx.Column{
		xlsx.Column{Name: "Col1", Width: 10},
		xlsx.Column{Name: "Col2", Width: 10},
	}

	sh := xlsx.NewSheetWithColumns(c)

	for i := 0; i < 100000; i++ {

		r := sh.NewRow()

		r.Cells[0] = xlsx.Cell{
			Type:  xlsx.CellTypeNumber,
			Value: "1",
		}
		r.Cells[1] = xlsx.Cell{
			Type:  xlsx.CellTypeNumber,
			Value: "2",
		}

		sh.AppendRow(r)
    }

    err := sh.SaveToFile("test.xlsx")
    _ = err
}
