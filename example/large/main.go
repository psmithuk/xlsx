package main

import (
	"github.com/sean-duffy/xlsx"
	"strconv"
)

func main() {

	c := []xlsx.Column{
		xlsx.Column{Name: "Col1", Width: 10},
		xlsx.Column{Name: "Col2", Width: 10},
	}

	sh := xlsx.NewSheetWithColumns(c, "MySheet")

	for i := 0; i < 10; i++ {

		r := sh.NewRow()

		r.Cells[0] = xlsx.Cell{
			Type:  xlsx.CellTypeNumber,
			Value: strconv.Itoa(i + 1),
		}
		r.Cells[1] = xlsx.Cell{
			Type:  xlsx.CellTypeNumber,
			Value: "1",
		}

		sh.AppendRow(r)
	}

	err := sh.SaveToFile("test.xlsx")
	_ = err
}
