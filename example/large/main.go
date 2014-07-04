package main

import (
	"bufio"
	"github.com/sean-duffy/xlsx"
	"os"
	"strconv"
)

func main() {

	outputfile, err := os.Create("test.xlsx")

	w := bufio.NewWriter(outputfile)
	ww := xlsx.NewWorkbookWriter(w)

	c := []xlsx.Column{
		xlsx.Column{Name: "Col1", Width: 10},
		xlsx.Column{Name: "Col2", Width: 10},
	}

	sh := xlsx.NewSheetWithColumns(c)
	sh.Title = "MySheet"

	sw, err := ww.NewSheetWriter(&sh)

	for i := 0; i < 1000000; i++ {

		r := sh.NewRow()

		r.Cells[0] = xlsx.Cell{
			Type:  xlsx.CellTypeNumber,
			Value: strconv.Itoa(i + 1),
		}
		r.Cells[1] = xlsx.Cell{
			Type:  xlsx.CellTypeNumber,
			Value: "1",
		}

		err = sw.WriteRows([]xlsx.Row{r})
	}

	err = ww.Close()
	defer w.Flush()

	if err != nil {
		panic(err)
	}
}
