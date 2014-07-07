package main

import (
	"bufio"
	"os"
	"strconv"

	"github.com/psmithuk/xlsx"
)

func main() {
	err := WriteStreaming()
	//err := WriteNoStreaming()

	if err != nil {
		panic(err)
	}
}

// Write a simple 1,000,000 row spreadsheet using streaming
// This has a maximum resident set of ~6.6MB
func WriteStreaming() error {
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
			Type:  xlsx.CellTypeInlineString,
			Value: "Test",
		}

		err = sw.WriteRows([]xlsx.Row{r})
	}

	err = ww.Close()
	defer w.Flush()

	return err
}

// Write a simple 1,000,000 row spreadsheet without using streaming
// This has a maximum resident set of ~240MB
func WriteNoStreaming() error {
	c := []xlsx.Column{
		xlsx.Column{Name: "Col1", Width: 10},
		xlsx.Column{Name: "Col2", Width: 10},
	}

	sh := xlsx.NewSheetWithColumns(c)
	sh.Title = "MySheet"

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

		sh.AppendRow(r)
	}

	err := sh.SaveToFile("test.xlsx")

	return err
}
