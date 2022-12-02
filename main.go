package main

import (
	"fmt"
	"strconv"

	"github.com/xuri/excelize/v2"
)

func main() {
	// import excel file
	f, err := excelize.OpenFile("./dummies-migration-tests.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	sheetName := "transactions"
	// get data inside the file
	rows, err := f.GetRows(sheetName)
	if err != nil {
		fmt.Println(err)
	}

	createData(rows)
}

func createData(data [][]string) {
	f := excelize.NewFile()

	i := f.NewSheet("datamart-1")

	loopRow := 1
	for _, val := range data {
		cell := ""
		loopCell := 1
		for _, colCell := range val {
			switch {
			case loopCell == 1:
				cell = "A"
			case loopCell == 2:
				cell = "B"
			case loopCell == 3:
				cell = "C"
			case loopCell == 4:
				cell = "D"
			case loopCell == 5:
				cell = "E"
			case loopCell == 6:
				cell = "F"
			case loopCell == 7:
				cell = "G"
			case loopCell == 8:
				cell = "H"
			case loopCell == 9:
				cell = "I"
			case loopCell == 10:
				cell = "J"
			case loopCell == 11:
				cell = "K"
			case loopCell == 12:
				cell = "L"
			case loopCell == 13:
				cell = "M"
			case loopCell == 14:
				cell = "N"
			case loopCell == 15:
				cell = "O"
			case loopCell == 16:
				cell = "P"
			case loopCell == 17:
				cell = "Q"
			case loopCell == 18:
				cell = "R"
			case loopCell == 19:
				cell = "S"
			case loopCell == 20:
				cell = "T"
			case loopCell == 21:
				cell = "U"
			}

			if loopCell == 21 {
				loopCell = 1
				loopRow++
			}

			f.SetCellValue("datamart-1", cell+strconv.Itoa(loopRow), colCell)
			f.SetActiveSheet(i)
			fmt.Println(cell + strconv.Itoa(loopRow))
			loopCell++
		}
	}
	fmt.Println(i)
	if err := f.SaveAs("data-migration.xlsx"); err != nil {
		fmt.Println(err)
	}
}
