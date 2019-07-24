package goExcelUtil

import (
	"fmt"
	"log"
	"os"
	"strconv"

	"github.com/tealeg/xlsx"
)

// Replace function
func Replace() {

	//log.Println(os.Args[0] + os.Args[1] + os.Args[2] + os.Args[3] + os.Args[4])
	fileName := os.Args[1]

	sheetNumber, err := strconv.Atoi(os.Args[2])
	if err != nil {
		log.Fatal(fmt.Errorf("Error Converting %s to type Int.\n Trace:\n %s", os.Args[2], err))
	}
	if sheetNumber == 0 {
		log.Fatal("Sheet number musn't be zero")
	}

	rowNumber, err := strconv.Atoi(os.Args[3])
	if err != nil {
		log.Fatal(fmt.Errorf("Error Converting %s to type Int.\n Trace:\n %s", os.Args[3], err))
	}
	if rowNumber == 0 {
		log.Fatal("Row number musn't be zero")
	}

	columnNumber, err := strconv.Atoi(os.Args[4])
	if err != nil {
		log.Fatal(fmt.Errorf("Error Converting %s to type Int.\nTrace:\n %s", os.Args[4], err))
	}
	if columnNumber == 0 {
		log.Fatal("Column number musn't be zero")
	}

	replacementString := os.Args[5]

	xlFile, err := xlsx.OpenFile(fileName)
	if err != nil {
		log.Fatal(fmt.Errorf("Error opening file %s.\nTrace:\n%s", fileName, err))
	}

	cell := xlFile.Sheets[sheetNumber-1].Rows[rowNumber-1].Cells[columnNumber-1]
	log.Printf("Replacing '%s' with '%s'", cell.String(), replacementString)
	cell.Value = replacementString

	err = xlFile.Save(fileName)
	if err != nil {
		log.Fatal(fmt.Errorf("Error saving file %s.\nTrace:\n%s", fileName, err))
	}

}
