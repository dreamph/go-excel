package main

import (
	"fmt"
	"log"
	"path/filepath"

	"github.com/dreamph/go-excel"
	"github.com/xuri/excelize/v2"
)

type Data struct {
	ID   string
	Name string
	Age  int
}

func main() {
	filePath := "example/read/read.xlsx"

	excelFile, err := excel.OpenFile(filepath.Clean(filePath))
	if err != nil {
		log.Fatalf(err.Error())
	}
	defer func(excelFile *excelize.File) {
		err := excelFile.Close()
		if err != nil {

		}
	}(excelFile)

	excelReader := excel.NewReader(excelFile)

	var result []Data
	err = excelReader.Read("Sheet1", 1, func(rowIndex int, rowData []string) error {
		data := Data{}
		data.ID = rowData[0]
		data.Name = rowData[1]
		result = append(result, data)
		return nil
	})
	if err != nil {
		fmt.Print(err.Error())
	}

	fmt.Println(len(result))

}
