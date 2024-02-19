package main

import (
	"log"
	"path/filepath"

	"fmt"

	"github.com/dreamph/go-excel"
)

func main() {
	filePath := "example/write/write.xlsx"
	excelFile, err := excel.OpenFile(filepath.Clean(filePath))
	if err != nil {
		log.Fatalf(err.Error())
	}
	defer excel.Close(excelFile)

	excelWriter := excel.NewWriter(excelFile)
	var list []excel.DataRef
	for i := 1; i < 5; i++ {
		list = append(list, excel.DataRef{SheetName: "Sheet1", CellName: fmt.Sprintf("C%d", i), Data: "Invalid Format", TextColor: "#1265BE"})
	}

	err = excelWriter.WriteList(list)

	if err != nil {
		log.Fatalf(err.Error())
	}

	err = excelWriter.SaveAs("example/write/write_1.xlsx")
	if err != nil {
		log.Fatalf(err.Error())
	}

	/*
		fileBytes, err := excel.ToBytes(excelFile)
		if err != nil {
			log.Fatalf(err.Error())
		}
		fmt.Println(fileBytes)
	*/
}
