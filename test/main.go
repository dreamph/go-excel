package main

import (
	"go-excel/pkg/core/excel"
	"os"

	"fmt"
)

/*
func main() {
	filePath := "/Users/dream/Data/go/src/sodau/TM101.xlsx"
	excelFile, err := excel.OpenFile(filePath)
	if err != nil {
		fmt.Print(err.Error())
		panic(err.Error())
	}
	excelWriter := excel.NewExcelWriter(excelFile)
	err = excelWriter.Write("Certificate", "A1", "Test")

	if err != nil {
		fmt.Print(err.Error())
		panic(err.Error())
	}

	err =  excelWriter.SaveAs("TM101_new.xlsx")
	if err != nil {
		fmt.Print(err.Error())
		panic(err.Error())
	}
}

*/

func main() {
	filePath := "/Users/dream/Data/go/src/sodau/TM101.xlsx"
	file, err := os.Open(filePath)
	if err != nil {
		panic(err)
	}
	excelFile, err := excel.OpenReader(file)
	if err != nil {
		panic(err)
	}
	defer excel.Close(excelFile)
	excelWriter := excel.NewExcelWriter(excelFile)
	var list []excel.DataRef
	for i := 1; i < 5; i++ {
		list = append(list, excel.DataRef{SheetName: "Certificate", CellName: fmt.Sprintf("G%d", i), Data: "Invalid Format", TextColor: "#1265BE"})
	}

	err = excelWriter.WriteList(list)

	if err != nil {
		fmt.Print(err.Error())
		panic(err.Error())
	}

	err = excelWriter.SaveAs("TM101_new.xlsx")
	if err != nil {
		fmt.Print(err.Error())
		panic(err.Error())
	}

	fileBytes, err := excel.ToBytes(excelFile)
	if err != nil {
		fmt.Print(err.Error())
		panic(err.Error())
	}
	fmt.Println(fileBytes)
}
