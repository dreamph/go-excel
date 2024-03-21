package main

import (
	"fmt"
	"log"
	"os"

	"github.com/dreamph/go-excel"
)

type Customer struct {
	Name     string `json:"name"`
	MobileNo string `json:"mobileNo"`
}

func GenerateCustomerData() *[]Customer {
	var list []Customer
	for i := 1; i <= 10; i++ {
		list = append(list, Customer{
			Name:     fmt.Sprintf("Name%d", i),
			MobileNo: fmt.Sprintf("00000000%d", i),
		})
	}
	return &list
}

func main() {
	var configs []excel.GenerateExcelConfig[Customer]
	configs = append(configs, excel.GenerateExcelConfig[Customer]{
		Header: "Name",
		Value: func(obj *Customer) string {
			return obj.Name
		},
	})
	configs = append(configs, excel.GenerateExcelConfig[Customer]{
		Header: "MobileNo",
		Value: func(obj *Customer) string {
			return obj.MobileNo
		},
	})
	dataList := GenerateCustomerData()
	dataBytes, err := excel.GenerateExcelAsBytes("Data", excel.FirstRowIndex, &configs, dataList)
	if err != nil {
		log.Fatalln(err)
	}

	err = WriteFile("test.xlsx", dataBytes)
	if err != nil {
		log.Fatalln(err)
	}
}

func WriteFile(filePath string, bytes []byte) error {
	file, err := os.Create(filePath)
	if err != nil {
		return err
	}
	defer func(file *os.File) {
		_ = file.Close()
	}(file)

	_, err = file.Write(bytes)
	if err != nil {
		return err
	}
	return nil
}
