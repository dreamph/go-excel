package excel

import (
	"github.com/xuri/excelize/v2"
)

type OnRow func(rowIndex int, rowData []string) error

type Reader interface {
	Read(sheetName string, startRow int, onRow OnRow) error
	GetCellValue(sheetName string, cellName string) (string, error)
	ExistsSheet(sheetName string) bool
}

type excelReader struct {
	File *excelize.File
}

func (r *excelReader) GetCellValue(sheetName string, cellName string) (string, error) {
	return r.File.GetCellValue(sheetName, cellName)
}

func NewReader(file *excelize.File) Reader {
	return &excelReader{File: file}
}

func (r *excelReader) Read(sheetName string, startRow int, onRow OnRow) error {
	return r.read(sheetName, startRow, onRow)
}

func (r *excelReader) read(sheetName string, startRow int, onRow OnRow) error {
	rows, err := r.File.GetRows(sheetName)
	if err != nil {
		return err
	}
	for rowIndex, row := range rows {
		if len(row) <= 0 {
			break
		}
		if rowIndex < startRow {
			continue
		}
		err := onRow(rowIndex, row)
		if err != nil {
			return err
		}
	}
	return nil
}

func (r *excelReader) ExistsSheet(sheetName string) bool {
	index, err := r.File.GetSheetIndex(sheetName)
	if err != nil {
		return false
	}
	return index != -1
}
