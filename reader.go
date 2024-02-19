package sodau

import (
	"github.com/xuri/excelize/v2"
)

type OnRow func(rowIndex int, rowData []string) error

type ExcelReader interface {
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

func NewExcelReader(file *excelize.File) ExcelReader {
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

type Data struct {
	ID   string
	Name string
	Age  int
}

/*
func main() {
	filePath := "/Users/dream/Data/go/src/sodau/Book2.xlsx"
	excelReader := NewExcelReader()

	var datas []Data
	err := excelReader.Read(filePath, "Sheet1", 1, func(rowIndex int, rowData []string) error {
		data := Data{}
		data.ID = rowData[0]
		data.Name = rowData[1]
		data.Age, _ = strconv.Atoi(rowData[2])
		datas = append(datas, data)
		return nil
	})
	if err != nil {
		fmt.Print(err.Error())
	}
	fmt.Println(len(datas))

}

*/
