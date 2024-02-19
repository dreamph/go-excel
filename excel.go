package sodau

import (
	"fmt"
	"io"

	"github.com/xuri/excelize/v2"
)

const (
	FirstRowIndex = 0
)

func OpenFile(filePath string) (*excelize.File, error) {
	file, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}
	return file, nil
}

func OpenReader(reader io.Reader) (*excelize.File, error) {
	file, err := excelize.OpenReader(reader)
	if err != nil {
		return nil, err
	}
	return file, nil
}

func ToBytes(file *excelize.File) ([]byte, error) {
	buf, err := file.WriteToBuffer()
	if err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func GetSheetName(file *excelize.File, index int) string {
	return file.GetSheetName(index)
}

func Close(f *excelize.File) {
	if f != nil {
		_ = f.Close()
	}
}

func ToChar(i int) string {
	return string(rune('A' + i))
}

func ToCellName(cellIndex int, rowIndex int) string {
	return fmt.Sprint(ToChar(cellIndex), rowIndex+1)
}

type GenerateExcelConfig[T any] struct {
	Header string
	Value  func(obj *T) string
}

type DataRef struct {
	SheetName   string
	CellName    string
	Data        interface{}
	TextColor   string
	CellFormula string
}

type ErrorDataRow struct {
	CellErrors *[]DataRef
	Detail     *DataRef
}

type RowInfo struct {
	SheetName string
	RowIndex  int
	RowData   []string
}
