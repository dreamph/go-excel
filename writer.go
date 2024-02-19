package sodau

import (
	"github.com/xuri/excelize/v2"
)

type ExcelWriter interface {
	Write(sheetName string, CellName string, data interface{}) error
	WriteList(list []DataRef) error
	SaveAs(filePath string) error
	ToBytes() ([]byte, error)
}

type excelWriter struct {
	file *excelize.File
}

func (e *excelWriter) ToBytes() ([]byte, error) {
	return ToBytes(e.file)
}

func (e *excelWriter) SaveAs(filePath string) error {
	return e.file.SaveAs(filePath)
}

func NewExcelWriter(file *excelize.File) ExcelWriter {
	return &excelWriter{file: file}
}

func (e *excelWriter) Write(sheetName string, colName string, data interface{}) error {
	sheetVisible, _ := e.file.GetSheetVisible(sheetName)
	if !sheetVisible {
		err := e.file.SetSheetName("Sheet1", sheetName)
		if err != nil {
			return err
		}
	}

	return e.file.SetCellValue(sheetName, colName, data)
}

func (e *excelWriter) WriteList(list []DataRef) error {
	for _, data := range list {
		if data.TextColor != "" {
			style, err := e.file.NewStyle(&excelize.Style{Font: &excelize.Font{Color: data.TextColor}})
			if err != nil {
				return err
			}
			err = e.file.SetCellStyle(data.SheetName, data.CellName, data.CellName, style)
			if err != nil {
				return err
			}
		}
		if data.CellFormula != "" {
			err := e.file.SetCellFormula(data.SheetName, data.CellName, data.CellFormula)
			if err != nil {
				return err
			}
		}

		err := e.Write(data.SheetName, data.CellName, data.Data)
		if err != nil {
			return err
		}
	}
	return nil
}

/*
	GenerateExcelAsBytes for create excel by configs
    var configs []excel.GenerateExcelConfig[models.ActivityData]
	configs = append(configs, excel.GenerateExcelConfig[models.ActivityData]{
		Header: "Test1",
		Value: func(obj *models.ActivityData) string {
			return obj.ID
		},
	})
	configs = append(configs, excel.GenerateExcelConfig[models.ActivityData]{
		Header: "Test1",
		Value: func(obj *models.ActivityData) string {
			return obj.ID
		},
	})
	dataList, _, err := repositoryRegistry.ActivityRepository.List(ctx, &models.ActivityListRequest{
		Limit: coremodels.MaxLimitForQuery,
	})
	dataBytes, err := excel.GenerateExcelAsBytes("Data", excel.FirstRowIndex, &configs, dataList)
	utils.WriteFile("test.xlsx", dataBytes)
*/

func GenerateExcelAsBytes[T any](sheetName string, startAtRowIndex int, configs *[]GenerateExcelConfig[T], dataList *[]T) ([]byte, error) {
	f := excelize.NewFile()
	defer Close(f)
	excelWriter := NewExcelWriter(f)

	var list []DataRef
	currentRow := startAtRowIndex
	for cellIndex, config := range *configs {
		list = append(list, DataRef{SheetName: sheetName, CellName: ToCellName(cellIndex, currentRow), Data: config.Header})
	}

	if dataList != nil && len(*dataList) > 0 {
		for _, row := range *dataList {
			data := row
			currentRow++
			for cellIndex, config := range *configs {
				list = append(list, DataRef{SheetName: sheetName, CellName: ToCellName(cellIndex, currentRow), Data: config.Value(&data)})
			}
		}
	}

	err := excelWriter.WriteList(list)
	if err != nil {
		return nil, err
	}

	fileBytes, err := excelWriter.ToBytes()
	if err != nil {
		return nil, err
	}
	return fileBytes, nil
}
