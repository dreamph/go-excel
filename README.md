## Basic Usage

### Installation

```bash
go install github.com/dreamph/go-excel
```

# Example
# Read
```go
package main

import (
"fmt"
"log"
"path/filepath"

	"github.com/dreamph/go-excel"
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
	defer excel.Close(excelFile)

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
```


# Write
```go
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

```