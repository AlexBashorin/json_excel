package excel_json

import (
	"encoding/json"
	"fmt"
	"net/http"

	excelize "github.com/xuri/excelize/v2"
)

func WriteJson(rw http.ResponseWriter, req *http.Request) {
	//========== READ ==========//
	err := req.ParseMultipartForm(32 << 20)
	if err != nil {
		fmt.Println(err)
	}

	fhs := req.MultipartForm.File["file"]
	file_name := fhs[0].Filename

	f, err := excelize.OpenFile(file_name)
	if err != nil {
		fmt.Println(err)
		return
	}

	// ROWS
	sheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		fmt.Println(err)
		return
	}

	// object: first row is array of keys, other rows is values
	bodyRow := rows[1:]

	var result = []map[string]string{}
	for _, row := range bodyRow {
		// берем всегда из первой строки ключ для создания объекта со строкой
		key := rows[0]
		var jsonExcel = make(map[string]string)
		for i := 0; i < len(row); i++ {
			jsonExcel[key[i]] = row[i]
		}
		result = append(result, jsonExcel)
	}
	jsonBytes, err := json.Marshal(result)
	if err != nil {
		fmt.Println("Error marshaling JSON:", err)
		return
	}

	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	rw.Header().Set("Content-Type", "text/json")
	rw.Write([]byte(jsonBytes))
}
