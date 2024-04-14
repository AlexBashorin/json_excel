package excel_json

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"os"

	excelize "github.com/xuri/excelize/v2"
)

func WriteJson(rw http.ResponseWriter, req *http.Request) {
	//========== READ ==========//
	err := req.ParseMultipartForm(32 << 20)
	if err != nil {
		fmt.Println("Parse", err)
	}

	fhs := req.MultipartForm.File["file"]
	// file_name := fhs[0].Filename

	nf, err := os.Create("ex.xlsx")
	if err != nil {
		fmt.Println("nf:", err)
	}
	defer nf.Close()

	opened, err := fhs[0].Open()
	if err != nil {
		fmt.Println(err)
	}

	buf := bytes.NewBuffer(nil)
	if _, err := io.Copy(buf, opened); err != nil {
		fmt.Println(err)
	}

	newFile, err := nf.Write(buf.Bytes())
	if err != nil {
		fmt.Println(err)
	}
	if err != nil {
		fmt.Println(err)
	}
	fmt.Println("writed: ", newFile)

	// gg, err = fhs[0].Open()
	// if err != nil {
	// 	fmt.Println(err)
	// }

	f, err := excelize.OpenFile("ex.xlsx")
	if err != nil {
		fmt.Println("cannot excelize.OpenFile: ", err)
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
			fmt.Println("Close", err)
		}
	}()

	// rw.Header().Set("Content-Type", "text/json")
	rw.Write([]byte(jsonBytes))

	os.Remove(nf.Name())
}
