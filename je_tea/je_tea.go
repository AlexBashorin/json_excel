package jetea

import (
	"bufio"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"os"
	"reflect"

	"github.com/tealeg/xlsx"
)

type Entry struct {
	Key   string
	Value string
}

type HeinSession struct {
	Session         string `json:"session"`
	Service         string `json:"service"`
	Code            string `json:"code"`
	Date_create     string `json:"date_create"`
	Date_close      string `json:"date_close"`
	Channel         string `json:"channel"`
	Date_appoint    string `json:"date_appoint"`
	Date_last_act   string `json:"date_last_act"`
	Date_read       string `json:"date_read"`
	Clients         string `json:"clients"`
	Operator        string `json:"operator"`
	Group_operators string `json:"group_operators"`
	Appeal          string `json:"appeal"`
	Priority        string `json:"priority"`
	Comment         string `json:"comment"`
}

func Jetea(w http.ResponseWriter, r *http.Request) {
	body, err := io.ReadAll(r.Body)
	if err != nil {
		panic(err)
	}

	// Открытие JSON
	var data []HeinSession
	if err := json.Unmarshal([]byte(body), &data); err != nil {
		fmt.Println("Ошибка при разборе JSON:", err)
		return
	}

	// Создание нового файла Excel
	file := xlsx.NewFile()
	sheet, err := file.AddSheet("Sheet1")
	if err != nil {
		fmt.Println("Ошибка при добавлении листа:", err)
		return
	}

	// Запись заголовков (ключей) в первую строку
	for i := 0; i < len(data); i++ {
		row := sheet.AddRow()
		structReflect := reflect.TypeOf(data[i])
		for j := 0; j < structReflect.NumField(); j++ {
			// field := structReflect.Field(j)
			fieldValue := reflect.ValueOf(data[i]).Field(j).Interface()
			cell := row.AddCell()
			cell.Value = fieldValue.(string)
		}
	}

	// Сохранение файла Excel
	tempFilename := "output.xlsx"
	if err := file.Save(tempFilename); err != nil {
		http.Error(w, "Ошибка при сохранении Excel файла", http.StatusInternalServerError)
		return
	}

	a, _ := os.Open(tempFilename)
	reader := bufio.NewReader(a)
	content, _ := io.ReadAll(reader)
	encoded := base64.StdEncoding.EncodeToString(content)
	a.Close()

	defer os.Remove(tempFilename)

	w.Header().Set("Content-Type", "text/json")
	w.Write([]byte(encoded))
}
