package main

import (
	"log"
	"net/http"

	"example.com/m/excel_json"
	// "example.com/m/json_excel"
)

func main() {
	// http.HandleFunc("/json_excel", json_excel.WriteExcel)
	// log.Fatal(http.ListenAndServe(":8082", nil))

	http.HandleFunc("/excel_json", excel_json.WriteJson)
	log.Fatal(http.ListenAndServe(":4444", nil))
}
