package main

import (
	"log"
	"net/http"

	"example.com/m/excel_json"
	jetea "example.com/m/je_tea"
	"example.com/m/json_excel"
	"github.com/gorilla/mux"
)

func main() {
	r := mux.NewRouter()

	// JSON to EXCEL
	r.HandleFunc("/json_excel", json_excel.WriteExcel)

	// JSON to JSON
	r.HandleFunc("/jetea", jetea.Jetea)

	// EXCEL to JSON
	r.HandleFunc("/excel_json", excel_json.WriteJson)

	log.Fatal(http.ListenAndServe(":4444", r))
}
