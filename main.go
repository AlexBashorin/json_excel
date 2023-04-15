package main

import (
	"log"
	"net/http"

	"example.com/m/json_excel"
)

func main() {
	http.HandleFunc("/json_excel", json_excel.WriteExcel)
	log.Fatal(http.ListenAndServe(":8082", nil))

	// http.HandleFunc("/json_excel", excel_json.WriteJson)
	// log.Fatal(http.ListenAndServe(":4044", nil))
}
