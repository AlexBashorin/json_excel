package json_excel

import (
	"bufio"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"math"
	"net/http"
	"os"
	"strconv"

	excelize "github.com/xuri/excelize/v2"
)

type JD struct {
	jsdata []map[string]interface{} `json:"data"`
}

type Data struct {
	NameSheet string                   `json:"nameSheet"`
	NameFile  string                   `json:"nameFile"`
	JsonData  []map[string]interface{} `json:"data"`
}

func WriteExcel(rw http.ResponseWriter, req *http.Request) {
	body, err := ioutil.ReadAll(req.Body)
	if err != nil {
		panic(err)
	}
	log.Println(string(body))

	var wholeData Data
	err = json.Unmarshal(body, &wholeData)
	if err != nil {
		fmt.Println(err)
	}

	var objmap = wholeData.JsonData

	// get keys of structure
	var keys []string
	for k := range objmap[0] {
		keys = append(keys, k)
	}
	fmt.Println(keys)

	// WRITE
	f := excelize.NewFile()
	index, err := f.NewSheet(wholeData.NameSheet)
	if err != nil {
		fmt.Println(err)
	}

	// set names of excel's columns
	aplphabet := []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
	var colNames []string
	colNames = make([]string, len(keys), len(keys))
	for p := 0; p < len(keys); p++ {
		if p > 26 {
			var count float64 = float64(p / 26)
			first := math.Floor(count) - 1
			var last int = p - (26 * int(first))
			colNames[p] = aplphabet[int(first)] + aplphabet[last]
		} else {
			colNames[p] = aplphabet[p]
		}
	}

	for i := 0; i < len(objmap); i++ {
		for inde, v := range keys {
			column := colNames[inde] + strconv.Itoa(i+1)
			f.SetCellValue(wholeData.NameSheet, column, objmap[i][v])
		}
	}
	f.SetActiveSheet(index)
	// Save spreadsheet by the given path.
	nameFile := wholeData.NameFile + ".xlsx"
	if err := f.SaveAs(nameFile); err != nil {
		fmt.Println(err)
	}

	a, _ := os.Open(nameFile)
	reader := bufio.NewReader(a)
	content, _ := ioutil.ReadAll(reader)
	encoded := base64.StdEncoding.EncodeToString(content)

	rw.Header().Set("Content-Type", "text/json")
	rw.Write([]byte(encoded))

	defer a.Close()

	errF := os.Remove(nameFile)
	if errF != nil {
		fmt.Println(errF)
	}
}
