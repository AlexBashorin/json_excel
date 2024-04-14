package json_excel

import (
	"bufio"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"io"
	"math"
	"net/http"
	"os"
	"sort"
	"strconv"

	"github.com/xuri/excelize/v2"
)

type Data struct {
	NameSheet string `json:"nameSheet"`
	NameFile  string `json:"nameFile"`
	// JsonData  []map[string]interface{} `json:"data"`
	JsonData []map[string]*json.RawMessage `json:"data"`
}

func WriteExcel(rw http.ResponseWriter, req *http.Request) {
	body, err := io.ReadAll(req.Body)
	if err != nil {
		panic(err)
	}

	var wholeData Data
	err = json.Unmarshal(body, &wholeData)
	if err != nil {
		fmt.Println(err)
	}

	var objmap = wholeData.JsonData

	// get keys
	// var keys []string
	// for i, _ := range objmap[0] {
	// 	keys = append(keys, i)
	// }
	// uniqueKeys := make(map[string]struct{})
	var keys []string
	// keys := make([]string, 0, len(objmap[0]))

	for k, _ := range objmap[0] {
		keys = append(keys, k)
	}

	sort.Slice(keys, func(i, j int) bool {
		return keys[i] < keys[j]
	})
	fmt.Println(keys)

	// WRITE
	f := excelize.NewFile()
	index, err := f.NewSheet(wholeData.NameSheet)
	if err != nil {
		fmt.Println(err)
	}

	// if columns > 26 (eng alphabet)
	aplphabet := []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
	colNames := make([]string, len(keys))
	for p := 0; p < len(keys); p++ {
		if p < 26 {
			colNames[p] = aplphabet[p]
		} else {
			var count float64 = float64(p / 26)
			first := math.Floor(count) - 1
			qua := 26 * count
			pp := p - int(qua)
			colNames[p] = aplphabet[int(first)] + aplphabet[pp]
		}
	}

	// Write rows
	for i := 0; i < len(objmap); i++ {
		for k := 0; k < len(keys); k++ {
			column := colNames[k] + strconv.Itoa(i+1)
			f.SetCellValue(wholeData.NameSheet, column, objmap[i][keys[k]])
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
	content, _ := io.ReadAll(reader)
	encoded := base64.StdEncoding.EncodeToString(content)
	defer a.Close()

	rw.Header().Set("Content-Type", "text/json")
	rw.Write([]byte(encoded))

	// Delete created file
	// os.Remove(nameFile)
}
