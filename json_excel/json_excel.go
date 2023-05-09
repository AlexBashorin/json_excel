package json_excel

import (
	"archive/zip"
	"bufio"
	"bytes"
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

	// write rows
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
	content, _ := ioutil.ReadAll(reader)

	// Compress file
	buf := new(bytes.Buffer)
	w := zip.NewWriter(buf)
	fArch, err := w.Create("archive.zip")
	if err != nil {
		log.Fatal(err)
	}
	_, err = fArch.Write(content)
	if err != nil {
		log.Fatal(err)
	}
	errW := w.Close()
	if errW != nil {
		log.Fatal(errW)
	}

	archive, _ := os.Open("archive.zip")
	readerArchive := bufio.NewReader(archive)
	contentArchive, _ := ioutil.ReadAll(readerArchive)

	encoded := base64.StdEncoding.EncodeToString(contentArchive)
	fmt.Println(contentArchive)

	rw.Header().Set("Content-Type", "text/json")
	rw.Write([]byte(encoded))

	defer a.Close()

	errF := os.Remove(nameFile)
	if errF != nil {
		fmt.Println(errF)
	}
}
