package main

import (
	"encoding/json"
	"fmt"
	"github.com/xuri/excelize/v2"
)

var file = "books.xlsx"

func openXlsxFile() [][]string {
	f, err := excelize.OpenFile(file)
	if err != nil {
		panic(err.Error())
	}

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		panic(err.Error())
	}

	return rows
}

func createKey(row []string) map[int]string {
	excelHeader := make(map[int]string)

	for i, col := range row {
		excelHeader[i] = col
	}

	return excelHeader
}

func main() {
	rows := openXlsxFile()
	excelHeader := createKey(rows[0])
	excelSheet := make([]interface{}, 0)

	for _, row := range rows[1:] {
		excelRow := make(map[string]interface{})
		for i, colCell := range row {
			excelRow[excelHeader[i]] = colCell
		}
		excelSheet = append(excelSheet, excelRow)
	}

	json, err := json.Marshal(excelSheet)
	if err != nil {
		panic(err.Error())
	}

	fmt.Println(string(json))
}
