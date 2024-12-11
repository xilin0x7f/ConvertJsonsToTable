// Author: 赩林, xilin0x7f@163.com

package main

import (
	"encoding/json"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"sort"
)

func flattenJSON(data map[string]interface{}, prefix string, flatMap map[string]interface{}) {
	for key, value := range data {
		fullKey := key
		if prefix != "" {
			fullKey = prefix + "." + key
		}
		switch v := value.(type) {
		case map[string]interface{}:
			flattenJSON(v, fullKey, flatMap)
		default:
			flatMap[fullKey] = v
		}
	}
}

func WriteJson2XLSX(filesName []string, dstFileName, sheetName, start string) error {
	keysMap := make(map[string]int)

	// First pass: collect keys
	for _, fileName := range filesName {
		file, _ := os.Open(fileName)
		var jsonData map[string]interface{}
		decoder := json.NewDecoder(file)
		_ = decoder.Decode(&jsonData)
		_ = file.Close()

		flatData := make(map[string]interface{})
		flattenJSON(jsonData, "", flatData)

		for key := range flatData {
			keysMap[key]++
		}
	}

	// Extract and sort keys
	var keys []string
	for key := range keysMap {
		keys = append(keys, key)
	}
	sort.Strings(keys)

	// Second pass: collect data
	resMap := make(map[string][]interface{})
	for _, fileName := range filesName {
		file, _ := os.Open(fileName)
		var jsonData map[string]interface{}
		decoder := json.NewDecoder(file)
		_ = decoder.Decode(&jsonData)
		_ = file.Close()

		flatData := make(map[string]interface{})
		flattenJSON(jsonData, "", flatData)

		for _, key := range keys {
			resMap[key] = append(resMap[key], flatData[key])
		}
	}

	// Build result matrix
	res := make([][]interface{}, len(resMap[keys[0]])+1)
	for idx := range res {
		res[idx] = make([]interface{}, len(keys)+1)
	}
	res[0][0] = "file"
	for idx, key := range keys {
		res[0][idx+1] = key
	}
	for rowIdx := range resMap[keys[0]] {
		res[rowIdx+1][0] = filesName[rowIdx]
		for colIdx := range keys {
			res[rowIdx+1][colIdx+1] = resMap[keys[colIdx]][rowIdx]
		}
	}

	// Write to Excel
	err := Write2XLSX(dstFileName, sheetName, start, res)
	return err
}
func Write2XLSX(fileName, sheetName, start string, data [][]interface{}) error {
	f := excelize.NewFile()
	index, err := f.NewSheet(sheetName)
	if err != nil {
		return err
	}
	f.SetActiveSheet(index)
	for idx := range data {
		if err := f.SetSheetRow(sheetName, fmt.Sprint(start, idx+1), &data[idx]); err != nil {
			return err
		}
	}
	if err := f.SaveAs(fileName); err != nil {
		return err
	}
	return nil
}
