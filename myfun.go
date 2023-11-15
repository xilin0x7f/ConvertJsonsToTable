// Author: 赩林, xilin0x7f@163.com

package main

import (
	"encoding/json"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
	"os"
	"sort"
)

func WriteJson2XLSX(filesName []string, dstFileName, sheetName, start string) error {
	keysMap := make(map[string]int)
	for _, fileName := range filesName {
		file, _ := os.Open(fileName)
		var jsonData map[string]interface{}
		reader := io.Reader(file)
		decoder := json.NewDecoder(reader)
		_ = decoder.Decode(&jsonData)
		_ = file.Close()
		for key := range jsonData {
			keysMap[key]++
		}
	}
	var keys []string
	for key := range keysMap {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	resMap := make(map[string][]interface{})
	for _, fileName := range filesName {
		file, _ := os.Open(fileName)
		var jsonData map[string]interface{}
		reader := io.Reader(file)
		decoder := json.NewDecoder(reader)
		_ = decoder.Decode(&jsonData)
		_ = file.Close()
		for _, key := range keys {
			resMap[key] = append(resMap[key], jsonData[key])
		}
	}
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
