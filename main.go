// Author: 赩林, xilin0x7f@163.com

package main

import (
	"fmt"
	"os"
)

func main() {
	if len(os.Args) < 3 {
		fmt.Println("Usage: ConvertJsonsToTable <output.xlsx> <file.json> <file2.json> ...")
		os.Exit(1)
	}
	xlsxFile := os.Args[1]
	jsonFiles := os.Args[2:]
	_ = WriteJson2XLSX(jsonFiles, xlsxFile, "Sheet1", "A")
	fmt.Println("\n祝君好运！！！")
}
