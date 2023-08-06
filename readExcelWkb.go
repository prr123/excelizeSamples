package main

import (
    "fmt"
	"os"
	"log"

    util "github.com/prr123/utility/utilLib"
    "github.com/xuri/excelize/v2"
)

func main() {

    numarg := len(os.Args)
    dbg := false
    flags:=[]string{"dbg","excel"}

    // default file
    xlsFilnam := ""

    useStr := "./readExcelWkb /excel=excelfile [/dbg]"
    helpStr := "program that reads an Excel file\n"

    if numarg > len(flags) +1 {
        fmt.Println("too many arguments in cl!")
        fmt.Println("usage: ./template [/flag1=] [/flag2]\n", useStr)
        os.Exit(-1)
    }

    if numarg > 1 && os.Args[1] == "help" {
        fmt.Printf("help: %s\n", helpStr)
        fmt.Printf("usage is: %s\n", useStr)
        os.Exit(1)
    }

    flagMap, err := util.ParseFlags(os.Args, flags)
    if err != nil {log.Fatalf("util.ParseFlags: %v\n", err)}

    _, ok := flagMap["dbg"]
    if ok {dbg = true}
    if dbg {
        fmt.Printf("dbg -- flag list:\n")
        for k, v :=range flagMap {
            fmt.Printf("  flag: /%s value: %s\n", k, v)
        }
    }

    xlsval, ok := flagMap["excel"]
    if !ok {
        log.Fatalf("error -- no excel flag provided!")
    } else {
        if xlsval.(string) == "none" {log.Fatalf("error: no excel file provided with /excel flag!")}
        xlsFilnam = xlsval.(string)
//        if dbg {log.Printf("excel file: %s\n", xlsFilnam)}
    }

    log.Printf("debug: %t\n", dbg)
    log.Printf("Using excel file: %s\n", xlsFilnam)

    xlsfil, err := excelize.OpenFile(xlsFilnam)
    if err != nil {log.Fatalf("error -- open excel file: %v\n", err)}
    defer xlsfil.Close()

	// list all sheets
	sheetList := xlsfil.GetSheetList()
	for i:=0; i< len(sheetList); i++ {
		fmt.Printf("  sheet[%d]: %s\n", i+1,sheetList[i]) 
	}

	sheet1 := sheetList[0]
    // Get value from cell by given worksheet name and cell reference.
    cell, err := xlsfil.GetCellValue(sheet1, "A1")
    if err != nil {log.Fatalf("error -- getcellvalue: %v\n", err)}
    fmt.Printf("printing cell A1: %s\n", cell)


    // Get all the rows in the Sheet1.
	// prr How do we know which rows these are?
    rows, err := xlsfil.GetRows(sheet1)
    if err != nil {log.Fatalf("error -- get rows: %v\n", err)}
	fmt.Printf("rows: %v\n", rows)

	rowNum := 0
    for _, row := range rows {
		rowNum ++
		fmt.Printf("  row[%d]: %v|\t",rowNum, row)
        for _, colCell := range row {
            fmt.Printf("%v \t",colCell)
        }
        fmt.Println()
    }

	// testing cell types
	valStr, err := xlsfil.GetCellValue(sheet1,"B3")
	if err !=nil {log.Printf("error getcellvalue B3: %v\n", err)}
	fmt.Printf("cell value B3: %s\n",valStr)
	celltyp, err := xlsfil.GetCellType(sheet1,"B3")
	if err !=nil {log.Printf("error getcelltype B3: %v\n", err)}
	fmt.Printf("cell type B3: %v\n",celltyp)

}
