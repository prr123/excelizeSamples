package main

import (
    "fmt"
	"os"
	"log"
	"strconv"

    util "github.com/prr123/utility/utilLib"
    "github.com/xuri/excelize/v2"
)

func main() {

    numarg := len(os.Args)
    dbg := false
	dbgLev := 0
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

    dbgval, ok := flagMap["dbg"]
    if ok {
		dbg = true
		if dbgval.(string) != "none" {
			dbgLev, err = strconv.Atoi(dbgval.(string))
			if err != nil {
				log.Printf("error -- dbg level: %v\n", err)
				dbgLev = 0
			}
		}
	}
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

    log.Printf("debug: %t debug level: %d\n", dbg, dbgLev)
    log.Printf("Using excel file: %s\n", xlsFilnam)

    xlsfil, err := excelize.OpenFile(xlsFilnam)
    if err != nil {log.Fatalf("error -- open excel file: %v\n", err)}
    defer xlsfil.Close()

	// list all sheets
	sheetList := xlsfil.GetSheetList()
	sheet1 := sheetList[0]
	if dbgLev > 0 {
		fmt.Println("*********** sheets *************")
		for i:=0; i< len(sheetList); i++ {
			fmt.Printf("  sheet[%d]: %s\n", i+1,sheetList[i]) 
		}
		fmt.Println("**** testin cell A1 *****")
    	// Get value from cell by given worksheet name and cell reference.
    	cell, err := xlsfil.GetCellValue(sheet1, "A1")
    	if err != nil {log.Fatalf("error -- getcellvalue: %v\n", err)}
    	fmt.Printf("cell A1: %s\n", cell)
	}

    // Get all the rows in the Sheet1.
	// prr How do we know which rows these are?
    rows, err := xlsfil.GetRows(sheet1)
    if err != nil {log.Fatalf("error -- get rows: %v\n", err)}

	if dbgLev> 0 {
		fmt.Println("************** display rows for sheet 1 *******")
		fmt.Printf("rows: %v\n", rows)
	}

	rowNum := 0
    for _, row := range rows {
		rowNum ++
		if dbgLev> 0 {
			fmt.Printf("  row[%d]:|\t",rowNum)
        	for _, colCell := range row {
            	fmt.Printf("%v \t",colCell)
        	}
        	fmt.Println()
		}
    }

	// testing cell types
	if dbgLev > 1 {
		fmt.Printf(" **** testing row 3 ***** ")
		for icol:=1; icol< 14; icol ++ {
			cellAdr := string(icol+64) + "3"
			cellVal, err := xlsfil.GetCellValue(sheet1,cellAdr)
			if err !=nil {log.Printf("error getcellvalue %s: %v\n", cellAdr, err)}
			cellTyp, err := xlsfil.GetCellType(sheet1,cellAdr)
			if err !=nil {log.Printf("error getcelltype %s: %v\n", cellAdr, err)}
			fmt.Printf("cell[%s]: %-15s| %d\n",cellAdr, cellVal, cellTyp)
		}
	}
}
