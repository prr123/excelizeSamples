
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

    useStr := "./createExcelWkb /excel=excelfile [/dbg]"
    helpStr := "program that creates an Excel file\n"

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



    excelfil := excelize.NewFile()
    defer func() {
        if err := excelfil.Close(); err != nil {
            log.Fatalf("error -- closing excel file: %v\n",err)
        }
    }()

    // Create a new sheet.
    index, err := excelfil.NewSheet("Sheet2")
    if err != nil {log.Fatalf("error -- creating a Sheet: %v", err)}
	log.Printf("index: %d\n",index)

    // Set value of a cell.
    excelfil.SetCellValue("Sheet2", "A1", "Hello Sheet1: " + xlsFilnam)
    excelfil.SetCellValue("Sheet1", "B1", "number")
    excelfil.SetCellValue("Sheet1", "B2", 100)
    // Set active sheet of the workbook.
    excelfil.SetActiveSheet(index)
    // Save spreadsheet by the given path.
    if err := excelfil.SaveAs(xlsFilnam); err != nil {
        log.Fatalf("error -- saving excel file: %v\n", err)
    }

}
