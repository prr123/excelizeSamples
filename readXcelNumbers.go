// readXcelNumbers
// input file: excelTstValues
// Author: prr, Azul Software
// Date: 7 Aug 2023
// copyright 2023 prr, azul software
//
// program that reads numbers if different formats from xcel and converts the numbers into type formats (int, float, etc)
//

package main

import (
    "fmt"
	"os"
	"log"
	"strconv"
	"bytes"
	"math"

    util "github.com/prr123/utility/utilLib"
    "github.com/xuri/excelize/v2"
)

func main() {

    numarg := len(os.Args)
    dbg := false
    flags:=[]string{"dbg","excel"}

    // default file
    xlsFilnam := ""

    useStr := "./readXcelNumbers /excel=excelfile [/dbg]"
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

/*
	// list all sheets
	sheetList := xlsfil.GetSheetList()
	for i:=0; i< len(sheetList); i++ {
		fmt.Printf("  sheet[%d]: %s\n", i+1,sheetList[i]) 
	}
*/
	sheet1 := "numbers"
    // Get value from cell by given worksheet name and cell reference.
	colB := 'B'
	colBval := int(colB) - 64
	fmt.Printf("col B: %d\n",colBval)

//	os.Exit(1)

	colEnd := int('N') - 64

	for irow :=3; irow< 4; irow++ {
		cellRowStr := fmt.Sprintf("%d",irow)
		for icol :=2; icol<colEnd; icol++ {
			cellColStr := string(icol+64)
			cellAdrStr := cellColStr + cellRowStr
		fmt.Println("*********************************************")
			fmt.Printf("cellStr: %s\n", cellAdrStr)
    		cellStr, err := xlsfil.GetCellValue(sheet1, cellAdrStr)
			if err != nil {log.Fatalf("error -- getcellvalue: %v\n", err)}
    		fmt.Printf("cell[%d,%d]: %s\n", irow, icol, cellStr)
			cellval, err := ParseCellStr(cellStr)
			if err != nil {fmt.Printf("cell parse error: %v\n",err)}
			PrintCell(cellval)
		}
	}


}

func ParseCellStr(str string)(val interface{}, err error) {
	fmt.Printf("*** parseCell: %s ***\n", str)
	l := len(str)
	endStr := l-1
	byt := []byte(str)

	if l ==0 {
		val = str
		return val, nil
	}

	if l ==1 {
		if util.IsNumeric(byt[0]) {
			val = int(byt[0]) - 47
			return val, nil
		} else {
			val = str
			return val, nil
		}
	}

	// first check for chars
	// test the first and the last char
	percent := false
	if byt[endStr] == '%' {
		percent = true
		endStr += -1
	}

	firstNum := false

	if util.IsNumeric(byt[0]) {
		firstNum = true
	} else {
		if (byt[0] == '-') && util.IsNumeric(byt[1]) {firstNum = true}
	}

	num := false
	if firstNum && util.IsNumeric(byt[endStr]) {num = true}

	if !num {
		val = str
		return val, nil
	}

	// we have a number
	// check for float
	floatNum := false

	// replace European with US notation
	idx := bytes.IndexByte(byt, ',')
	if idx>0 {byt[idx] = '.'}

	idx = bytes.IndexByte(byt, '.')
	if idx>0 {floatNum = true}

	// check for exp notation
	expPos := bytes.IndexAny(byt,"eE")
	if expPos> 0 {
		if percent {
            val = str
            return val, fmt.Errorf("impossible percent & exponential")
		}
		baseStr := str[0:expPos]
		baseVal, err := strconv.ParseFloat(baseStr, 64)
		if err != nil {
            val = str
            return val, fmt.Errorf("strconv base: %v", err)
		}

		expStr := str[(expPos + 1):]
		expVal, err := strconv.ParseInt(expStr, 10, 64)
		if err != nil {
            val = str
            return val, fmt.Errorf("strconv exp: %v", err)
		}

		val = baseVal * math.Pow10(int(expVal))
		return val, nil
	}

	// int
	if !floatNum && !percent {
		valInt, err := strconv.Atoi(str)
		if err != nil {
			val = str
			return val, nil
		}
		val = valInt
		return val, nil
	}

	// float
	valFloat, err := strconv.ParseFloat(string(byt[:endStr+1]), 64)
	if err != nil {
		val = str
		return val, nil
	}

	if percent {valFloat = valFloat/100.0}
	val = valFloat
	return val, nil
}

func parseCellStrOld(str string)(val interface{}, err error) {

	fmt.Printf("*** parseCell: %s ***\n", str)
	l := len(str)
	state:=0
	typ := 's'
	num:=-1
	bstr := []byte(str)
//	stnum := -1
	endnum := -1
	floatNum := 0.0
	fract := 0.0
	fmt.Printf("  %d %v\n",l, bstr)
	for i:= l-1; i>-1; i-- {
		fmt.Printf(" %d: %q\n", i, bstr[i])
		switch bstr[i] {
		case ' ':
			if state != 0 {
				val = str
				return val, fmt.Errorf("invalid wsp!")
			}
		case '0','1','2','3','4','5','6','7','8','9':
			if state == 0 {
				state = 1
				endnum = i
				if typ == 's' {typ ='n'}
			}
			if state == 1 {
				state = 2
			}
			if state == 3 {
				state = 4
				endnum = i
			}
			if state == 4 {
				state = 5
			}
			num, err = strconv.Atoi(string(bstr[i:endnum+1]))
			if err!=nil {val = str; return val, nil}
			fmt.Printf(" num state %d: %d: num: %d \n", state, i, num)
			if typ == 'f' {
				floatNum = float64(num) + fract
			}
		case '.':
			if typ == 'n' {
				typ = 'f'
				fracnum, err := strconv.Atoi(string(bstr[i+1:endnum+1]))
				if err!=nil {val = str; return val, fmt.Errorf("fract conv: %v", err)}
			fmt.Printf(" fract state %d: %d: frac: %d \n", state, i, fracnum)
				div := math.Pow(10, float64(endnum - i))
				fract = float64(fracnum)/div
			fmt.Printf(" float: %6.3f\n", fract)
				state = 3
			}

		case ',':

		case '+':
// exponent
		case '%':
			if state == 0 {
				state = 1
				typ = '%'
				endnum = i
			} else {
				val = str
				return val, fmt.Errorf("% in pos %d\n", i)
			}
		default:
			val = str
			return val, nil
		}
	}


	switch typ {
	case '%':
		floatNum = float64(num)/100.0
		val = floatNum
	case 'n':
		val = num
	case 'f':
		val = floatNum

	default:
		val = str
	}

	fmt.Printf("*** end parseCell ***\n")

	return val, nil
}

func PrintCell(cell interface{}) {

	switch cell.(type) {
	case int:
		fmt.Printf("-- cell int: %d\n", cell.(int))

	case float64:
		fmt.Printf("-- cell float: %10.2f\n", cell.(float64))

	default:
		fmt.Printf("-- cell string: %s\n", cell.(string))
	}
}
