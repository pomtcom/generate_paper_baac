package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"strconv"
)

type record struct {
	number     int
	branchName string
	district   string
	province   string
	quantity   int
}

func main() {
	//f, err := excelize.OpenFile("./info.xlsx")
	//if err != nil {
	//	fmt.Println(err)
	//	return
	//}
	//// Get value from cell by given worksheet name and axis.
	//cell, err := f.GetCellValue("record", "B2")
	//if err != nil {
	//	fmt.Println(err)
	//	return
	//}
	//fmt.Println(cell)
	fmt.Println("Generate print paper baac program is starting")

	data_record, err := readRecord()
	if err != nil {
		fmt.Println(err)
		return
	}

	//fmt.Println("data record is ", data_record)

	writeOutput(data_record)
	fmt.Println("Please any key to exit")
	fmt.Scanln() // wait for Enter Key

}

func writeOutput(data_record []record){


	f, err := excelize.OpenFile("./template/template.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	index := 1
	for _, item := range data_record {
		// index is the index where we are
		// element is the element from someSlice for where we are
		fmt.Println("Writing record of ", item.number)
		for i := 1; i <= item.quantity; i++ {

			branchNameIndex := "E" + strconv.Itoa(index+1)
			f.SetCellValue("Sheet1", branchNameIndex, item.branchName)

			districtIndex := "B" + strconv.Itoa(index+2)
			f.SetCellValue("Sheet1", districtIndex, item.district)

			provinceIndex := "B" + strconv.Itoa(index+3)
			f.SetCellValue("Sheet1", provinceIndex, item.province)

			quantityIndexRunning := "C" + strconv.Itoa(index+4)
			f.SetCellValue("Sheet1", quantityIndexRunning, i)

			quantityIndexLast := "E" + strconv.Itoa(index+4)
			f.SetCellValue("Sheet1", quantityIndexLast, item.quantity)
			index = index + 5
		}
	}


	// Set active sheet of the workbook.
	//f.SetActiveSheet(index)
	// Save xlsx file by the given path.
	err_savefile := f.SaveAs("./output/output.xlsx")
	if err_savefile != nil {
		fmt.Println(err_savefile)
	}
	fmt.Println("save file is completed, the last row of record is ", index-1)
}

func readRecord() ([]record, error) {
	fmt.Println("read input from source")
	f, err := excelize.OpenFile("./source_record/source.xlsx")
	if err != nil {
		fmt.Println(err)
		return nil, err
	}

	// Get all the rows in the record sheet.
	rows, err := f.GetRows("Sheet1")
	recordOfInfo := []record{}
	rowNumber := 0

	for _, row := range rows {
		rowNumber++
		orderOfCell := 0
		tempRecord := record{}
		fmt.Println("processing row number ", rowNumber)
		for _, colCell := range row {
			orderOfCell = orderOfCell + 1
			if orderOfCell == 1 {
				if _, err := strconv.Atoi(colCell); err != nil {
					//check if the first column is not a number
					//fmt.Printf("%q is not a number.\n", colCell)
					fmt.Println("row number ", rowNumber, " is not data row -> skip this row")
					break
				}
				tempRecord.number, err = strconv.Atoi(colCell)
				if err != nil {
					fmt.Println(colCell, " is not a number please check")
				}
			} else if orderOfCell == 2 {
				tempRecord.branchName = colCell
			} else if orderOfCell == 3 {
				tempRecord.district = colCell
			} else if orderOfCell == 4 {
				tempRecord.province = colCell
			} else if orderOfCell == 5 {
				tempRecord.quantity, err = strconv.Atoi(colCell)
				if err != nil {
					fmt.Println(colCell, " is not a number please check")
				}
			}
			//fmt.Print(colCell, "\t")
		}
		if tempRecord.number != 0 && tempRecord.branchName != "" && tempRecord.district != "" &&
			tempRecord.province != ""  && tempRecord.quantity != 0 {
			recordOfInfo = append(recordOfInfo, tempRecord)
		}else{
			//fmt.Println("record of ", tempRecord.number, " is not valid -> ignore this record")
		}

		//fmt.Println()
	}

	//fmt.Println("recordOfInfo is ", recordOfInfo)
	return recordOfInfo, nil
}
