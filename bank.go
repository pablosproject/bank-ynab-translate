package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/extrame/xls"
	"github.com/tealeg/xlsx"
)

const (
	date    = "Date"
	inflow  = "Inflow"
	outflow = "Outflow"
	memo    = "Memo"
	circuit = "Circuit"
)

type bankParse struct {
	name              string
	rowDeletionTop    int
	rowDeletionBottom int
	excludedCircuit   string
	mapping           map[string]int
}

func main() {
	fileName := flag.String("file", "", "The name of the file to parse")
	parseType := flag.String("type", "fineco", "The type of the file to parse")
	output := flag.String("output", "remapped-bank-statement.csv", "The name of the output file")

	flag.Parse()

	log.Printf("Parsing file: %s", *fileName)
	log.Printf("Parse type: %s", *parseType)
	if *fileName == "" {
		// Get the first xls file from current directory
		files, err := os.ReadDir(".")
		if err != nil {
			log.Fatalf("Error reading directory: %v", err)
		}
		var filePath string
		for _, file := range files {
			// Check if extension is xlsx
			if strings.Contains(file.Name(), ".xls") {
				filePath = file.Name()
			}
		}

		if filePath == "" {
			log.Fatalf("No Excel file found in the current directory")
		}

		fileName = &filePath
	}

	if *parseType != "fineco" && *parseType != "mastercard" {
		log.Fatalf("Unknown parse type: %s", *parseType)
	}

	processData(parseType, fileName, output)
}

func processData(parseType *string, fileName *string, output *string) {
	var parse bankParse
	switch *parseType {
	case "fineco":
		parse = bankParse{
			name:              "fineco",
			rowDeletionTop:    7,
			rowDeletionBottom: 0,
			mapping: map[string]int{
				date:    0,
				inflow:  1,
				outflow: 2,
				memo:    4,
			},
		}
	case "mastercard":
		parse = bankParse{
			name:              "mastercard",
			rowDeletionTop:    3,
			rowDeletionBottom: 3,
			excludedCircuit:   "BANCOMAT",
			mapping: map[string]int{
				date:    3,
				inflow:  0,
				outflow: 10,
				memo:    5,
				circuit: 8,
			},
		}
	default:
		log.Fatalf("Unknown parse type: %s", *parseType)
	}

	remappedData := mapToCSV(*fileName, parse)

	err := saveToCSV(remappedData, *output)
	if err != nil {
		log.Fatalf("Error saving remapped data to CSV: %v", err)
	}

	fmt.Printf("Remapped data saved to: %s\n", *output)
}

func mapToCSV(fileName string, parse bankParse) [][]string {
	var remappedData [][]string

	if parse.name == "fineco" {
		remappedData = mapXlsx(fileName, parse)
		return remappedData
	}

	if parse.name == "mastercard" {
		remappedData = mapXls(fileName, parse)
		return remappedData
	}

	return remappedData
}

func mapXls(fileName string, parse bankParse) [][]string {
	var remappedData [][]string
	remappedData = append(remappedData, []string{"Date", "Inflow", "Outflow", "Memo", "Payee", "Category"})

	workBook, err := xls.Open(fileName, "utf-8")
	if err != nil {
		log.Fatalf("Error opening Excel file: %v", err)
	}

	sheet := workBook.GetSheet(0)
	if sheet == nil {
		log.Fatal("Sheet not found.")
	}

	for i := 0; i <= int(sheet.MaxRow); i++ {
		if i < parse.rowDeletionTop || i >= int(sheet.MaxRow)-parse.rowDeletionBottom {
			continue
		}

		row := sheet.Row(i)
		rawDateString := row.Col(parse.mapping[date])
		rawDateInt, _ := strconv.Atoi(rawDateString)
		dateTime := excelDateToTime(rawDateInt)
		date := dateTime.Format("02/01/2006")
		inflow := row.Col(parse.mapping[inflow])
		outflow := row.Col(parse.mapping[outflow])
		memo := row.Col(parse.mapping[memo])
		circuit := row.Col(parse.mapping[circuit])

		if parse.excludedCircuit != "" && circuit == parse.excludedCircuit {
			continue
		}

		// Remove the unneded '-' sign
		if len(outflow) != 0 && outflow[0] == '-' {
			outflow = outflow[1:]
		} else if len(outflow) != 0 && outflow[0] != '-' {
			// We can sometimes have reimbursement on credit card
			inflow = outflow
			outflow = ""
		}

		log.Printf("Date: %s, Inflow: %s, Outflow: %s, Memo: %s, Circuit: %s", date, inflow, outflow, memo, circuit)
		if len(outflow) != 0 && outflow[0] == '-' {
			outflow = outflow[1:]
		}

		remappedRow := []string{date, inflow, outflow, memo, "", ""}
		remappedData = append(remappedData, remappedRow)
	}
	return remappedData
}

func mapXlsx(fileName string, parse bankParse) [][]string {
	var remappedData [][]string
	remappedData = append(remappedData, []string{"Date", "Inflow", "Outflow", "Memo", "Payee", "Category"})
	xlFile, err := xlsx.OpenFile(fileName)
	log.Printf("Opening file: %s", fileName)
	if err != nil {
		log.Fatalf("Error opening Excel file: %v", err)
	}

	sheet := xlFile.Sheets[0]
	for index, row := range sheet.Rows {
		if index < parse.rowDeletionTop || index >= len(sheet.Rows)-parse.rowDeletionBottom {
			continue
		}

		date := row.Cells[parse.mapping[date]].String()
		inflow := row.Cells[parse.mapping[inflow]].String()
		outflow := row.Cells[parse.mapping[outflow]].String()
		memo := row.Cells[parse.mapping[memo]].String()

		// Remove the unneded '-' sign
		if len(outflow) != 0 && outflow[0] == '-' {
			outflow = outflow[1:]
		}

		log.Printf("Date: %s, Inflow: %s, Outflow: %s, Memo: %s, circuit: %s", date, inflow, outflow, memo, circuit)

		remappedRow := []string{date, inflow, outflow, memo, "", ""}
		remappedData = append(remappedData, remappedRow)
	}
	return remappedData
}

func excelDateToTime(excelDate int) time.Time {
	// Excel uses December 30, 1899, as day 0, but Go's time package considers January 1, 1970, as day 0
	// We need to adjust for this difference
	baseDate := time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)
	days := time.Hour * 24 * time.Duration(excelDate)
	return baseDate.Add(days)
}

func saveToCSV(data [][]string, filePath string) error {
	// Create or overwrite the CSV file
	file, err := os.Create(filePath)
	if err != nil {
		return err
	}
	defer file.Close()

	// Create a CSV writer
	writer := csv.NewWriter(file)
	defer writer.Flush()

	// Write the data to the CSV file
	err = writer.WriteAll(data)
	if err != nil {
		return err
	}

	return nil
}
