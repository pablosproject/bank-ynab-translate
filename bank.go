package main

import (
	"encoding/csv"
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/tealeg/xlsx"
)

const (
	date    = "Date"
	inflow  = "Inflow"
	outflow = "Outflow"
	memo    = "Memo"
)

type bankParse struct {
	rowDeletionTop    int
	rowDeletionBottom int
	mapping           map[string]int
}

func main() {
	// Get the first xls file from current directory
	files, err := os.ReadDir(".")
	if err != nil {
		log.Fatalf("Error reading directory: %v", err)
	}
	var filePath string
	for _, file := range files {
		// Check if extension is xlsx
		if strings.Contains(file.Name(), ".xlsx") {
			filePath = file.Name()
		}
	}

	if filePath == "" {
		log.Fatalf("No Excel file found in the current directory")
	}

	finecoParse := bankParse{
		rowDeletionTop:    7,
		rowDeletionBottom: 0,
		mapping: map[string]int{
			date:    0,
			inflow:  1,
			outflow: 2,
			memo:    4,
		},
	}

	// Open the Excel file
	xlFile, err := xlsx.OpenFile(filePath)
	if err != nil {
		log.Fatalf("Error opening Excel file: %v", err)
	}

	sheet := xlFile.Sheets[0]
	cleanData(sheet, finecoParse)
	remappedData := mapToCSV(sheet, finecoParse)

	// Save the remapped data to a CSV file
	csvFilePath := "remapped-bank-statement.csv"
	err = saveToCSV(remappedData, csvFilePath)
	if err != nil {
		log.Fatalf("Error saving remapped data to CSV: %v", err)
	}

	fmt.Printf("Remapped data saved to: %s\n", csvFilePath)
}

func mapToCSV(sheet *xlsx.Sheet, parse bankParse) [][]string {
	var remappedData [][]string

	remappedData = append(remappedData, []string{"Date", "Inflow", "Outflow", "Memo", "Payee", "Category"})

	// [A:Date] [B:Inflow][C:Outflow][E:Memo]
	// Assuming the columns are A, B, C, and E
	log.Printf("Mapping: %v", len(sheet.Rows))

	for _, row := range sheet.Rows {
		date := row.Cells[parse.mapping[date]].String()
		inflow := row.Cells[parse.mapping[inflow]].String()
		outflow := row.Cells[parse.mapping[outflow]].String()
		memo := row.Cells[parse.mapping[memo]].String()

		log.Printf("Date: %s, Inflow: %s, Outflow: %s, Memo: %s", date, inflow, outflow, memo)
		if len(outflow) != 0 && outflow[0] == '-' {
			outflow = outflow[1:]
		}

		remappedRow := []string{date, inflow, outflow, memo, "", ""}
		remappedData = append(remappedData, remappedRow)
	}
	return remappedData
}

func cleanData(sheet *xlsx.Sheet, parse bankParse) {
	// Remove the specified number of rows from the beginning of the sheet
	sheet.Rows = sheet.Rows[parse.rowDeletionTop:]

	// Remove the specified number of rows from the end of the sheet
	sheet.Rows = sheet.Rows[:len(sheet.Rows)-parse.rowDeletionBottom]
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
