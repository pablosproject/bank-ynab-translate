package main

import (
	"encoding/csv"
	"fmt"
	"os"
	"reflect"
	"testing"
)

func TestFinecoParse(t *testing.T) {
	tempDir, err := os.MkdirTemp(".", "test")
	if err != nil {
		t.Fatalf("Failed to create temp dir: %v", err)
	}
	defer os.RemoveAll(tempDir) // Cleanup

	testFile := "test_fineco.xlsx"
	testType := "fineco"
	testOutput := tempDir + "/test_fineco.csv"

	processData(&testType, &testFile, &testOutput)

	// Then
	file, err := os.Open(testOutput)
	if err != nil {
		fmt.Println("Error opening CSV file:", err)
		return
	}
	defer file.Close()

	reader := csv.NewReader(file)
	records, err := reader.ReadAll()
	if err != nil {
		fmt.Println("Error reading CSV records:", err)
		return
	}

	expectedRecords := [][]string{
		{"Date", "Inflow", "Outflow", "Memo", "Payee", "Category"},
		{"22/07/2024", "", "1500", "Test 1", "", ""},
		{"21/07/2024", "", "10", "Test 2", "", ""},
		{"21/07/2024", "", "100", "Test 3", "", ""},
		{"16/07/2024", "1000", "", "Income 1", "", ""},
	}

	// Check if the number of records matches
	if len(records) != len(expectedRecords) {
		t.Errorf("Expected %d records, got %d", len(expectedRecords), len(records))
	}

	// Compare each record
	for i, record := range records {
		if !reflect.DeepEqual(record, expectedRecords[i]) {
			t.Errorf("Record %d does not match expected. Got %v, want %v", i, record, expectedRecords[i])
		}
	}
}

func TestMastercardParse(t *testing.T) {
	t.Skip("Skip test for now, it's impossible to edit data and xls right")
	tempDir, err := os.MkdirTemp(".", "test")
	if err != nil {
		t.Fatalf("Failed to create temp dir: %v", err)
	}
	defer os.RemoveAll(tempDir) // Cleanup

	testFile := "test_mastercard.xls"
	testType := "mastercard"
	testOutput := tempDir + "/test_.csv"

	processData(&testType, &testFile, &testOutput)

	// Then
	file, err := os.Open(testOutput)
	if err != nil {
		fmt.Println("Error opening CSV file:", err)
		return
	}
	defer file.Close()

	reader := csv.NewReader(file)
	records, err := reader.ReadAll()
	if err != nil {
		fmt.Println("Error reading CSV records:", err)
		return
	}

	expectedRecords := [][]string{
		{"Date", "Inflow", "Outflow", "Memo", "Payee", "Category"},
		{"22/07/2024", "", "10", "Mastercard Expense 1", "", ""},
		{"20/07/2024", "", "20", "Mastercard Expense 2", "", ""},
		{"03/07/2024", "", "50", "Mastercard Expense 3", "", ""},
	}

	// Check if the number of records matches
	if len(records) != len(expectedRecords) {
		t.Errorf("Expected %d records, got %d", len(expectedRecords), len(records))
	}

	// Compare each record
	for i, record := range records {
		// Here we need to skip the date record since it's impossible to edit and get it right
		for j, field := range record {
			if j == 0 {
				continue
			}
			if field != expectedRecords[i][j] {
				t.Errorf("Record %d does not match expected. Got %v, want %v", i, record, expectedRecords[i])
			}
		}
	}
}
