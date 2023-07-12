package main

import (
	"fmt"
	"log"
	"strings"

	"github.com/tealeg/xlsx"
)

func main() {
	// Provide the names and paths of both zipped Excel files
	sourceFilePath := "cashcopy.xlsx"
	destinationFilePath := "cash.xlsx"

	// Provide the names of the sheets to be copied
	sheetNames := []string{"MayYTD", "May"}

	// Open the source Excel file
	sourceFile, err := xlsx.OpenFile(sourceFilePath)
	if err != nil {
		log.Fatalf("Failed to open source file: %s", err)
	}

	// Open the destination Excel file
	destinationFile, err := xlsx.OpenFile(destinationFilePath)
	if err != nil {
		log.Fatalf("Failed to open destination file: %s", err)
	}

	// Iterate over each sheet in the source file
	for _, sourceSheet := range sourceFile.Sheets {
		// Check if the current sheet is in the list of sheets to be copied
		if contains(sheetNames, sourceSheet.Name) {
			// Check if the sheet already exists in the destination file
			destinationSheet := destinationFile.Sheet[sourceSheet.Name]
			if destinationSheet == nil {
				// Create a new sheet in the destination file with the same name as the source sheet
				destinationSheet, err = destinationFile.AddSheet(sourceSheet.Name)
				if err != nil {
					log.Fatalf("Failed to create destination sheet: %s", err)
				}
			} else {
				// Clear existing data from the destination sheet
				destinationSheet.Rows = []*xlsx.Row{}
			}

			// Copy data from source sheet to destination sheet
			for _, sourceRow := range sourceSheet.Rows {
				destinationRow := destinationSheet.AddRow()

				for _, sourceCell := range sourceRow.Cells {
					destinationCell := destinationRow.AddCell()
					destinationCell.Value = sourceCell.Value
				}
			}
		}
	}

	// Save the modified destination Excel file
	if err := destinationFile.Save(destinationFilePath); err != nil {
		log.Fatalf("Failed to save destination file: %s", err)
	}

	fmt.Println("Data copied successfully!")
}

// Helper function to check if a string is present in a string slice
func contains(slice []string, str string) bool {
	for _, s := range slice {
		if strings.EqualFold(s, str) {
			return true
		}
	}
	return false
}
