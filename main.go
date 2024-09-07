package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	// Folder containing images
	imageFolder := "./images"

	// Create a new Excel file or open an existing one
	f := excelize.NewFile()

	// Read all image files from the folder
	files, err := os.ReadDir(imageFolder)
	if err != nil {
		log.Fatalf("Failed to read image folder: %s", err)
	}

	// Iterate over image files and insert them into cells
	row := 1 // Starting row to insert images
	for _, file := range files {
		if file.IsDir() {
			continue // Skip directories
		}

		// Get the file extension and normalize it to lowercase
		filePath := filepath.Join(imageFolder, file.Name())
		fileExt := strings.ToLower(filepath.Ext(filePath))

		// Check if the file is a supported image format
		if fileExt == ".jpg" || fileExt == ".jpeg" || fileExt == ".png" {
			cellName := fmt.Sprintf("A%d", row)

			// Boolean values for LockAspectRatio and Locked fields
			lockAspectRatio := true
			locked := false

			// Insert image into the specified cell with GraphicOptions
			err := f.AddPicture("Sheet1", cellName, filePath, &excelize.GraphicOptions{
				OffsetX:         10,
				OffsetY:         10,
				LockAspectRatio: lockAspectRatio,
				Locked:          &locked,
			})
			if err != nil {
				log.Printf("Failed to insert image: %s, error: %s\n", file.Name(), err)
				continue
			}
			fmt.Printf("Inserted %s into cell %s\n", file.Name(), cellName)

			row++ // Move to the next row
		} else {
			fmt.Printf("Skipping unsupported file format: %s\n", file.Name())
		}
	}

	// Save the Excel file
	if err := f.SaveAs("images_in_excel.xlsx"); err != nil {
		log.Fatalf("Failed to save Excel file: %s", err)
	}

	fmt.Println("Excel file created with images.")
}
