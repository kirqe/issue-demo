package main

import (
	"fmt"
	_ "image/png"

	"github.com/xuri/excelize/v2"
)

func FieldsSection(sectionName string) [][]interface{} {
	fields := [][]interface{}{
		{nil, sectionName, nil},
		{nil, nil, nil},
		{nil, nil, nil},
		{nil, nil, nil},
		{nil, nil, nil},
		{nil, nil, nil},
		{nil, nil, nil},
		{nil, nil, nil},
		{nil, nil, nil},
		{nil, nil, nil},
		{nil, nil, nil},
		{nil, nil, nil},
	}

	return fields
}

func BuildFields() [][]interface{} {
	fields := [][]interface{}{{nil, nil, "---"}}
	fields = append(fields, FieldsSection("Section 1")...)
	fields = append(fields, [][]interface{}{{nil, nil, "---"}}...)
	fields = append(fields, FieldsSection("Section 2")...)
	fields = append(fields, [][]interface{}{{nil, nil, "---"}}...)

	return fields
}

func main() {
	f := excelize.NewFile()

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	fields := BuildFields()
	for i, row := range fields {
		cell, err := excelize.CoordinatesToCellName(1, i+1)

		if err != nil {
			fmt.Println(err)
			return
		}

		defaultRowStyle, _ := f.NewStyle(&excelize.Style{
			Alignment: &excelize.Alignment{
				Horizontal: "center",
				Vertical:   "center",
				WrapText:   true,
			},
		})

		defaultRowStyleL, _ := f.NewStyle(&excelize.Style{
			Alignment: &excelize.Alignment{
				Horizontal: "left",
				Vertical:   "center",
				WrapText:   true,
			},
		})

		defaultRowStyleR, _ := f.NewStyle(&excelize.Style{
			Alignment: &excelize.Alignment{
				Horizontal: "right",
				Vertical:   "center",
				WrapText:   true,
			},
		})

		hintStyle, _ := f.NewStyle(&excelize.Style{
			Alignment: &excelize.Alignment{
				Horizontal: "center",
				Vertical:   "top",
				WrapText:   true,
			},
		})

		f.SetRowStyle("Sheet1", 1, 27, defaultRowStyle)

		f.SetCellStyle("Sheet1", "C1", "C1", defaultRowStyleL)
		f.SetCellStyle("Sheet1", "D1", "D1", defaultRowStyleR)
		f.SetCellStyle("Sheet1", "C3", "C3", hintStyle)

		f.SetCellStyle("Sheet1", "C4", "C4", defaultRowStyle)
		f.SetCellStyle("Sheet1", "D4", "D4", defaultRowStyle)
		f.SetCellStyle("Sheet1", "C5", "C5", hintStyle)
		f.SetCellStyle("Sheet1", "D5", "D5", hintStyle)

		f.SetCellStyle("Sheet1", "C6", "C6", defaultRowStyle)
		f.SetCellStyle("Sheet1", "C7", "C7", hintStyle)

		f.SetCellStyle("Sheet1", "D9", "C9", hintStyle)
		f.SetCellStyle("Sheet1", "D11", "C11", hintStyle)

		f.SetCellStyle("Sheet1", "C14", "C14", defaultRowStyleL)
		f.SetCellStyle("Sheet1", "D14", "D14", defaultRowStyleR)

		f.MergeCell("Sheet1", "C2", "D2")
		f.MergeCell("Sheet1", "C3", "D3")
		f.MergeCell("Sheet1", "C6", "D6")
		f.MergeCell("Sheet1", "C7", "D7")
		f.MergeCell("Sheet1", "C8", "D8")
		f.MergeCell("Sheet1", "C9", "D9")
		f.MergeCell("Sheet1", "C10", "D10")
		f.MergeCell("Sheet1", "C11", "D11")
		f.MergeCell("Sheet1", "C12", "D12")

		// Second part

		f.SetCellStyle("Sheet1", "C14", "C14", defaultRowStyleL)
		f.SetCellStyle("Sheet1", "D14", "D14", defaultRowStyleR)
		f.SetCellStyle("Sheet1", "C16", "C16", hintStyle)

		f.SetCellStyle("Sheet1", "C17", "C17", defaultRowStyle)
		f.SetCellStyle("Sheet1", "D17", "D17", defaultRowStyle)
		f.SetCellStyle("Sheet1", "C18", "C18", hintStyle)
		f.SetCellStyle("Sheet1", "D18", "D18", hintStyle)

		f.SetCellStyle("Sheet1", "C19", "C19", defaultRowStyle)
		f.SetCellStyle("Sheet1", "C20", "C20", hintStyle)

		f.SetCellStyle("Sheet1", "D22", "C22", hintStyle)
		f.SetCellStyle("Sheet1", "D24", "C24", hintStyle)

		f.SetCellStyle("Sheet1", "C27", "C27", defaultRowStyleL)
		f.SetCellStyle("Sheet1", "D27", "D27", defaultRowStyleR)

		f.MergeCell("Sheet1", "C15", "D15")
		f.MergeCell("Sheet1", "C16", "D16")
		f.MergeCell("Sheet1", "C19", "D19")
		f.MergeCell("Sheet1", "C20", "D20")
		f.MergeCell("Sheet1", "C21", "D21")
		f.MergeCell("Sheet1", "C22", "D22")
		f.MergeCell("Sheet1", "C23", "D23")
		f.MergeCell("Sheet1", "C24", "D24")
		f.MergeCell("Sheet1", "C25", "D25")

		// Cols width
		f.SetColWidth("Sheet1", "A", "A", 5)
		f.SetColWidth("Sheet1", "B", "B", 40)
		f.SetColWidth("Sheet1", "C", "C", 45)
		f.SetColWidth("Sheet1", "D", "D", 45)

		f.SetRowHeight("Sheet1", i+1, 30) // height of each row
		f.SetSheetRow("Sheet1", cell, &row)

		if err := f.AddPicture("Sheet1", "A12", "test.png",
			&excelize.GraphicOptions{ScaleX: 0.2, ScaleY: 0.2}); err != nil {
			fmt.Println(err)
		}
	}

	// Save spreadsheet by the given path.

	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}
