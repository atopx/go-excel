package goexcel

import (
	"fmt"
	"reflect"

	"github.com/xuri/excelize/v2"
)

func New[T any](sheetName string) *Excel[T] {
	f := excelize.NewFile()
	_, _ = f.NewSheet(sheetName)
	if defaultSheet := "Sheet1"; len(sheetName) > 0 && sheetName != defaultSheet {
		_ = f.SetSheetName("Sheet1", sheetName)
	} else {
		sheetName = defaultSheet
	}
	return &Excel[T]{
		File:      f,
		SheetName: sheetName,
	}
}

func (e *Excel[T]) Write(rows []*T, styles ...*excelize.Style) error {
	if len(rows) == 0 {
		return nil
	}

	typeRef := reflect.TypeOf(rows[0]).Elem()
	colNum := typeRef.NumField()

	if colNum == 0 {
		return StructEmptyError
	}

	// set styles
	switch len(styles) {
	case 1:
		style, _ := e.File.NewStyle(styles[0])
		endCell, _ := excelize.CoordinatesToCellName(colNum, len(rows)+1)
		_ = e.File.SetCellStyle(e.SheetName, "A1", endCell, style)
	case 2:
		headStyle, _ := e.File.NewStyle(styles[0])
		endCell, _ := excelize.CoordinatesToCellName(colNum, 1)
		_ = e.File.SetCellStyle(e.SheetName, "A1", endCell, headStyle)
		valueStyle, _ := e.File.NewStyle(styles[1])
		endCell, _ = excelize.CoordinatesToCellName(colNum, len(rows)+1)
		_ = e.File.SetCellStyle(e.SheetName, "A2", endCell, valueStyle)
	}

	tag, err := parseTag(typeRef.Field(0))
	if err != nil {
		return err
	}

	// 写入 headlines
	for colIdx := 0; colIdx < colNum; colIdx++ {
		tag, err = parseTag(typeRef.Field(colIdx))
		if err != nil {
			return fmt.Errorf("parse tag failed: %w", err)
		}
		cell, _ := excelize.CoordinatesToCellName(colIdx+1, 1)
		if err = e.File.SetCellStr(e.SheetName, cell, tag.Head); err != nil {
			return fmt.Errorf("writing headline failed: %w", err)
		}

		if tag.Width > 0 {
			col := CoordCol(colIdx + 1)
			fmt.Println(col, tag.Width)
			err = e.File.SetColWidth(e.SheetName, col, col, tag.Width)
			if err != nil {
				return fmt.Errorf("set col width failed: %w", err)
			}
		}
	}

	// 写入values
	for rowIdx, row := range rows {
		valueRef := reflect.ValueOf(row).Elem()
		for colIdx := 0; colIdx < colNum; colIdx++ {
			cell, _ := excelize.CoordinatesToCellName(colIdx+1, rowIdx+2)
			_ = e.File.SetCellValue(e.SheetName, cell, valueRef.Field(colIdx).String())
		}
	}
	return nil
}
