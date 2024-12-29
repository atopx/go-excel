package goexcel

import (
	"fmt"
	"io"
	"reflect"

	"github.com/xuri/excelize/v2"
)

func Open[T any](reader io.Reader, sheetName string) (*Excel[T], error) {
	f, err := excelize.OpenReader(reader)
	if err != nil {
		return nil, fmt.Errorf("open excel: %w", err)
	}
	return &Excel[T]{
		File:      f,
		SheetName: sheetName,
	}, nil
}

func (e *Excel[T]) Read() ([]*T, error) {

	ref := reflect.TypeOf(new(T)).Elem()
	colNum := ref.NumField()

	rows, err := e.File.GetRows(e.SheetName)
	if err != nil {
		return nil, fmt.Errorf("get rows failed: %w", err)
	}

	if len(rows) == 0 {
		return nil, RecordNotFound
	}

	for i := 0; i < colNum; i++ {
		tag, err := parseTag(ref.Field(i))
		if err != nil {
			return nil, fmt.Errorf("parse tag failed: %w", err)
		}

		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		if rows[0][i] != tag.Head {
			return nil, fmt.Errorf("head inconsistent, cell=%s, expected=%s, actual=%s",
				cell, tag.Head, rows[0][i])
		}
	}

	data := make([]*T, len(rows)-1)
	for rowIdx, row := range rows[1:] {
		obj := new(T)
		valueRef := reflect.ValueOf(obj).Elem()
		for colIdx := 0; colIdx < colNum; colIdx++ {
			f := valueRef.Field(colIdx)
			if f.CanSet() {
				f.SetString(row[colIdx])
			}
		}
		data[rowIdx] = obj
	}
	return data, nil
}
