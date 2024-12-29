package goexcel

import (
	"github.com/xuri/excelize/v2"
)

type Excel[T any] struct {
	File      *excelize.File
	SheetName string
}

func (e *Excel[T]) Save(filepath string) error {
	return e.File.SaveAs(filepath)
}

func CoordCol(n int) (col string) {
	for n > 0 {
		n--
		col = string(rune('A'+(n%26))) + col
		n /= 26
	}
	return col
}
