package goexcel

import "github.com/xuri/excelize/v2"

func (e *Excel[T]) DefaultStyles() []*excelize.Style {
	border := []excelize.Border{
		{Type: "left", Color: "000000", Style: 1},
		{Type: "top", Color: "000000", Style: 1},
		{Type: "right", Color: "000000", Style: 1},
		{Type: "bottom", Color: "000000", Style: 1},
	}
	return []*excelize.Style{
		{
			Font:      &excelize.Font{Bold: true, Size: 10, Family: "微软雅黑"},
			Fill:      excelize.Fill{Type: "pattern", Color: []string{"2dece4"}, Pattern: 1},
			Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
			Border:    border,
		},
		{
			Font:   &excelize.Font{Bold: false, Size: 10, Family: "微软雅黑"},
			Border: border,
		},
	}
}
