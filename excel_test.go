package goexcel_test

import (
	goexcel "github.com/atopx/go-excel"
	"github.com/stretchr/testify/assert"
	"os"
	"testing"
)

type WorkSheet struct {
	Code string `excel:"head:编号"`
	Name string `excel:"head:名称"`
	Age  string `excel:"head:年龄"`
}

var data = []*WorkSheet{
	{Code: "001", Name: "小明", Age: "30"},
	{Code: "002", Name: "小红", Age: "29"},
}

const (
	filename = "test.xlsx"
	sheet    = "Sheet1"
)

func TestWorkSheet(t *testing.T) {
	excel := goexcel.New[WorkSheet](sheet)

	err := excel.Write(data, excel.DefaultStyles()...)
	if err != nil {
		t.Error(err)
	}

	if err = excel.Save(filename); err != nil {
		t.Error(err)
	}
}

func TestExcelImport(t *testing.T) {
	file, _ := os.Open(filename)
	excel, err := goexcel.Open[WorkSheet](file, sheet)
	if err != nil {
		t.Error(err)
	}
	dst, err := excel.Read()
	if err != nil {
		t.Error(err)
	}
	assert.Equal(t, len(data), len(dst))
	for i := 0; i < len(dst); i++ {
		assert.Equal(t, data[i].Code, dst[i].Code)
		assert.Equal(t, data[i].Name, dst[i].Name)
		assert.Equal(t, data[i].Age, dst[i].Age)
	}
}
