# go-excel

excel相关工具包, 基于 github.com/xuri/excelize/v2

# 安装

```shell
go get -u github.com/atopx/go-excel 
```

# examples

### 1. 导出

```go
package main

import (
	goexcel "github.com/atopx/go-excel"
)

type WorkSheet struct {
	Code string `excel:"head:编号;width:20"`
	Age  string `excel:"head:年龄"`
	Name string `excel:"head:名称;width:30;"`
}

const (
	filename = "test.xlsx"
	sheet    = "Sheet1"
)

func main() {
	var data = []*WorkSheet{
		{Code: "001", Name: "小明", Age: "30"},
		{Code: "002", Name: "小红", Age: "29"},
	}

	excel := goexcel.New[WorkSheet](sheet)
	err := excel.Write(data, excel.DefaultStyles()...)
	if err != nil {
		panic(err)
	}

	if err = excel.Save(filename); err != nil {
		panic(err)
	}
}
```

### 2. 导入

```go
package main

import (
	"fmt"
	goexcel "github.com/atopx/go-excel"
	"os"
)

type WorkSheet struct {
	Code string `excel:"head:编号"`
	Name string `excel:"head:名称"`
	Age  string `excel:"head:年龄"`
}

const (
	filename = "test.xlsx"
	sheet    = "Sheet1"
)

func main() {
	file, _ := os.Open(filename)

	excel, err := goexcel.Open[WorkSheet](file, sheet)
	if err != nil {
		panic(err)
	}

	data, err := excel.Read()
	if err != nil {
		panic(err)
	}

	for _, row := range data {
		fmt.Printf("%+v\n", row)
	}
}
```
