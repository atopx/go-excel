package goexcel

import (
	"errors"
	"fmt"
)

var (
	StructEmptyError = errors.New("struct is empty")
	RecordNotFound   = errors.New("record not found")
	HeadInconsistent = errors.New("head tag inconsistent")

	TagNotFound = fmt.Errorf("not found struct tag %s", TagNameBase)
)
