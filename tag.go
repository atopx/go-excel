package goexcel

import (
	"reflect"
	"strconv"
	"strings"
)

const (
	TagNameBase   = "excel"
	TagNameHead   = "head:"
	TagNameWidth  = "width:"
	TagNameHeight = "height:"
)

type Tag struct {
	Head   string
	Width  float64
	Height float64
}

func parseTag(field reflect.StructField) (*Tag, error) {
	tag := new(Tag)
	str, ok := field.Tag.Lookup(TagNameBase)
	if !ok {
		return nil, TagNotFound
	}
	for _, tagStr := range strings.Split(str, ";") {
		if strings.HasPrefix(tagStr, TagNameHead) {
			tag.Head = strings.TrimPrefix(tagStr, TagNameHead)
		}
		if strings.HasPrefix(tagStr, TagNameWidth) {
			value := strings.TrimPrefix(tagStr, TagNameWidth)
			tag.Width, _ = strconv.ParseFloat(value, 64)
		}
		if strings.HasPrefix(tagStr, TagNameHeight) {
			value := strings.TrimPrefix(tagStr, TagNameHeight)
			tag.Height, _ = strconv.ParseFloat(value, 64)
		}
	}
	if tag.Head == "" {
		tag.Head = field.Name
	}
	return tag, nil
}
