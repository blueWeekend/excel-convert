package utils

import (
	"errors"
	"os"
	"reflect"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type TmplCheckMode int

const (
	_                TmplCheckMode = iota
	TmplCheckStrict                //严格校验文件标题与初始化传参数组是否一致
	TmplCheckLenient               //校验文件标题是否包含初始化数组
	TmplCheckDisable               //关闭模板校验
)

type ExcelConverter struct {
	tmplHeader         []string
	columnIndexMap     map[string]int
	tagName            string
	tmplCheckMode      TmplCheckMode
	tmplCheckLineLimit int
	ptrKind            map[reflect.Kind]struct{}
}

type ExcelMarshaler interface {
	MarshalExcel() string
}

type ExcelUnmarshaler interface {
	UnmarshalExcel(string) error
}

type Option func(opt *ExcelConverter)

func SetTagName(tagName string) Option {
	return func(opt *ExcelConverter) {
		opt.tagName = tagName
	}
}

func SetTmplCheckMode(mode TmplCheckMode) Option {
	return func(opt *ExcelConverter) {
		opt.tmplCheckMode = mode
	}
}

func NewExcelConverter(columnNames []string, opts ...Option) *ExcelConverter {
	c := &ExcelConverter{
		tagName:            "excel",
		tmplCheckMode:      TmplCheckLenient,
		tmplCheckLineLimit: 5,
		tmplHeader:         columnNames,
		ptrKind: map[reflect.Kind]struct{}{
			reflect.Ptr:   {},
			reflect.Map:   {},
			reflect.Slice: {},
		},
	}
	for i := range opts {
		opts[i](c)
	}
	c.columnIndexMap = make(map[string]int, len(columnNames))
	for i := range columnNames {
		_, ok := c.columnIndexMap[columnNames[i]]
		if ok {
			panic("repeat column name:" + columnNames[i])
		}
		c.columnIndexMap[columnNames[i]] = i
	}
	return c
}

func (e *ExcelConverter) checkTmpl(rows [][]string) int {
	if e.tmplCheckMode == TmplCheckDisable {
		return 0
	}
	for i := 0; i < len(rows) && i < e.tmplCheckLineLimit; i++ {
		if len(rows[i]) < len(e.tmplHeader) {
			continue
		}
		for j := range e.tmplHeader {
			if e.tmplCheckMode == TmplCheckStrict && rows[i][j] != e.tmplHeader[j] {
				return -1
			}
			if e.tmplCheckMode == TmplCheckLenient && !strings.Contains(rows[i][j], e.tmplHeader[j]) {
				return -1
			}
		}
		return i
	}
	return -1
}

func (e *ExcelConverter) getSliceInfo(slicePtr any) (sliceValue reflect.Value, elemType reflect.Type, err error) {
	ptrValue := reflect.ValueOf(slicePtr)
	if ptrValue.Kind() != reflect.Ptr {
		return reflect.Value{}, nil, errors.New("slicePtr must be a pointer to a slice")
	}
	sliceValue = ptrValue.Elem()
	if sliceValue.Kind() != reflect.Slice {
		err = errors.New("slicePtr must be a pointer to a slice")
		return
	}
	elemType = sliceValue.Type().Elem()
	if elemType.Kind() == reflect.Ptr {
		err = errors.New("pointer slice is not supported, use slice of struct instead")
		return
	}
	return
}

func (e *ExcelConverter) ReadAll(fileName string, slicePtr any) error {
	_, err := os.Stat(fileName)
	if err != nil {
		return err
	}
	sliceValue, elemType, err := e.getSliceInfo(slicePtr)
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		return err
	}
	defer f.Close()
	sheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return err
	}
	headerLine := e.checkTmpl(rows)
	if headerLine < 0 {
		return errors.New("invalid tmpl")
	}
	rows = rows[headerLine+1:]
	newSlice := reflect.MakeSlice(sliceValue.Type(), 0, len(rows))
	for i := range rows {
		zeroValue := reflect.New(elemType).Elem()
		err = e.setFields(zeroValue, rows[i])
		if err != nil {
			return err
		}
		newSlice = reflect.Append(newSlice, zeroValue)
	}
	sliceValue.Set(newSlice)
	return nil
}

func (e *ExcelConverter) unmarshal(rows []string, importItem any) error {
	targetValue := reflect.ValueOf(importItem)
	if targetValue.Kind() != reflect.Ptr {
		return errors.New("importItem is must be pointer")
	}
	targetValue = targetValue.Elem()
	if targetValue.Kind() != reflect.Struct {
		return errors.New("target must be struct of assignStructFields")
	}
	return e.setFields(targetValue, rows)
}
func (e *ExcelConverter) setFields(val reflect.Value, rows []string) error {
	valType := val.Type()
	for i := 0; i < val.NumField(); i++ {
		field := val.Field(i)
		fieldType := valType.Field(i)
		_, isPtr := e.ptrKind[field.Kind()]
		tag := fieldType.Tag.Get(e.tagName)
		if isPtr && field.IsNil() {
			switch field.Kind() {
			case reflect.Ptr:
				newPtr := reflect.New(fieldType.Type.Elem())
				field.Set(newPtr)
			case reflect.Slice:
				newSlice := reflect.MakeSlice(field.Type(), 0, 0)
				field.Set(newSlice)
			case reflect.Map:
				newMap := reflect.MakeMap(field.Type())
				field.Set(newMap)
			}
		}
		//无tag为结构体嵌套，否则为自定义字段
		if tag == "" && (field.Kind() == reflect.Struct || (field.Kind() == reflect.Ptr && field.Elem().Kind() == reflect.Struct)) {
			var err error
			if field.Kind() == reflect.Ptr {
				err = e.setFields(field.Elem(), rows)
			} else {
				err = e.setFields(field, rows)
			}
			if err != nil {
				return err
			}
			continue
		}
		if tag == "" || !field.CanSet() {
			continue
		}
		index, ok := e.columnIndexMap[tag]
		if !ok {
			continue
		}
		if index >= len(rows) {
			continue
		}
		if rows[index] == "" {
			continue
		}
		switch field.Kind() {
		case reflect.String:
			field.SetString(rows[index])
		case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
			val, err := strconv.ParseInt(rows[index], 10, 64)
			if err != nil {
				return err
			} else {
				field.SetInt(val)
			}
		case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
			val, err := strconv.ParseUint(rows[index], 10, 64)
			if err != nil {
				return err
			} else {
				field.SetUint(val)
			}
		case reflect.Bool:
			val, err := strconv.ParseBool(rows[index])
			if err != nil {
				return err
			} else {
				field.SetBool(val)
			}
		case reflect.Float32, reflect.Float64:
			if rows[index] != "" {
				val, err := strconv.ParseFloat(rows[index], 64)
				if err != nil {
					return err
				} else {
					field.SetFloat(val)
				}
			}
		default:
			method, ok := field.Interface().(ExcelUnmarshaler)
			if ok {
				err := method.UnmarshalExcel(rows[index])
				if err != nil {
					return err
				}
			}
		}
	}
	return nil
}

func (e *ExcelConverter) WriteExcel(header []string, fileName, sheetName string, datas any) error {
	targetValue := reflect.ValueOf(datas)
	if targetValue.Kind() != reflect.Slice {
		return errors.New("datas must be slice of writeValidDataToExcel")
	}
	if sheetName == "" {
		sheetName = "Sheet1"
	}
	file := excelize.NewFile()
	defer file.Close()
	streamWriter, err := file.NewStreamWriter(sheetName)
	if err != nil {
		return err
	}
	cell, err := excelize.CoordinatesToCellName(1, 1)
	if err != nil {
		return err
	}
	headerRow := make([]any, len(header))
	for i := range header {
		headerRow[i] = header[i]
	}
	if err := streamWriter.SetRow(cell, headerRow); err != nil {
		return err
	}
	length := targetValue.Len()
	for i := 0; i < length; i++ {
		elemValue := targetValue.Index(i)
		columnValMap := make(map[string]columnItem, elemValue.NumField())
		err := e.marshal(elemValue, columnValMap, 1)
		if err != nil {
			return err
		}
		row := make([]any, 0, len(columnValMap))
		for j := range header {
			row = append(row, columnValMap[header[j]].data)
		}
		cell, err := excelize.CoordinatesToCellName(1, i+2)
		if err != nil {
			return err
		}
		if err := streamWriter.SetRow(cell, row); err != nil {
			return err
		}
	}
	if err := streamWriter.Flush(); err != nil {
		return err
	}
	if err := file.SaveAs(fileName); err != nil {
		return err
	}
	return nil
}

type columnItem struct {
	data  any
	depth uint8
}

func (c *ExcelConverter) marshal(val reflect.Value, columnValMap map[string]columnItem, depth uint8) error {
	valType := val.Type()
	for i := 0; i < val.NumField(); i++ {
		field := val.Field(i)
		fieldType := valType.Field(i)
		tag := fieldType.Tag.Get(c.tagName)
		if tag == "" && (field.Kind() == reflect.Struct || (field.Kind() == reflect.Ptr && !field.IsNil() && field.Elem().Kind() == reflect.Struct)) {
			var err error
			if field.Kind() == reflect.Ptr {
				err = c.marshal(field.Elem(), columnValMap, depth+1)
			} else {
				err = c.marshal(field, columnValMap, depth+1)
			}
			if err != nil {
				return err
			}
			continue
		}
		if !field.CanSet() {
			continue
		}
		if tag == "" {
			continue
		}
		_, ok := c.columnIndexMap[tag]
		if !ok {
			continue
		}
		if item, ok := columnValMap[tag]; ok && item.depth < depth {
			continue
		}
		switch field.Kind() {
		case reflect.String:
			columnValMap[tag] = columnItem{data: field.String(), depth: depth}
		case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
			val := field.Int()
			if val != 0 {
				columnValMap[tag] = columnItem{data: val, depth: depth}
			} else {
				columnValMap[tag] = columnItem{data: "", depth: depth}
			}
		case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
			val := field.Uint()
			if val != 0 {
				columnValMap[tag] = columnItem{data: val, depth: depth}
			} else {
				columnValMap[tag] = columnItem{data: "", depth: depth}
			}
		case reflect.Float32, reflect.Float64:
			val := field.Float()
			if val != 0 {
				columnValMap[tag] = columnItem{data: val, depth: depth}
			} else {
				columnValMap[tag] = columnItem{data: "", depth: depth}
			}
		default:
			method, ok := field.Interface().(ExcelMarshaler)
			if ok {
				columnValMap[tag] = columnItem{data: method.MarshalExcel(), depth: depth}
			}
		}

	}
	return nil
}
