package xlsxt

import (
    "io"
    //"fmt"
    "errors"
    "regexp"
    "reflect"
    "strings"
    "strconv"
    "github.com/tealeg/xlsx"
    "github.com/aymerick/raymond"
)

var (
    rxTemplateItem  = regexp.MustCompile(`\{\{\s*([\w|\.]+)\s*\}\}`)
	rxMergeCellV    = regexp.MustCompile(`\[\s?v-merge:\s?(\d+)\s?\]`)
)


// XlsxTemplateFile - файл шаблонизатора
type XlsxTemplateFile struct {
    template *xlsx.File
    result *xlsx.File
}

// Save (XlsxTemplateFile) - сохраняем результат
func (s *XlsxTemplateFile) Save(path string) error {
	if s.result != nil {
        return s.result.Save(path)		
    } else if s.template != nil {
        return s.template.Save(path)
    }
    return errors.New("Not load template xlsx file")
}

// Save (XlsxTemplateFile) - пишем результат в io.Writer
func (s *XlsxTemplateFile) Write(writer io.Writer) error {
    if s.result != nil {
        return s.result.Write(writer)
    } else if s.template != nil {
        return s.template.Write(writer)
    }
    return errors.New("Not load template xlsx file")
}

// OpenTemplate - открыть файл шаблона
func OpenTemplate(filename string) (*XlsxTemplateFile, error) {
    file, err := xlsx.OpenFile(filename)
    if err != nil {
        return nil, err
    }
    // Пробигаемся по ячейкам шаблона и проводим тестирование фона
    for _, sheet := range file.Sheets {
        for _, row := range sheet.Rows {
            for _, cell := range row.Cells {
                if style := cell.GetStyle(); style != nil {
                    if len(style.Fill.FgColor) < 1 {
                        style.Fill.FgColor = "FFFFFFFF"
                        cell.SetStyle(style)
                    }
                }
            }
        }
    }
    return &XlsxTemplateFile{template: file}, nil
}

// RenderTemplate (XlsxTemplateFile) рендер интрефейса в шаблон
func (s *XlsxTemplateFile) RenderTemplate(v interface{}) error {
    if s.template != nil {
        s.result = xlsx.NewFile()
        // Проходимся по вкладкам        
        for sheetIndex, sheet := range s.template.Sheets {
            newSheet, err := s.result.AddSheet(sheet.Name)
            if err != nil {
                s.result = nil
                return err
            }
            cloneSheet(sheet, newSheet)
            // Получаем объект
            obj := getObject(v, sheetIndex)
            // Проходимся по строкам
            for _, row := range sheet.Rows {
                // Проверка на массив или срез
                if !haveArrayInRow(row, obj) {                    
                    newRow := newSheet.AddRow()
                    cloneRow(row, newRow)
                    err := renderRow(newRow, obj)
                    if err != nil {
                        return nil
                    }
                    continue
                }
                // Если это массив или срез
            }            
        }
    }
    return errors.New("Not load template xlsx file")
}

/* Вспомогательные функции */

func getObject(v interface{}, index int) interface{} {
    val := reflect.ValueOf(v)
    if val.Type().Kind() == reflect.Ptr {
        val = val.Elem()
    }
    if val.Type().Kind() == reflect.Slice || val.Type().Kind() == reflect.Array {
        val = val.Index(index)
    }
    return val.Interface()
}

// cloneCell - клонирование ячейки
func cloneCell(from, to *xlsx.Cell) {
	to.Value = from.Value
	style := from.GetStyle()	
	to.SetStyle(style)
	to.HMerge = from.HMerge
	to.VMerge = from.VMerge
	to.Hidden = from.Hidden
	to.NumFmt = from.NumFmt
}

// cloneRow - клонирование строки
func cloneRow(from, to *xlsx.Row) {
	to.Height = from.Height
	for _, cell := range from.Cells {
		newCell := to.AddCell()
		cloneCell(cell, newCell)
	}
}

// cloneSheet - клонирование вкладки
func cloneSheet(from, to *xlsx.Sheet) {
	for _, col := range from.Cols {
		newCol := xlsx.Col{}
		style := col.GetStyle()
		newCol.SetStyle(style)
		newCol.Width = col.Width
		newCol.Hidden = col.Hidden
		newCol.Collapsed = col.Collapsed
		newCol.Min = col.Min
		newCol.Max = col.Max
		to.Cols = append(to.Cols, &newCol)
	}
}

// renderCell - рендер ячейки
func renderCell(cell *xlsx.Cell, v interface{}) error {	    
    template, err := raymond.Parse(cell.Value)
	if err != nil {
		return err
	}
    out, err := template.Exec(v)
	if err != nil {
		return err
	}
    cell.Value = out  
    // После преобразования выполняем   
    if rxMergeCellV.MatchString(cell.Value) {
        vals := rxMergeCellV.FindStringSubmatch(cell.Value)
        if len(vals) > 1 {
            if val, err := strconv.Atoi(vals[1]); err == nil && val > 0 {
                cell.VMerge = val
            }
        }
        cell.Value = rxMergeCellV.ReplaceAllString(cell.Value, "")
    }
	return nil
}

// Рендер строки
func renderRow(row *xlsx.Row, v interface{}) error {
	for _, cell := range row.Cells {
		err := renderCell(cell, v)
		if err != nil {
			return err
		}
	}
	return nil
}

// haveArrayInRow - содержится ли массив в строке
func haveArrayInRow(row *xlsx.Row, v interface{}) bool {
    if row != nil {
        // Обнодим все ячейки в строке и проверяем наличие массива
        for _, cell := range row.Cells {
            if match := rxTemplateItem.FindStringSubmatch(cell.Value); match != nil && len(match) > 1 {                
                names := strings.Split(match[1], ".")
                if len(names) > 0 {
                    t := reflect.TypeOf(v)
                    for _, name := range names {
                        t := findInType(t, name)
                        if t != nil {
                            if t.Kind() == reflect.Array || t.Kind() == reflect.Slice {
                                return true
                            }
                        } else {
                            break
                        }
                    }
                }
            }
        }
    }
    return false
}

// findInType - получаем тип по имени
func findInType(t reflect.Type, name string) reflect.Type {
    kind := t.Kind()
    // Если это ссылка, то получаем истенный тип
    if kind == reflect.Ptr || kind == reflect.Interface {
        t = t.Elem()
    }
    if kind == reflect.Struct {
        if field, ok := t.FieldByName(name); ok {
            return field.Type
        }
    } 
    return nil
}