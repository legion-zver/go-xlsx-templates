package xlsxt

import (
    "io"    
    "fmt"
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
	rxMergeCellV    = regexp.MustCompile(`\[\s?v-merge\s?\]`)
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
            // Раскладываем объект на граф
            graph := new(node)
            graph.FromObject(v)
            // Вывод для теста
            lines := graph.ListMap()            
            // Проходимся по строкам
            for _, row := range sheet.Rows {
                // Проверка на массив или срез
                if !haveArrayInRow(row, obj) {                    
                    newRow := newSheet.AddRow(); cloneRow(row, newRow)
                    if err := renderRow(newRow, obj); err != nil {
                        return nil
                    }
                    continue
                }
                for i := 0; i < len(lines); i++ {
                    newRow := newSheet.AddRow()
                    cloneRow(row, newRow)                    
                    if err := renderRow(newRow, lines[i]); err != nil {
                        return nil
                    }
                }
            }
            graph = nil            
        }
        return nil
    }
    return errors.New("Not load template xlsx file")
}

/* Вспомогательные функции */

type cacheMergeCell struct {

}

// node - элмент графа
type node struct {
    name     string
    values   map[string]interface{}    
    nodes []*node
}

func (n *node) String() string {    
    return fmt.Sprintln(n.ListMap()) 
}

func (n *node) ListMap() []map[string]interface{} {
    var m = make([]map[string]interface{}, 0)
    n.toListMap("", make(map[string]interface{}), &m)
    return m
} 

func (n *node) toListMap(name string, current map[string]interface{}, out *[]map[string]interface{}) {
    for key, value := range n.values {
        if len(n.name) > 0 {
            key = n.name + "_" + key
        } 
        if len(name) > 0 {
            key = name + "_" + key
        }
        current[key] = value
    }
    if len(n.nodes) > 0 {
        path := name
        if len(n.name) > 0 {
            if len(path) > 0 {
                path += "_"
            } 
            path += n.name
        }
        // Проходимся и формируем счетчик
        var counter = make(map[string]int)        
        for _, node := range n.nodes {
            nname := path
            if len(nname) > 0 {
                if len(node.name) > 0 {
                    nname += "_"
                }                 
            }
            nname += node.name            
            if cv, ok := counter[nname+"_length"]; ok {
                counter[nname+"_length"] = cv + 1
            } else { 
                counter[nname+"_length"] = 1
            }
        }
        for key, cv := range counter {
            current[key] = cv
        }
        // Формируем строки
        for _, node := range n.nodes {
            var nw = make(map[string]interface{})
            for k,v := range current {
                nw[k] = v
            }       
            node.toListMap(path, nw, out)
        }
    } else {
        *out = append(*out, current)
    }
}

func (n* node) FromObject(obj interface{}) {
    if n.values == nil {
        n.values = make(map[string]interface{}, 0)
    }
    if n.nodes == nil {
        n.nodes = make([]*node, 0)
    }
    val := reflect.ValueOf(obj)
    kind := val.Type().Kind()
    if kind == reflect.Ptr || kind == reflect.Interface {
        val  = val.Elem()
        kind = val.Type().Kind()
    }
    if kind == reflect.Struct {
        t := val.Type()
        for i := 0; i < t.NumField(); i++ {
            field := t.Field(i)
            fv := val.FieldByIndex(field.Index)
            kind = fv.Type().Kind()
            if kind == reflect.Ptr || kind == reflect.Interface {
                fv  = fv.Elem()
                kind = fv.Type().Kind()
            }
            if kind == reflect.Map {
                node := new(node)
                node.name = field.Name
                node.FromObject(fv.Interface())
                n.nodes = append(n.nodes, node)
            } else if kind == reflect.Array || kind == reflect.Slice {                
                for j := 0; j < fv.Len(); j++ {
                    node := new(node)
                    node.name = field.Name
                    node.FromObject(fv.Index(j).Interface())
                    n.nodes = append(n.nodes, node)
                }
            } else {
                n.values[field.Name] = val.FieldByIndex(field.Index).Interface()
            }
        }
    } else if kind == reflect.Map {
        for _, key := range val.MapKeys() {
            mapItem := val.MapIndex(key)
            kind = mapItem.Type().Kind()
            if kind == reflect.Ptr || kind == reflect.Interface {
                mapItem  = mapItem.Elem()
                kind = mapItem.Type().Kind()
            }
            if kind == reflect.Map {
                node := new(node)
                node.name = key.String()
                node.FromObject(mapItem.Interface())
                n.nodes = append(n.nodes, node)
            } else if kind == reflect.Array || kind == reflect.Slice {                
                for j := 0; j < mapItem.Len(); j++ {
                    node := new(node)
                    node.name = key.String()
                    node.FromObject(mapItem.Index(j).Interface())
                    n.nodes = append(n.nodes, node)
                }
            } else {
                n.values[key.String()] = mapItem.Interface()
            }
        }
    } else if kind == reflect.Array || kind == reflect.Slice {
        for i := 0; i < val.Len(); i++ {
            node := new(node)
            node.name = n.name
            if len(node.name) > 0 {
                node.name += "_"+strconv.FormatInt(int64(i),10)
            }
            node.FromObject(val.Index(i).Interface())
            n.nodes = append(n.nodes, node)
        }
    }
}


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

func indexRow(row *xlsx.Row) int {
    if row != nil && row.Sheet != nil {
        for i, r := range row.Sheet.Rows {
            if r == row {
                return i
            }
        }
    } 
    return -1
}

func indexCell(cell *xlsx.Cell) int {
    if cell != nil && cell.Row != nil {
        for i, c := range cell.Row.Cells {
            if c == cell {
                return i
            }
        }
    } 
    return -1
}

// renderCell - рендер ячейки
func renderCell(cell *xlsx.Cell, v interface{}) error {	    
    // Правки для совместимости шаблонизатора
    tpl := strings.Replace(cell.Value, "{{", "{{{", -1)
	tpl = strings.Replace(tpl, "}}", "}}}", -1)
    tpl = strings.Replace(tpl,".","_",-1)
    tpl = strings.Replace(tpl,":length","_length",-1)
    // Обработка контента
    out, err := raymond.Render(tpl, v)
	if err != nil {
		return err
	}
    cell.Value = out  
    // Если у поля есть флаг авто мерджинга, то начинаем проверку   
    if rxMergeCellV.MatchString(cell.Value) {
        cell.Value = rxMergeCellV.ReplaceAllString(cell.Value, "")
        if len(strings.TrimSpace(cell.Value)) > 0 {
            // Проверяем значения
            ic, ir := indexCell(cell), indexRow(cell.Row)
            if ir >= 0 && ic >= 0 {
                var lastRow *xlsx.Row
                for i := (ir-1); i >= 0; i-- {
                    row := cell.Row.Sheet.Rows[i]
                    if row.Cells[ic].Value == cell.Value {
                        lastRow = row
                    } else {
                        break
                    }
                }
                if lastRow != cell.Row {
                    ilr := indexRow(lastRow)
                    if ilr >= 0 {
                        lastRow.Cells[ic].VMerge = ir-ilr
                    }
                }
            }
        }
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
                        t := findType(t, name)
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


// findType - получаем тип по имени
func findType(t reflect.Type, name string) reflect.Type {
    kind := t.Kind()
    // Если это ссылка, то получаем истенный тип
    if kind == reflect.Ptr || kind == reflect.Interface {
        t = t.Elem()
    }
    kind = t.Kind()
    if kind == reflect.Struct {
        if field, ok := t.FieldByName(name); ok {
            return field.Type
        }
    } 
    return nil
}

// findValue - получаем тип по имени
func findValue(v reflect.Value, name string) (reflect.Value, bool) {
    kind := v.Type().Kind()
    // Если это ссылка, то получаем истенный тип
    if kind == reflect.Ptr || kind == reflect.Interface {
        v = v.Elem()
    }
    kind = v.Type().Kind()
    if kind == reflect.Struct {
        v := v.FieldByName(name)
        if v.IsValid() {
            return v, true
        }        
    } 
    return v, false
}