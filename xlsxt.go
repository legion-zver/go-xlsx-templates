package xlsxt

import (
    "io"    
    "fmt"
    "errors"
    "regexp"
    "reflect"    
    "strings"   
    "strconv"  
    "io/ioutil"     
    "github.com/tealeg/xlsx"
    "github.com/legion-zver/gopdf"
    "github.com/aymerick/raymond"
)

var (
    rxTemplateItem  = regexp.MustCompile(`\{\{\s*([\w|\.]+)\s*\}\}`)
	rxMergeCellV    = regexp.MustCompile(`\[\s?v-merge\s?\]`)
    rxMergeIndex    = regexp.MustCompile(`\[\s?index\s?:\s?[\d|\.|\,]+\s?\]`)
    rxBrCellV       = regexp.MustCompile(`\[\s?BR\s?\]`)
)


// XlsxTemplateFile - файл шаблонизатора
type XlsxTemplateFile struct {
    template *xlsx.File
    result *xlsx.File
    fontDir string
}

// SetFontDir (XlsxTemplateFile)
func (s *XlsxTemplateFile) SetFontDir(path string) {
    s.fontDir = path
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

// SaveToHTML (XlsxTemplateFile) - сохраняем результат в PDF
func (s *XlsxTemplateFile) SaveToHTML(path string) error {
    var html string
	if s.result != nil {
        html = convertXlsxToHTML(s.result, true)
    } else if s.template != nil {
        html = convertXlsxToHTML(s.template, true)
    }
    if len(html) > 0 {        
        err := ioutil.WriteFile(path, []byte(html), 0655)
        if err != nil {
            return err
        }
        return nil
    }
    return errors.New("Not load template xlsx file")
}

// WriteToHTML (XlsxTemplateFile) - сохраняем результат в PDF
func (s *XlsxTemplateFile) WriteToHTML(writer io.Writer) error {
    var html string
	if s.result != nil {
        html = convertXlsxToHTML(s.result, true)
    } else if s.template != nil {
        html = convertXlsxToHTML(s.template, true)
    }
    if len(html) > 0 {        
        _, err := writer.Write([]byte(html))
        if err != nil {
            return err
        }
        return nil
    }
    return errors.New("Not load template xlsx file")
}

// SaveToPDF (XlsxTemplateFile) - сохраняем результат в PDF
func (s *XlsxTemplateFile) SaveToPDF(path string) error {
    var pdf *gopdf.GoPdf
	if s.result != nil {
        pdf = convertXlsxToPdf(s.result, s.fontDir)
    } else if s.template != nil {
        pdf = convertXlsxToPdf(s.template, s.fontDir)
    }
    if pdf != nil {        
        pdf.WritePdf(path)
        return nil
    }
    return errors.New("Not load template xlsx file")
}

// WriteToPDF (XlsxTemplateFile) - пишем результат в io.Writer
func (s *XlsxTemplateFile) WriteToPDF(writer io.Writer) error {
    var pdf *gopdf.GoPdf
	if s.result != nil {
        pdf = convertXlsxToPdf(s.result, s.fontDir)
    } else if s.template != nil {
        pdf = convertXlsxToPdf(s.template, s.fontDir)
    }
    if pdf != nil {
        bytes, err := pdf.GetBytesPdfReturnErr()
        if err != nil {
            return err
        }
        _, err = writer.Write(bytes)
        if err != nil {
            return err
        }
        return nil 
    }
    return errors.New("Not load template xlsx file")
}

// removeMergeCells
func removeMergeCells(file *xlsx.File) {
    if file != nil {
        for _, sheet := range file.Sheets {
            for rowIndex, row := range sheet.Rows {
                for cellIndex, cell := range row.Cells {
                    if cell.HMerge > 0 {
                        for x := 1; x <= cell.HMerge; x++ {
                            c := row.Cells[cellIndex+x]
                            if c != nil {
                                c.Value = ""
                                c.Hidden = true
                            }
                        }
                    }
                    if cell.VMerge > 0 {
                        for y := 1; y <= cell.VMerge; y++ {
                            r := sheet.Rows[rowIndex+y]
                            if r != nil {
                                c := r.Cells[cellIndex]
                                if c != nil {
                                    c.Value = ""
                                    c.Hidden = true
                                }                                
                            }
                        }
                    }
                }
            }
        }
    }
}

// convertXlsxToHTML - в HTML
func convertXlsxToHTML(file *xlsx.File, landscape bool) string {
    html := ""
    removeMergeCells(file)
    if file != nil {
        html += "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\">\n"
        html += "<html>\n<head>\n"
        html += "\t<meta http-equiv=\"content-type\" content=\"text/html; charset=utf-8\"/>\n\t<title></title>\n"
        html += "\t<style type=\"text/css\">\n"
        if landscape {
            html += "\t\t@page { size: landscape }\n"
        }
        html += "\t\ttd p, p { font-family: \"Liberation Sans\"; font-size: 10pt }\n"        
        html += "\t</style>\n"
        html += "</head>\n"
        html += "<body lang=\"ru-RU\" dir=\"ltr\">\n"
        pageWidthCM := 21.0*37.795276
        if landscape {
            pageWidthCM = 29.7*37.795276
        }
        for _, sheet := range file.Sheets {
            html += "\t<table cellpadding=\"2\" cellspacing=\"0\" style=\"page-break-before: always\">\n"            
            sizeWidth := getSheetWidth(sheet)
            // Устанавливаем размеры в процентах            
            for i, col := range sheet.Cols {
                if i < sheet.MaxCol {
                    html += "\t\t<col width=\""+strconv.FormatInt(int64((col.Width/sizeWidth)*pageWidthCM), 10)+"\">\n"
                } else {
                    break
                }
            }            
            // Проходимся по строкам
            for _, row := range sheet.Rows {
                html += "\t\t<tr height=\""+strconv.FormatInt(int64(row.Height), 10)+"\">\n"
                for _, cell := range row.Cells {
                    if !cell.Hidden {
                        style := cell.GetStyle()
                        html += "\t\t\t<td"
                        // Параметры ячейки
                        if cell.HMerge > 0 {
                            html += " colspan=\""+strconv.FormatInt(int64(cell.HMerge+1),10)+"\""
                        }
                        if cell.VMerge > 0 {
                            html += " rowspan=\""+strconv.FormatInt(int64(cell.VMerge+1),10)+"\""
                        }
                        if style != nil {
                            // Выравнивани по высоте внутри ячейки
                            if style.ApplyAlignment {
                                if len(style.Alignment.Vertical) > 0 && style.Alignment.Vertical != "none" {
                                    html += " valign=\""+style.Alignment.Vertical+"\""
                                }
                            }
                            // Бордер
                            html += " style=\""
                            if style.ApplyBorder {
                                if len(style.Border.Top) > 0 && style.Border.Top != "none" {
                                    html += "border-top: "+style.Border.Top+" solid "+style.Border.TopColor+"; "                                    
                                } else {
                                    html += "border-top: none; "    
                                }
                                if len(style.Border.Bottom) > 0 && style.Border.Bottom != "none" {
                                    html += "border-bottom: "+style.Border.Bottom+" solid "+style.Border.BottomColor+"; "                                    
                                } else {
                                    html += "border-bottom: none; "    
                                }
                                if len(style.Border.Left) > 0 && style.Border.Left != "none" {
                                    html += "border-left: "+style.Border.Left+" solid "+style.Border.LeftColor+"; "                                    
                                } else {
                                    html += "border-left: none; "    
                                }
                                if len(style.Border.Right) > 0 && style.Border.Right != "none" {
                                    html += "border-right: "+style.Border.Right+" solid "+style.Border.RightColor+"; "                                    
                                } else {
                                    html += "border-right: none; "    
                                }
                            } else {
                                html += "border: none; "                                
                            }                            
                            html += "padding: 0.05cm\""
                        }
                        html += ">\n" 
                        // Контент
                        html += "\t\t\t\t<p"
                        if style != nil {
                            if style.ApplyAlignment {
                                if len(style.Alignment.Horizontal) > 0 && style.Alignment.Horizontal != "none" {
                                    html += " align=\""+style.Alignment.Horizontal+"\""
                                }
                            }
                        }
                        html +=">"
                        if style != nil {
                            if style.ApplyFont {
                                html += "<font"
                                if len(style.Font.Name) > 0 {
                                    html += " face=\""+style.Font.Name+"\""
                                }
                                html += ">"
                                if style.Font.Bold {
                                    html += "<b>"    
                                }
                                if style.Font.Italic {
                                    html += "<i>"    
                                }
                                if style.Font.Underline {
                                    html += "<u>"    
                                }
                            }
                            html += cell.Value
                            if style.ApplyFont {
                                if style.Font.Underline {
                                    html += "</u>"    
                                }
                                if style.Font.Italic {
                                    html += "</i>"    
                                }
                                if style.Font.Bold {
                                    html += "</b>"    
                                }
                                html += "</font>"
                            }
                        } else {
                            html += cell.Value
                        }
                        html += "</p>\n"
                        html += "\t\t\t</td>\n"
                    }    
                }
                html += "\t\t</tr>\n"
            }
            html += "\t</table>\n"
        }
        html += "</body></html>"
    }
    return html
}

// convertXlsxToPdf - конвертирование XLSX в PDF
func convertXlsxToPdf(file *xlsx.File, fontDir string) *gopdf.GoPdf {
    removeMergeCells(file)
    if file != nil {
        pdf := gopdf.GoPdf{}
        w, h := 841.89, 595.28        
        pdf.Start(gopdf.Config{Unit: "pt", PageSize: gopdf.Rect{W: w, H: h}})        
        var addFonts = make(map[string]bool)
        for _, sheet := range file.Sheets {
            pdf.AddPage()
            pdf.SetX(0);pdf.SetY(0)            
            x, y, kW := 0.0, 0.0, w/getSheetWidth(sheet)
            for _, row := range sheet.Rows {
                // Анализ и правка высоты ячейки
                // Выставление шрифтов
                for i, cell := range row.Cells {                    
                    if !cell.Hidden {                        
                        style := cell.GetStyle()
                        if style != nil {
                            fontName := toPdfFont(style)                           
                            if !addFonts[fontName] {
                                err := pdf.AddTTFFont(fontName, fontDir+"/"+fontName+".ttf")
                                if err != nil {
                                    fmt.Println("Error load font: ", fontDir+"/"+fontName + ".ttf ", err)
                                    return nil
                                }
                                addFonts[fontName] = true
                            }
                            err := pdf.SetFont(fontName, getPdfFontStyleFromXLSXStyle(style), style.Font.Size)
                            if err != nil {                                
                                fmt.Println("Error set font: ", fontDir+"/"+fontName + ".ttf ", err)
                                return nil
                            }
                            // Только для WrapText
                            if style.Alignment.WrapText {                       
                                mergeWidth, mergeHeight := getMergeSizesFromCell(cell)
                                cellWidth := (sheet.Cols[i].Width+mergeWidth)*kW
                                if textWidth, err := pdf.MeasureTextWidth(cell.Value); err == nil {                            
                                    if textWidth > cellWidth {
                                        // Меняем выравнивание
                                        style.Alignment.Vertical = "top"
                                        // Разбиваем по словам и начинаем сложение                                        
                                        words := strings.Split(cell.Value, " ")
                                        line  := ""; countLines := 1; cell.Value = ""
                                        for _, word := range words {
                                            if tw, err := pdf.MeasureTextWidth(line+" "+word); err == nil {
                                                if tw > cellWidth {
                                                    countLines++
                                                    if len(cell.Value) > 0 {
                                                        cell.Value += "\n"
                                                    }
                                                    cell.Value += line
                                                    line = word
                                                } else {
                                                    if len(line) > 0 {
                                                        line += " "
                                                    }
                                                    line += word
                                                }
                                            } else {
                                                if len(line) > 0 {
                                                    line += " "
                                                }
                                                line += word
                                            }                                            
                                        }
                                        if len(line) > 0 {
                                            countLines++
                                            if len(cell.Value) > 0 {
                                                cell.Value += "\n"
                                            }
                                            cell.Value += line
                                            line = ""
                                        }
                                        // Проверка высоты
                                        if _, h, err := pdf.MeasureText("Z"); err == nil {
                                            if h*float64(countLines) > row.Height+mergeHeight {
                                                row.Height = h*float64(countLines)-mergeHeight
                                            }
                                        } else {
                                            if float64(style.Font.Size*countLines) > row.Height+mergeHeight {
                                                row.Height = float64(style.Font.Size*countLines)-mergeHeight
                                            }
                                        }                                      
                                    }                                
                                }
                            }
                        }
                    }
                }                
                cellHeigth := row.Height
                for i, cell := range row.Cells {
                    cellWidth := sheet.Cols[i].Width*kW
                    if !cell.Hidden {
                        style := cell.GetStyle()
                        if style != nil {
                            fontName := toPdfFont(style)                            
                            err := pdf.SetFont(fontName, getPdfFontStyleFromXLSXStyle(style), style.Font.Size)
                            if err != nil {
                                fmt.Println("Error set(2) font: ", fontDir+"/"+fontName + ".ttf ", err)
                                return nil
                            }
                        }
                        mergeWidth, mergeHeight := getMergeSizesFromCell(cell)
                        lines := strings.Split(cell.Value, "\n")
                        for lineIndex, line := range lines {
                            line = strings.Replace(line, "₽", "р.",-1)
                            if lineIndex < 1 {
                                pdf.CellWithOption(&gopdf.Rect{
                                W: cellWidth+mergeWidth*kW,
                                H: cellHeigth+mergeHeight}, line, toPdfCellOption(style, false))
                            } else {
                                if _,h, err := pdf.MeasureText(line); err == nil {                                    
                                    pdf.Br(h);pdf.SetX(x)
                                    
                                    pdf.CellWithOption(&gopdf.Rect{
                                    W: cellWidth+mergeWidth*kW, H: h}, line, toPdfCellOption(style, true))
                                } else {
                                    pdf.Br(float64(style.Font.Size));pdf.SetX(x)
                                    pdf.Text(line)
                                }                                                                                                
                            }
                        }                                                    
                    }                 
                    x += cellWidth; pdf.SetX(x); pdf.SetY(y)
                }
                y += cellHeigth; x = 0.0                 
                if y+cellHeigth >= h {
                    y = 0.0; pdf.AddPage()                    
                }                
                pdf.SetX(x);pdf.SetY(y)
            }
        }
        return &pdf
    }
    return nil
}

func toPdfFont(style *xlsx.Style) string {
    fontName := style.Font.Name        
    if style.Font.Bold {        
        fontName += "Bold"        
    } 
    if style.Font.Italic {
        fontName += "Italic"
    }
    return fontName
}

func toPdfCellOption(style *xlsx.Style, skipBorder bool) gopdf.CellOption {
    opt := gopdf.CellOption{}
    if style != nil {
        if style.Alignment.Horizontal == "center" {
            opt.Align = opt.Align | gopdf.Center
        } else if style.Alignment.Horizontal == "left" {
            opt.Align = opt.Align | gopdf.Left
        } else if style.Alignment.Horizontal == "right" {
            opt.Align = opt.Align | gopdf.Right
        }
        if style.Alignment.Vertical == "center" || style.Alignment.Vertical == "middle" {
            opt.Align = opt.Align | gopdf.Middle
        } else if style.Alignment.Vertical == "top" {
            opt.Align = opt.Align | gopdf.Top
        } else if style.Alignment.Vertical == "bottom" {
            opt.Align = opt.Align | gopdf.Bottom
        }       
        if !skipBorder { 
            if len(style.Border.Bottom) > 0 && style.Border.Bottom != "none" {
                opt.Border = opt.Border | gopdf.Bottom
            }
            if len(style.Border.Top) > 0 && style.Border.Top != "none" {
                opt.Border = opt.Border | gopdf.Top
            }  
            if len(style.Border.Left) > 0 && style.Border.Left != "none" {
                opt.Border = opt.Border | gopdf.Left
            }
            if len(style.Border.Right) > 0 && style.Border.Right != "none" {
                opt.Border = opt.Border | gopdf.Right
            }
        }
    }
    return opt
}

func getMergeSizesFromCell(cell *xlsx.Cell) (w, h float64) {
    w, h = 0.0, 0.0
    if cell != nil {
        sheet := cell.Row.Sheet
        if sheet != nil {
            cellIndex := indexCell(cell)
            rowIndex  := indexRow(cell.Row)
            if cell.HMerge > 0 {
                for x := 1; x <= cell.HMerge; x++ {
                    col := sheet.Cols[cellIndex+x]
                    if col != nil {
                        w += col.Width
                    }
                }  
            }
            if cell.VMerge > 0 {
                for y := 1; y <= cell.VMerge; y++ {
                    row := sheet.Rows[rowIndex+y]
                    if row != nil {
                        h += row.Height
                    }
                }
            }
        }
    }
    return
}

func getPdfFontStyleFromXLSXStyle(style *xlsx.Style) string {
    if style != nil {
        fontStyle := ""        
        if style.Font.Underline {            
            fontStyle += "U"
        }        
        return fontStyle
    }
    return ""
}

// getSheetWidth - длина sheet (сумма длин всех колонок)
func getSheetWidth(sheet *xlsx.Sheet) float64 {
    width := 0.0
    if sheet != nil {
        for i, col := range sheet.Cols {
            if i < sheet.MaxCol {
                width += col.Width   
            } else {
                break
            }      
        }
    }
    if width > 0.0 {
        return width
    }
    return 1.0
}


// Write (XlsxTemplateFile) - пишем результат в io.Writer
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
                    style.Alignment.WrapText = true
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

            // Убираем индексы [index:1] и проверяем на [BR]         
            for _,row := range newSheet.Rows {
                boldRight := false
                for _,cell := range row.Cells {
                    if cell != nil {
                        if len(cell.Value) > 0 {
                            if rxMergeIndex.MatchString(cell.Value) {                        
                                cell.Value = rxMergeIndex.ReplaceAllString(cell.Value, "")
                            }
                            if rxBrCellV.MatchString(cell.Value) {
                                cell.Value = rxBrCellV.ReplaceAllString(cell.Value, "")
                                boldRight = !boldRight                            
                            }
                            if boldRight && len(cell.Value) > 0 {    
                                if style := cell.GetStyle(); style != nil {                                    
                                    boldRightStyle := xlsx.NewStyle()                                    
                                    boldRightStyle.ApplyAlignment         = style.ApplyAlignment
                                    boldRightStyle.ApplyBorder            = style.ApplyBorder
                                    boldRightStyle.ApplyFill              = style.ApplyFill
                                    boldRightStyle.ApplyFont              = style.ApplyFont                                    

                                    if !boldRightStyle.ApplyFont {
                                        boldRightStyle.ApplyFont = true
                                    } 

                                    boldRightStyle.Border.Bottom        = style.Border.Bottom
                                    boldRightStyle.Border.BottomColor   = style.Border.BottomColor
                                    boldRightStyle.Border.Left          = style.Border.Left
                                    boldRightStyle.Border.LeftColor     = style.Border.LeftColor
                                    boldRightStyle.Border.Top           = style.Border.Top
                                    boldRightStyle.Border.TopColor      = style.Border.TopColor
                                    boldRightStyle.Border.Right         = style.Border.Right
                                    boldRightStyle.Border.RightColor    = style.Border.RightColor 

                                    boldRightStyle.Alignment.Horizontal   = style.Alignment.Horizontal
                                    boldRightStyle.Alignment.Indent       = style.Alignment.Indent
                                    boldRightStyle.Alignment.ShrinkToFit  = style.Alignment.ShrinkToFit
                                    boldRightStyle.Alignment.TextRotation = style.Alignment.TextRotation
                                    boldRightStyle.Alignment.Vertical     = style.Alignment.Vertical
                                    boldRightStyle.Alignment.WrapText     = style.Alignment.WrapText                                    

                                    boldRightStyle.Fill.BgColor     = style.Fill.BgColor
                                    boldRightStyle.Fill.FgColor     = style.Fill.FgColor
                                    boldRightStyle.Fill.PatternType = style.Fill.PatternType 

                                    boldRightStyle.Font.Bold      = true
                                    boldRightStyle.Font.Charset   = style.Font.Charset
                                    boldRightStyle.Font.Color     = style.Font.Color
                                    boldRightStyle.Font.Family    = style.Font.Family
                                    boldRightStyle.Font.Italic    = style.Font.Italic
                                    boldRightStyle.Font.Name      = style.Font.Name
                                    boldRightStyle.Font.Size      = style.Font.Size
                                    boldRightStyle.Font.Underline = style.Font.Underline   
                                    cell.SetStyle(boldRightStyle)                                                           
                                }                                                                                    
                            }
                        }
                    }
                }
            }         
        }
        return nil
    }
    return errors.New("Not load template xlsx file")
}

/* Вспомогательные функции */

// node - элeмент графа
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