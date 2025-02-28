package main

import (
    "fmt"
    "github.com/legion-zver/go-xlsx-templates"
)

type SubItem struct {
    Name string
}

type Item struct {
    Name string
    SubItems []SubItem
}

type Test struct {
    Items []Item
}

func main()  {    
    file, err := xlsxt.OpenTemplate("example.xlsx")
    file.SetFontDir("./fonts")
    if err != nil {
        fmt.Println(err)
        return
    }
    test := Test{Items: []Item{
        Item{Name: "item1", SubItems: []SubItem{ SubItem{ Name: "1"}, SubItem{ Name: "2"}}},
        Item{Name: "item2", SubItems: []SubItem{ SubItem{ Name: "1"}, SubItem{ Name: "2"}, SubItem{ Name: "3"}}}}}

    err = file.RenderTemplate(&test)
    if err != nil {
        fmt.Println(err)
        return
    }
    err = file.Save("result.xlsx")
    if err != nil {
        fmt.Println(err)
    }
    err = file.SaveToPDF("result.pdf")
    if err != nil {
        fmt.Println(err)
    }
}