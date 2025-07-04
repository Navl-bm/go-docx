package main

import (
	"fmt"

	godocx "github.com/Navl-bm/go-docx"
)

// Пример использования
func main() {
	replacements := map[string]interface{}{
		"{data1}": "text1",
		"{data2}": "text2",
		"{data3}": "text3-1\ntext3-2\ntext3-3",
		"{data4}": []string{"text4-1", "text4-2", "text4-3"},
		"{data5}": "text5",
		"{data6}": "text6",
	}

	outputPath, err := godocx.GenerateDocx("template.docx", replacements)
	if err != nil {
		fmt.Println("Ошибка:", err)
	} else {
		fmt.Println("Создан документ:", outputPath)
	}
}
