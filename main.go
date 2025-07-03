package main

import (
	"archive/zip"
	"fmt"
	"io"
	"io/fs"
	"os"
	"path/filepath"
	"strings"

	"github.com/beevik/etree"
)

// Разархивирует DOCX в указанную директорию
func UnzipDocx(src, dest string) error {
	r, err := zip.OpenReader(src)
	if err != nil {
		return err
	}
	defer r.Close()

	for _, f := range r.File {
		err := extractFile(f, dest)
		if err != nil {
			return err
		}
	}
	return nil
}

// Извлекает отдельный файл из архива
func extractFile(f *zip.File, dest string) error {
	rc, err := f.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	path := filepath.Join(dest, f.Name)
	if f.FileInfo().IsDir() {
		return os.MkdirAll(path, os.ModePerm)
	}

	if err = os.MkdirAll(filepath.Dir(path), os.ModePerm); err != nil {
		return err
	}

	out, err := os.Create(path)
	if err != nil {
		return err
	}
	defer out.Close()

	_, err = io.Copy(out, rc)
	return err
}

// Создает DOCX файл из распакованной директории
func ZipDocx(source, target string) error {
	zipFile, err := os.Create(target)
	if err != nil {
		return err
	}
	defer zipFile.Close()

	writer := zip.NewWriter(zipFile)
	defer writer.Close()

	return filepath.Walk(source, func(path string, info fs.FileInfo, err error) error {
		if err != nil {
			return err
		}

		relPath, err := filepath.Rel(source, path)
		if err != nil {
			return err
		}

		// Пропускаем корневую директорию и системные файлы
		if relPath == "." || strings.HasPrefix(relPath, "__") {
			return nil
		}

		header, err := zip.FileInfoHeader(info)
		if err != nil {
			return err
		}

		// Сохраняем оригинальную структуру путей
		header.Name = filepath.ToSlash(relPath)
		if info.IsDir() {
			header.Name += "/"
		} else {
			header.Method = zip.Deflate
		}

		entry, err := writer.CreateHeader(header)
		if err != nil {
			return err
		}

		if !info.IsDir() {
			file, err := os.Open(path)
			if err != nil {
				return err
			}
			defer file.Close()

			_, err = io.Copy(entry, file)
			if err != nil {
				return err
			}
		}
		return nil
	})
}

// SafeReplace заменяет текст в XML-файле с сохранением пространств имен и префиксов
func SafeReplace(xmlPath, oldText string, newText interface{}) error {
	doc := etree.NewDocument()

	// Настройки записи
	doc.WriteSettings = etree.WriteSettings{
		CanonicalAttrVal: true,
		CanonicalText:    true,
		CanonicalEndTags: true,
	}

	if err := doc.ReadFromFile(xmlPath); err != nil {
		return fmt.Errorf("ошибка чтения XML: %v", err)
	}

	// Функция для создания элемента с префиксом w
	createElement := func(tag string) *etree.Element {
		return &etree.Element{
			Tag: "w:" + tag,
		}
	}

	// Находим все параграфы
	for _, p := range doc.FindElements("//w:p") {
		// Собираем весь текст параграфа
		fullText := ""
		for _, t := range p.FindElements(".//w:t") {
			fullText += t.Text()
		}

		// Проверяем, содержится ли шаблон в собранном тексте
		if strings.Contains(fullText, oldText) {
			// Копируем свойства параграфа
			pPr := p.SelectElement("w:pPr")

			// Копируем свойства run (берем из первого run)
			var rPr *etree.Element
			if firstRun := p.SelectElement("w:r"); firstRun != nil {
				rPr = firstRun.SelectElement("w:rPr")
			}

			switch v := newText.(type) {
			case string:
				// Обработка строки с переносами
				parts := strings.Split(v, "\n")

				// Создаем новый run
				newR := createElement("r")
				if rPr != nil {
					newR.AddChild(rPr.Copy())
				}

				// Добавляем текст с переносами
				for i, part := range parts {
					if i > 0 {
						// Добавляем перенос строки
						br := createElement("br")
						newR.AddChild(br)
					}

					newT := createElement("t")
					newT.SetText(part)
					newR.AddChild(newT)
				}

				// Создаем новый параграф
				newP := createElement("p")
				if pPr != nil {
					newP.AddChild(pPr.Copy())
				}
				newP.AddChild(newR)

				// Заменяем исходный параграф новым
				if parent := p.Parent(); parent != nil {
					for idx, child := range parent.Child {
						if child == p {
							parent.InsertChildAt(idx, newP)
							parent.RemoveChild(p)
							break
						}
					}
				}

			case []string:
				// Обработка массива строк
				var newParagraphs []*etree.Element
				for _, line := range v {
					newP := createElement("p")

					if pPr != nil {
						newP.AddChild(pPr.Copy())
					}

					newR := createElement("r")

					if rPr != nil {
						newR.AddChild(rPr.Copy())
					}

					newT := createElement("t")
					newT.SetText(line)

					newR.AddChild(newT)
					newP.AddChild(newR)
					newParagraphs = append(newParagraphs, newP)
				}

				// Заменяем исходный параграф новыми
				if parent := p.Parent(); parent != nil {
					for idx, child := range parent.Child {
						if child == p {
							parent.RemoveChild(p)
							for i, newP := range newParagraphs {
								parent.InsertChildAt(idx+i, newP)
							}
							break
						}
					}
				}

			default:
				return fmt.Errorf("неподдерживаемый тип newText")
			}
		}
	}

	doc.Indent(2)
	return doc.WriteToFile(xmlPath)
}

// ReplaceInAllFiles обрабатывает все XML-файлы в DOCX
func ReplaceInAllFiles(dir, oldText string, newText interface{}) error {
	files := []string{
		"word/document.xml",
		"word/footer1.xml",
		"word/header1.xml",
		"word/footnotes.xml",
		"word/endnotes.xml",
	}

	for _, file := range files {
		path := filepath.Join(dir, file)
		if _, err := os.Stat(path); os.IsNotExist(err) {
			continue
		}
		if err := SafeReplace(path, oldText, newText); err != nil {
			return fmt.Errorf("ошибка в файле %s: %v", file, err)
		}
	}
	return nil
}

// Пример использования
func main() {
	// Распаковка документа
	err := UnzipDocx("template.docx", "unzipped")
	if err != nil {
		panic(err)
	}
	defer os.RemoveAll("unzipped")

	// Замена текста во всех частях документа
	err = ReplaceInAllFiles("unzipped", "{tableName1}", "Иван Иванов\nПетр Петров")
	if err != nil {
		panic(err)
	}
	err = ReplaceInAllFiles("unzipped", "{data1}", "дата1")
	if err != nil {
		panic(err)
	}
	err = ReplaceInAllFiles("unzipped", "{data2}", "дата2")
	if err != nil {
		panic(err)
	}
	err = ReplaceInAllFiles("unzipped", "{data3}", "дата3")
	if err != nil {
		panic(err)
	}
	err = ReplaceInAllFiles("unzipped", "{data4}", "дата4")
	if err != nil {
		panic(err)
	}
	err = ReplaceInAllFiles("unzipped", "{data5}", "дата5")
	if err != nil {
		panic(err)
	}

	// err = ReplaceInAllFiles("unzipped", "{template2}", []string{"Иван Иванов", "Петр Петров"})
	// if err != nil {
	// 	panic(err)
	// }

	// Создание нового документа
	err = ZipDocx("unzipped", "modified.docx")
	if err != nil {
		panic(err)
	}

	fmt.Println("DOCX успешно изменен!")
}
