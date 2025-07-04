package godocx

import (
	"archive/zip"
	"fmt"
	"io"
	"io/fs"
	"math/rand"
	"os"
	"path/filepath"
	"strings"
	"time"

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

// SafeReplace заменяет текст в XML-файле с сохранением контекста, пробелов и форматирования
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
		// Собираем весь текст параграфа вместе с пробелами
		fullText := ""
		var textNodes []*etree.Element
		var spaceNodes []*etree.Element // Для узлов с xml:space="preserve"

		// Собираем все текстовые элементы в параграфе
		for _, t := range p.FindElements(".//w:t") {
			text := t.Text()
			fullText += text

			// Проверяем, является ли узел пробельным с сохранением
			if spaceAttr := t.SelectAttr("xml:space"); spaceAttr != nil && spaceAttr.Value == "preserve" {
				spaceNodes = append(spaceNodes, t)
			}

			textNodes = append(textNodes, t)
		}

		// Проверяем, содержится ли шаблон в собранном тексте
		start := strings.Index(fullText, oldText)
		if start == -1 {
			continue
		}
		end := start + len(oldText)

		// Ищем узлы, содержащие часть шаблона
		var startNode, endNode *etree.Element
		var startOffset, endOffset int
		currentPos := 0

		for _, t := range textNodes {
			text := t.Text()
			textLen := len(text)

			// Проверяем, находится ли начало шаблона в этом узле
			if startNode == nil && currentPos <= start && start < currentPos+textLen {
				startNode = t
				startOffset = start - currentPos
			}

			// Проверяем, находится ли конец шаблона в этом узле
			if endNode == nil && currentPos < end && end <= currentPos+textLen {
				endNode = t
				endOffset = end - currentPos
			}

			currentPos += textLen
		}

		// Если нашли начало и конец шаблона
		if startNode != nil && endNode != nil {
			// Находим родительский run для вставки
			r := startNode.Parent()
			if r == nil {
				continue
			}

			// Копируем свойства run
			rPr := r.SelectElement("w:rPr")

			switch v := newText.(type) {
			case string:
				// Обработка строки с переносами
				parts := strings.Split(v, "\n")
				trimmedParts := make([]string, len(parts))
				for i, part := range parts {
					trimmedParts[i] = strings.TrimSpace(part)
				}

				// Создаем новые элементы для вставки
				var newElements []*etree.Element
				for i, part := range trimmedParts {
					if i > 0 {
						// Добавляем перенос строки
						br := createElement("br")
						newElements = append(newElements, br)
					}

					if part != "" {
						newT := createElement("t")
						newT.SetText(part)

						// Сохраняем атрибут xml:space если он был в оригинале
						for _, sn := range spaceNodes {
							if sn == startNode || sn == endNode {
								newT.CreateAttr("xml:space", "preserve")
								break
							}
						}

						newElements = append(newElements, newT)
					}
				}

				// Заменяем шаблон новыми элементами
				if startNode == endNode {
					// Шаблон полностью в одном узле
					text := startNode.Text()
					prefix := text[:startOffset]
					suffix := text[endOffset:]

					// Сохраняем пробел перед заменой
					if prefix != "" && strings.HasSuffix(prefix, " ") {
						spaceT := createElement("t")
						spaceT.CreateAttr("xml:space", "preserve")
						spaceT.SetText(" ")
						newElements = append([]*etree.Element{spaceT}, newElements...)
					}

					// Сохраняем пробел после замены
					if suffix != "" && strings.HasPrefix(suffix, " ") {
						spaceT := createElement("t")
						spaceT.CreateAttr("xml:space", "preserve")
						spaceT.SetText(" ")
						newElements = append(newElements, spaceT)
						suffix = strings.TrimPrefix(suffix, " ")
					}

					// Создаем элементы для префикса и суффикса
					if prefix != "" {
						prefixT := createElement("t")
						prefixT.SetText(prefix)
						newElements = append([]*etree.Element{prefixT}, newElements...)
					}

					if suffix != "" {
						suffixT := createElement("t")
						suffixT.SetText(suffix)
						newElements = append(newElements, suffixT)
					}

					// Находим позицию узла у родителя
					parent := startNode.Parent()
					for idx, child := range parent.Child {
						if child == startNode {
							// Удаляем старый узел
							parent.RemoveChildAt(idx)

							// Вставляем новые элементы
							for i, el := range newElements {
								parent.InsertChildAt(idx+i, el)
							}
							break
						}
					}
				} else {
					// Шаблон разбит на несколько узлов

					// 1. Обновляем начальный узел
					startText := startNode.Text()
					startNode.SetText(startText[:startOffset])

					// Сохраняем атрибут xml:space если он был
					if spaceAttr := startNode.SelectAttr("xml:space"); spaceAttr != nil {
						startNode.CreateAttr("xml:space", "preserve")
					}

					// 2. Вставляем новые элементы после начального узла
					parent := startNode.Parent()
					startIdx := -1
					for idx, child := range parent.Child {
						if child == startNode {
							startIdx = idx
							break
						}
					}

					if startIdx != -1 {
						// Вставляем новые элементы
						for i, el := range newElements {
							parent.InsertChildAt(startIdx+1+i, el)
						}
					}

					// 3. Обновляем конечный узел
					endText := endNode.Text()
					endNode.SetText(endText[endOffset:])

					// Сохраняем атрибут xml:space если он был
					if spaceAttr := endNode.SelectAttr("xml:space"); spaceAttr != nil {
						endNode.CreateAttr("xml:space", "preserve")
					}

					// 4. Удаляем промежуточные узлы, но сохраняем пробельные
					startFound := false
					var nodesToRemove []*etree.Element
					for _, t := range textNodes {
						if t == startNode {
							startFound = true
							continue
						}

						if startFound {
							if t == endNode {
								break
							}

							// Не удаляем узлы с xml:space="preserve"
							if spaceAttr := t.SelectAttr("xml:space"); spaceAttr == nil || spaceAttr.Value != "preserve" {
								nodesToRemove = append(nodesToRemove, t)
							}
						}
					}

					for _, t := range nodesToRemove {
						parent := t.Parent()
						if parent != nil {
							parent.RemoveChild(t)
						}
					}
				}

			case []string:
				// Обработка массива строк
				if len(v) == 0 {
					continue
				}

				// Модифицируем первую строку с учетом пробела перед шаблоном
				if start > 0 && fullText[start-1] == ' ' {
					v[0] = " " + v[0]
				}

				// Модифицируем последнюю строку с учетом пробела после шаблона
				if end < len(fullText) && fullText[end] == ' ' {
					v[len(v)-1] = v[len(v)-1] + " "
				}

				// Копируем свойства параграфа
				pPr := p.SelectElement("w:pPr")

				// Заменяем шаблон первой строкой в текущем параграфе
				if startNode == endNode {
					text := startNode.Text()
					startNode.SetText(text[:startOffset] + v[0] + text[endOffset:])
				} else {
					// Для разбитого шаблона
					startText := startNode.Text()
					startNode.SetText(startText[:startOffset] + v[0])

					// Удаляем промежуточные узлы
					startFound := false
					var nodesToRemove []*etree.Element
					for _, t := range textNodes {
						if t == startNode {
							startFound = true
							continue
						}

						if startFound {
							if t == endNode {
								// Обновляем конечный узел
								endText := t.Text()
								t.SetText(endText[endOffset:])
								break
							}
							nodesToRemove = append(nodesToRemove, t)
						}
					}

					for _, t := range nodesToRemove {
						parent := t.Parent()
						if parent != nil {
							parent.RemoveChild(t)
						}
					}
				}

				// Создаем новые параграфы для остальных строк
				if len(v) > 1 {
					parent := p.Parent()
					if parent == nil {
						continue
					}

					// Находим позицию текущего параграфа
					idx := -1
					for i, child := range parent.Child {
						if child == p {
							idx = i
							break
						}
					}

					if idx != -1 {
						// Создаем новые параграфы
						for i := 1; i < len(v); i++ {
							newP := createElement("p")

							if pPr != nil {
								newP.AddChild(pPr.Copy())
							}

							newR := createElement("r")

							if rPr != nil {
								newR.AddChild(rPr.Copy())
							}

							newT := createElement("t")
							newT.SetText(v[i])

							newR.AddChild(newT)
							newP.AddChild(newR)

							parent.InsertChildAt(idx+i, newP)
						}
					}
				}
			}
		}
	}

	// Постобработка: удаление пустых пробельных узлов во всем документе
	for _, t := range doc.FindElements("//w:t") {
		if t.Text() == "" {
			if spaceAttr := t.SelectAttrValue("xml:space", ""); spaceAttr == "preserve" {
				parent := t.Parent()
				if parent != nil {
					t.SetText(" ")
				}
			}
		}
	}

	doc.Indent(2)
	return doc.WriteToFile(xmlPath)
}

// GenerateDocx создает новый DOCX файл с заменой шаблонов
func GenerateDocx(templatePath string, replacements map[string]interface{}) (string, error) {
	// Генерация уникального имени выходного файла
	rand.Seed(time.Now().UnixNano())
	outputFile := fmt.Sprintf("output_%d_%d.docx", time.Now().UnixNano(), rand.Intn(10000))
	absOutput, err := filepath.Abs(outputFile)
	if err != nil {
		return "", fmt.Errorf("ошибка получения абсолютного пути: %v", err)
	}

	// Создание временной директории
	tmpDir, err := os.MkdirTemp("", "docx_*")
	if err != nil {
		return "", fmt.Errorf("ошибка создания временной директории: %v", err)
	}
	defer os.RemoveAll(tmpDir)

	// Распаковка документа
	if err := UnzipDocx(templatePath, tmpDir); err != nil {
		return "", fmt.Errorf("ошибка распаковки: %v", err)
	}

	// Применение замен
	if err := ReplaceMultiple(tmpDir, replacements); err != nil {
		return "", err
	}

	// Создание нового документа
	if err := ZipDocx(tmpDir, absOutput); err != nil {
		return "", fmt.Errorf("ошибка упаковки: %v", err)
	}

	return absOutput, nil
}

// ProcessDocx обрабатывает документ с использованием временной директории
func ProcessDocx(templatePath, outputPath string, replacements map[string]interface{}) error {
	// Создаем уникальную временную директорию
	tmpDir, err := os.MkdirTemp("", "docx_*")
	if err != nil {
		return fmt.Errorf("ошибка создания временной директории: %v", err)
	}
	defer os.RemoveAll(tmpDir) // Удаляем временную директорию после завершения

	// Распаковываем документ
	if err := UnzipDocx(templatePath, tmpDir); err != nil {
		return fmt.Errorf("ошибка распаковки: %v", err)
	}

	// Применяем замены
	if err := ReplaceMultiple(tmpDir, replacements); err != nil {
		return err
	}

	// Создаем новый документ
	return ZipDocx(tmpDir, outputPath)
}

// ReplaceMultiple заменяет несколько шаблонов в документе
func ReplaceMultiple(dir string, replacements map[string]interface{}) error {
	for oldText, newText := range replacements {
		if err := ReplaceInAllFiles(dir, oldText, newText); err != nil {
			return fmt.Errorf("ошибка замены '%s': %v", oldText, err)
		}
	}
	return nil
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
