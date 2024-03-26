package customexcelwriter

import (
	"fmt"
	"strconv"

	"github.com/xuri/excelize/v2"
)

const DEFAULT_SHEET_NAME = ""

type ExcelWriter struct {
	file        *excelize.File //Файл с которым мы работаем
	activeSheet string         //Текущий активный лист

	row int //Текущая строка
	col int //Текущая колонка

	writeBlock WorkBlock //Лежат данные последнего блока, который писали
}

func (ew *ExcelWriter) CreateFile() *ExcelWriter {
	ew.file = excelize.NewFile()
	ew.file.SetDefaultFont("Calibri")
	ew.activeSheet = "Sheet1"
	ew.SetCursor(1, 1)
	return ew
}

// Удаляем все листы с книги
func (ew *ExcelWriter) RemoveAllSheets(left ...string) *ExcelWriter {
	for _, sheet := range ew.file.GetSheetList() {
		var keepIt bool
		for _, lName := range left {
			if lName == sheet {
				keepIt = true
				break
			}
		}
		if !keepIt {
			ew.file.DeleteSheet(sheet)
		}
	}
	return ew
}

// Создаем лист с нужным именем и делаем его активным
func (ew *ExcelWriter) CreateSheet(sName ...string) *ExcelWriter {
	var sheetName string
	if len(sName) > 0 {
		sheetName = sName[0]
	} else {
		sheetName = "Worksheet"
	}

	ew.file.NewSheet(sheetName)
	ew.activeSheet = sheetName
	ew.SetCursor(1, 1)

	var defaultRowHeight float64 = 14
	ew.file.SetSheetProps(sheetName, &excelize.SheetPropsOptions{
		DefaultRowHeight: &defaultRowHeight,
	})
	return ew
}

// Создаем листи и оставляем его единственным
func (ew *ExcelWriter) CreateLonelySheet(sName string) *ExcelWriter {
	ew.CreateSheet(sName)
	ew.RemoveAllSheets(sName)
	return ew
}

// Сохраняем файл
func (ew *ExcelWriter) SaveFile(path string) error {
	return ew.file.SaveAs(path)
}

func (ew *ExcelWriter) GetFileContext() ([]byte, error) {
	content, err := ew.file.WriteToBuffer()
	if err != nil {
		return nil, err
	}
	return content.Bytes(), nil
}

// Применяем стиль к блоку ячеек
func (ew *ExcelWriter) ApplyStyle(style *CellStyle, block *WorkBlock) {
	if style == nil {
		return
	}
	var wb WorkBlock
	if block == nil {
		wb = ew.writeBlock
	} else {
		wb = *block
	}
	hcell, wcell := wb.GetFirstLastCells()
	styleIndex, err := ew.file.NewStyle(style.calculateStyle())
	if err != nil {
		return
	}
	ew.file.SetCellStyle(ew.activeSheet, hcell, wcell, styleIndex)
}

// Выровнять высоту заданных или всех строк всех листов файла
// nHeight - высота в пикселях. По умолчанию - 14
func (ew *ExcelWriter) AlignFileRows(nHeight ...int) {
	var height int = 14
	for _, h := range nHeight {
		height = h
	}

	for _, sheet := range ew.file.GetSheetList() {
		rows, err := ew.file.GetRows(sheet)
		if err != nil {
			continue
		}
		for row := 1; row <= len(rows); row++ {
			ew.file.SetRowHeight(sheet, row, float64(height))
		}
	}
}

// Установить ширину колонок
// Если не указываем номер колонки - тогда все колонки блока
func (ew *ExcelWriter) SetColumnsWidth(widht int, cols ...string) {
	if len(cols) > 0 {
		for _, col := range cols {
			ew.file.SetColWidth(ew.activeSheet, col, col, float64(widht))
		}
	} else {
		listColumns := *ew.writeBlock.ColumnsList()
		ew.file.SetColWidth(ew.activeSheet, listColumns[0], listColumns[len(listColumns)-1], float64(widht))
	}
}

// Установить ширину строки
// Если не указано - тогда все строки блока
func (ew *ExcelWriter) SetRowsHeight(height int, rows ...int) {
	if len(rows) > 0 {
		for _, row := range rows {
			ew.file.SetRowHeight(ew.activeSheet, row, float64(height))
		}
	} else {
		for i := ew.writeBlock.RowStart; i <= ew.writeBlock.RowEnd; i++ {
			ew.file.SetRowHeight(ew.activeSheet, i, float64(height))
		}
	}
}

// Перевести каретку в начало строки со сдвигом на rows строк
// По умолчанию rows = 1
func (ew *ExcelWriter) CursorNextLine(rows ...int) {
	cRows := 1
	for _, r := range rows {
		cRows = r
	}
	ew.SetCursor(1, ew.row+cRows)
}

// Даем пользователю текущий рабочий блок
func (ew *ExcelWriter) WorkBlock() *WorkBlock {
	return &ew.writeBlock
}

// Переводим номер колонки (число) в строку
func (ew ExcelWriter) ColumnIndexToString(index int) (string, error) {
	return excelize.ColumnNumberToName(index)
}

// SetPageAutofilters - встановити автофільтри на сторінці
func (ew *ExcelWriter) SetPageAutofilters() error {
	var maxRow, lastColumn string = "1", "A"
	{
		if rows, err := ew.file.GetRows(ew.activeSheet); err == nil {
			maxRow = strconv.Itoa(len(rows))
			lastColumn, _ = ew.ColumnIndexToString(len(rows[0]))
		}
	}

	return ew.file.AutoFilter(ew.activeSheet, fmt.Sprintf("A1:%s%s", lastColumn, maxRow), []excelize.AutoFilterOptions{})
}

// freezePanes заморожує рядки та стовпці вище та лівіше заданих координат.
func (ew *ExcelWriter) FreezePanes(row, col int) error {
	// Перетворюємо координати рядка та стовпця на нотацію Excel.
	cell, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return fmt.Errorf("не вдалося перетворити координати на ім'я комірки: %w", err)
	}

	// Встановлюємо панелі, щоб заморозити рядки та стовпці вище та лівіше заданої комірки.
	if err := ew.file.SetPanes(ew.activeSheet, &excelize.Panes{
		Freeze:      true,       // Увімкнути заморожування
		Split:       false,      // Вимкнути розділення
		XSplit:      0,          // Немає горизонтального розділення
		YSplit:      row,        // Вертикальне розділення на рядку
		TopLeftCell: cell,       // Верхня ліва комірка замороженої області
		ActivePane:  "topRight", // Встановити активну панель у верхньому правому куті (незаморожена область)
	}); err != nil {
		return fmt.Errorf("не вдалося встановити панелі: %w", err)
	}

	return nil
}
