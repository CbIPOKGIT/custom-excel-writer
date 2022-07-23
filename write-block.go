package customexcelwriter

import "github.com/xuri/excelize/v2"

type WorkBlock struct {
	ColStart int
	ColEnd   int
	RowStart int
	RowEnd   int
}

//Устанавливаем данные текущего рабочего блока
func (wb *WorkBlock) calculateBlockData(startCol, startRow, width, height int) {
	wb.ColStart = startCol
	wb.RowStart = startRow
	if width == 0 {
		width = 1
	}
	wb.ColEnd = startCol + width - 1
	if height == 0 {
		height = 1
	}
	wb.RowEnd = startRow + height - 1
}

//Получаем координаты первой и последней по диагонали ячейки
func (wb WorkBlock) GetFirstLastCells() (string, string) {
	return coordinatesToString(wb.ColStart, wb.RowStart), coordinatesToString(wb.ColEnd, wb.RowEnd)
}

//Возвращаем ширину блока
func (wb WorkBlock) Width() int {
	return wb.ColEnd - wb.ColStart + 1
}

//Возвращаем список колонок блока
func (wb WorkBlock) ColumnsList() *[]string {
	columns := make([]string, 0, wb.Width())
	for col := wb.ColStart; col <= wb.ColEnd; col++ {
		if colString, err := excelize.ColumnNumberToName(col); err == nil {
			columns = append(columns, colString)
		}
	}
	return &columns
}
