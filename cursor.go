package customexcelwriter

import "github.com/xuri/excelize/v2"

//Переводим координаты в int в строку для Excel
func coordinatesToString(col, row int) string {
	c, _ := excelize.CoordinatesToCellName(col, row)
	return c
}

//Переводим координаты в int в строку для Excel
func (ew *ExcelWriter) cursorToString() string {
	return coordinatesToString(ew.col, ew.row)
}

//Переводим курсор на начало следующего блока
//тоесть на след колонку в которую будем писать
func (ew *ExcelWriter) moveCursorNext() {
	ew.col = ew.writeBlock.ColEnd + 1
}
