package customexcelwriter

import (
	"errors"
)

//Устанавливаем курсор вручную
func (ew *ExcelWriter) SetCursor(col, row int) *ExcelWriter {
	if col > 0 {
		ew.col = col
	}
	if row > 0 {
		ew.row = row
	}
	return ew
}

// func (ew *ExcelWriter)

/*
Пишем значение в ячейку:
value - значение для записи,
widht - ширина блока (количество колонок для объединения),
height - высота блока (количество строк для объединения)
*/
func (ew *ExcelWriter) SetCellValue(value interface{}, sizes ...int) *WorkBlock {
	width, height := 1, 1
	if len(sizes) >= 1 {
		width = sizes[0]
	}
	if len(sizes) >= 2 {
		height = sizes[1]
	}

	ew.file.SetCellValue(ew.activeSheet, ew.cursorToString(), value)
	ew.calculateWriteBlock(width, height)

	if width > 1 || height > 1 {
		ew.mergeWorkBlockCells()
	}

	ew.moveCursorNext()
	return &ew.writeBlock
}

/*
Обєднуємо декілька ячеєк в одну
*/
func (ew *ExcelWriter) MergeCells(colStart, rowStart, colEnd, rowEnd int) error {
	start := coordinatesToString(colStart, rowStart)
	end := coordinatesToString(colEnd, rowEnd)

	return ew.file.MergeCell(ew.activeSheet, start, end)
}

/*
Обєднуємо ячейки останнього блоку
*/
func (ew *ExcelWriter) mergeWorkBlockCells() error {
	return ew.MergeCells(ew.writeBlock.ColStart, ew.writeBlock.RowStart, ew.writeBlock.ColEnd, ew.writeBlock.RowEnd)
}

/**
* Пишем значение в ячейку
* Координаты устанавливаются относительно последнего блока
 */
func (ew *ExcelWriter) SetCellValueRelatively(value interface{}, offsetY, offsetX, width, height int) (*WorkBlock, error) {
	var tempBlock WorkBlock
	tempBlock.ColStart = ew.writeBlock.ColStart + offsetX
	tempBlock.RowStart = ew.writeBlock.RowStart + offsetY
	if tempBlock.ColStart < 1 || tempBlock.RowStart < 1 {
		return nil, errors.New("Cell is out of sheet")
	}

	tempBlock.ColEnd = tempBlock.ColEnd + width - 1
	tempBlock.RowEnd = tempBlock.RowStart + height - 1

	ew.file.SetCellValue(ew.activeSheet, coordinatesToString(tempBlock.ColStart, tempBlock.RowStart), value)

	return &tempBlock, nil
}

//Вычисляем размеры текущего блока с которым работаем
func (ew *ExcelWriter) calculateWriteBlock(width, height int) {
	ew.writeBlock.calculateBlockData(ew.col, ew.row, width, height)
}
