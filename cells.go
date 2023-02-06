package customexcelwriter

import (
	"errors"
)

// Устанавливаем курсор вручную
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
Додаємо hyperlink до активного блоку
link - посилання
cell - Координати ячейки. Якщо нема, приміняємо до активного блоку
*/
func (ew *ExcelWriter) SetHyperLink(link string, cell ...string) {
	var coord string
	if len(cell) == 1 {
		coord = cell[0]
	} else {
		coord = ew.writeBlock.GetFirstCell()
	}

	ew.file.SetCellHyperLink(ew.activeSheet, coord, link, "External")
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
func (ew *ExcelWriter) mergeWorkBlockCells(blocks ...*WorkBlock) error {
	block := &ew.writeBlock
	if len(blocks) > 0 {
		block = blocks[0]
	}
	return ew.MergeCells(block.ColStart, block.RowStart, block.ColEnd, block.RowEnd)
}

/**
* Пишем значение в ячейку
* Координаты устанавливаются относительно последнего блока
 */
func (ew *ExcelWriter) SetCellValueRelatively(value interface{}, offsetY, offsetX int, sizes ...int) (*WorkBlock, error) {
	var tempBlock WorkBlock
	tempBlock.ColStart = ew.writeBlock.ColStart + offsetX
	tempBlock.RowStart = ew.writeBlock.RowStart + offsetY
	if tempBlock.ColStart < 1 || tempBlock.RowStart < 1 {
		return nil, errors.New("Cell is out of sheet")
	}

	tempBlock.ColEnd = tempBlock.ColStart
	if len(sizes) > 0 {
		if sizes[0] > 1 {
			tempBlock.ColEnd = tempBlock.ColStart + sizes[0] - 1
		}
	}

	tempBlock.RowEnd = tempBlock.RowStart
	if len(sizes) > 1 {
		if sizes[1] > 1 {
			tempBlock.RowEnd = tempBlock.RowStart + sizes[1] - 1
		}
	}

	ew.file.SetCellValue(ew.activeSheet, coordinatesToString(tempBlock.ColStart, tempBlock.RowStart), value)

	return &tempBlock, ew.mergeWorkBlockCells(&tempBlock)
}

// Вычисляем размеры текущего блока с которым работаем
func (ew *ExcelWriter) calculateWriteBlock(width, height int) {
	ew.writeBlock.calculateBlockData(ew.col, ew.row, width, height)
}
