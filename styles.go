package customexcelwriter

import (
	"strings"

	"github.com/xuri/excelize/v2"
)

const DEFAULT_FONT_SIZE = 9

type CellStyle struct {
	BordersList []string //Если пусто - значит все, иначе только со списка
	BorderStyle string   //Пусто - нет рамки, иначе - (thin, thick, dotted, dashed)
	BorderColor string   //Цвет (rgb)

	FontSize   int    //Размер шрифта
	FontBold   bool   //Жирный
	FontNormal bool   //Нормальный вес шрифта
	FontColor  string //Цвет шрифта (rgb)

	FillColor string //Цвет заливки (rgb)

	HorizontalAlignment string //Горизонтальное выравнивание (left, center, right)
	VerticalAlignment   string //Вертикальное выравнивание (top, center)
	DontWrap            bool   //Не переносить текст
}

func (ew ExcelWriter) GetHeaderStyle() *CellStyle {
	return &CellStyle{
		FillColor:           "d6e5cb",
		HorizontalAlignment: "center",
		VerticalAlignment:   "center",
		FontBold:            true,
		BorderStyle:         "thick",
		BorderColor:         "a0a0a0",
	}
}

func (ew ExcelWriter) GetCommonStyle() *CellStyle {
	return &CellStyle{
		FillColor:           "fffad9",
		HorizontalAlignment: "left",
		VerticalAlignment:   "top",
		BorderStyle:         "thin",
		BorderColor:         "a0a0a0",
	}
}

func (cs CellStyle) calculateStyle() *excelize.Style {
	exstyle := new(excelize.Style)

	//Границы
	exstyle.Border = make([]excelize.Border, 0)
	if cs.BorderStyle != "" {
		var bStyle int
		switch cs.BorderStyle {
		case "thin":
			bStyle = 1
		case "thick":
			bStyle = 2
		case "dotted":
			bStyle = 4
		case "dashed":
			bStyle = 3
		default:
			bStyle = 0
		}
		var bColor string
		if cs.BorderColor == "" {
			bColor = rgbColor("000000")
		} else {
			bColor = rgbColor(cs.BorderColor)
		}
		var bList []string
		if len(cs.BordersList) == 0 {
			bList = []string{"top", "left", "bottom", "right"}
		} else {
			bList = cs.BordersList
		}
		for _, bType := range bList {
			exstyle.Border = append(exstyle.Border, excelize.Border{
				Style: bStyle,
				Color: bColor,
				Type:  bType,
			})
		}
	}

	//Заливка
	if cs.FillColor != "" {
		exstyle.Fill = excelize.Fill{Color: []string{rgbColor(cs.FillColor)}, Pattern: 1, Type: "pattern"}
	}

	//Шрифт
	font := new(excelize.Font)
	if cs.FontSize != 0 {
		font.Size = float64(cs.FontSize)
	} else {
		font.Size = DEFAULT_FONT_SIZE
	}
	font.Bold = cs.FontBold && !cs.FontNormal
	font.Color = rgbColor(cs.FontColor)
	exstyle.Font = font

	var alignment excelize.Alignment

	if cs.HorizontalAlignment == "" {
		alignment.Horizontal = "left"
	} else {
		alignment.Horizontal = cs.HorizontalAlignment
	}
	if cs.VerticalAlignment == "" {
		alignment.Vertical = "top"
	} else {
		alignment.Vertical = cs.VerticalAlignment
	}

	//Перенос слов
	alignment.WrapText = !cs.DontWrap

	exstyle.Alignment = &alignment

	return exstyle
}

func rgbColor(c string) string {
	parts := strings.Split(c, "")
	if len(parts) == 0 || parts[0] == "#" {
		return c
	}
	return "#" + c
}
