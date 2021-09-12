package excel

import (
	"os"
	"strconv"

	"github.com/xuri/excelize/v2"
)

func CreateDefaultTable(savePathName string, tableHead Header, data [][]interface{}, sheet string) error {
	file, err := os.OpenFile(
		savePathName,
		os.O_CREATE|os.O_APPEND, os.ModePerm,
	)
	if err != nil {
		return err
	}
	defer file.Close()

	f := excelize.NewFile()
	index := f.NewSheet(sheet)

	// Добавляю шапку
	if err := CreateHead(f, tableHead, sheet); err != nil {
		return err
	}

	// Наполняю таблицу данными
	if err := SetDataToRows(f, data, sheet, 2, 18, []string{GreenRowStyle, LightGreenRowStyle}); err != nil {
		return err
	}

	f.SetActiveSheet(index)
	f.DeleteSheet("Sheet1")

	if err := f.SaveAs(savePathName); err != nil {
		return err
	}
	return nil
}

// Создать шапку таблицы
func CreateHead(f *excelize.File, p Header, sheet string) error {
	// Устанавливаю параметры в столбцы шапки
	for _, v := range p.ColParams {
		f.SetColWidth(sheet, v.StartCol, v.EndCol, float64(v.Width))
	}
	// Устанавливаю значения в столбцы шапки
	for _, v := range p.CellParams {
		f.SetCellValue(sheet, v.Axis, v.Value)
	}
	// Устанавливаю значения ряд шапки
	for _, v := range p.RowParams {
		f.SetRowHeight(sheet, v.Row, float64(v.Height))
	}
	// Создаю стиль для шапки
	style, err := f.NewStyle(p.Style)
	if err != nil {
		return err
	}
	// Устанавливаю стиль для шапки
	if err := f.SetCellStyle(sheet, p.CellParams[0].Axis,
		p.CellParams[len(p.CellParams)-1].Axis,
		style,
	); err != nil {
		return err
	}
	return nil
}

// Создать объект колонок из переданного списка названий
func CreateHeaderCell(cellsList []string, row string) []HeaderCell {
	var cells []HeaderCell
	for i, v := range cellsList {
		cells = append(cells, HeaderCell{
			Axis:  ColCoordinates[i] + row,
			Value: v,
		})
	}
	return cells
}

// Записать данные в таблицу
func SetDataToRows(f *excelize.File, data [][]interface{}, sheet string, startFrom int, rHeight float64, style []string) error {
	var lastColl string
	for i, v := range data {
		for ii, vv := range v {
			f.SetCellValue(sheet, ColCoordinates[ii]+strconv.Itoa(i+startFrom), vv)
			lastColl = ColCoordinates[ii]
		}
		// Высота ряда
		f.SetRowHeight(sheet, i+startFrom, rHeight)
		// Устанавливаю стиль
		if err := setStyleForRow(f, sheet, i, 2, style[0], style[1], []string{"A", lastColl}); err != nil {
			return err
		}
	}
	return nil
}

// Установить стиль для ряда
func setStyleForRow(f *excelize.File, sheet string, i, startFrom int, style1, style2 string, startEndCells []string) error {
	var style int
	if i%2 == 0 {
		s, err := f.NewStyle(style1)
		if err != nil {
			return err
		}
		style = s
	} else {
		s, err := f.NewStyle(style2)
		if err != nil {
			return err
		}
		style = s
	}
	// Устанавливаю стиль
	if err := f.SetCellStyle(
		sheet,
		startEndCells[0]+strconv.Itoa(i+startFrom),
		startEndCells[1]+strconv.Itoa(i+startFrom),
		style,
	); err != nil {
		return err
	}
	return nil
}
