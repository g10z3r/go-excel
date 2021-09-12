package excel

import (
	"github.com/xuri/excelize/v2"
)

func NewFile(pathName string) error {
	f := excelize.NewFile()
	if err := f.SaveAs(pathName); err != nil {
		return err
	}
	return nil
}
