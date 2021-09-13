package excel

import "fmt"

// Стили для шапки таблиц
const (
	GreenHeadStyle string = `{
		"border":[
			{"type":"left","color":"ffffff","style":1},	
			{"type":"right","color":"ffffff","style":1}
		],		
		"font": {"bold":true,"size":11, "color":"083c01"},
		"alignment": {"horizontal":"center", "vertical":"center", "wrap_text":true},
		"fill":{"type":"pattern","color":["#9BB958"],"pattern":1}
	}`

	BlueHeadStyle string = `{
		"border":[
			{"type":"left","color":"ffffff","style":1},	
			{"type":"right","color":"ffffff","style":1}
		],		
		"font": {"bold":true,"size":11, "color":"FFFFFF"},
		"alignment": {"horizontal":"center", "vertical":"center", "wrap_text":true},
		"fill":{"type":"pattern","color":["#4F81BD"],"pattern":1}
	}`
)

func CustomHeadStyle(fontColor, fillColor string) string {
	return fmt.Sprintf(`{
		"border":[
			{"type":"left","color":"ffffff","style":1},	
			{"type":"right","color":"ffffff","style":1}
		],		
		"font": {"bold":true,"size":11, "color":"%s"},
		"alignment": {"horizontal":"center", "vertical":"center", "wrap_text":true},
		"fill":{"type":"pattern","color":["%s"],"pattern":1}
	}`, fontColor, fillColor)
}

// Варианты стилей для рядов
const (
	GreenRowStyle string = `{
		"border":[
			{"type":"top","color":"ffffff","style":1},	
			{"type":"bottom","color":"ffffff","style":1},
			{"type":"left","color":"ffffff","style":1},	
			{"type":"right","color":"ffffff","style":1}
		],	
		"alignment": {"horizontal":"center", "vertical":"center", "wrap_text":true},
		"fill":{"type":"pattern","color":["#D8E5BA"],"pattern":1}
	}`

	LightGreenRowStyle string = `{
		"border":[
			{"type":"top","color":"ffffff","style":1},	
			{"type":"bottom","color":"ffffff","style":1},
			{"type":"left","color":"ffffff","style":1},	
			{"type":"right","color":"ffffff","style":1}
		],
		"alignment": {"horizontal":"center", "vertical":"center", "wrap_text":true},
		"fill":{"type":"pattern","color":["#ECF0E1"],"pattern":1}
	}`

	WhiteRowStyle string = `{
		"border":[
			{"type":"top","color":"ffffff","style":1},	
			{"type":"bottom","color":"ffffff","style":1},
			{"type":"left","color":"ffffff","style":1},	
			{"type":"right","color":"ffffff","style":1}
		],	
		"alignment": {"horizontal":"center", "vertical":"center", "wrap_text":true},
		"fill":{"type":"pattern","color":["#FFFFFF"],"pattern":1}
	}`

	LightBlueRowStyle string = `{
		"border":[
			{"type":"top","color":"ffffff","style":1},	
			{"type":"bottom","color":"ffffff","style":1},
			{"type":"left","color":"ffffff","style":1},	
			{"type":"right","color":"ffffff","style":1}
		],
		"alignment": {"horizontal":"center", "vertical":"center", "wrap_text":true},
		"fill":{"type":"pattern","color":["#DCE5F0"],"pattern":1}
	}`
)

func CustomRowStyle(fontColor, fillColor string) string {
	return fmt.Sprintf(`{
		"border":[
			{"type":"top","color":"ffffff","style":1},	
			{"type":"bottom","color":"ffffff","style":1},
			{"type":"left","color":"ffffff","style":1},	
			{"type":"right","color":"ffffff","style":1}
		],
		"font": {"bold":true,"size":11, "color":"%s"},
		"alignment": {"horizontal":"center", "vertical":"center", "wrap_text":true},
		"fill":{"type":"pattern","color":["%s"],"pattern":1}
	}`, fontColor, fillColor)
}
