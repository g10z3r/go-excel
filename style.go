package excel

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
)

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
)
