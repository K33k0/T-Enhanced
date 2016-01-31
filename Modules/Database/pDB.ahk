global IBM := {"Base":["41A2902MAINBRD","CLFCMOSBAT","CLFMEM1GB","HDDSATAFOOD","RISERCRD-41A2903","WW153703","44V1999-PSU","44V2026 I/O CARD","44V2036 RISER","44V2129 MAINBOARD","CLFCMOSBAT","MEM-1GB DDR2"]
,"Printer":["part1","part2"]
,"PocketPC":[]
,"Server":[]}

PartCenter := {"mess": test()}

Partcenter.mess()

test(){
msgbox % IBM.Base[12]
}
