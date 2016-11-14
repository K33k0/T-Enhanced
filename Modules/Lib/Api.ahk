#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.

;~ debug:=0
;~ if debug {
;~ #include Functions.ahk
;~ global TesseractVersion:="5.40.14"
;~ }

;~ RO := 
;~ MsgBox % getRO("41-fa516","1770049581")

getRO(SerialNumber,ProductCode){
	temp:=IEvGet(title)
	url= http://hypappbs005/SC5/SC_SerProd/aspx/serprod_modify.aspx?SER_NUM=%SerialNumber%&PROD_NUM=%ProductCode%&SITE_NUM=ZULU
	temp.Navigate2(url, 4096)
	while (RO = "") {
		try {
			temp := IEGet("serprod_modify - " settings.Tesseract )
			RO:=temp.document.getElementById("txtSerReference2").value
			SerNum := temp.document.getElementById("txtSerNum").value 
			OutputDebug, [TE] %RO% - we're are in here you know - %SerNum%
		} catch {
			OutputDebug, [TE] Failed to grab RO attempt %A_index%
			attempts += 1
			sleep, 100
			if (attempts = 100) {
				return Failed
				break
			}
		}
	}
	temp.quit()
return RO	
}

;~ msgbox % GetThreshold("1770049581")

GetThreshold(ProductCode){
	temp:=IEvGet(title)
	url = http://hypappbs005/SC5/SC_Part/aspx/Part_modify.aspx?PART_NUM=%ProductCode%
	temp.Navigate2(url, 4096)
	while (Charge = "") {
		try {
			temp := IEGet("Part_modify - " settings.Tesseract)
			Charge:=temp.document.getElementById("txtPartRepairEnd").value
			OutputDebug, [TE] %RO% - we're are in here you know - %SerNum%
		} catch {
			OutputDebug, [TE] Failed to grab RO attempt %A_index%
			attempts += 1
			sleep, 100
			if (attempts = 100) {
				return Failed
				break
			}
		}
	}
	temp.quit()
	SetFormat,float,0.2
	Threshold := Charge * 2
	
	return Threshold
}