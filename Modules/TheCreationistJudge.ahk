#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

AltCreate:
Pwb.document.getElementById("cmdNext") .click
IELoad(Pwb)
Pwb.document.getElementById("cboCallSerNum") .value := SerialNumber
Pwb.document.getElementsByTagName("INPUT")[58] .click
if (A_IsCompiled = 1){
run,Modules\TZAltThread.exe
}else{
run,Modules\TZAltThread.ahk
}
Pwb.document.getElementsByTagName("IMG")[19] .click
Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
SerialNumber:= Pwb.document.getElementById("cboCallSerNum").value
StringUpper, SerialNumber, SerialNumber
ProdCode:= Pwb.document.getElementById("cboJobPartNum").value

If (ProdCode = ""){
Msgbox,4,SerialNumber Not Found,Would you like to attempt the Serial Number again?
IfMsgBox yes
{
InputBox,SerialNumber,Insert Serial Number,Please input the serial number
If SerialNumber
goto, InsertSN
}
TrayTip,Create Wizard,Failed!
return
}
while (MyRO = ""){
	Loop {
		InputBox,myRO,Repair Order Input,Insert your repair Order Number
		IfInString,myRO, 480
			RealRO := true
	}until (realRO = true)
}

RealRO :=""

Pwb.document.getElementById("txtJobRef6") .value :=MyRO

Pwb.document.getElementById("cmdNext") .click
IELoad(Pwb)
SiteNumber:=Pwb.document.getElementById("cboCallSiteNum") .value
Pwb.document.getElementById("cmdNext") .click
IELoad(Pwb)
Pwb.document.getElementById("cboShipSiteName") .value :=""
Pwb.document.getElementsByTagName("Input")[31] .value :=""
Pwb.document.getElementsByTagName("Input")[32] .value :=""
Pwb.document.getElementsByTagName("TEXTAREA")[2] .value :=""
Pwb.document.getElementsByTagName("Input")[35] .value :=""
Pwb.document.getElementsByTagName("Input")[36] .value :=""
Pwb.document.getElementsByTagName("Input")[37] .value :=""
if  (JobType = "ZR2"){
	ShipSite = IMACREP
} else {
	ShipSite = STOKGOODS
}
Pwb.document.getElementById("cboShipSiteNum") .value := ShipSite
if (A_IsCompiled = 1){
run,Modules\TZAltThread.exe
}else{
run,Modules\TZAltThread.ahk
}
Pwb.document.getElementsByTagName("IMG")[15] .click
IELoad(Pwb)
Pwb.document.getElementById("cmdNext") .click
IELoad(Pwb)

