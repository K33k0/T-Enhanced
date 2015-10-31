SN := ""
PC := ""
Note := ""
Count :=
IniRead, CustomVersion,%config%,VersionNumbers,MyVersions
gui:
gosub, datagrab
if (SN = ""){
SN:="Insert Serial Number"
PC:="Insert Product Code"
}

Gui, Input01: Margin, 0, 0
Gui, Input01:Add,text,,Product Code
Gui, Input01:Add, Edit, vPC, %PC%
Gui, Input01:Add,text,,Call Number
Gui, Input01:Add, Edit, vCall +Disabled, %Call%
Gui, Input01:add,text,,Version Number
Gui, Input01:Add, ComboBox, vNotes +Right, %CustomVersion%
Gui, Input01:Add, Button, gSaveNote -TabStop, Save
Gui, Input01:Add, Button,  gDeleteNote -TabStop, Delete
Gui, Input01:add,text,,Serial Number
Gui, Input01:Add, Edit, vSN, %SN%
Gui, Input01:add,text,,Label Count
Gui, Input01:Add, Edit, vCount,
Gui, Input01:Add, UpDown,
Gui, Input01:Add, Button, gPrint,Print
Gui, Input01:  +AlwaysOnTop +ToolWindow +Owner%MasterWindow%
X:=GetWinPosX("T-Enhanced Manual Print Window")
Y:=GetWinPosY("T-Enhanced Manual Print Window")
if (X = "" OR Y = "" OR X= "Error" OR Y="Error"){
Gui, Input01: Show, ,T-Enhanced Manual Print Window
} else {
Gui, Input01: Show, X%x% Y%y%  ,T-Enhanced Manual Print Window
}
return
Input01GuiClose:
Input01GuiEscape:
SaveWinPos("T-Enhanced Manual Print Window")
Gui,Input01:Destroy
return

DataGrab:
try {
Pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion)
frame := Pwb.document.all(10).contentWindow
SN:=frame.document.getElementById("cboCallSerNum").value
PC:=frame.document.getElementById("cboJobPartNum").value
Call:=frame.document.getElementById("txtCallNum").value
}
if (SN  = "") {
try {
Pwb := IEGet("Repair Shipping Wizard - " TesseractVersion)
SN:=Pwb.document.getElementsByTagName("INPUT")[40] .value
PC:=Pwb.document.getElementsByTagName("INPUT")[39] .value
}
}
if (SN = ""){
if not pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion){
SN:="Insert Serial Number"
PC:="Insert Product Code"
}else{
frame := pwb.document.all(11).contentWindow
SN:=frame.document.getElementById("cboFSRSerNum").value
PC:=frame.document.getElementById("cboFSRPartNum").value
}
}
return


Print:
Gui,Input01: Submit
SaveWinPos("T-Enhanced Manual Print Window")
gui,Input01: destroy
StringUpper, SN, SN
StringUpper, PC, PC
if (StrLen(SN) > 15 Or StrLen(PC) > 13){
DymoAddIn.Open("Modules/Workshop Codes Large.label")
MsgBox,48 , Workshop Labels,Insert Large labels. `nclick continue once inserted
} else {
DymoAddIn.Open("Modules/Workshop Codes.label")
}
DymoAddin.StartPrintJob()
DymoLabel.SetField( 1, SN)
DymoLabel.SetField( 2, PC)
DymoLabel.SetField( "JobNumber", Call)
DymoLabel.SetField( 3, Notes)
/*
Gui,Small:+ToolWindow +alwaysontop -caption +lastfound
	Gui,Small: Color, EEAA99
	WinSet, TransColor, EEAA99
	Gui,Small: show,x0 y0 w500 h150  noactivate, DymoDrawSmall
	HWND:=Winexist("DymoDrawSmall")
	hDCS := DllCall("GetDCEx", "UInt", hwnd, "UInt", 0, "UInt", 1|2)
	DymoEngine.DrawLabel(hdcS)
*/
DymoAddIn.Print( Count, TRUE )
DymoAddin.EndPrintJob()
if (StrLen(SN) > 15 Or StrLen(PC) > 13){
MsgBox,48 , Workshop Labels, Reinsert small labels once labels finsih printing, 10
}
Gui,Input01:Destroy
Pwb:=""
return
SaveNote:
Gui, Input01:Submit, Nohide
if (CustomVersion = ""){
CustomVersion = %Notes%
}Else{
CustomVersion =%CustomVersion%|%Notes%
}
StringReplace, CustomVersion, CustomVersion,||,|, All
IniWrite, %CustomVersion%, %config%,VersionNumbers,MyVersions
IniRead, CustomVersion, %config%,VersionNumbers,MyVersions
GuiControl,,Notes,|%CustomVersion%
return
DeleteNote:
Empty:=""
Gui, Input01:Submit, Nohide
StringReplace, CustomVersion, CustomVersion, %Notes%|
IniWrite, %CustomVersion%, %config%,VersionNumbers,MyVersions
Notes:=""
CustomVersion:=""
GuiControl,,Notes,%CustomVersion%
MsgBox, Change will be mode the next time this loads
return