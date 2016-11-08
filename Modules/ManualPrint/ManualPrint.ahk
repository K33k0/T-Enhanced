IniRead, CustomVersion,%config%,VersionNumbers,MyVersions
Notes := ""
gui:
gosub, datagrab
if (getKeyState("Alt","P") = 1 && settings.Engineer = "406"){
	DymoAddIn.Open("Modules/WorkshopCode_Small.label")
	StringUpper, SN, SN
	StringUpper, PC, PC
	DymoLabel.SetField( "Call_Number", Call)
	DymoLabel.SetField( "Part_Code", PC)
	DymoLabel.SetField( "Serial_Number", SN)
	DymoLabel.SetField( "Version_Text", Notes)
	DymoAddIn.Print( 2, TRUE )
	return
	
}



if (SN = ""){
	SN:="Insert Call Number"
	PC:="Insert Product Code"
}



DymoAddIn.Open("Modules/WorkshopCode_Small.label")
DymoLabel.SetField( "Call_Number", Call)
DymoLabel.SetField( "Part_Code", PC)
DymoLabel.SetField( "Serial_Number", SN)
DymoLabel.SetField( "Version_Text", Notes)

Gui,PriSmall:+ToolWindow +alwaysontop +Owner%MasterWindow%
Gui,PriSmall: add,picture,  +hwndcontainer w210 h130,



Gui,PriLarge:+ToolWindow +alwaysontop +Owner%MasterWindow%
Gui,PriLarge: add,picture,  +hwndcontainer2 w340 h130,


Gui,PriInfo: +AlwaysOnTop +ToolWindow +Owner%MasterWindow%
Gui,PriInfo:add,text,,Label Size

;Label Size Options
Gui,Priinfo:Add, Radio, gswitchLabel vrad1 checked,Small Preview
Gui,Priinfo:Add, Radio, gSwitchLabel vrad2,Large Preview

;Product Code
Gui,PriInfo:add,text,,Product Code
Gui,Priinfo:add,edit,vPC gUpdate,%PC%
;Call Number
Gui,PriInfo:add,text,,Call Number
Gui,Priinfo:add,edit,vCall gUpdate,%Call%
;Serial Number
Gui,PriInfo:add,text,,Serial Number
Gui,Priinfo:add,edit,vSN gUpdate,%SN%
;Version Text
Gui,PriInfo:add,text,,Version
Gui,Priinfo:add,edit,vNotes gUpdate,%Notes%
;Print Count
Gui,PriInfo:add,text,,Total Printoffs
Gui,Priinfo:add,edit,vCount
Gui,Priinfo:add,updown

Gui,Priinfo:add,button, gPriPrint,Print

X:=GetWinPosX("T-Enhanced Manual Print Window")
Y:=GetWinPosY("T-Enhanced Manual Print Window")
if (X = "" OR Y = "" OR X= "Error" OR Y="Error"){
	Gui, Priinfo: Show, ,T-Enhanced Manual Print Window
} else {
	Gui, Priinfo: Show, X%x% Y%y%  ,T-Enhanced Manual Print Window
}

switchLabel:
gui,Priinfo:submit,nohide
if (rad1 = 1){
	DymoAddIn.Open("Modules/WorkshopCode_Small.label")
	hDCS := DllCall("GetDCEx", "UInt", container, "UInt", 0, "UInt", 1|2)
	DymoEngine.DrawLabel(hdcS)
	X:=GetWinPosX("Small Preview")
	Y:=GetWinPosY("Small Preview")
	if (X = "" OR Y = "" OR X= "Error" OR Y="Error"){
		X:=0
		Y:=0
	}
	Gui,PriLarge:cancel
	Gui,PriSmall: show, autosize noactivate X%x% Y%y%,Small Preview
	
}else if (rad2 = 1){
	DymoAddIn.Open("Modules/WorkshopCode_Large.label")
	hDCS2 := DllCall("GetDCEx", "UInt", container2, "UInt", 0, "UInt", 1|2)
	DymoEngine.DrawLabel(hdcS2)
	X:=GetWinPosX("Large Preview")
	Y:=GetWinPosY("Large Preview")
	if (X = "" OR Y = "" OR X= "Error" OR Y="Error"){
		X:=0
		Y:=0
	}
	Gui,PriSmall:cancel
	Gui,PriLarge: show, autosize noactivate X%x% Y%y%,Large Preview
}
gosub, Update
return


PriPrint:
gui,Priinfo:submit,nohide
DymoAddIn.Print( Count, TRUE )
return

Update:
gui,Priinfo:submit,nohide
StringUpper, SN, SN
StringUpper, PC, PC
DymoLabel.SetField( "Call_Number", Call)
DymoLabel.SetField( "Part_Code", PC)
DymoLabel.SetField( "Serial_Number", SN)
DymoLabel.SetField( "Version_Text", Notes)
if (rad1 = 1){
	
	DymoEngine.DrawLabel(hdcS)
	
	
}else if (rad2 = 1){
	
	DymoEngine.DrawLabel(hdcS2)
	
}
return

Man_printClose:
PriinfoGuiClose:
PriinfoGuiEscape:
IfWinExist, Small Preview
	SaveWinPos("Small Preview")
IfWinExist, Large Preview
	SaveWinPos("Large Preview")
SaveWinPos("T-Enhanced Manual Print Window")
gui,Priinfo:destroy
gui,PriSmall:destroy
gui,PriLarge:destroy
Pwb:=""
return

DataGrab:
try {
	Pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion)
	frame := Pwb.document.all(10).contentWindow
	SN:=frame.document.getElementById("cboCallSerNum").value
	PC:=frame.document.getElementById("cboJobPartNum").value
	Call:=frame.document.getElementById("txtCallNum").value
	StringUpper, SN, SN
	StringUpper, PC, PC
}

return
