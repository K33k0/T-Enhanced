IniRead, CustomVersion,%config%,VersionNumbers,MyVersions
gui:
gosub, datagrab
if (getKeyState("Alt","P") = 1 && eng = "406bk"){
DymoAddIn.Open("Modules/Workshop Codes.label")
StringUpper, SN, SN
StringUpper, PC, PC
DymoLabel.SetField( 1, SN)
DymoLabel.SetField( 2, PC)
DymoLabel.SetField( "JobNumber", Call)
DymoLabel.SetField( 3, Notes)
DymoAddIn.Print( 2, TRUE )
return

}
	


if (SN = ""){
SN:="Insert Serial Number"
PC:="Insert Product Code"
}



DymoAddIn.Open("Modules/Workshop Codes.label")
DymoLabel.SetField( 1, SN)
DymoLabel.SetField( 2, PC)
DymoLabel.SetField( "JobNumber", Call)
DymoLabel.SetField( 3, Notes)

	Gui,PriSmall:+ToolWindow +alwaysontop
	Gui,PriSmall: add,picture,  +hwndcontainer w210 h130,

	
	
	Gui,PriLarge:+ToolWindow +alwaysontop 
	Gui,PriLarge: add,picture,  +hwndcontainer2 w340 h130,

	
	
	Gui,PriInfo:add,text,,Label Size
	Gui,Priinfo:Add, Radio, gswitchLabel vrad1 checked,Small Preview
	Gui,Priinfo:Add, Radio, gSwitchLabel vrad2,Large Preview
	Gui,PriInfo:add,text,,Product Code
	Gui,Priinfo:add,edit,vPC gUpdate,%PC%
	Gui,PriInfo:add,text,,Serial Number
	Gui,Priinfo:add,edit,vSN gUpdate,%SN%
	Gui,PriInfo:add,text,,Call Number
	Gui,Priinfo:add,edit,vCall gUpdate,%Call%
	Gui,PriInfo:add,text,,Version
	Gui,Priinfo:add,edit,vNotes gUpdate,%Notes%
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
	DymoAddIn.Open("Modules/Workshop Codes.label")
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
	DymoAddIn.Open("Modules/Workshop Codes Large.label")
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
StringUpper, SN, SN
StringUpper, PC, PC
DymoAddIn.Print( Count, TRUE )
return

Update:
gui,Priinfo:submit,nohide
DymoLabel.SetField( 1, SN)
DymoLabel.SetField( 2, PC)
DymoLabel.SetField( "JobNumber", Call)
DymoLabel.SetField( 3, Notes)
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

