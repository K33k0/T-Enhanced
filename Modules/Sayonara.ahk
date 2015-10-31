Pwb := ""
CallNum := ""
ProdCode := ""
SerialNumber := ""
CustomerDamageCheck := ""
ShipSite := ""
GuiWidth := ""
height := ""
Guixpos := ""
CustomerDamage := ""
IniRead, CustomVersion,%config%,VersionNumbers,MyVersions
WinActivate,ahk_class IEFrame
if not Pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion){
	MsgBox Error accessing page
	return
} else {
	frame := Pwb.document.all(10).contentWindow
	CallNum:= frame.document.getElementsByTagName("INPUT")[0] .value
	sleep, 250
	ProdCode:= frame.document.getElementByID("cboJobPartNum").value
	SerialNumber:= frame.document.getElementByID("cboCallSerNum").value
	CustomerDamageCheck:= frame.document.getElementByID("cboCallProbCode").value
	If (frame.document.getElementById("cboJobShipSiteNum").value = "ZULU") {
		goto ZuluShip
		return
}
	if not Pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion)
		MsgBox Error accessing page
	else {
		sleep, 250
		frame := Pwb.document.all(10).contentWindow
		ShipSite:= frame.document.getElementById("cboJobShipSiteNum") .value
		sleep, 250
		frame := Pwb.document.all(9).contentWindow
		frame.document.getElementById("lblJobShipOutWizard") .click
	}
	IELoad(pwb)
	Pwb.document.getElementById("txtInputJobNum") .value :=CallNum
	Pwb.document.getElementById("cmdAddJobNum") .click
	Pwb.document.getElementById("cmdNext") .click
	IELoad(Pwb)
	Pwb.document.getElementById("cmdNext") .click
	IELoad(Pwb)
	Pwb.document.getElementsByTagName("INPUT")[48] .click

	Iniread,Height,%config%,T-Enhanced Shipout Window Position,GuiY
	Iniread,Guixpos,%config%,T-Enhanced Shipout Window Position,GuiX
	if (height = "error" or GuiXpos = "error") {
		GuiWidth := 534
		Height := Taskbar(150)
		Guixpos := A_ScreenWidth - GuiWidth
	}
	gui,ShipoutGui:Add, Picture,y0 x247 h10 w-1 gmovegui 0x4000000, Modules\Img\move.png
	Gui, ShipoutGui:Font, s11
	gui,ShipoutGui:Add, Text, x0 y0 w247 center, Shipping %CallNum% to %ShipSite%
	Gui, ShipoutGui:Font, s8
	Gui, ShipoutGui:Add, Text, x0 y+5 w267 center BackgroundTrans,Enter Version Number
	Gui, ShipoutGui:Add, ComboBox, x33 y+5 w200 vNotes, %CustomVersion%
	Gui, ShipoutGui:Add, Button, X181 y+5 w50 gSaveNoteShipout -TabStop, Save
	Gui, ShipoutGui:Add, Button, X34 ym+60 w50 gDeleteNoteShipout -TabStop, Delete
	Gui, ShipoutGui:Add, Button, X207 y+30 w50 gNext vShipOutGuiNext +Default Disabled, Ship
	Gui, ShipoutGui:Add, Button, X7 ym+112 w50 gCancelShip, Cancel
	if (CustomerDamageCheck = "CDAM"){
		CustomerDamage = This call is marked as customer damage
	} else {
		CustomerDamage = This call is not marked as customer damage
	}
	Gui, ShipoutGui:Font, s11
	Gui, ShipoutGui:Add, Text, x65 y100 w130 cFF0000 Center, %CustomerDamage%
	Gui, ShipoutGui:Font, s8
	Gui, ShipoutGui:-Caption +Border +AlwaysOnTop +ToolWindow +Owner%MasterWindow%
	Gui, ShipoutGui:Show, x%Guixpos% y%height% h150 w267, T-Enhanced Shipout Window
	gosub, ConfirmGuiWait
	return
	ConfirmGuiWait:
	IniRead,russ,%Config%,Engineer,Number
	If (russ = "430BK"){
		if (shipsite != "STOKGOODS") {
			ComObjCreate("SAPI.SpVoice").Speak("You are not shipping to Stoke Goods")
		}
	}
	sleep, 3000
	GuiControl,ShipoutGui:Enabled,ShipOutGuiNext
	return
	Next:
	gui,ShipoutGui:submit
	SerialNumber:=""
	ProdCode:=""
	SerialNumber:=Pwb.document.getElementbyID("cbaListCallSerNumLineArray").value
	ProdCode:=Pwb.document.getElementbyID("cbaListJobPartNumLineArray").value
	if(Prodcode = "STOKVP"){
	}else{
		If (strLen(SerialNumber)>15 Or strlen(Prodcode) > 13){
			DymoAddIn.Open("Modules\Workshop Codes Large.label")
		MsgBox,48 , Workshop Labels,Insert Large labels. `nclick continue once inserted
		}else{
			DymoAddIn.Open("Modules\Workshop Codes.label")
		}
	}
	Pwb.document.getElementById("cmdFinish") .click
	if (A_IsCompiled = 1){
		run,Modules\TZAltThread.exe
	}else{
		run,Modules\TZAltThread.ahk
	}
	IELoad(Pwb)
	Pwb.document.getElementsByTagName("INPUT")[40] .click
	Pwb.document.getElementById("cmdFinish") .click
	pwb:=""
	if(Prodcode != "STOKVP"){
		DymoAddin.StartPrintJob()
		DymoLabel.SetField( 1, SerialNumber )
		DymoLabel.SetField( 2, ProdCode )
		DymoLabel.SetField( 3, Notes )
		DymoAddIn.Print( 2, TRUE )
		DymoAddin.EndPrintJob()
	If (strLen(SerialNumber)>15 Or strlen(Prodcode) > 13){
		msgbox,48,Workshop Labels, Reinsert small labels
	}
	Gui,ShipoutGui:Destroy
	;~ gosub, UpdateCheck
	pwb:=""
	return
	}
	Gui,ShipoutGui:Destroy
	;~ gosub, UpdateCheck
	pwb:=""
	return
}
cancelShip:
Process,Close,Msgbox_Close.exe
Gui,ShipoutGui:Destroy
OutputDebug, Shipout Canceled within cancel ship label
TP_Show("Automatic Shipout Cancelled")
return
SaveNoteShipout:
Gui, ShipoutGui:Submit, Nohide
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
DeleteNoteShipout:
Empty:=""
Gui, ShipoutGui:Submit, Nohide
StringReplace, CustomVersion, CustomVersion, %Notes%|
IniWrite, %CustomVersion%, %config%,VersionNumbers,MyVersions
Notes:=""
CustomVersion:=""
GuiControl,,Notes,%CustomVersion%
MsgBox, Change will be mode the next time this loads
return

ZuluShip:
OutputDebug, [T-Enhanced] Zulu Ship function activated
Loop{
Try{
frame := Pwb.document.all(10).contentWindow
PageLoaded:= frame.document.getElementsByTagName("Label")[0].innertext
}
}Until (PageLoaded = "Job Details")
PageLoaded:=""
frame := Pwb.document.all(6).contentWindow
f6td1 = <TD width="25`%"><DIV style="Color:Red; height:100`%; text-Align:center; font:20">Powered by <br>T-Enhanced</br></DIV></TD>
frame.document.getElementsBytagName("td")[1].innerhtml := f6td1
frame := Pwb.document.all(10).contentWindow
frame.document.getElementByID("cboJobFlowCode").value:="ZULUAW"
frame.document.getElementByID("cboCallAreaCode").value:="WSF"

sleep, 250
frame := Pwb.document.all(7).contentWindow
Loop{
Try{
PageLoaded:= frame.document.getElementByID("cmdSubmit").value
}
}Until (PageLoaded = "submit")
PageLoaded:=""
frame.document.getElementById("cmdSubmit").click
frame := Pwb.document.all(10).contentWindow
frame.document.getElementById("cboCallUpdAreaCode").value := "WSF"
if (A_IsCompiled = 1){
run,Modules\TZAltThread.exe
}else{
run,Modules\TZAltThread.ahk
}
frame.document.getElementsByTagName("IMG")[35].click
WinWaitClose,Popup List -- Webpage Dialog
if (A_IsCompiled = 1){
run,Modules\TZAltThread.exe
}else{
run,Modules\TZAltThread.ahk
}
frame := Pwb.document.all(7).contentWindow
frame.document.getElementById("cmdSubmit").click
Pwb:=""
WinWaitClose, Message from webpage
OutputDebug, [T-Enhanced] Zulu Ship function Ended
return