406Launcher:
gui,master:submit, nohide
msgbox, hi Kieran
return


!1::
;{ check stats
if  pwb:=IEgeturl("http://hypappbs005/SC5/SC_RepairJob/aspx/RepairJob_frameset.aspx") {
	msgbox, this function is not available on this page
	pwb:=""
	return
}
FormatTime,Today, YYYYMMDD, dd/MM/yyyy
pwb:=IEVget(Title)
pwb.Navigate2("http://hypappbs005/SC5/SC_RepairJob/aspx/RepairJob_frameset.aspx",2048)
sleep 2500
if not pwb:=IEGeturl("http://hypappbs005/SC5/SC_RepairJob/aspx/RepairJob_frameset.aspx") {
	msgbox,Page not accessable
}
Loop {
	try {
		sleep,250
		frame := Pwb.document.all(10).contentWindow
		PageLoad:= frame.document.getElementsByTagName("LABEL")[0].innertext
		frame := Pwb.document.all(7).contentWindow
		PageLoad2:= frame.document.getElementsByTagName("Input")[0].value
	}
}until (PageLoad = "Job Details" AND PageLoad2 = "submit")

try {
	frame := Pwb.document.all(10).contentWindow
	frame.document.getElementByID("datJobCDate").Value:= Today
	frame.document.getElementByID("cboCallEmployNum").Value:= "406"
	frame := Pwb.document.all(7).contentWindow
	frame.document.getElementByID("cmdSubmit").click
} catch {
	msgbox , failed to pull %value%
}

Loop {
	try {
		frame := Pwb.document.all(10).contentWindow
		TotalRecords:=frame.document.getElementByID("lblRecordCount").innertext
		SingleRecord:=frame.document.getElementByID("cboJobWorkshopSiteNum").value
	}
}Until (TotalRecords != "" OR SingleRecord ="STOWS")

If (SingleRecord = "STOWS") {
	Jobs := 1
}
else
{
	Jobs := TotalRecords
}
pwb.quit
MsgBox % TotalRecords
return
;}

!2::
PartList:=""
Iniread,Height,%config%/Modules.ini,T-Enhanced Bench Kit Window Position,GuiY
Iniread,Guixpos,%config%/Modules.ini,T-Enhanced Bench Kit Window Position,GuiX
if (height = "error" or GuiXpos = "error") {
	GuiWidth := 267
	Height := Taskbar(150)
	Guixpos := A_ScreenWidth - GuiWidth
}
Gui,benchkit:add,Listview, x2 w264 h100 gPartCopy,Part Description|Qty
Gui,benchkit:add,button,w50 h30 gclosebenchkit,quit
Gui,benchkit:Default
Gui,benchkit: -caption +border +ToolWindow +alwaysontop
gui,benchkit:show, w267 h150 Y%Height% x%guixpos%, T-Enhanced Bench Kit Window
pwb:=IEvGet(title)
PartList:=""
iniRead,Eng,%Config%,Engineer,Number
Pwb.Navigate2("http://hypappbs005/SC5/SC_StockControl/aspx/StockControl_Frameset.aspx",2048)
sleep, 2500
Loop 15{
	sleep, 250
	if pwb:=IEGeturl("http://hypappbs005/SC5/SC_StockControl/aspx/StockControl_Frameset.aspx"){
		break
	}
}
if not pwb:=IEGeturl("http://hypappbs005/SC5/SC_StockControl/aspx/StockControl_Frameset.aspx") {
	Gui,Benchkit:Destroy
	PWB:=""
	return
}
loop{
	try{
		frame := Pwb.document.all(10).contentWindow
		PartsPageLoad:=frame.document.getElementByID("lblstock").innertext
	}
}until (PartsPageLoad != "")
if (PartsPageLoad = ""){
	MsgBox, Could Not Access Frame
	return
}
PartsPageLoad := ""
Loop {
	sleep, 100
	try{
		frame.document.getElementByID("cboStockSiteNo").value := settings.Benchkit
		PartsPageLoad:=frame.document.getElementsByTagName("INPUT")[4].value
	}
}until (PartsPageLoad = settings.Benchkit)
PartsPageLoad := ""
frame := Pwb.document.all(7).contentWindow
frame.document.getElementByID("cmdSubmit").click
frame := Pwb.document.all(10).contentWindow
loop{
	try{
		TotalRecs:= frame.document.getElementByID("lblRecordCount").innertext
		if (totalRecs = ""){
			TotalRecs:= frame.document.getElementByID("cmdAdjustCode").Value
		}
	}
}until (TotalRecs !="")
TagNo := 17 + 18 * TotalRecs
If (TotalRecs = "Adjust Code"){
	Goto, AlternatePage2
}
try{
	loop{
		PartDesc:= frame.document.getElementsByTagName("TD")[TagNo].innertext
		Tagno2 := Tagno - 14
		PartCount:=frame.document.getElementsByTagName("TD")[TagNo2].innertext
		TagNo := TagNo - 18
		if (PartDesc != "") {
			LV_Add(vis,PartDesc,PartCount)
			LV_Modifycol(1,"autohdr")
			LV_Modifycol(2,"autohdr integer")
		}
	}Until (TagNo < 34)
}
pwb.quit
return
AlternatePage2:
PartDesc:= frame.document.getElementsByTagName("Input")[8].Value
PartCount:=frame.document.getElementByID("txtTotalQty").Value
if (PartDesc != "") {
	LV_Add(vis,PartDesc,PartCount)
	LV_Modifycol(1,"autohdr")
	LV_Modifycol(2,"autohdr integer")
}
pwb.quit()
return
closebenchkit:
gui,Benchkit:destroy
return
PartCopy:
if A_GUIEvent = doubleclick
{
	LV_GetText(RowText, A_Eventinfo)
	Clipboard := RowText
}
return
