PartList:=""
Iniread,Height,%config%/Modules.ini,T-Enhanced Bench Kit Window Position,GuiY
Iniread,Guixpos,%config%/Modules.ini,T-Enhanced Bench Kit Window Position,GuiX
if (height = "error" or GuiXpos = "error") {
GuiWidth := 267
Height := Taskbar(150)
Guixpos := A_ScreenWidth - GuiWidth
}
iniRead,Eng,%Config%,Engineer,Number
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
frame.document.getElementByID("cboStockSiteNo").value :=Eng
PartsPageLoad:=frame.document.getElementsByTagName("INPUT")[4].value
}
}until (PartsPageLoad = Eng)
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