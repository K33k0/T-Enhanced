if not pwb := getServiceReportPointer()
	goto, Service_cancel

ProdCode:=SERVICEREPORT_getProdCode(pwb)
JobType:=SERVICEREPORT_getJobType(pwb)
RO:=SERVICEREPORT_getRoNumber(pwb)
if not SERVICEREPORT_bIsNewReport(pwb){
	SERVICEREPORT_openNewReport(pwb)
}
SERVICEREPORT_InjectEnhanced(pwb)
SERVICEREPORT_getTimeSinceLastReport()
SERVICEREPORT_getAverageRepairTime(ProdCode)

Gui,ServiceReportGui:  +AlwaysOnTop +ToolWindow +Owner%MasterWindow%
Gui,ServiceReportGui: Add, Text, x80 y10 w149 h13, Choose the repair code
Gui,ServiceReportGui: Add, Text, x80 y55 w146 h13, Choose the fault code
Gui,ServiceReportGui: Add, Text, x95 y100 w146 h13, Enter your solution
Gui,ServiceReportGui: Add, Text, x250 y205 w50 h13, Next area
Gui,ServiceReportGui: Add, Text, x10 y205 w140 h15,Time Taken (minutes)
Gui,ServiceReportGui: Add, Text, x70 y220 w170 h25 Center, % "Time Since last completed job`n" SERVICEREPORT_getTimeSinceLastReport() " minutes"
Gui,ServiceReportGui: Add, Text, x70 y+10 w170 h25 Center BackgroundTrans,% "Usually these take you`n" SERVICEREPORT_getAverageRepairTime(ProdCode) " minutes"
Gui,ServiceReportGui: Font, Norm
Gui,ServiceReportGui: Add, DropDownList, x10 y30 w300 Sort vRep, Repaired|Cleaned|Replaced HDD|Reimaged|Replaced Part|No Fault Found|Flashed Firmware|BER|Awaiting Spares|Warranty Repairs|Datalogic ELF Power Issue|Damaged Due To CLF|CLF Not Attempted|Reconfigured
Gui,ServiceReportGui: Add, DropDownList, x10 y75 w300 Sort vFault, Epos Fault|Server Fault|Pocket PC Fault|Printer Consumable Fault|Printer Fault|Self Checkout Issue|Software Epos Fault|Software Workstation Fault
Gui,ServiceReportGui: Add, Edit, x10 y120 w300 h80 vsolution,
Gui,ServiceReportGui: Add, edit, x10 y220 w50 h20,
Gui,ServiceReportGui: Add,UpDown,vVarTime Range1-1000, % SERVICEREPORT_getTimeSinceLastReport()
Gui,ServiceReportGui: Add, DropDownList, x240 y220 w70 Sort vNextArea, WREP|WSB|WSF|3RDP|APC1|BW3RP
Gui,ServiceReportGui: Add, Checkbox, x15 y245 w75 h13 vAddParts gReadPartsInStock, Add Parts?
Gui,ServiceReportGui: Add, Checkbox, x215 y245 w92 h13 vItemRepaired, Item Repaired?
Gui,ServiceReportGui: Font, Bold
Gui,ServiceReportGui: font, s12
Gui,ServiceReportGui: Font, norm
Gui,ServiceReportGui: font, s8
Gui,ServiceReportGui: Add, Button, x255 y265 w54 h23 gCompleteServiceReport, Continue
X:=GetWinPosX("T-Enhanced Service Report Window")
Y:=GetWinPosY("T-Enhanced Service Report Window")
if (X = "" OR Y = "" OR X= "Error" OR Y="Error"){
	Gui, ServiceReportGui: Show, , T-Enhanced Service Report Window
} else {
	Gui, ServiceReportGui: Show, X%x% Y%y%  , T-Enhanced Service Report Window
}
return

CompleteServiceReport:
Gui,ServiceReportGui:Submit,Nohide
SaveWinPos("T-Enhanced Service Report Window")
if (Rep = "" OR Fault = "" OR Solution = "" OR VarTime = 0 OR NextArea = ""){
	Msgbox, Complete all fields to continue
	return
}
Gui,ServiceReportGui:Destroy


if not pwb := getServiceReportPointer()
	goto, Service_cancel



frame := Pwb.document.all(11).contentWindow
frame.document.getElementsByTagName("INPUT")[13].value :=  SERVICEREPORT_getCurrentTime(varTime)
frame.document.getElementsByTagName("INPUT")[16].value := SERVICEREPORT_getEndTime()
Call:=frame.document.getElementsByTagName("INPUT")[0].value
frame.document.getElementById("cboFSREmployNum").value := settings.engineer
frame.document.getElementById("cboFSRSympCode").value := "WSHOP"
ModalDialogue()
frame.document.getElementsByTagName("IMG")[11].click
frame.document.getElementById("cboFSRFaultCode").value := SERVICEREPORT_TranslateFault(Fault)
ModalDialogue()
frame.document.getElementsByTagName("IMG")[12].click
frame.document.getElementById("cboFSRRepCode").value := ((rep = "Repaired")?(3):(((rep = "Cleaned")?(4):(((rep = "Reconfigured")?(5):(((rep = "Replaced HDD")?(15):(((rep = "Reimaged")?(16):(((rep = "Replaced Part")?(17):(((rep = "No Fault Found")?(18):(((rep = "Flashed Firmware")?(19):(((rep = "BER")?(21):(((rep = "Awaiting Spares")?(31):(((rep = "Warranty Repairs")?(32):(((rep= "Datalogic ELF Power Issue")?(40):(((rep = "Damaged Due To CLF")?(44):(((rep = "CLF Not Attempted")?(46):())))))))))))))))))))))))))))
CLF:=((rep = "Damaged Due To CLF")?(True):(((rep = "CLF Not Attempted")?(True):(False))))
ModalDialogue()
frame.document.getElementsByTagName("IMG")[13].click

;additional information in solution
IfWinExist, Part Add 
{
	Gui,AddPartsGui:submit,NoHide
	AdditionalInfo = `n---[T-Enhanced]---`n
	I=1
	while (PartDesc%I% != "")
	{
		if (Reason%I% = "") {
			Reason%I% := "No Information given"
		}
		AdditionalInfo := AdditionalInfo .  "[--- Part" . I . "---]`n Replaced " . PartDesc%I% . "`nReason - " . Reason%I% . "`n"
		
		I+=1
	}
	Solution := Solution . AdditionalInfo
}

frame.document.getElementsByTagName("TEXTAREA")[0].value := solution
SerialNumber:=frame.document.getElementByID("cboFSRSerNum").value
ProdCode:=frame.document.getElementByID("cboFSRPartNum").Value
frame := Pwb.document.all(8).contentWindow
frame.document.getElementById("cmdSubmit").click
IELoad(pwb)
frame := Pwb.document.all(11).contentWindow

if (ItemRepaired = 1){
	checkboxDisabled =<INPUT onclick="PopulateCheckBox('chkJobComplete')" tabIndex=460 disabled class=XL_formcheckbox type=checkbox value="" name=chkJobComplete>
	
	
	Loop{
		VerifyRepairBefore:=frame.document.getElementsByTagName("INPUT")[52].OuterHTML
		OutputDebug, Checkbox = %VerifyRepairBefore%
		if (VerifyRepairBefore = checkboxDisabled){
			skipcheckbox := true
			OutputDebug, skipcheckbox = %skipcheckbox%
		}
		if (Rep != "ber" || Rep != "Awaiting Spares") {
			frame.document.getElementsByTagName("INPUT")[52].click
		}
		VerifyRepairAfter:=frame.document.getElementsByTagName("INPUT")[52].OuterHTML
	}until (VerifyRepairBefore != VerifyRepairAfter OR SkipCheckbox = True)
	
	iniread,TotalCompleted,Modules\ProductsAverage.ini,%ProdCode%,Total
	iniread,TotalTime,Modules\ProductsAverage.ini,%ProdCode%,TotalTime
	if (totalCompleted = "error") {
		totalCompleted = 0
	}
	if (TotalTime = "error") {
		TotalTime = 0
	}
	TotalCompleted+=1
	TotalTime+=VarTime
	AverageTime:=TotalTime / TotalCompleted
	IniWrite,%TotalCompleted%,Modules\ProductsAverage.ini,%Prodcode%,Total
	iniWrite,%TotalTime%,Modules\ProductsAverage.ini,%ProdCode%,TotalTime
	IniWrite,%AverageTime%,Modules\ProductsAverage.ini,%ProdCode%,Average
	RegWrite,REG_SZ,HKEY_CURRENT_USER,Software\TesseractZoanthropy,Last_SP,%A_Now%
}
sleep, 500
frame.document.getElementById("cboCallAreaCode").value := NextArea
ModalDialogue()
frame.document.getElementsByTagName("IMG")[18].click
WinWaitClose, Popup List -- Webpage Dialog
frame := Pwb.document.all(8).contentWindow
PageAlert()

frame.document.getElementById("cmdSubmit").click
WinWaitClose, message from webpage,,5
sleep,1500
frame := Pwb.document.all(11).contentWindow

/* if (ItemRepaired = 1 && JobType != "FRN" OR Rep = "BER" ) {
	* 	gosub, ChargeCodes
	* }
*/
If (AddParts = 1){
	Gosub, AddParts
	goto,EndServiceReport
	return
}
If ( ItemRepaired  != 1) {
	frame := Pwb.document.all(10).contentWindow
	frame.document.getElementById("lblRepairJob").click
} else {
	frame := Pwb.document.all(10).contentWindow
	frame.document.getElementById("lblRepairJob").click
}
goto,EndServiceReport
return
EndServiceReport:
If (CLF = True OR Rep = "Awaiting Spares"){
	Goto, CLF
}
If (Rep = "BER" or Rep = 21){
	Filecopy,%A_ScriptDir%/modules/BerForm.docx,%A_Temp%/%RO%.docx,
	Run,%A_Temp%/%RO%.docx
}
gosub, Service_cancel
return
ReadPartsInStock:
Gui,ServiceReportGui:Submit, Nohide
If (AddParts !=1){
	try{
		Gui,AddPartsGui:Destroy
		PartList:=""
		return
	}
}
PartList:=""
Pwb.Navigate2("http://hypappbs005/SC5/SC_StockControl/aspx/StockControl_Frameset.aspx",2048)
sleep, 2500
Loop 15{
	sleep, 250
	if pwb:=IEGeturl("http://hypappbs005/SC5/SC_StockControl/aspx/StockControl_Frameset.aspx"){
		break
	}
}
if not pwb:=IEGeturl("http://hypappbs005/SC5/SC_StockControl/aspx/StockControl_Frameset.aspx") {
	msgbox, Error at 347 Service Report
	gosub, Service_cancel
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
		frame.document.getElementByID("cboStockSiteNo").value := settings.BenchKit
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
	Goto, AlternatePage
}
try{
	loop{
		PartDesc:= frame.document.getElementsByTagName("TD")[TagNo].innertext
		Tagno2 := Tagno - 14
		PartCount:=frame.document.getElementsByTagName("TD")[TagNo2].innertext
		TagNo := TagNo - 18
		PartList=%PartDesc%|%PartList%
	}Until (TagNo < 34)
}catch{
	If (Attempted = True){
		Msgbox, Error on parts read
		Pwb.Quit()
		return
	}
	Attempted:= True
	goto, ReloadAddParts
}
goto, AddPartsGui
return
AlternatePage:
PartDesc:= frame.document.getElementsByTagName("Input")[8].Value
PartCount:=frame.document.getElementByID("txtTotalQty").Value
PartList=%PartDesc%|%PartList%
goto,AddPartsGui
AddPartsGui:
pwb.quit()
pwb:=""
Gui,ServiceReportGui:Submit, Nohide
If (AddParts !=1){
	try{
		Gui,AddPartsGui:Destroy
		PartList:=""
		return
	}
}
LineNo := 1

;Gui,AddPartsGui:Add,Button,y5 x5 w50 h20 vAddPartsReload gReloadAddParts, Reload
Gui,AddPartsGui:Add,Text, yp+10 xp+50 h20 w60 BackgroundTrans, Select Part
Gui,AddPartsGui:Add,Text, yp xp+113 h20 w60 BackgroundTrans, Qunatity
Gui,AddPartsGui:Add, DDL, y30 x5 w160 vPartDesc%LineNo% gDynamicAddPartsLines, %PartList%
Gui,AddPartsGui:Add, edit, Yp w40 h20 x+5
Gui,AddPartsGui:Add, Updown, vQuantity%LineNo%
Gui,AddPartsGui:add, checkbox,yp x+5 vBillable%LineNo%, Billable?
Gui,AddPartsGui:Add, Text, Yp+28 w40 h20 x5, Reason
Gui,AddPartsGui:Add, edit, Yp-3 w160 h20 x+5 vReason%LineNo%,
Gui,AddPartsGui:  +AlwaysOnTop +Border +Owner%MasterWindow%
X:=GetWinPosX("Part Add")
Y:=GetWinPosY("Part Add")
if (X = "" OR Y = "" OR X= "Error" OR Y="Error"){
	Gui, AddPartsGui: Show, , Part Add
} else {
	Gui, AddPartsGui: Show, X%x% Y%y%  , Part Add
}
return
AddParts:
If Not Pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion){
	MsgBox Error accessing page
	Pwb:=""
	return
}

frame := Pwb.document.all(10).contentWindow
frame.document.getElementById("lblServiceReportLines").click

IELoad(pwb)
SaveWinPos("Part Add")
Gui,AddPartsGui:Submit

I:=1
LineNo-=1
loop, %LineNo%{
	Loop{
		If (Quantity%I% = 0 OR Quantity%I% = ""){
			inputbox, Quantity%I%, Error - Invalid Quantity, Insert Quantity
		}
	}until (Quantity%I% != 0 AND Quantity%I% != "" )
	StringReplace,PartDesc%I%,PartDesc%I%,%A_SPACE%,`%,All
	Pwb := IEGet("FSRL_Create_Wzd - " TesseractVersion)
	ModalDialogue()
	sleep,500
	Pwb.document.getElementById("cboWZPartDesc").value :=PartDesc%I%
	Pwb.document.getElementById("cboWZPartDesc_Container").click
	check:=Pwb.document.getElementById("cboWZPartNum").value
	if(check = ""){
		MsgBox,Part not found. Input Manually
		IEload(Pwb)
	}else{
		Pwb.document.getElementById("cmdNext").Click
		IEload(Pwb)
	}
	iniread,PreQuantity,Modules\Database\Parts.ini,%Prodcode%,Quantity
	If (PreQuantity = "error"){
		PreQuantity = 0
	}
	PreQuantiy+=Quantity%i%
	iniwrite,%PreQuantiy%,Modules\Database\Parts.ini,%ProdCode%,%check%
	if(Pwb.document.getElementById("chkAllowNewStockUsed").outerHTML = 	"<INPUT id=chkAllowNewStockUsed type=checkbox name=chkAllowNewStockUsed>"){
		Pwb.document.getElementById("chkAllowNewStockUsed").Click
	}
	Pwb.document.getElementById("txtFSRLQty").value :=Quantity%I%
	if(Billable%I% = 1){
		if(Pwb.document.getElementById("chkFSRLBillable").outerHTML = "<INPUT id=chkFSRLBillable CHECKED type=checkbox name=chkFSRLBillable>"){
		}else{
			Pwb.document.getElementById("chkFSRLBillable").Click
		}
	}else{
		if(Pwb.document.getElementById("chkFSRLBillable").outerHTML = "<INPUT id=chkFSRLBillable CHECKED type=checkbox name=chkFSRLBillable>"){
			Pwb.document.getElementById("chkFSRLBillable").Click
		}
	}
	Pwb.document.getElementById("cmdFinish").click
	IEload(Pwb)
	I+=1
	if (PartDesc%I% = ""){
		Pwb.document.getElementById("optNextActionModifyCall").click
		Pwb.document.getElementById("cmdFinish").click
		Gui,AddPartsGui:Destroy
		return
	}else{
		Pwb.document.getElementById("cmdFinish").click
		IEload(pwb)
	}
}
Gui,AddPartsGui:destroy
Pwb:=""
return
ReloadAddParts:

Gui,AddPartsGui:Destroy
PartList:=""
Goto, ReadPartsInStock
return
DynamicAddPartsLines:
Gui,AddPartsGui:submit, nohide
If (PartDesc%LineNo% = ""){
	return
}
LineNo:=LineNo + 1
Gui,AddPartsGui:Add, DDL, yp+30 x5 w160 vPartDesc%LineNo% gDynamicAddPartsLines, %PartList%
Gui,AddPartsGui:Add, edit, Yp w40 h20 x+5
Gui,AddPartsGui:Add, Updown, vQuantity%LineNo%
Gui,AddPartsGui:add, checkbox,yp x+5 vBillable%LineNo%, Billable?
Gui,AddPartsGui:Add, Text, Yp+28 w40 h20 x5, Reason
Gui,AddPartsGui:Add, edit, Yp-3 w160 h20 x+5 vReason%LineNo%,
if (X = "" OR Y = "" OR X= "Error" OR Y="Error"){
	Gui, AddPartsGui: Show, AutoSize, Part Add
} else {
	Gui, AddPartsGui: Show, X%x% Y%y%  AutoSize, Part Add
}
return
CLF:
OutputDebug, [T-Enhanced] CLF function activated
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
if (rep = "Damaged Due To CLF"){
	frame.document.getElementsByTagName("INPUT")[99].click
	frame.document.getElementsByTagName("INPUT")[95].click
	frame.document.getElementsByTagName("INPUT")[90].value := "Y"
}
If (rep = "CLF Not Attempted"){
	frame.document.getElementsByTagName("INPUT")[90].value := "Y"
	frame.document.getElementsByTagName("INPUT")[96].click
}

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
frame.document.getElementById("cboCallUpdAreaCode").value := NextArea
ModalDialogue()
frame.document.getElementsByTagName("IMG")[35].click
WinWaitClose,Popup List -- Webpage Dialog
PageAlert()
frame := Pwb.document.all(7).contentWindow
frame.document.getElementById("cmdSubmit").click
Pwb:=""
WinWaitClose, Message from webpage
OutputDebug, [T-Enhanced] CLF function Ended
return

ChargeCodes: ; No call to here
OutputDebug, [T-Enhanced] Begin Charge Code Function
If Not Pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion){
	MsgBox Error accessing page
	Pwb:=""
	return
}
frame := Pwb.document.all(10).contentWindow
frame.document.getElementById("lblServiceReportLines").click
OutputDebug, [T-Enhanced] Loading Page
IELoad(pwb)
OutputDebug, [T-Enhanced] Page Loaded
ReadyCheck:=""
try{
	Readycheck:=Frame.document.getElementByID("cmdAdjustCode").Value
}
If (ReadyCheck !=""){
	frame := Pwb.document.all(9).contentWindow
	frame.document.getElementById("lblServiceReportLines").click
	IELoad(pwb)
}
Readycheck:=""

Pwb := IEGet("FSRL_Create_Wzd - " TesseractVersion)
if (A_IsCompiled = 1){
	run,Modules\TZAltThread.exe
}else{
	run,Modules\TZAltThread.ahk
}
sleep,500
Pwb.document.getElementById("cboWZChargeCode").value :="ZULU%"
Pwb.document.getElementById("cboWZChargeCode_Container").click
Pwb.document.getElementById("cmdNext").Click
IEload(Pwb)

if(Pwb.document.getElementById("chkFSRLBillable").outerHTML = "<INPUT id=chkFSRLBillable CHECKED type=checkbox name=chkFSRLBillable>"){
}else{
	Pwb.document.getElementById("chkFSRLBillable").Click
}

Pwb.document.getElementById("cmdFinish").click
IEload(Pwb)
Pwb.document.getElementById("optNextActionModifyFSRL").click
Pwb.document.getElementById("cmdFinish").click
OutputDebug, [T-Enhanced] Page Loading
IELoad(pwb)
OutputDebug, [T-Enhanced] Page Loaded
OutputDebug, [T-Enhanced] Completed Charge Code Function
return

ServiceReportGuiGuiClose:
ServiceReportGuiGuiEscape:
Service_cancel:
pwb:=""
LineNo:=""
MultipleRecords:=""
LastSP:=""
PartList:=""
NextArea:=""
ProdCode:=""
SerialNumber:=""
Solution:=""
Rep:=""
Time_Now:=""
TimeSinceSP:=""
TimeEnd:=""
VarTime:=""
TotalRecs:=""

gui,ServiceReportGui:destroy
gui,AddPartsGui:destroy
return

getServiceReportPointer(){
	If Not Pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion){
		MsgBox Error accessing page
		return false
	} else {
		return Pwb
	}
}

SERVICEREPORT_bIsNewReport(pwb){
	frame := Pwb.document.all(9).contentWindow
	frame.document.getElementById("lblRepairJobFSR").click
	Loop{
		frame:=Pwb.document.all(11).contentWindow
		sleep, 100
		try{
			MultipleRecords:=frame.document.getElementById("lblRecordCount").innertext
			if (MultipleMultipleRecords != ""){
				MultipleRecords:=""
				return false
			}
			SingleCloseCall:=frame.document.getElementsByTagName("INPUT")[1].value
			if (SingleCloseCall != ""){
				SingleCloseCall:=""
				return false
			}
			OpenServiceReport:=frame.document.getElementById("lblComplete").innertext
			if (OpenServiceReport = "Complete"){
				return true
			}
		}
	}
}
SERVICEREPORT_getProdCode(pwb){
	frame := Pwb.document.all(10).contentWindow
	return frame.document.getElementById("cboJobPartNum").value
}
SERVICEREPORT_getJobType(pwb){
	frame := Pwb.document.all(10).contentWindow
	return frame.document.getElementById("cboCallCalTCode").value
}
SERVICEREPORT_getRoNumber(pwb){
	frame := Pwb.document.all(10).contentWindow
	return frame.document.getElementById("txtJobRef6").value
}
#Include lib\SERVICEREPORT_bIsNewReport.ahk
SERVICEREPORT_openNewReport(pwb){
	frame:=Pwb.document.all(10).contentWindow
	frame.document.getElementById("lblJobServiceReportAdd").click
	loop{
		Try{
			frame:=Pwb.document.all(11).contentWindow
			PageLoaded:=frame.document.getElementById("lblComplete").innertext
		}
	}until (PageLoaded = "Complete")
	return true
}
SERVICEREPORT_InjectEnhanced(pwb){
	frame := Pwb.document.all(7).contentWindow
	f6td1 = <TD width="25`%"><DIV style="Color:RED; height:100`%; text-Align:center; font:20">Powered by <br>T-Enhanced</br></DIV></TD>
	frame.document.getElementsBytagName("td")[1].innerhtml := f6td1
}
SERVICEREPORT_getTimeSinceLastReport(){
	RegRead,LastSP,HKCU,Software\TesseractZoanthropy,Last_SP
	TimeSinceSP:=A_now
	Envsub,TimeSinceSP,LastSP,minits
	return TimeSinceSP
}
SERVICEREPORT_getAverageRepairTime(ProdCode){
	IniRead,AverageTime,modules\ProductsAverage.ini,%ProdCode%,Average
	if (AverageTime = "error"){
		AverageTime = 0
	}
	Averagetime:=Floor(Averagetime)
	return AverageTime
}
SERVICEREPORT_TranslateFault(Fault){
	if (Fault = "Epos Fault"){
		Fault = HEP
	}else if (Fault = "Server Fault"){
		Fault = HSV
	}else if (Fault = "Pocket PC Fault"){
		Fault = PPC
	}else if (Fault = "Printer Consumable Fault"){
		Fault = PRC
	}else if (Fault = "Printer Fault"){
		Fault = PRF
	}else if (Fault = "Self Checkout Issue"){
		Fault = SCO
	}else if (Fault = "Software Epos Fault"){
		Fault = SEF
	}else if (Fault = "Software Workstation Fault"){
		Fault = SSF
	}
	return Fault
}
SERVICEREPORT_getCurrentTime(varTime){
	Time_Now:=A_Now
	EnvAdd, Time_Now, -VarTime ,Minits
	FormatTime, Time_Now,%Time_Now%, HH:mm
	return Time_Now
}
SERVICEREPORT_getEndTime(){
	FormatTime, TimeEnd,, HH:mm
	return TimeEnd
}