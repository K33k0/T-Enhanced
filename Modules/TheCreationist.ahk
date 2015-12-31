FormatTime,curDate,, yyyy-MM-dd

	CurrentCallNo := ""
	Engineer := ""
	f6td1 := ""
	fault := ""
	FinishedTime := ""
	JobType := ""
	ProbCode := ""
	ProdCode := "" 
	RO := ""
	AdditionalInfo:= ""

Create_EngineerNumber() {
	IniRead,Engineer,%Config%,Engineer,Number
	StringTrimRight,Engineer,Engineer,2
	OutputDebug, [T-Enhanced] Engineer var updated to - %Engineer%
	return Engineer
}

Create_NavigateToPage() {
	 If  Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion ) {
			OutputDebug, [T-Enhanced] Supported page
			return true
		} else if  Pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion) {
		try {
			frame := Pwb.document.all(9).contentWindow
			if (frame.document.getElementsByTagName("LABEL")[1].innertext = "job create wizard") {
				frame.document.getElementById("lblJobCreateWizard").click
				OutputDebug, [T-Enhanced] Navigated automatically to Create Wizard
				return TRUE
		}
	}
} else {
	return false
}
}

Create_KillPopup() {
	if (A_IsCompiled = 1){
		run,Modules\TZAltThread.exe
	}else{
		run,Modules\TZAltThread.ahk
	}
}

if not Engineer := Create_EngineerNumber() {
	outputDebug, Engineer Number set to %Engineer%
	msgbox, Failed to read Engineer Number.
	Reload
	return
} else {
	outputDebug, Engineer Number set to %Engineer%
}

If not Create_NavigateToPage() {
	OutputDebug, Failed to detect relevant page.
	MsgBox, Navigate manually
	return
}


CreateGui:
Gui, CreateGuint: Margin, 0, 0
Gui, CreateGuint:Add, Text,center BackgroundTrans,Serial Number
Gui, CreateGuint:Add, Edit, vSerialNumber,
Gui, CreateGuint:Add, Text, center BackgroundTrans,Problem Code
Gui, CreateGuint:Add, DropDownList,  sort vProbCode, Customer Damage|Distribution|Epos|Handheld|Printer|Self Checkout|Server
Gui, CreateGuint:Add, Text,  center BackgroundTrans,Job Type
Gui, CreateGuint:Add, DropDownList,  sort vJobType, Zulu Repair|Imac Refurb|Adhoc TCG|Consumables|Zulu VP
;Gui, CreateGuint:Add, Text,  center BackgroundTrans,Repair order Number
;Gui, CreateGuint:Add, edit, vRepOrdNo
Gui, CreateGuint:Add, Button, x65 gCreateContinue, Continue
Gui, CreateGuint: +AlwaysOnTop  +Owner%MasterWindow% +ToolWindow

X:=GetWinPosX("T-Enhanced Create Job Window")
Y:=GetWinPosY("T-Enhanced Create Job Window")
if (X = "" OR Y = "" OR X= "Error" OR Y="Error"){
Gui, CreateGuint: Show, , T-Enhanced Create Job Window
} else {
Gui, CreateGuint: Show, X%x% Y%y%  , T-Enhanced Create Job Window
}

OutputDebug, [T-Enhanced] Created the Create GUI
return
CreateCancel:
CreateGuintGuiClose:
CreateGuintGuiEscape:
gosub,Create_cancel
OutputDebug, [T-Enhanced] Destroyed the gui
return


CreateContinue:
StartTime := A_Now
Gui,CreateGuint:Submit
SaveWinPos(" T-Enhanced Create Job Window")
if (Create_EngineerNumber() = "406"){
	msgbox,4,Customer Damage Verification, Is there ANY sign of customer damage?
	IfMsgBox,Yes
	{
		ProbCode = Customer Damage
	}
}
		
If (ProbCode = "Epos") {
	ProbCode = HEP
} Else if (ProbCode = "HandHeld") {
	Probcode = HHT
}Else if (ProbCode = "Printer") {
	Probcode = HPR
}Else if (ProbCode = "Server") {
	Probcode = HSV
}Else if (ProbCode = "Self Checkout") {
	Probcode = SCO
}Else if (ProbCode = "Customer Damage") {
	Probcode = CDAM
}Else if (ProbCode = "Distribution") {
	Probcode = RDC
}
If (JobType = "Zulu Repair"){
	JobType = ZR1
} else if (JobType = "Imac Refurb") {
	JobType = ZR2
} else if (JobType = "Adhoc TCG") {
	JobType = ZR3
} else if (JobType = "Consumables") {
	JobType = ZR4
}else if (JobType = "Zulu VP") {
	JobType = VP
}

Create_Page1:
Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
FormatTime, Times,, HH:mm
Pwb.document.getElementsByTagName("INPUT")[4] .value := Times
Pwb.document.getElementsByTagName("INPUT")[8] .value := Times
OutputDebug, [T-Enhanced] book in Time input [%TIMES%]
Pwb.document.getElementById("cboJobWorkshopSiteNum") .value := "STOWS"
OutputDebug, [T-Enhanced] Workshop Site = Stows
Create_KillPopup()
Pwb.document.getElementsByTagName("IMG")[0] .click
Pwb.document.getElementById("cmdNext") .click
OutputDebug, [T-Enhanced] Page Load Initiated [ref 161]
IELoad(Pwb)
OutputDebug, [T-Enhanced] Page Load Success [ref 161]

Create_Page2:
if (JobType = "ZR1") {
Pwb.document.getElementById("cboCallSiteNum") .value := "ZULU"
Create_KillPopup()
Pwb.document.getElementsByTagName("IMG")[5] .click 
Loop {
	try {
	if  not Pwb.document.getElementById("txtCallSiteAddress") .value
		sleep 100
	else
		break
}
}

Pwb.document.getElementById("cmdNext") .click
IELoad(Pwb)
} else {
	If (JobType = "VP") {
		JobType=ZR1
	}
	#Include Modules/TheCreationistJudge.ahk
	gosub,Create_Page4
	return
}





Create_Page3:
Pwb.document.getElementById("cmdNext") .click
OutputDebug, [T-Enhanced] Page Load Initiated [ref 166]
IELoad(Pwb)
OutputDebug, [T-Enhanced] Page Load Success [ref 166]
InsertSN:
Pwb.document.getElementById("cboCallSerNum") .value := SerialNumber
Pwb.document.getElementsByTagName("INPUT")[58] .click
Create_KillPopup()
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
if  (JobType = "ZR1") {
Pwb.document.getElementById("txtJobRef6") .value :=GetRO(SerialNumber,ProdCode)
if (Pwb.document.getElementById("txtJobRef6") .value = "") {
	MsgBox,Requires Order Number
	gosub,Create_Cancel
	return
}
}

Pwb.document.getElementById("cmdNext") .click
OutputDebug, [T-Enhanced] Page Load Initiated [ref 198]
IELoad(Pwb)
OutputDebug, [T-Enhanced] Page Load Success [ref 198]

Create_Page4:
Pwb.document.getElementById("cboCallCalTCode") .value := JobType
OutputDebug, [T-Enhanced] Successfully inputted Job Type
Create_KillPopup()
Pwb.document.getElementsByTagName("IMG")[22] .click

Pwb.document.getElementById("cboJobFlowCode") .value := "SWBENCH"
OutputDebug, [T-Enhanced] Successfully inputted FlowCode
Create_KillPopup()
Pwb.document.getElementsByTagName("IMG")[23] .click

Pwb.document.getElementById("cboCallAreaCode") .value := "BFCA1"
OutputDebug, [T-Enhanced] Successfully inputted Area
Create_KillPopup()
Pwb.document.getElementsByTagName("IMG")[23] .click

Pwb.document.getElementById("cboCallEmployNum") .value := Engineer
OutputDebug, [T-Enhanced] Successfully inputted Engineer Number
Create_KillPopup()
Pwb.document.getElementsByTagName("IMG")[25] .click

Create_KillPopup()
Pwb.document.getElementById("cboCallProbCode") .value := ProbCode
OutputDebug, [T-Enhanced] Successfully inputted Problem Code
Pwb.document.getElementsByTagName("IMG")[26] .click
if (JobType = "ZR1"){
msgbox % "Your maximum budget is - £"  GetThreshold(ProdCode)

}
FinishedTime:=A_Now
EnvSub,FinsihedTime,StartTime,Seconds
Pwb.document.getElementsByTagName("TEXTAREA")[4] .value := Fault . "`n---------------[T-Enhanced]---------------`n Job created in "FinsihedTime " Seconds"
OutputDebug, [T-Enhanced] Successfully inputted customer fault
FinsihedTime:=""
StartTime:=""
msgbox,4,Creation Confirmation,Are you happy to continue. `nMistakes may lead to stock anomalies
IfMsgBox, No
	{
	Gosub, Create_cancel
	return
}
ifMsgbox, Abort
		{
	Gosub, Create_cancel
	return
}
IfMsgBox, cancel
		{
	Gosub, Create_cancel
	return
}

Create_Confirmed:
Create_KillPopup()
Pwb.document.getElementById("cmdFinish") .click
OutputDebug, [T-Enhanced] Page Load Initiated [ref 534]
IELoad(Pwb)

Create_Page5:
OutputDebug, [T-Enhanced] Page Load Success [ref 534]
WinWaitClose, Message from webpage
Pwb.document.getElementsByTagName("INPUT")[119] .click
Pwb.document.getElementById("cmdFinish") .click
OutputDebug, [T-Enhanced] Page Load Initiated [ref 539]
IELoad(Pwb) ;[ref 539]
OutputDebug, [T-Enhanced] Page Load Success [ref 539]

Create_Page6:
if not wb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion){
MsgBox Error accessing page, click 'reload T-Enhanced'
gosub,Create_cancel
return
}
frame := Pwb.document.all(6).contentWindow
f6td1 = <TD width="25`%"><DIV style="Color:Red; height:100`%; text-Align:center; font:20">Powered by <br>T-Enhanced</br></DIV></TD>
frame.document.getElementsBytagName("td")[1].innerhtml := f6td1
frame := wb.document.all(10).contentWindow
if (JobType = "ZR2"){
call = IMAC
frame.document.getElementById("txtJobApproveRef").value:= call
}
CurrentCallNo:=frame.document.getElementsByTagName("INPUT")[0].value
frame := wb.document.all(7).contentWindow
frame.document.getElementById("cmdSubmit").click
frame := wb.document.all(10).contentWindow
frame.document.getElementById("cboCallUpdAreaCode").value := "WSB"
Create_KillPopup()
frame := wb.document.all(7).contentWindow
frame.document.getElementById("cmdSubmit").click
WinWaitClose, Message from webpage

Create_cancel:
OutputDebug,[T-Enhanced]%A_thislabel% Started
Pwb:=""
wb:=""
gosub,Create_VarCleanup
Gui,CreateGuint:destroy
Gui,CreateOTFinfo:destroy
OutputDebug,[T-Enhanced]%A_thislabel% Finished
return

Create_VarCleanup:
	OutputDebug,[T-Enhanced]%A_thislabel% Started
	CurrentCallNo := ""
	Engineer := ""
	f6td1 := ""
	fault := ""
	FinishedTime := ""
	JobType := ""
	ProbCode := ""
	ProdCode := "" 
	RepOrdNo := ""
	AdditionalInfo:= ""
	OutputDebug,[T-Enhanced]%A_thislabel% Finsihed
return