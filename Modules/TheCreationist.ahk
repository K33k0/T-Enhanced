FormatTime,curDate,, yyyy-MM-dd

Create_EngineerNumber() {
	IniRead,Engineer,%Config%,Engineer,Number
	StringTrimRight,Engineer,Engineer,2
	OutputDebug, [T-Enhanced] Engineer var updated to - %Engineer%
	return Engineer
}


Create_NavigateToPage() {
		try {
		Pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion)
		frame := Pwb.document.all(9).contentWindow
		if (frame.document.getElementsByTagName("LABEL")[1].innertext = "job create wizard") {
			frame.document.getElementById("lblJobCreateWizard").click
			OutputDebug, [T-Enhanced] Navigated automatically to Create Wizard
			return TRUE
		}else{
			If Not Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion ) {
				OutputDebug, [T-Enhanced] Unsupported page
				return False
			}
		}
	}catch{
		return false
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
Gui, CreateGuint:Add, Text,center BackgroundTrans,Enter the Serial Number
Gui, CreateGuint:Add, Edit, vSerialNumber,
Gui, CreateGuint:Add, Text, center BackgroundTrans,Choose the Problem Code
Gui, CreateGuint:Add, DropDownList,  sort vProbCode, Customer Damage|Distribution|Epos|Handheld|Printer|Self Checkout|Server
Gui, CreateGuint:Add, Text,  center BackgroundTrans, Choose the Job Type
Gui, CreateGuint:Add, DropDownList,  sort vJobType, Food|Healthcare|East Of England|Generic|Food Refurb|Scales|Farm|Distribution
Gui, CreateGuint:Add, Text,  center BackgroundTrans,Repair order Number
Gui, CreateGuint:Add, edit, vRepOrdNo
Gui, CreateGuint:Add, Button,gCreateContinue, Continue
Gui, CreateGuint:Add, Button,gCreateCancel -TabStop , Quit
Gui, CreateGuint: +AlwaysOnTop +ToolWindow
Gui, CreateGuint:Show, , T-Enhanced Create Job Window
OutputDebug, [T-Enhanced] Created the Create GUI
return
CreateCancel:
CreateGuintGuiClose:
CreateGuintGuiEscape:
Gui,CreateGuint:destroy
OutputDebug, [T-Enhanced] Destroyed the gui
return


CreateContinue:
StartTime := A_Now
Gui,CreateGuint:Submit
ProbCode := ((ProbCode = "Epos") ? ("HEP") :(((ProbCode = "HandHeld") ? ("HHT") : (((Probcode = "Printer") ? ("HPR") : (((Probcode = "Server") ? ("HSV") : (((Probcode = "Self Checkout") ? ("SCO") : (((ProbCode = "Customer Damage") ? ("CDAM") : (((ProbCode = "Distribution") ? ("RDC") : (return))))))))))))))
JobType := ((JobType = "Food") ? ("W1F") : (((JobType = "Healthcare") ? ("W1H") : (((Jobtype = "East Of England") ? ("W1E") : (((Jobtype = "Generic") ? ("W1G") : (((JobType = "Food Refurb") ? ("FRN") : (((JobType = "Scales") ? ("WFS") : (((JobType = "Farm") ? ("WSF") : (((JobType = "Distribution") ? ("WSD") : (return))))))))))))))))

Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
FormatTime, Times,, HH:mm
Pwb.document.getElementsByTagName("INPUT")[4] .value := Times
OutputDebug, [T-Enhanced] book in Time input [%TIMES%]
Pwb.document.getElementsByTagName("INPUT")[8] .value := Times
OutputDebug, [T-Enhanced] approve Time input [%TIMES%]
Pwb.document.getElementById("cboJobWorkshopSiteNum") .value := "STOWS"
OutputDebug, [T-Enhanced] Workshop Site = Stows
if (A_IsCompiled = 1){
run,Modules\TZAltThread.exe
}else{
run,Modules\TZAltThread.ahk
}
Pwb.document.getElementsByTagName("IMG")[0] .click
Pwb.document.getElementById("cmdNext") .click
OutputDebug, [T-Enhanced] Page Load Initiated [ref 161]
IELoad(Pwb)
OutputDebug, [T-Enhanced] Page Load Success [ref 161]

Pwb.document.getElementById("cboCallSiteNum") .value := "ZULU"
if (A_IsCompiled = 1){
run,Modules\TZAltThread.exe
}else{
run,Modules\TZAltThread.ahk
}
Pwb.document.getElementsByTagName("IMG")[5] .click ;need to find the right number
Pwb.document.getElementById("cmdNext") .click
IELoad(Pwb)

Pwb.document.getElementById("cmdNext") .click
OutputDebug, [T-Enhanced] Page Load Initiated [ref 166]
IELoad(Pwb)
OutputDebug, [T-Enhanced] Page Load Success [ref 166]
InsertSN:
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
Pwb.document.getElementById("txtJobRef6") .value := RepOrdNo
if (RepOrdNo = "") {
	MsgBox,Requires Order Number
	Pwb:=""
	wb:=""
	Gui,CreateGuint:destroy
	Gui,CreateOTFinfo:destroy
	return
}
Pwb.document.getElementById("cmdNext") .click
OutputDebug, [T-Enhanced] Page Load Initiated [ref 198]
IELoad(Pwb)
OutputDebug, [T-Enhanced] Page Load Success [ref 198]
SiteNumber:=Pwb.document.getElementById("cboCallSiteNum") .value
IfInString,SiteNumber, VK
{
msgbox,,Stock Anomaly, Give to your Team Leader immediately
Pwb:=""
wb:=""
Gui,CreateGuint:destroy
Gui,CreateOTFinfo:destroy
return
}

Pwb.document.getElementById("cmdNext") .click
OutputDebug, [T-Enhanced] Page Load Initiated [ref 214]
IELoad(Pwb)
OutputDebug, [T-Enhanced] Page Load Success [ref 214]
if (A_IsCompiled = 1){
run,Modules\TZAltThread.exe
}else{
run,Modules\TZAltThread.ahk
}
Pwb.document.getElementById("cmdNext") .click
OutputDebug, [T-Enhanced] Page Load Initiated  [ref 237]
IELoad(Pwb)
OutputDebug, [T-Enhanced] Page Load Success  [ref 237]

Pwb.document.getElementById("cboCallCalTCode") .value := JobType
OutputDebug, [T-Enhanced] Successfully inputted Job Type
if (A_IsCompiled = 1){
run,Modules\TZAltThread.exe
}else{
run,Modules\TZAltThread.ahk
}
Pwb.document.getElementsByTagName("IMG")[22] .click

Pwb.document.getElementById("cboCallEmployNum") .value := Engineer
OutputDebug, [T-Enhanced] Successfully inputted Engineer Number
if (A_IsCompiled = 1){
run,Modules\TZAltThread.exe
}else{
run,Modules\TZAltThread.ahk
}
Pwb.document.getElementsByTagName("IMG")[25] .click

if (A_IsCompiled = 1){
run,Modules\TZAltThread.exe
}else{
run,Modules\TZAltThread.ahk
}
Pwb.document.getElementById("cboCallProbCode") .value := ProbCode
OutputDebug, [T-Enhanced] Successfully inputted Problem Code
Pwb.document.getElementsByTagName("IMG")[26] .click

if (Fault = ""){
	Fault = No information available
}
if (ProdCode = "STOKVP"){
	OutputDebug, [T-Enhanced] Waiting for user input
	InputBox,Fault,Call Details, Input details of job
	if (Fault = ""){
		Fault = No Information given
	}
}
FinishedTime:=A_Now
EnvSub,FinsihedTime,StartTime,Seconds
Pwb.document.getElementsByTagName("TEXTAREA")[4] .value := Fault . "`n---------------[TG]---------------`n Job created in "FinsihedTime " Seconds"
OutputDebug, [T-Enhanced] Successfully inputted customer fault
FinsihedTime:=""
StartTime:=""
return
ConfirmationGuiClose:
No:
OutputDebug, [T-Enhanced] Confirmation close [NO]
TrayTip, Job Create Wizard,Terminated,
Gui,CreateGuint:destroy
Gui,CreateOTFinfo:destroy
sleep, 2000
TrayTip
return
yes:
OutputDebug, [T-Enhanced] Confirmation closed [YES]
if (A_IsCompiled = 1){
run,Modules\TZAltThread.exe
}else{
run,Modules\TZAltThread.ahk
}
Pwb.document.getElementById("cmdFinish") .click
OutputDebug, [T-Enhanced] Page Load Initiated [ref 534]
IELoad(Pwb)
OutputDebug, [T-Enhanced] Page Load Success [ref 534]
WinWaitClose, Message from webpage
Pwb.document.getElementsByTagName("INPUT")[119] .click
Pwb.document.getElementById("cmdFinish") .click
OutputDebug, [T-Enhanced] Page Load Initiated [ref 539]
IELoad(Pwb) ;[ref 539]
OutputDebug, [T-Enhanced] Page Load Success [ref 539]
if (Precall = "")
Precall = N/A
try {
if not wb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion){
MsgBox Error accessing page, click 'reload Tesseract Enhancned'
Process,Close,TE_Create_msgbox.exe
Process,Close,IEinterupt.exe
Gui,CreateGuint:destroy
Gui,CreateOTFinfo:destroy
return
}else {
frame := Pwb.document.all(6).contentWindow
f6td1 = <TD width="25`%"><DIV style="Color:Red; height:100`%; text-Align:center; font:20">Powered by <br>T-Enhanced</br></DIV></TD>
frame.document.getElementsBytagName("td")[1].innerhtml := f6td1
frame := wb.document.all(10).contentWindow
if (JobType = "FRN"){
call = IMAC
}
frame.document.getElementsByTagName("INPUT")[7].value := ((call = "0")?("Not Found"):(call))
CurrentCallNo:=frame.document.getElementsByTagName("INPUT")[0].value
frame.document.getElementsByTagName("INPUT")[87].value :=PreEngNo
frame.document.getElementsByTagName("INPUT")[88].value :=SiteNo
If (CLF = True){
frame.document.getElementsByTagName("INPUT")[90].value :="Y"
}Else{
frame.document.getElementsByTagName("INPUT")[90].value :="N"
}
frame := wb.document.all(7).contentWindow
frame.document.getElementById("cmdSubmit").click
frame := wb.document.all(10).contentWindow
frame.document.getElementById("cboCallUpdAreaCode").value := "WSB"
if (A_IsCompiled = 1){
run,Modules\TZAltThread.exe
}else{
run,Modules\TZAltThread.ahk
}
frame := wb.document.all(7).contentWindow
frame.document.getElementById("cmdSubmit").click
WinWaitClose, Message from webpage

}
}catch e{
MsgBox, Unable to access Page. Click 'reload teseract Zoanthropy '`n[3151]
}
Pwb:=""
wb:=""
Gui,CreateGuint:destroy
Gui,CreateOTFinfo:destroy
return
