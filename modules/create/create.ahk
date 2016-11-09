
Gui, Create:New,, T-Enhanced Create Job Window
Gui, Create:Add, Text,center BackgroundTrans,Serial Number
Gui, Create:Add, Edit, vSerialNumber,
Gui, Create:Add, Text, center BackgroundTrans,Problem Code
Gui, Create:Add, DropDownList,  sort vProblemCode, Customer Damage|Distribution|Epos|Handheld|Printer|Self Checkout|Server
Gui, Create:Add, Text,  center BackgroundTrans,Job Type
Gui, Create:Add, DropDownList,  sort vJobType, Zulu Repair|Imac Refurb|Adhoc TCG|Consumables|Zulu VP|Mid Counties
Gui, Create:Add, Text,  center BackgroundTrans,Repair order Number
Gui, Create:Add, edit, vRONumber
Gui, Create:Add, Button, x65 gCreate_Continue, Continue
Gui, Create: +AlwaysOnTop  +Owner%MasterWindow% +ToolWindow
X:=GetWinPosX("T-Enhanced Create Job Window")
Y:=GetWinPosY("T-Enhanced Create Job Window")
if (X = "" OR Y = "" OR X= "Error" OR Y="Error"){
} else {
	Gui, Create: Show, X%x% Y%y%
}
return

create_init(){
	if Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion){
		return pwb
	} else if PWB := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion) {
		try {
			frame := Pwb.document.all(9).contentWindow
			if (frame.document.getElementsByTagName("LABEL")[1].innertext = "job create wizard") {
				frame.document.getElementById("lblJobCreateWizard").click
				OutputDebug, [T-Enhanced] Navigated automatically to Create Wizard
				return pwb
			}
		}
	} else {
		return False
	}
}
pwb := create_init()


Create_continue:
SaveWinPos("T-Enhanced Create Job Window")
gui, submit
StartTime := A_Now
probCode := convertProblemCode(ProblemCode)
JobType := convertJobType(JobType)
if not setPageOne()
	return
if (JobType = "ZR1"){
	if (!setPageTwo_Zulu()){
		msgbox, failed to set page 2 [zulu]
		return
	} else {
		if (!setPageThree(JobType,SerialNumber,RONumber )){
			msgbox, failed to set page 3
			return
		}
	}
} else {
	if (!setPageTwo_Alt(JobType,SerialNumber,RONumber )){
		msgbox, failed to set page 2 [alt]
		return
	}
}
if (!setPageFour(JobType,settings,Probcode)){
	return
}
if (!confirmed()){
	return
}
return

convertProblemCode(ProblemCode){
	If (ProblemCode = "Epos") {
		return "HEP"
	} Else if (ProblemCode = "HandHeld") {
		return "HHT"
	} Else if (ProblemCode = "Printer") {
		return "HPR"
	} Else if (ProblemCode = "Server") {
		return "HSV"
	} Else if (ProblemCode = "Self Checkout") {
		return "SCO"
	} Else if (ProblemCode = "Customer Damage") {
		return "CDAM"
	} Else if (ProblemCode = "Distribution") {
		return "RDC"
	} else {
		return false
	}
}
convertJobType(JobType){
	If (JobType = "Zulu Repair"){
		return "ZR1"
	} else if (JobType = "Imac Refurb") {
		return "ZR2"
	} else if (JobType = "Adhoc TCG") {
		return "ZR3"
	} else if (JobType = "Consumables") {
		return "ZR4"
	} else if (JobType = "Zulu VP") {
		return "VP"
	} else if (JobType = "Mid Counties"){
		return "ZR6"
	} else {
		return false
	}
}
setPageOne(){
	if not Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
		return false
	FormatTime, Times,, HH:mm
	Pwb.document.getElementsByTagName("INPUT")[4] .value := Times
	Pwb.document.getElementsByTagName("INPUT")[8] .value := Times
	Pwb.document.getElementById("cboJobWorkshopSiteNum") .value := "STOWS"
	ModalDialogue()
	Pwb.document.getElementsByTagName("IMG")[0] .click
	Pwb.document.getElementById("cmdNext") .click
	IELoad(Pwb)
	return true
}
setPageTwo_Zulu(){
	if not Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
		return false
	Pwb.document.getElementById("cboCallSiteNum") .value := "ZULU"
	ModalDialogue()
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
	return true
}
setPageTwo_Alt(JobType, SerialNumber, RONumber){
	if not Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
		return false
	If (JobType = "VP") {
		JobType:="ZR1"
	}
	Pwb.document.getElementById("cmdNext") .click
	IELoad(Pwb)
	Pwb.document.getElementById("cboCallSerNum") .value := SerialNumber
	Pwb.document.getElementsByTagName("INPUT")[58] .click
	ModalDialogue()
	Pwb.document.getElementsByTagName("IMG")[19] .click
	Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
	SerialNumber:= Pwb.document.getElementById("cboCallSerNum").value
	ProdCode:= Pwb.document.getElementById("cboJobPartNum").value
	If (ProdCode = ""){
		return false
	}
	if (JobType = "ZR1" && RONumber = ""){
		return false
	}
	Pwb.document.getElementById("txtJobRef6") .value := RONumber
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
	} else if (JobType = "ZR6"){
		shipsite = STOFSLGDSI
	} else {
		ShipSite = STOKGOODS
	}
	Pwb.document.getElementById("cboShipSiteNum") .value := ShipSite
	ModalDialogue()
	Pwb.document.getElementsByTagName("IMG")[15] .click
	IELoad(Pwb)
	Pwb.document.getElementById("cmdNext") .click
	IELoad(Pwb)
	return True
}
setPageThree(JobType, SerialNumber, RONumber){
	if not this.Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
		return false
	Pwb.document.getElementById("cmdNext") .click
	IELoad(Pwb)
	Pwb.document.getElementById("cboCallSerNum") .value := SerialNumber
	Pwb.document.getElementsByTagName("INPUT")[58] .click
	ModalDialogue()
	Pwb.document.getElementsByTagName("IMG")[19] .click
	Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
	SerialNumber:= Pwb.document.getElementById("cboCallSerNum").value
	ProdCode:= Pwb.document.getElementById("cboJobPartNum").value
	If (ProdCode = ""){
		return false
	}
	if  (JobType = "ZR1") {
		Pwb.document.getElementById("txtJobRef6") .value := GetRO(SerialNumber,ProdCode) = roVerify ? roVerify : ""
		if (Pwb.document.getElementById("txtJobRef6") .value = "") {
			return false
		}
	}
	Pwb.document.getElementById("cmdNext") .click
	IELoad(Pwb)
	return True
}
setPageFour(JobType,settings,Probcode){
if not Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
	return false
pwb := Pwb
Pwb.document.getElementById("cboCallCalTCode") .value := JobType
ModalDialogue()
Pwb.document.getElementsByTagName("IMG")[22] .click
Pwb.document.getElementById("cboJobFlowCode") .value := "SWBENCH"
ModalDialogue()
Pwb.document.getElementsByTagName("IMG")[23] .click
Pwb.document.getElementById("cboCallAreaCode") .value := "BFCA1"
ModalDialogue()
Pwb.document.getElementsByTagName("IMG")[24] .click
Pwb.document.getElementById("cboCallEmployNum") .value := settings.engineer
ModalDialogue()
Pwb.document.getElementsByTagName("IMG")[25] .click
ModalDialogue()
Pwb.document.getElementById("cboCallProbCode") .value := ProbCode
Pwb.document.getElementsByTagName("IMG")[26] .click
Pwb.document.getElementsByTagName("TEXTAREA")[4] .value := "Created automatically via T-Enhanced by Kieran Wynne"
msgbox,4,Creation Confirmation,Are you happy to continue. `nMistakes may lead to stock anomalies
IfMsgBox, Yes
	return true
else
	return false
}
confirmed(){
			if not Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
				return false
			PageAlert()
			Pwb.document.getElementById("cmdFinish") .click
			IELoad(Pwb)
			WinWaitClose, Message from webpage
			Pwb.document.getElementsByTagName("INPUT")[119] .click
			Pwb.document.getElementById("cmdFinish") .click
			IELoad(Pwb)
			return True
		}
