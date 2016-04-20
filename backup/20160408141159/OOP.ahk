TEnhanced := new TEnhanced(settings)

class TEnhanced {
	
	
	Class AutoLogin {
		
		__New(settings,Tab){
			OutputDebug,[T-Enhanced]  Quick login started
			if (A_GuiControl = "Tab"){
				if (Tab = "Engineer") {
					IniRead,UserHash,%Config%,Login,UserName
					IniRead,PassHash,%Config%,Login,Password
					If (UserHash = "" OR UserHash = "Error"){
						return
					}
					if not PWB:= IEGET("Service Centre 5 Login") {
						pwb:=""
						return
					} else {
						pwb.document.getElementById("txtUserName").value := settings.decrypt("username")
						pwb.document.getElementById("txtPassword").value := settings.decrypt("password")
						pwb.document.getElementsByTagName("IMG")[7].click
						pwb:=""
					}
				}
			}
			OutputDebug,[T-Enhanced]  Quick login ended
		}
	}
	
	Class Create {
		static SerialNumber
		static ProblemCode
		static JobType
		static RONumber
		static pwb
		static StartTime
		static settings
		__New(settings){
			this.settings := settings
			if (!this.getPage()){
				msgbox, Failed to navigate to Job Create Wizard
				return False
			}
			if (!this.createGui()) {
				return false
			}
			if (!this.ConvertHumanText()){
				msgbox, Failed to convert text.
				return false
			}
			if (!this.setPageOne()){
				msgbox, failed to set page one
				return false
			}
			if (this.JobType = "ZR1"){
				if (!this.setPageTwo_Zulu()){
					msgbox, failed to set page 2 [zulu]
					return false
				} else {
					if (!this.setPageThree()){
						msgbox, failed to set page 3
						return true
					}
				}
			} else {
				if (!this.setPageTwo_Alt()){
					msgbox, failed to set page 2 [alt]
					return False
				}
			}
			if (!this.setPageFour()){
				return False
			}
			if (!this.confirmed()){
				return False
			}
			return true
		}
		__Delete(){
			gui,Create:Destroy
			this:= ""
		}
		getPage(){
			if Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion){
				return this.pwb := pwb
			} else if PWB := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion) {
				try {
					frame := Pwb.document.all(9).contentWindow
					if (frame.document.getElementsByTagName("LABEL")[1].innertext = "job create wizard") {
						frame.document.getElementById("lblJobCreateWizard").click
						OutputDebug, [T-Enhanced] Navigated automatically to Create Wizard
						return this.pwb := pwb
					}
				}
			} else {
				return False
			}
		}
		createGui(){
			static SerialNumber
			static ProblemCode
			static JobType
			static RONumber
			Gui, Create:New,, T-Enhanced Create Job Window
			Gui, Create:Add, Text,center BackgroundTrans,Serial Number
			Gui, Create:Add, Edit, vSerialNumber,
			Gui, Create:Add, Text, center BackgroundTrans,Problem Code
			Gui, Create:Add, DropDownList,  sort vProblemCode, Customer Damage|Distribution|Epos|Handheld|Printer|Self Checkout|Server
			Gui, Create:Add, Text,  center BackgroundTrans,Job Type
			Gui, Create:Add, DropDownList,  sort vJobType, Zulu Repair|Imac Refurb|Adhoc TCG|Consumables|Zulu VP
			Gui, Create:Add, Text,  center BackgroundTrans,Repair order Number
			Gui, Create:Add, edit, vRONumber
			Gui, Create:Add, Button, x65 gContinue, Continue
			Gui, Create: +AlwaysOnTop  +Owner%MasterWindow% +ToolWindow
			X:=GetWinPosX("T-Enhanced Create Job Window")
			Y:=GetWinPosY("T-Enhanced Create Job Window")
			if (X = "" OR Y = "" OR X= "Error" OR Y="Error"){
				Gui, Create: Show
			} else {
				Gui, Create: Show, X%x% Y%y%
			}
			WinWaitClose T-Enhanced Create Job Window
			if (this.serialNumber || this.problemCode || this.jobType){
				return true
			} else {
				return false
			}
			continue:
			SaveWinPos("T-Enhanced Create Job Window")
			gui, submit
			this.StartTime := A_Now
			this.serialNumber := SerialNumber
			this.problemCode := ProblemCode
			this.jobType := JobType
			this.roNumber:= RONumber
			return
		}
		ConvertHumanText(){
			If (this.ProblemCode = "Epos") {
				this.ProbCode := "HEP"
			} Else if (this.ProbCode = "HandHeld") {
				this.Probcode := "HHT"
			}Else if (this.ProbCode = "Printer") {
				this.Probcode := "HPR"
			}Else if (this.ProbCode = "Server") {
				this.Probcode := "HSV"
			}Else if (this.ProbCode = "Self Checkout") {
				this.Probcode := "SCO"
			}Else if (this.ProbCode = "Customer Damage") {
				this.Probcode := "CDAM"
			}Else if (this.ProbCode = "Distribution") {
				this.Probcode := "RDC"
			} else {
				return false
			}
			If (this.JobType = "Zulu Repair"){
				this.JobType := "ZR1"
			} else if (this.JobType = "Imac Refurb") {
				this.JobType := "ZR2"
			} else if (this.JobType = "Adhoc TCG") {
				this.JobType := "ZR3"
			} else if (this.JobType = "Consumables") {
				this.JobType := "ZR4"
			} else if (this.JobType = "Zulu VP") {
				this.JobType := "VP"
			} else {
				return false
			}
			return true
		}
		setPageOne(){
			if not this.Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
				return false
			pwb := this.Pwb
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
			if not this.Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
				return false
			pwb := this.Pwb
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
		setPageTwo_Alt(){
			if not this.Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
				return false
			pwb := this.Pwb
			If (this.JobType = "VP") {
				this.JobType:="ZR1"
			}
			Pwb.document.getElementById("cmdNext") .click
			IELoad(Pwb)
			Pwb.document.getElementById("cboCallSerNum") .value := this.SerialNumber
			Pwb.document.getElementsByTagName("INPUT")[58] .click
			ModalDialogue()
			Pwb.document.getElementsByTagName("IMG")[19] .click
			Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
			SerialNumber:= Pwb.document.getElementById("cboCallSerNum").value
			ProdCode:= Pwb.document.getElementById("cboJobPartNum").value
			If (ProdCode = ""){
				return false
			}
			if (this.JobType = "ZR1" && this.RONumber = ""){
				return false
			}
			Pwb.document.getElementById("txtJobRef6") .value := this.RONumber
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
		setPageThree(){
			if not this.Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
				return false
			pwb := this.Pwb
			Pwb.document.getElementById("cmdNext") .click
			IELoad(Pwb)
			InsertSN:
			Pwb.document.getElementById("cboCallSerNum") .value := this.SerialNumber
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
		setPageFour(){
			if not this.Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
				return false
			pwb := this.Pwb
			Pwb.document.getElementById("cboCallCalTCode") .value := this.JobType
			ModalDialogue()
			Pwb.document.getElementsByTagName("IMG")[22] .click
			Pwb.document.getElementById("cboJobFlowCode") .value := "SWBENCH"
			ModalDialogue()
			Pwb.document.getElementsByTagName("IMG")[23] .click
			Pwb.document.getElementById("cboCallAreaCode") .value := "BFCA1"
			ModalDialogue()
			Pwb.document.getElementsByTagName("IMG")[24] .click
			Pwb.document.getElementById("cboCallEmployNum") .value := this.settings.engineer
			ModalDialogue()
			Pwb.document.getElementsByTagName("IMG")[25] .click
			ModalDialogue()
			Pwb.document.getElementById("cboCallProbCode") .value := this.ProbCode
			Pwb.document.getElementsByTagName("IMG")[26] .click
			FinishedTime:= A_Now
			EnvSub,FinishedTime,this.StartTime,Seconds
			Pwb.document.getElementsByTagName("TEXTAREA")[4] .value := "T-Enhanced - created in " FinishedTime " seconds"
			msgbox,4,Creation Confirmation,Are you happy to continue. `nMistakes may lead to stock anomalies
			IfMsgBox, Yes
				return true
			else
				return false
		}
		confirmed(){
			if not this.Pwb := IEGet("Repair Job Creation Wizard - " TesseractVersion )
				return false
			pwb := this.Pwb
			PageAlert()
			Pwb.document.getElementById("cmdFinish") .click
			IELoad(Pwb)
			WinWaitClose, Message from webpage
			Pwb.document.getElementsByTagName("INPUT")[119] .click
			Pwb.document.getElementById("cmdFinish") .click
			IELoad(Pwb)
			return True
		}
		
	}
	
	Class Shipout {
		static pwb
		__New() {
			if not this.Pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion){
				return False
			}
			if !(JobType:= this.getJobType()){
				return false
			}
			if (JobType = "ZULUAW"){
				msgbox, Already shipped
			}
			if !(ShipSite:= this.getShipSite()){
				return False
			}
			if (shipSite = "ZULU"){
				this.zuluShipout()
				return true
			} else {
				this.oldShipout()
			}
		}
		
		__Delete(){
			this:= ""
		}
		
		getJobType(){
			pwb:= this.pwb
			frame := Pwb.document.all(10).contentWindow
			return (frame.document.getElementByID("cboJobFlowCode").value)
		}
		
		getShipSite(){
			pwb:= this.pwb
			frame := Pwb.document.all(10).contentWindow
			return (frame.document.getElementById("cboJobShipSiteNum").value)
		}
		
		zuluShipout(){
			pwb := this.pwb
			Loop{
				Try{
					frame := Pwb.document.all(10).contentWindow
					PageLoaded:= frame.document.getElementsByTagName("Label")[0].innertext
				}
			}Until (PageLoaded = "Job Details")
			PageLoaded:=""
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
			ModalDialogue()
			frame.document.getElementsByTagName("IMG")[35].click
			sleep, 1500
			WinWaitClose,Popup List -- Webpage Dialog,,5
			frame := Pwb.document.all(7).contentWindow
			PageAlert()
			frame.document.getElementById("cmdSubmit").click
			WinWaitClose, Message from webpage
			return
		}
		
		oldShipout(){
			if not Pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion)
				return false
			sleep, 250
			frame := Pwb.document.all(10).contentWindow
			ShipSite:= frame.document.getElementById("cboJobShipSiteNum") .value
			sleep, 250
			frame := Pwb.document.all(9).contentWindow
			frame.document.getElementById("lblJobShipOutWizard") .click
			IELoad(pwb)
			Pwb.document.getElementById("txtInputJobNum") .value :=CallNum
			Pwb.document.getElementById("cmdAddJobNum") .click
			Pwb.document.getElementById("cmdNext") .click
			IELoad(Pwb)
			Pwb.document.getElementById("cmdNext") .click
			IELoad(Pwb)
			Pwb.document.getElementsByTagName("INPUT")[48] .click
			SerialNumber:=Pwb.document.getElementbyID("cbaListCallSerNumLineArray").value
			ProdCode:=Pwb.document.getElementbyID("cbaListJobPartNumLineArray").value
			Pwb.document.getElementById("cmdFinish") .click
			PageAlert()
			IELoad(Pwb)
			Pwb.document.getElementsByTagName("INPUT")[40] .click
			Pwb.document.getElementById("cmdFinish") .click
			return true
		}
	}
	
}

