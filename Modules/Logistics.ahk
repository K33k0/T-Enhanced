Logistics := new Logistics()

class Logistics{
	__New() {
	}
	class Bookin {
		static StartTime		
		__New(){
			static SerialNumber
			static newRO
			static productCode
			gui, KMN:add, text,,Insert Serial Number
			Gui, KMN:Add, Edit, vSerialNumber
			gui, KMN:add, text,,Insert Repair Order Number
			Gui, KMN:Add, Edit, vnewRO
			gui, KMN:add, text,,Insert Product Code
			Gui, KMN:Add, Edit, vproductCode,
			Gui, KMN:add, button, gKMN_Submit, Continue
			Gui, KMN: +AlwaysOnTop +ToolWindow +LastFound
			Gui, KMN:show,, Book In
			WinWaitClose, Book In
			return "Gui Closed"
			kmn_submit:
			gui, KMN:submit, NoHide
			if (!SerialNumber || !newRO || !productCode) {
				return False
			}
			gui, KMN:submit
			if (!this.ROisFree(newRO)){
				msgbox, RO is in use
				return false
			}
			StringUpper, SerialNumber, SerialNumber
			this.StartTime := A_Now
			
			if not History := this.HistoryCheck(SerialNumber) {
				return False
			} else {
				if (History = "Zulu"){
					if (!this.ZuluStock(NewRO)){
						return false
					}
				} else if (History = "Not Zulu") {
					if (! this.InstalledElsewhere()){
						return false
					}
					if (!this.ProductAdd(SerialNumber, NewRO, ProductCode)){
						return false
					}
				} else if (History = "Not Installed") {
					if (!this.ProductAdd(SerialNumber, NewRO, ProductCode)){
						return false
					}
				} else {
					return false
				}
			}
			if Call:=this.Create(SerialNumber,newRO) {
				OutputDebug % call
				this.print(newRO,Call)
				return true
			} else {
				return False
			}
		}
	
		ROisFree(RO){
			wb:=IEVget(Title) ;Gets active IE window
			wb.Navigate2("http://hypappbs005/SC5/SC_SerProd/aspx/serprod_main.aspx") ;navigates selected window to this url
			Loop{ ;begin loop to (wait for page to load)
				try{
			frame := wb.document.all(7).contentWindow
			pageTitle :=  frame.document.getElementById("txtFunctionText").innertext
			sleep 250
			}
			} until (pageTitle = "serialised product query" && wb.busy = 0)
			;loop has ended because page is correct and browser is reporting that loading has finished
			;if for whatever reason the script has reached here it and the page title is wrong then you get an error
			if  (pageTitle != "serialised product query"){
				msgbox, Incorrect page found - %pageTitle%
				return False
			}
			
			frame := wb.document.all(10).contentWindow ;select relevant frame
			frame.document.getElementById("txtSerReference2").value := RO ;insert serial number
			frame := wb.document.all(7).contentWindow ;select frame
			frame.document.getElementById("cmdsubmit").click ;click submit
			
			loop{ ;loop waiting for page to load whilst figuring out the landing page
				try{
					 frame := wb.document.all(10).contentWindow ;select frame
					 SiteNo := frame.document.getElementById("txtSerSiteNum").value ;check for site number
					 Records :=  frame.document.getElementById("lblRecordCount").innerText ;check for total records
					 
				}
			}until (SiteNo != ""|| Records != "") ;stop loop as soon as one exists
			wb:=
			if (records = 0){
				return true
			} else {
				return false
			}
		}
	
		HistoryCheck(SerialNumber){
			wb:=IEVget(Title) ;Gets active IE window
			wb.Navigate2("http://hypappbs005/SC5/SC_SerProd/aspx/serprod_main.aspx") ;navigates selected window to this url
			Loop{ ;begin loop to (wait for page to load)
				try{
			frame := wb.document.all(7).contentWindow
			pageTitle :=  frame.document.getElementById("txtFunctionText").innertext
			OutputDebug % wb.busy
			sleep 250
			}
			} until (pageTitle = "serialised product query" && wb.busy = 0)
			;loop has ended because page is correct and browser is reporting that loading has finished
			;if for whatever reason the script has reached here it and the page title is wrong then you get an error
			if  (pageTitle != "serialised product query"){
				msgbox, Incorrect page found - %pageTitle%
				return
			}
			frame := wb.document.all(10).contentWindow ;select relevant frame
			frame.document.getElementById("txtSerNum").value := SerialNumber ;insert serial number
			frame := wb.document.all(7).contentWindow ;select frame
			frame.document.getElementById("cmdsubmit").click ;click submit
			loop{ ;loop waiting for page to load whilst figuring out the landing page
				try{
					 frame := wb.document.all(10).contentWindow ;select frame
					 SiteNo := frame.document.getElementById("txtSerSiteNum").value ;check for site number
					 Records :=  frame.document.getElementById("lblRecordCount").innerText ;check for total records
					 
				}
			}until (SiteNo != ""|| Records != "") ;stop loop as soon as one exists
			if (SiteNo = "ZULU" && Records = ""){
				return "ZULU"
			} else if (siteNo != "" && siteNo != "ZULU" && Records = ""){
				return "Not Zulu"
			} else if (Records = "0" && SiteNo = ""){
				return "Not Installed"
			} else {
				return False
			}
			
		}
		
		ZuluStock(NewRO) {
			wb:=IEVget(Title)
			frame := wb.document.all(10).contentWindow
			
			if (ExistingRO := frame.document.getElementById("txtSerReference2").value) {
			frame.document.getElementById("txtSerReference1").value := ExistingRO ;move the existing RO number to the last RO field
		}
			
			frame.document.getElementById("txtSerReference2").value := NewRO
			NewRO := frame.document.getElementById("txtSerReference2").value
			IfNotInString,NewRO,480
			{
			msgbox,4,RO Number verify,Something seem's off with your RO number`nare you sure you want to continue with %NewRO%
				IfMsgBox, No 
				{
					return false
				}
			}
			
			Product:=frame.document.getElementById("cboSerProdNum").value
			FormatTime, KMN_Date, YYYYMMDD, dd/MM/yyyy
			frame.document.getElementById("dtpSerInstallDate").value := KMN_Date
			frame.document.getElementByid("cboSerSeStatCode").value := "REP"
			
			IfNotInString, Product, 177 
			{
				MsgBox,4,Product Code Verify, Product Code is showing as %Product%. `nIs this Correct
				IfMsgBox, No
				{
					return false
				}
			}
			frame := wb.document.all(7).contentWindow
			;=================================
			PageAlert()
			frame.document.getElementById("cmdSubmit").click
			pageloading(wb)
			;=================================
			return true
		}
			
		InstalledElsewhere(){
			wb:=IEVget(Title)
			OutputDebug, Install in a site that isn't ZULU
			frame := wb.document.all(7).contentWindow ;select frame
			;=================================
			frame.document.getElementById("cmdDelete").click
			WinWaitClose, Message from webpage
			pageloading(wb)
			;=================================
			return true
		}
		
		productAdd(SerialNumber, NewRO, ProductCode){
			wb:=IEVget(Title)
			frame := wb.document.all(9).contentWindow ;select frame
			frame.document.getElementById("lblSerProdAdd").click  ;add the part
			Loop{ ;begin loop to (wait for page to load)
				try{
					frame := wb.document.all(7).contentWindow
					pageTitle :=  frame.document.getElementById("txtFunctionText").innertext
					sleep 250
				}
			} until (pageTitle = "serialised product add" && wb.busy = 0)
			OutputDebug, productAdd - reached marker 1
			
			frame := wb.document.all(10).contentWindow ;select frame
			frame.document.getElementById("txtSerNum").value :=SerialNumber ;check for site number
			
			IfNotInString,NewRO,480
				{
				msgbox,4,RO Number verify,Something seem's off with your RO number`nare you sure you want to continue with %NewRO%
					IfMsgBox, No 
					{
						return false
					}
				}
			frame.document.getElementById("txtSerReference2").value := newRO
			OutputDebug, New RO added to field
			frame.document.getElementByid("cboSerSeStatCode").value := "REP"
			OutputDebug, Status set to REP
			IfNotInString, ProductCode, 177 
				{
					MsgBox,4,Product Code Verify, Product Code is showing as %productCode%. `nIs this Correct
					IfMsgBox, No
					{
						return false
					}
				}
			productCode:= LTrim(productCode, "0")
			OutputDebug, removed trailing zero's from product
			frame.document.getElementById("cboSerProdNum").value := productCode
			OutputDebug, Product Code inserted
			frame.document.getElementById("txtSerSiteNum").value := "ZULU"
			OutputDebug, site number set to zulu
			frame := wb.document.all(7).contentWindow
				;=================================
				PageAlert()
				frame.document.getElementById("cmdSubmit").click
				OutputDebug, submit has been hiti for product add
				PageLoading(wb)
				;=================================
			return true
		}
	
		Create(SerialNumber,newRO){
			wb:=IEVget(Title)
			wb.Navigate2("http://hypappbs005/SC5/SC_RepairJob/aspx/repairjob_create_wzd.aspx")
			sleep, 2500
			while ((wb.document.getElementById("cboJobWorkshopSiteNum") .value := "STOWS") != "STOWS")
				sleep, 100
			ModalDialogue()
			wb.document.getElementsByTagName("IMG")[0] .click
			wb.document.getElementById("cmdNext") .click
			PageLoading(wb)
			;=============
			wb.document.getElementById("cboCallSiteNum") .value := "ZULU"
			ModalDialogue()
			wb.document.getElementsByTagName("IMG")[5] .click 
			wb.document.getElementById("cmdNext") .click
			PageLoading(wb)
			;=============
			wb.document.getElementById("cmdNext") .click
			PageLoading(wb)
			;=============
			while ((wb.document.getElementById("cboCallSerNum") .value := SerialNumber) != SerialNumber)
				sleep, 100
			wb.document.getElementsByTagName("INPUT")[58] .click
			ModalDialogue()
			wb.document.getElementsByTagName("IMG")[19] .click
			if (wb.document.getElementById("cboJobPartNum").value = ""){
				msgbox, not found
				return false
			}
			while ((wb.document.getElementById("txtJobRef6") .value := newRO) != newRO)
				sleep, 100
			wb.document.getElementById("cmdNext") .click
			PageLoading(wb)
			;=============
			wb.document.getElementById("cboCallCalTCode") .value := "ZR1"
			ModalDialogue()
			wb.document.getElementsByTagName("IMG")[22] .click

			wb.document.getElementById("cboJobFlowCode") .value := "SWBOOKIN"
			ModalDialogue()

			FinishedTime:=A_Now
			EnvSub,FinsihedTime,this.StartTime,Seconds
			wb.document.getElementsByTagName("TEXTAREA")[4] .value := "==== Booking in finsihed in: " . FinsihedTime . " seconds ====`n==== Repair Order Number: " . newRO . " ====`n=======Powered by T-Enhanced======="
			;===================================
			PageAlert()
			wb.document.getElementById("cmdFinish") .click
			Pageloading(wb)
			wb.document.getElementsByTagName("INPUT")[119] .click
			wb.document.getElementById("cmdFinish") .click
			Pageloading(wb)
			frame := wb.document.all(10).contentWindow
			newCall:=frame.document.getElementsByTagName("INPUT")[0].value
			return newCall
		}
		
		print(newRO,newCall){
			global DymoLabel
			global DymoAddIn
			global DymoEngine
			OutputDebug % "newRO = " newRO " - SerialNumber = " SerialNumber
			DymoAddIn.Open("Modules/Zulu-book-in.label")
			DymoLabel.SetField( "RO-Number", newRO)
			DymoLabel.SetField( "Call-Number", newCall)
			DymoAddIn.Print( 1, TRUE )
			return true
		}
	
		__Delete(){
			this := ""
			Gui, KMN:Destroy
		}
	}
	
	
	class BookOut {
		__New() {
			static manifest
			static CallNumber
			gui, KMN:add, text, w150,Insert Call Number
			Gui, KMN:Add, Edit, w150 vCallNumber
			gui, KMN:add, text,w150,Insert Manifest Number
			FormatTime,Today,,dMyyyy
			Gui, KMN:Add, Edit, w150 vmanifest, 1%Today%
			Gui, KMN:add, button, gKMN_BookOut, Continue
			Gui, KMN: +AlwaysOnTop +ToolWindow +LastFound
			Gui, KMN:show,autosize, Book out
			WinWaitClose, Book out
			return "Gui Close"
			
			KMN_BookOut:
			gui,KMN:submit, NoHide
			if (!manifest || !CallNumber) {
				return False
			}
			gui,KMN:Submit
			if not RoNumber:= this.GetRO(CallNumber)
				return False
			if (!this.shipOut(CallNumber,RONumber,Manifest)){
				return False
			} else {
				return True
			}
	}
		__Delete(){
			this := ""
			Gui, KMN:Destroy
		}
		getRO(callNumber){
			URL := "http://hypappbs005/SC5/SC_RepairJob/aspx/RepairJob_Modify.aspx?CALL_NUM="callNumber
			wb:=IEVget(Title) ;Gets active IE window
			wb.Navigate2(URL) ;navigates selected window to this url
			while callNumberVerify != callNumber {
				try {
					callNumberVerify := wb.document.getElementById("txtCallNum").value
					sleep 250
				}
			}
				
			RONumber := wb.document.getElementById("txtJobRef6").value
			if !RONumber
				return False
			else
				return RONumber
		}		
		ShipOut(CallNumber,RONumber,Manifest){
			url := "http://hypappbs005/SC5/SC_RepairJob/aspx/repairjob_ship_wzd.aspx"
			wb:=IEVget(Title) ;Gets active IE window
			wb.Navigate2(URL) ;navigates selected window to this url
			while LoadVerify != "STOWS" {
				try {
					LoadVerify := wb.document.getElementById("cboJobNumWorkshopSiteNum").value := "STOWS"
					sleep 250
				}
			}
			wb.document.getElementById("txtInputJobNum").value := CallNumber
			wb.document.getElementById("cmdAddJobNum").click
			wb.document.getElementById("cmdNext").click
			PageLoading(wb)
			wb.document.getElementById("txtJobShipRef").value := Manifest
			wb.document.getElementById("txtJobShipRef1").value := RONumber
			wb.document.getElementById("cmdNext").click
			PageLoading(wb)
			if ((wb.document.getElementByID("cbaListJobPartNumLineArray").value) = ""){
				return false
			} else {
				PageAlert()
				wb.document.getElementById("cmdFinish").click
				WinWaitClose, Message from webpage
				return true
		}
		
	}
	}
}