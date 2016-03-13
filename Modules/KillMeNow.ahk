try {
gui,KMN:Destroy
}
Logistics := new Logistics()
Logistics.Bookin := new Logistics.Bookin()
gui, KMN:add, text,,Insert Serial Number
Gui, KMN:Add, Edit, vSerialNumber
gui, KMN:add, text,,Insert Repair Order Number
Gui, KMN:Add, Edit, vnewRO
gui, KMN:add, text,,Insert Product Code
Gui, KMN:Add, Edit, vproductCode,
Gui, KMN:add, button, gKMN_Submit, Continue
Gui, KMN: +AlwaysOnTop +ToolWindow
Gui, KMN:show
return

KMN_Submit:
gui, KMN:submit, NoHide
if (!SerialNumber || !newRO || !productCode) {
	return
}
gui, KMN:submit
if (!Logistics.Bookin.ROisFree(newRO)){
	msgbox, RO is in use
	gui, KMN:Destroy
	Logistics.Bookin:= ""
	return
}
StringUpper, SerialNumber, SerialNumber
StartTime := A_Now
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

if (SiteNo = "ZULU" && Records = ""){ ;if sitenumber is ZULU & records is nothingthen do some stuff relating to the part being in zulu
	
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
				return
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
				return
			}
		}
		frame := wb.document.all(7).contentWindow
		;=================================
		PageAlert()
		frame.document.getElementById("cmdSubmit").click
		pageloading(wb)
		;=================================
		gosub, KMN_Create
		return
} else if (siteNo != "" && siteNo != "ZULU" && Records = "") { ; if  site number isn't blank and it doesnt  equal zero and Records is blank then the unit is installed in a store
	OutputDebug, Install in a site that isn't ZULU
	frame := wb.document.all(7).contentWindow ;select frame
		;=================================
		frame.document.getElementById("cmdDelete").click
		WinWaitClose, Message from webpage
	
		pageloading(wb)
		;=================================
	gosub, productAdd
	return

}else if (Records = "0" && SiteNo = "") { ; if  records = 0 and no site number exists then the unit does not exist on the system
	gosub, productAdd
	return

} else { ;if the if statements havn't caught is already then failsafe
	msgbox, an error has occurred whilst trying to determine the units location
	return
}


productAdd:
	OutputDebug, running productAdd
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
	
	if (ExistingRO) {
		frame.document.getElementById("txtSerReference1").value := ExistingRO ;move the existing RO number to the last RO field
		outputDebug, inserting pre-existing RO into required field
	}
	IfNotInString,NewRO,480
		{
		msgbox,4,RO Number verify,Something seem's off with your RO number`nare you sure you want to continue with %NewRO%
			IfMsgBox, No 
			{
				return
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
				return
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
		gosub, KMN_Create
return

KMN_Create:
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
	return
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
EnvSub,FinsihedTime,StartTime,Seconds
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
;===================================

DymoAddIn.Open("Modules/Zulu-book-in.label")
DymoLabel.SetField( "RO-Number", newRO)
DymoLabel.SetField( "Call-Number", newCall)
DymoAddIn.Print( 1, TRUE )
gui, KMN:Destroy
Logistics.Bookin:= ""
return

class Logistics{
	
	__New() {
	}
	class Bookin {
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
			return
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
	}
	class BookOut {
		__New(){
			while not Manifest
				InputBox, Manifest, Manifest Input, Insert Manifest Number
			if not Pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion){
				MsgBox Error accessing page
				return
			} else {
				frame := Pwb.document.all(10).contentWindow
				CallNum:= frame.document.getElementsByTagName("INPUT")[0] .value
				RONum:= frame.document.getElementByID("txtJobRef6").value
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
				Pwb.document.getElementById("txtJobShipRef").value := Manifest
				Pwb.document.getElementById("txtJobRef1").value := RONum
				Pwb.document.getElementById("cmdNext") .click
				IELoad(Pwb)
				return
				Pwb.document.getElementsByTagName("INPUT")[48] .click
				Pwb.document.getElementById("cmdFinish") .click
				pageAlert()
				IELoad(Pwb)
				Pwb.document.getElementsByTagName("INPUT")[40] .click
				Pwb.document.getElementById("cmdFinish") .click
				return
			}
		}
	}
	
}