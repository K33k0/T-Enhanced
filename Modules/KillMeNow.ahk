try {
gui,KMN:Destroy
}

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
	msgbox, unit is in zulu
	
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
		;PageAlert()
		;frame.document.getElementById("cmdSubmit").click
		;=================================
		gosub, KMN_Create
		return
} else if (Records = 0 && SiteNo = "") { ; if  records = 0 and no site number exists then the unit does not exist on the system
	msgbox, unit is not in any customer assets
	gosub, productAdd
	return

	


}else if (siteNo != "" && siteNo != "ZULU" && Records = "") { ; if  site number isn't blank and it doesnt  equal zero and Records is blank then the unit is installed in a store
	msgbox, deleting out of store
	frame := wb.document.all(7).contentWindow ;select frame
		;=================================
		;PageAlert()
		;frame.document.getElementById("cmdDelete").click
		;=================================
	gosub, productAdd
	return
} else { ;if the if statements havn't caught is already then failsafe
	msgbox, an error has occurred whilst trying to determine the units location
	return
}


msgbox, all done dude

productAdd:
	frame := wb.document.all(9).contentWindow ;select frame
	frame.document.getElementById("lblSerProdAdd").click  ;add the part
	Loop{ ;begin loop to (wait for page to load)
		try{
			frame := wb.document.all(7).contentWindow
			pageTitle :=  frame.document.getElementById("txtFunctionText").innertext
			sleep 250
		}
	} until (pageTitle = "serialised product add" && wb.busy = 0)
	
	frame := wb.document.all(10).contentWindow ;select frame
	frame.document.getElementById("txtSerNum").value :=SerialNumber ;check for site number
	
	if (ExistingRO) {
		frame.document.getElementById("txtSerReference1").value := ExistingRO ;move the existing RO number to the last RO field
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
	frame.document.getElementByid("cboSerSeStatCode").value := "REP"
	IfNotInString, ProductCode, 177 
		{
			MsgBox,4,Product Code Verify, Product Code is showing as %productCode%. `nIs this Correct
			IfMsgBox, No
			{
				return
			}
		}
	productCode:= LTrim(productCode, "0")
	frame.document.getElementById("cboSerProdNum").value := productCode
	frame.document.getElementById("txtSerSiteNum").value := "ZULU"
	frame := wb.document.all(7).contentWindow
		;=================================
		;PageAlert()
		;frame.document.getElementById("cmdSubmit").click
		;=================================
		gosub, KMN_Create
return

!C::
StartTime := A_Now
SerialNumber = D11D05834
StringUpper, SerialNumber, SerialNumber
newRO = 1872349
productCode = 0001770049577
wb:=IEVget(Title)
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
;wb.document.getElementById("cmdSubmit") .click
;===================================
msgbox, done

return

