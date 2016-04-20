;#Include lib/functions.ahk

global TesseractVersion := "5.40.14"
Shipout := new ShipOut()


if !(SHIPOUT_POINTER := ShipOut.ConnectAndVerifySupportedPage()){
	msgbox Failed to hook into shipout
}
if (ShipOut.bIsCallZulu()){
	Gosub, Ship_to_ZULU
	return
} else {
	Gosub, Ship_to_STOKGOODS
	return
}

Ship_to_ZULU:

Loop{
	Try{
		frame := Pwb.document.all(10).contentWindow
		PageLoaded:= frame.document.getElementsByTagName("Label")[0].innertext
	}
}Until (PageLoaded = "Job Details")
PageLoaded:= ""
SHIPOUT_FRAME := SHIPOUT_POINTER.document.all(10).contentWindow
SHIPOUT_FRAME.document.getElementByID("cboJobFlowCode").value:="ZULUAW"
SHIPOUT_FRAME.document.getElementByID("cboCallAreaCode").value:="WSF"

sleep, 250
SHIPOUT_FRAME := SHIPOUT_POINTER.document.all(7).contentWindow
Loop{
	Try{
		PageLoaded:= SHIPOUT_FRAME.document.getElementByID("cmdSubmit").value
	}
}Until (PageLoaded = "submit")
PageLoaded:=""
SHIPOUT_FRAME.document.getElementById("cmdSubmit").click
pageloading(SHIPOUT_POINTER)
sleep, 500
SHIPOUT_FRAME := SHIPOUT_POINTER.document.all(10).contentWindow
SHIPOUT_FRAME.document.getElementById("cboCallUpdAreaCode").value := "WSF"
ModalDialogue()
SHIPOUT_FRAME.document.getElementsByTagName("IMG")[35].click
WinWaitClose,Popup List -- Webpage Dialog,,5
SHIPOUT_FRAME := SHIPOUT_POINTER.document.all(7).contentWindow
PageAlert()

return

Shipout_to_STOKGOODS:
SHIPOUT_CALLNUMBER := ShipOut.getCallNumber()
return


class ShipOut {
	static WebPointer:=""
	ConnectAndVerifySupportedPage(){
		if (this.WebPointer:= IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion)) {
			return true
		} else {
			msgbox, Failed to connect to a supported page
			return false
		}
	}
	static CallNumber:=""
	getCallNumber(){
		frame := this.WebPointer.document.all(10).contentWindow
		this.CallNumber := frame.document.getElementByID("txtCallNum").value
		return
	}
	static ShipSite:=""
	bIsCallZulu(){
		frame := this.WebPointer.document.all(10).contentWindow
		this.ShipSite:=frame.document.getElementById("cboJobShipSiteNum").value
		return (this.ShipSite == "ZULU")
	}
}