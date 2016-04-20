;#Include lib/functions.ahk

global TesseractVersion := "5.40.14"
Shipout := new ShipOut()


if !(ShipOut.ConnectAndVerifySupportedPage()){
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
PageLoaded:=""
frame := shipOut.webpointer.document.all(10).contentWindow
frame.document.getElementByID("cboJobFlowCode").value:="ZULUAW"
frame.document.getElementByID("cboCallAreaCode").value:="WSF"

sleep, 250
frame := shipOut.webpointer.document.all(7).contentWindow
Loop{
	Try{
		PageLoaded:= frame.document.getElementByID("cmdSubmit").value
	}
}Until (PageLoaded = "submit")
PageLoaded:=""
frame.document.getElementById("cmdSubmit").click
pageloading(shipOut.webpointer)
sleep, 500
frame := shipOut.webpointer.document.all(10).contentWindow
frame.document.getElementById("cboCallUpdAreaCode").value := "WSF"
ModalDialogue()
frame.document.getElementsByTagName("IMG")[35].click;
WinWaitClose,Popup List -- Webpage Dialog,,5
frame := shipOut.webpointer.document.all(7).contentWindow
PageAlert()

return
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