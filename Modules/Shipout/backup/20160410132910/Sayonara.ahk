#Include lib/functions.ahk

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
return

Shipout_to_STOKGOODS
return

ShipOut.getCallNumber()
ExitApp

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