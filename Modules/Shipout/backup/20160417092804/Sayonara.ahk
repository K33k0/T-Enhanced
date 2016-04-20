Pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion)
JobType:= getJobType(Pwb)
ShipSite:= getShipSite(Pwb)
if shipsite = "ZULU"
	zuluShipout()
else
	oldShipout()




getJobType(Pwb){
	frame := Pwb.document.all(10).contentWindow
	return (frame.document.getElementByID("cboJobFlowCode").value)
}

getShipSite(Pwb){
	frame := Pwb.document.all(10).contentWindow
	return (frame.document.getElementById("cboJobShipSiteNum").value)
}

Class Shipout {
	static pwb
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
		pageloading(pwb)
		sleep, 500
		frame := Pwb.document.all(10).contentWindow
		frame.document.getElementById("cboCallUpdAreaCode").value := "WSF"
		ModalDialogue()
		frame.document.getElementsByTagName("IMG")[35].click
		WinWaitClose,Popup List -- Webpage Dialog,,5
		frame := Pwb.document.all(7).contentWindow
		PageAlert()
		
		return
	}
	
	oldShipout(){
		if not Pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion)
			return false
		sleep, 250
		frame := Pwb.document.all(10).contentWindow
		CallNum := ShipSite:= frame.document.getElementById("txtCallNum") .value
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