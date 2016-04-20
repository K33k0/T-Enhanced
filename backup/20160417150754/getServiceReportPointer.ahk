getServiceReportPointer(){
	If Not Pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion){
		MsgBox Error accessing page
		return false
	} else {
		return Pwb
	}
}