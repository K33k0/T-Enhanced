FileCreateDir, Modules/Database
FileInstall, InstallMe/PartDescriptions.ini,Modules/Database/PartDescriptions.ini, 1
FileInstall, InstallMe/partList.ini,Modules/Database/partList.ini,1

readManufacturers(){
	IniRead, Sections, % settings.partList
	StringReplace, Sections, Sections, `n,|, 1
	return Sections
}
gui,Move2:add,Text,,Select Manufacturer
gui,Move2: add, DDL, vselectedManufacturer gManuUpdate w200, % readManufacturers()
gui, Move2: +AlwaysOnTop +ToolWindow +OwnDialogs -DPIScale 
gui, Move2:show,, Parts Movement
WinActivate, Parts Movement
return

ManuUpdate:
gui, Move2:submit,nohide
readManufacturerTypes(selectedManufacturer){
	IniRead, data, % settings.partList, % selectedManufacturer
	data := StrSplit(data, "`n")
	Loop % data.MaxIndex() {
		this_line := SubStr(data[a_index], 1, InStr(data[a_index], "=") - 1) 
		Types .= this_line . "|"
	}
	StringTrimRight, Types, Types, 1

	return Types
}

GuiControl, Move2:, SelectedType,% "|" . readManufacturerTypes(selectedManufacturer)
if (errorlevel) {
	gui,Move2:add,Text,w200,Select Unit Type ; <--- Error Same variable cant be used in more than one control
	gui, Move2:add, DDL, w200 vSelectedType gTypeUpdate,% readManufacturerTypes(selectedManufacturer)
}
gui, Move2: show, AutoSize

return

TypeUpdate:
gui, Move2:submit, NoHide
readFilteredParts(selectedManufacturer,SelectedType){
	IniRead, parts, % settings.partList, %selectedManufacturer% , %SelectedType%
	parts := Trim(parts)
	return parts
}

GuiControl,Move2:, SelectPartsText,Select Parts
if (errorLevel){
	;~ SelectPartsText does not exist. This is creating that section
	gui,Move2:add,Text, vSelectPartsText w200,Select Parts
	gui,Move2:add, button, w200 xm vgoButton gPartMoveGo Disabled, Submit
	gui,Move2:add, button, w200 xm vdescButton gPartMoveDesc, Description Lookup
	Loop % StrSplit(readFilteredParts(selectedManufacturer,SelectedType), "|").MaxIndex()
	{
		gui, Move2:add, DDL, w140 xm vSelectedKey%A_Index% genableSubmit, % readFilteredParts(selectedManufacturer,SelectedType)
		;~ This drop down works, it successfully changes the list of parts
		gui, Move2:Add,edit, w40 yp x+2
		gui, Move2:add,updown, vKeyQuantity%A_Index%
		gui, Move2:add, text, vstatusText%A_Index% w20 h20 yp0 x+4,
		
		
	}until A_index > 4
	
	
} else {
	;~ SelectPartsText exists already, just edit the values instead
	Loop % StrSplit(readFilteredParts(selectedManufacturer,SelectedType), "|").MaxIndex()
	{
		GuiControl,Move2:, SelectedKey%A_Index% , % "|" . readFilteredParts(selectedManufacturer,SelectedType)
		if (errorlevel){
			;~ if the a field needs adding this will do it, it also does something with the buttons...
			GuiControl,Move2:move, goButton,yp
			GuiControl,Move2:move, descButton,yp
			gui, Move2:add, DDL, w140 xm vSelectedKey%A_Index% genableSubmit, % readFilteredParts(selectedManufacturer,SelectedType)
			gui, Move2:Add,edit, w40 yp x+2
		}
	}until A_index > 4
}
gui, Move2: show, AutoSize
return

enableSubmit:
GuiControl, enable, goButton
return

PartMoveDesc:
if (!move2LV){
	Gui, Move2:add, ListView,w450 x220 ym r16 gPartLookup, Part Code|Description|Price
	Move2LV := true
}
LV_Delete()
thelist := readFilteredParts(selectedManufacturer,SelectedType)
Loop, parse, thelist , |
{
	IniRead,tempdescription, % settings.partlistDesc ,PartDescriptions,%A_LoopField%
	LV_Add("", A_LoopField , tempDescription, "£   -   ")
	LV_ModifyCol()  
}
LV_ModifyCol()  
gui, Move2: show, AutoSize
return

partLookup:
if (A_guievent = "DoubleClick"){
	LV_Modify(A_eventinfo,,,,priceCheck(A_EventInfo))
	LV_ModifyCol() 
}
return

priceCheck(RowNumber){
		LV_GetText(part, RowNumber)
		StringReplace,part,part,%A_space%,+
		bpwb:= ievget()
		baseUri:= "http://hypappbs005/SC5/SC_StockControl/aspx/StockControl_modify.aspx"
		uri := "?SiteNo=STOWPARTS&PartNo=" . part
		bpwb.Navigate2(baseUri . uri, 2048)
		loop {
			try{
                    bpwb := IEGetUrl(baseUri . uri)
                    Loaded := bpwb.document.GetElementByID("txtCost").value
			}
		}until (Loaded != "")
		value := bpwb.document.GetElementByID("txtCost").value
		bpwb.quit()
		
		return "£" . value
	}
return

PartMoveGo:
gui,Move2:submit, nohide
DymoAddIn.Open("Modules\Part Order.label")
DymoAddin.StartPrintJob()
loop, 5 {
	if (selectedKey%A_Index%) {
		currentPart := selectedKey%A_Index%
		selectedQuantity := KeyQuantity%A_Index%
		currentStatus := statusText%A_Index%
		Gui Font, cBlue s14
		GuiControl, Move2:Font, statusText%A_Index%
		GuiControl,Move2:text ,statusText%A_Index%, % chr(0x00221E)
		while not selectedQuantity
			InputBox, selectedQuantity,Select Quantity, Input correct quantity - current set to %selectedQuantity%
		
		if (MovePart(currentPart,selectedQuantity)) {
			Gui Font, cGreen
			GuiControl, Move2:Font, statusText%A_Index%
			GuiControl,Move2:text ,statusText%A_Index%, % chr(0x002714)
		} else {
			Gui Font, cRed
			GuiControl, Move2:Font, statusText%A_Index%
			GuiControl,Move2:text ,statusText%A_Index%, X
		}
	}
}
DymoAddin.EndPrintJob()

Move2GuiEscape:
Move2GuiClose:
PartMove := ""
move2LV := False
gui,Move2:destroy
return





MovePart(part,quantity){
		global DymoAddin
		global DymoLabel
	
		PartMovePointer:=IEVget(Title)
		URL=http://hypappbs005/SC5/SC_StockMove/aspx/stockmove_frameset.aspx ;set the url
		PartMovePointer.Navigate2(URL,2048) ;navigate the hijacked session to a new tab opening the set url
		Loop {
			try {
				PartMovePointer:=IEGetURL("http://hypappbs005/SC5/SC_StockMove/aspx/stockmove_frameset.aspx")  ;get session by url
				Frame:=PartMovePointer.document.all(9).contentwindow
				Frame.document.GetElementById("cboPartNum").value := Part ;set the value of the field
			}
		}Until (Frame.document.GetElementById("cboPartNum").value = Part) ;break the loop when the field is set to the correct field
		frame.document.getelementbyID("cboSourceSiteNum").value := "STOWPARTS"
		ModalDialogue() 
		frame.document.getElementsByTagName("IMG")[2] .click
		if (Frame.document.GetElementById("cboSourceSiteName").value = ""){
			PartMovePointer.quit
			return false
		}
		while (frame.document.getelementbyID("txtSourceTotalQty").value = "")
			sleep, 500
		sleep, 250
		
		SourceQuantity := frame.document.getelementbyID("txtSourceTotalQty").value
		if (SourceQuantity < quantity) {
			TrayTip, incorrect quantity
			InputBox,quantity,New Quantity, Not enough in stock.`nMax available: %SourceQuantity%`nSelect new amount
			WinWaitClose, New Quantity
			if (errorlevel = 1) {
				PartMovePointer.quit
				return false
			}
		} 
		frame.document.getelementbyID("cboDestSiteNum").value := settings.Benchkit ;input engineer number
		ModalDialogue() 
		frame.document.getElementsByTagName("IMG")[6] .click 
		
		while (frame.document.getelementbyID("txtSourceTotalNeed").value = "")
			sleep, 500
		frame.document.getelementbyID("txtMoveTotalQty").value := Quantity
		frame.document.getelementbyID("cboAdjustCode").value := "MV" ;set adjustment code
		frame.document.getelementbyID("txtReason").value := "Automated by T-Enhanced" ;inserted reason
		frame.document.getelementbyID("cboSourceSiteNum").value := "STOWPARTS" ;insert movement site
		ModalDialogue() 
		frame.document.getElementsByTagName("IMG")[2] .click 
		sleep, 500
		frame.document.getelementbyid("chkAllowNewStockFlag").click  ;check the flag
		PageAlert()
		frame := PartMovePointer.document.all(6).contentWindow
		
		frame.document.getElementByID("cmdSubmit").Click ;submit
		Frame:=PartMovePointer.document.all(9).contentwindow
		pageloading(PartMovePointer)
		while (Frame.document.GetElementById("cboPartNum").value)
			sleep, 500
		WinwaitClose,Message from webpage,,5
		
		PartMovePointer.Navigate("http://hypappbs005/SC5/SC_StockControl/aspx/StockControl_modify.aspx?SiteNo=STOWPARTS&PartNo=" . part) ;navigate the hijacked session to a new tab opening the set url
		Loop {
			try {
				PartMovePointer:=IEGetURL("http://hypappbs005/SC5/SC_StockControl/aspx/StockControl_modify.aspx?SiteNo=STOWPARTS&PartNo=" . part)  ;get session by url
				stockCheck := PartMovePointer.document.GetElementById("txtTotalQty").value ;set the value of the field
			}
		}Until (stockCheck != "") ;break the loop when the field is set to the correct field
		OutputDebug, [TE] %stockcheck%
		StockLocation :=  PartMovePointer.document.GetElementById("txtLocation").value
		OutputDebug, [TE] %StockLocation% 1
		if (stockLocation = "") {
			StockLocation :=  PartMovePointer.document.GetElementById("txtBinLoc").value
		}
		OutputDebug, [TE] %StockLocation% 2
		if  (StockLocation = "") {
			StockLocation := false
		}
		OutputDebug, [TE] %StockLocation% 3
		PartMovePointer.quit()
		PartMovePointer := ""
		
		DymoLabel.SetField( "Part1", part) 
		IniRead,description, % settings.partlistDesc ,PartDescriptions,%part%
		DymoLabel.SetField( "Description1", description) 
		DymoLabel.SetField( "Quantity1", quantity) 
		DymoLabel.SetField( "Engineer", settings.Engineer)
		DymoLabel.SetField( "Location1", StockLocation) 
		DymoAddIn.Print( 1, TRUE )
		return true
	}
