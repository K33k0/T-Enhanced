FileCreateDir, Modules/Database
FileInstall, InstallMe/PartDescriptions.ini,Modules/Database/PartDescriptions.ini, 1
FileInstall, InstallMe/partList.ini,Modules/Database/partList.ini,1
FileInstall, InstallMe/Parts-Request.msg,Modules/Parts-Request.msg,1
PartMove := new Movement(settings)
PartMove.ini := new PartMove.ini("default")


gui,Move2:add,Text,,Select Manufacturer
gui,Move2: add, DDL, vSelectedSection gManuUpdate w200, % PartMove.ini.Sections()
gui, Move2: +AlwaysOnTop +ToolWindow +OwnDialogs -DPIScale 
gui, Move2:show,, Parts Movement
WinActivate, Parts Movement
return

ManuUpdate:
gui, Move2:submit,nohide
;GuiControl,disable,SelectedSection

GuiControl, Move2:, SelectedKey,% "|" . PartMove.ini.SectionKeys(selectedSection)
if (errorlevel) {
	gui,Move2:add,Text,w200,Select Unit Type
	gui, Move2:add, DDL, w200 vSelectedKey gTypeUpdate,% PartMove.ini.SectionKeys(selectedSection)
}
gui, Move2: show, AutoSize

return

TypeUpdate:
gui, Move2:submit, NoHide

selectedKey :=  PartMove.ini.SectionkeyValues(SelectedSection, SelectedKey)

GuiControl,Move2:, textCheck,Select Parts
if (errorLevel){
	gui,Move2:add,Text, vtextCheck w200,Select Parts
	gui,Move2:add, button, w200 xm vgoButton gPartMoveGo Disabled, Submit
	gui,Move2:add, button, w200 xm vdescButton gPartMoveDesc, Description Lookup
	Loop % PartMove.ini.KeyValues(selectedKey).MaxIndex()
	{
		gui, Move2:add, DDL, w140 xm vSelectedKey%A_Index% genableSubmit, % PartMove.ini.filteredParts
		gui, Move2:Add,edit, w40 yp x+2
		gui, Move2:add,updown, vKeyQuantity%A_Index%
		gui, Move2:add, text, vstatusText%A_Index% w20 h20 yp0 x+4,
		
		
	}until A_index > 4
	
	
} else {
	Loop % PartMove.ini.KeyValues(selectedKey).MaxIndex()
	{
		GuiControl,Move2:, SelectedKey%A_Index% , % "|" . PartMove.ini.filteredParts
		if (errorlevel){
			GuiControl,Move2:move, goButton,yp
			GuiControl,Move2:move, descButton,yp
			gui, Move2:add, DDL, w140 xm vSelectedKey%A_Index% genableSubmit, % PartMove.ini.filteredParts
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
thelist := PartMove.ini.filteredParts
Loop, parse, thelist , |
{
	IniRead,tempdescription, Modules/Database/PartDescriptions.ini,PartDescriptions,%A_LoopField%
	LV_Add("", A_LoopField , tempDescription, "£   -   ")
	LV_ModifyCol()  
}
LV_ModifyCol()  
gui, Move2: show, AutoSize
return

partLookup:
if (A_guievent = "DoubleClick"){
	LV_Modify(A_eventinfo,,,,PartMove.priceCheck(A_EventInfo))
	LV_ModifyCol() 
}
return

PartMoveGo:
gui,Move2:submit, nohide
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
		
		if (PartMove.MovePart(currentPart,selectedQuantity)) {
			sleep, 150
			PartMove.queuePrint(currentPart,selectedQuantity)
			sleep, 150
			partMove.GetPartLocation(currentPart)
			sleep, 150
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

partmove.print()

Move2GuiClose:
Move2GuiEscape:
PartMove := ""
move2LV := False
gui,Move2:destroy
return


;~ #IfWinActive,Parts Movement
;~ {
   ;~ $WheelDown::
    ;~ if selectedKey
        ;~ return
    ;~ else
        ;~ Send {WheelDown}
    ;~ return

    ;~ $WheelUp::
    ;~ if selectedKey
       ;~ return
    ;~ else
        ;~ Send {WheelUp}
    ;~ return
;~ }


class Movement
{
	requestedPart := {}
	partLocation := {}
	
	__New(settings)
	{
		this.settings := settings
        ;ini := new this.ini("default")
        ;gui:= new this.gui("default")
	}
	
	__Delete()
	{
		RIni_Shutdown(1)
		loop, 5 {
			selectedKey%A_Index% := ""
			KeyQuantity%A_Index% := ""
		}
	}
	class ini
	{
		databasePath := A_ScriptDir . "/modules/Database/partList.ini"
		static filteredParts
		
		Sections()
		{
			partList := this.databasePath
			IniRead, iniContents, %partList%
			StringReplace, iniContents, iniContents, `n,|, 1 
			;StringTrimRight, iniContents, iniContents, 1
			return iniContents
		}
		
		SectionKeys(Manufacturer)
		{
			partList := this.databasePath
			IniRead, iniContents, %partList%, %manufacturer%
			
			iniContents := StrSplit(iniContents, "`n")
			Loop % iniContents.MaxIndex()
			{
				this_line := SubStr(iniContents[a_index], 1, InStr(iniContents[a_index], "=") - 1) 
				Types .= this_line . "|"
			}
			StringTrimRight, Types, Types, 1
			
			return Types
		}
		
		SectionkeyValues(Manufacturer, UnitType)
		{
			partList := this.databasePath
			IniRead, parts, %partlist%, %Manufacturer%, %UnitType%
			parts := Trim(parts)
			this.filteredParts := parts
			return parts
		}
		
		KeyValues(FilteredParts)
		{
			return  StrSplit(FilteredParts, "|")
		}
	}
	
	class gui
	{
		
	}
	
	MovePart(part,quantity)
	{
    ;preMoveStock := this.partVerify(part, this.settings.Benchkit)    
		sleep 250
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
		frame.document.getelementbyID("cboDestSiteNum").value := this.settings.Benchkit ;input engineer number
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
		;WinwaitClose,Message from webpage,,5
			PartMovePointer.quit()
			PartMovePointer := ""
			sleep, 100
		
    ;postMoveStock := this.partVerify(part, this.settings.Benchkit)
		return true
	}
	
	queuePrint(part,quantity) 
	{
		this.requestedPart[part] := quantity
		return true
	}
	
	GetpartLocation(part)
	{
		SecondaryPointer:=IEVget(Title)
		URL:="http://hypappbs005/SC5/SC_StockControl/aspx/StockControl_modify.aspx?SiteNo=STOWPARTS&PartNo=" . part ;set the url
		SecondaryPointer.Navigate2(URL,4096) ;navigate the hijacked session to a new tab opening the set url
		Loop {
			try {
				SecondaryPointer:=IEGetURL("http://hypappbs005/SC5/SC_StockControl/aspx/StockControl_modify.aspx?SiteNo=STOWPARTS&PartNo=" . part)  ;get session by url
				stockCheck := SecondaryPointer.document.GetElementById("txtTotalQty").value ;set the value of the field
			}
		}Until (stockCheck != "") ;break the loop when the field is set to the correct field
		StockLocation :=  SecondaryPointer.document.GetElementById("txtLocation").value
		if (stockLocation = "") {
			StockLocation :=  SecondaryPointer.document.GetElementById("txtBinLoc").value
		}
		if  (StockLocation = "") {
			StockLocation := false
		}
		this.partLocation[part] := StockLocation
			Sleep 100
			SecondaryPointer.quit()
			SecondaryPointer := ""
			Sleep 100
		return true
	}
	
	print()
	{
		global DymoAddin
		global DymoLabel
		DymoAddIn.Open("Modules\Part Order.label")
		DymoAddin.StartPrintJob()
		
		For key, value in this.requestedPart
		{
			OutputDebug % key . " = " value
			DymoLabel.SetField( "Part1", key) 
			IniRead,description, Modules/Database/PartDescriptions.ini,PartDescriptions,%key%
			DymoLabel.SetField( "Description1", description) 
			DymoLabel.SetField( "Quantity1", value) 
			DymoLabel.SetField( "Engineer", this.settings.Engineer)
			if (this.partLocation[key]) {
				DymoLabel.SetField( "Location1", this.partLocation[key]) 
			}
			DymoAddIn.Print( 1, TRUE )
		}
		DymoAddin.EndPrintJob()
	}
	
	partVerify(PartNo, StockLocation)
	{       
		bpwb:= ievget(Title)
		baseUri:= "http://hypappbs005/SC5/SC_StockControl/aspx/StockControl_modify.aspx"
		uri := "?SiteNo=" . StockLocation . "&PartNo=" . PartNo
		bpwb.Navigate2(baseUri . uri, 2048)
		loop {
			try{
                    bpwb.Navigate("javascript: alert = function(){};")
                    bpwb := IEGetUrl(baseUri . uri)
                    Loaded := bpwb.document.GetElementByID("lblTotalQuantity").innerText
			}
		}until (Loaded != "")
		value := bpwb.document.GetElementByID("txtTotalQty").value
		if (Value = ""){
			Value := 0
		}
		IfWinExist,Message from webpage
		{
			WinWaitClose, Message from webpage
		}
		bpwb.quit()
		return Value
	}
	
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
}