databasePath := "/modules/Database/DatabaseStart/partList.ini"

PartMove := new Movement("default")
PartMove.ini := new PartMove.ini("default")

        
gui,Move2:add,Text,,Select Manufacturer
gui,Move2: add, DDL, vSelectedSection gManuUpdate, % PartMove.ini.Sections()
gui, Move2: +OwnDialogs
Gui,Move2: show
return

ManuUpdate:
gui, Move2:submit,nohide
GuiControl,disable,SelectedSection
gui,Move2:add,Text,,Select Unit Type
gui, Move2:add, DDL, vSelectedKey gTypeUpdate,% PartMove.ini.SectionKeys(selectedSection)
gui, Move2: show, AutoSize
return

TypeUpdate:
gui, Move2:submit, NoHide
GuiControl,disable,SelectedKey
gui,Move2:add,Text,,Select Parts
selectedKey :=  PartMove.ini.SectionkeyValues(SelectedKey)

Loop % PartMove.ini.KeyValues().MaxIndex()
{
    gui, Move2:add, DDL, vSelectedKey%A_Index%, % PartMove.ini.SectionKeyValue
    gui, Move2:Add,edit
    gui, Move2:add,updown, vKeyQuantity%A_Index%
}until A_index > 4

gui,Move2:add, button, gPartMoveGo, Submit

gui, Move2: show, AutoSize
return

PartMoveGo:
gui,Move2:submit
loop, 5 {
    if (selectedKey%A_Index%) {
        currentPart := selectedKey%A_Index%
        selectedQuantity := KeyQuantity%A_Index%
        if  (selectedQuantity = "" OR selectedQuantity =  0) {
            InputBox, selectedQuantity,Select Quantity, Input correct quantity - current set to %selectedQuantity%
        }
        
        PartMove.MovePart(currentPart,selectedQuantity)
        
    }
   
}
For key, value in PartMove.parts
    MsgBox %key% = %value%

loop, 5 {
    selectedKey%A_Index% := ""
    KeyQuantity%A_Index% := ""
}
PartMove := ""
gui,Move2:destroy
return


class Movement
{
    Parts := {}
    __New()
    {
        ini := new this.ini("default")
        gui:= new this.gui("default")
    }
    
    __Delete()
    {
        this.destroy()
    }
    class ini
    {
        databasePath := "/modules/Database/DatabaseStart/partList.ini"
        instance := ""
        selectedSection := ""
        selectedKey:= ""
        SectionkeyValue:= ""
        __New() 
        {
            this.read(1)
        }
        Read(instance)
        {
            this.instance := instance
            RIni_Read(this.instance, A_ScriptDir . this.databasePath)
        }
        
        Sections()
        {
            return RIni_GetSections(this.instance,"|")
        }
        
        SectionKeys(SelectedSection)
        {
            this.SelectedSection := SelectedSection
            return RIni_GetSectionKeys(this.instance,this.SelectedSection, "|")
        }
        
        SectionkeyValues(key)
        {
            this.selectedKey:= key
            OutputDebug % this.selectedKey . " - " . this.SelectedSection
            tempsectionkeyvalue := Rini_GetKeyValue(this.instance, this.SelectedSection, this.SelectedKey)
            this.SectionkeyValue := Trim(tempsectionkeyvalue)
            return this.SectionKeyValue
        }
        
        KeyValues()
        {
            return  StrSplit(this.sectionKeyValue, "|")
        }
    }
    
    class gui
    {
        __New()
        {
            
        }
    }
    
    MovePart(part,quantity)
    {
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
    IniRead,Engineer,%Config%,Engineer,Number ;read engineer number
	
	frame.document.getelementbyID("cboDestSiteNum").value := Engineer ;input engineer number
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
	;frame := PartMovePointer.document.all(6).contentWindow
    PartMovePointer.quit
    return this.queuePrint(part,quantity)
}

    queuePrint(part,quantity) 
    {
    this.Parts[part] := quantity
    return this.Parts
    }
}