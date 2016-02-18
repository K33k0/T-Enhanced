databasePath := "/modules/Database/DatabaseStart/partList.ini"



RIni_Read(1, A_ScriptDir . databasePath)

SectionNames:= RIni_GetSections(1,"|") ;this will read each section
SectionKeys:= RIni_GetSectionKeys(1,SelectedSection) ;this will read each sections keys
SelectedKey=Printer
SectionKeyValue:= Rini_GetKeyValue(1, "HP", "Printer ")

gui,Move2:add,Text,,Select Manufacturer
gui,Move2: add, DDL, vSelectedSection gManuUpdate, %SectionNames%
Gui,Move2: show
return

ManuUpdate:
gui, Move2:submit,nohide
GuiControl,disable,SelectedSection
SectionKeys:= RIni_GetSectionKeys(1,SelectedSection, "|") ;this will read each sections keys
gui,Move2:add,Text,,Select Unit Type
gui, Move2:add, DDL, vSelectedKey gTypeUpdate,%SectionKeys%
gui, Move2: show, AutoSize

return
TypeUpdate:
gui, Move2:submit, NoHide
GuiControl,disable,SelectedKey
gui,Move2:add,Text,,Select Parts

SectionKeyValue:= Rini_GetKeyValue(1, SelectedSection, SelectedKey)
SectionKeyValue:= Trim(SectionKeyValue)
KeyValue := StrSplit(SectionKeyValue, "|")
Loop % KeyValue.MaxIndex()
{
    gui, Move2:add, DDL, vSelectedKey%A_Index%, %SectionKeyValue%
    gui, Move2:Add,edit
    gui, Move2:add,updown, vKeyQuantity%A_Index%
}until A_index > 4

gui,Move2:add, button, gPartMoveGo, Submit

gui, Move2: show, AutoSize
return

PartMoveGo:
gui,Move2:submit
loop, 5 {
    if selectedKey%A_Index% {
        
    }
}



loop, 5 {
    selectedKey%A_Index% := ""
}
gui,Move2:destroy
return

