
MultiLineAdd:
AddPartsRefresh:
Start:
LineNo := 1

Gui,ImacMultAdd:Add,Button,  gAddPartsDone, Done
Gui,ImacMultAdd:Add,Button, xp+75 yp vAddPartsReload gAddPartsReload, Reload
Gui,ImacMultAdd:Add,Text, xm BackgroundTrans, Input Partcode
Gui,ImacMultAdd:Add, Edit,vPartCode%LineNo% ,
Gui,ImacMultAdd:Add,Text,BackgroundTrans, Input Serial Number
Gui,ImacMultAdd:Add, edit,vSerial%LineNo% gAddNewLine
Gui,ImacMultAdd:  +AlwaysOnTop +ToolWindow +OwnDialogs -DPIScale

X:=GetWinPosX("T-Enhanced Multi Part Move")
Y:=GetWinPosY("T-Enhanced Multi Part Move")
if (X = "ERROR" || X= "" OR Y = "ERROR" || Y=""){
	Gui, ImacMultAdd: Show,AutoSize ,T-Enhanced Multi Part Move
} else {
	Gui, ImacMultAdd: Show, X%x% Y%y%  AutoSize,T-Enhanced Multi Part Move
}


return
AddNewLine:
Gui,ImacMultAdd:submit, nohide
If (Serial%LineNo% = ""){
	return
}
PartCode:=PartCode1
LineNo +=1
Gui,ImacMultAdd:Add, Edit,vSerial%LineNo% gAddNewLine ,
X:=GetWinPosX("T-Enhanced Multi Part Move")
Y:=GetWinPosY("T-Enhanced Multi Part Move")
Gui,ImacMultAdd: Show, x%x% y%y% AutoSize
return
#D::

Addpartsdone:
SaveWinPos("T-Enhanced Multi Part Move")
Gui,ImacMultAdd:Submit
I:=1
LineNo-=1
loop, %LineNo%{
	StringReplace,PartCode,PartCode,%A_SPACE%,`%,All
	Pwb := IEGet("FSRL_Create_Wzd - " TesseractVersion)
	ModalDialogue()
	Pwb.document.getElementById("cboWZPartNum").value :=PartCode
	Pwb.document.getElementById("cboWZPartNum_Container").click
	check:=Pwb.document.getElementById("cboWZPartDesc").value
	if(check = ""){
		MsgBox,Part not found. Input Manually
		IEload(Pwb)
		IEload(Pwb)
	}else{
		Pwb.document.getElementById("cmdNext").Click
		IEload(Pwb)
		Pwb.document.getElementById("cmdNext").Click
		IEload(Pwb)
	}
	if(Pwb.document.getElementById("chkAllowNewStockUsed").outerHTML = 	"<INPUT id=chkAllowNewStockUsed type=checkbox name=chkAllowNewStockUsed>"){
		Pwb.document.getElementById("chkAllowNewStockUsed").Click
	}
	ModalDialogue()
	pwb.document.getElementByID("cboFSRLStockSiteNum").Value:=""
	pwb.document.getElementByID("cboFSRLIDNum").Value:=Serial%I%
	pwb.document.getElementByID("cboFSRLIDNum_Container").Click
	Pwb.document.getElementById("cmdFinish").click
	IEload(Pwb)
	I+=1
	if (Serial%I% = ""){
		Pwb.document.getElementById("cmdFinish").click
		Gui,ImacMultAdd:Destroy
		return
	}else{
		Pwb.document.getElementById("cmdFinish").click
		IEload(pwb)
	}
}
Gui,ImacMultAdd:destroy
return
AddPartsReload:
if (reload != True){
	reload := true
	Guicontrol, , AddPartsReload,Confirm
	return
}
Gui,ImacMultAdd:Destroy
Reload := False
gosub, AddPartsRefresh
return


ImacMultiAddGuiClose:
ImacMultiAddGuiEscape:
Gui,ImacMultAdd:Destroy
return