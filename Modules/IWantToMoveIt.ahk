/*
Owner - Kieran Wynne
Module -I Like To Move It, Move It
Version - 1.0

----- Info -----
Due to the typical user's technical skill SQL has been avoided and instead opted for something much more pliable; INI files
*/
x:=""
y:=""
Listofmovedparts:=""

;=======Array Wipe==========
I=1
while	PartcodeQuantity%i% != "" {
PartcodeQuantity%i% := ""
I+=1
}
I := 1
while	Partcode%i% != ""  {
PartcodeCode%i% := ""
I+=1
}
;========================


global  PartsDatabase:="modules/database/PartsDataBase.ini"
FileInstall,C:/Users/kieran.wynn/Projects/Git/T-Enhanced [ZULU]/InstallMe/PartsDataBase.ini,modules/database/PartsDataBase.ini,1
FileInstall,C:/Users/kieran.wynn/Projects/Git/T-Enhanced [ZULU]/InstallMe/Parts-Request.msg,modules/Parts-Request.msg,1
;==============Get eligible parts from the ini db ====================
iniread,EligibleProducts,modules/database/PartsDataBase.ini,EligibleProducts,List
;=======================================================

;============== Select  product gui ==========================
gui,ProductSelector:add,text,,product
Gui,ProductSelector:add,ComboBox,0x100 vProductcode,%EligibleProducts%
gui,ProductSelector:add,button,gsubmitstage1,submit
gui,ProductSelector:add,button,gRequestParts,Request Part
Gui, ProductSelector:- +AlwaysOnTop +ToolWindow
X:=GetWinPosX("T-Enhanced Product Select")
Y:=GetWinPosY("T-Enhanced Product Select")
if (X = "" OR Y = "" OR X= "Error" OR Y="Error"){
Gui, ProductSelector: Show, ,T-Enhanced Product Select
} else {
Gui, ProductSelector: Show, X%x% Y%y%  ,T-Enhanced Product Select
}
return
;====================================================


submitstage1:
SaveWinPos("T-Enhanced Product Select")
gui,ProductSelector:submit
;================ Get Parts using the supplied product ======================================
iniread,Eligibleparts,%PartsDatabase%,PartsbyProduct,%ProductCode%
if (EligibleParts="error") {
		MsgBox, %Productcode% has no eligibleparts for this function.
		gosub,EndMovement
		return
}
;=============================================================================


;================= Gui for selecting your parts ======================
gui,Partselector:add,text,w254 x5 center,%Productcode%
stringsplit,Partmoveloop,Eligibleparts,|
I=1
if PartmoveLoop0 > 5
	newloop = 5
else
	newloop := partmoveloop0
loop, %newloop% {
	gui,Partselector:add,combobox,vPartcode%i% x5 yp+30 w195 ,%Eligibleparts%
	gui,Partselector:add,edit, xp+205 yp w50
	gui,partselector:add,updown,Range0-1000 vPartcodeQuantity%i%
	Height:= Height - 30
	i+=1
}
gui,Partselector:add,button,gViewdescriptions x5 w254, View Part Descriptions

gui,Partselector:add,button,gEndMovement x5 w254,close
gui,Partselector:add,button,gsubmitparts x5 w254,Move Parts
Gui, Partselector:+AlwaysOnTop +ToolWindow

X:=GetWinPosX("T-Enhanced Part Select")
Y:=GetWinPosY("T-Enhanced Part Select")
if (X = "" OR Y = "" OR X= "Error" OR Y="Error"){
Gui, Partselector: Show, ,T-Enhanced Part Select

} else {
Gui, Partselector: Show, X%x% Y%y%,  T-Enhanced Part Select
}
return
;============================================================

;============= Subroutine to show descriptions for parts filtered by product ========
Viewdescriptions:
gui,Partselector:submit,nohide
SaveWinPos("T-Enhanced Part Select")
I=1
loop, %partmoveloop0% {
	parttemp := partmoveloop%i%
	iniread,descriptiontemp,%PartsDatabase%,PartDescriptions,%parttemp%
	gui,Partdescriptions:add,text,,%PartTemp%  =  %descriptiontemp%
	i+=1
}

gui,Partdescriptions:show
return
;=============================================================

;============== Commence the movements =========================
SubmitParts:
SaveWinPos("T-Enhanced Part Select")
gui,Partselector:submit

I = 1
Part := PartCode%I%
Quantity := PartcodeQuantity%i%
while (Part != "") {
; code for the submission
PartMovePointer:=IEVget(Title) ;hijack session
URL=http://hypappbs005/SC5/SC_StockMove/aspx/stockmove_frameset.aspx ;set the url
	PartMovePointer.Navigate2(URL,2048) ;navigate the hijacked session to a new tab opening the set url
	Loop {
		try {
			PartMovePointer:=IEGetURL("http://hypappbs005/SC5/SC_StockMove/aspx/stockmove_frameset.aspx")  ;get session by url
			Frame:=PartMovePointer.document.all(9).contentwindow
			Frame.document.GetElementById("cboPartNum").value := Part ;set the value of the field
		}
	}Until (Frame.document.GetElementById("cboPartNum").value = Part) ;break the loop when the field is set to the correct field

frame := PartMovePointer.document.all(9).contentWindow

;================= Page Inputs =============================
frame.document.getelementbyID("cboPartNum").value := Part ;--input part

iniread,AltlocationList,%PartsDatabase%,STOKGOODS,List ;--get list of parts that come from other places
IfInString,AltLocationList,%Part%  ; if part is in the list then give it an alt location
{
MovementSite := "STOKGOODS"
}
else
{
MovementSite := "STOWPARTS"
}
frame.document.getelementbyID("cboSourceSiteNum").value := MovementSite ;set the movement site to whatever was decided above
OutputDebug, [COK]Movement site is set to %MovementSite%
ModalDialogue() 
frame.document.getElementsByTagName("IMG")[2] .click 
WinWaitClose, ahk_class Internet Explorer_TridentDlgFrame
sleep, 500

while (frame.document.getelementbyID("txtSourceSerialised").value = "")
	sleep, 250


if (frame.document.getelementbyID("txtSourceTotalQty").value != "" OR 0) {
	SourceQuantity := frame.document.getelementbyID("txtSourceTotalQty").value ;check current stock
	if (frame.document.getelementbyID("txtSourceTotalQty").value < Qunatity) {
		While (SourceQuantity < Qunatity) {
		SourceQuantity := frame.document.getelementbyID("txtSourceTotalQty").value
		Gui +LastFound +OwnDialogs +AlwaysOnTop
		inputbox, QUANTITY, Insufficient Stock, There is insufficient stock in STOWPARTS`n maximum available is %SourceQuantity%`ninput a new amount
		WinWaitClose
	}
	}
	IniRead,Engineer,%Config%,Engineer,Number ;read engineer number
	
	frame.document.getelementbyID("cboDestSiteNum").value := Engineer ;input engineer number
	ModalDialogue() 
	frame.document.getElementsByTagName("IMG")[6] .click 
	
	while (frame.document.getelementbyID("txtDestTotalNeed").value = "")
		sleep, 500
	
	frame.document.getelementbyID("txtMoveTotalQty").value := Quantity ;insert  quantity
	frame.document.getelementbyID("cboAdjustCode").value := "MV" ;set adjustment code
	frame.document.getelementbyID("txtReason").value := "Automated movement - TE" ;inserted reason
	frame.document.getelementbyID("cboSourceSiteNum").value := MovementSite ;insert movement site
	ModalDialogue() 
	frame.document.getElementsByTagName("IMG")[2] .click 
	sleep, 500
	frame.document.getelementbyid("chkAllowNewStockFlag").click  ;check the flag
	frame := PartMovePointer.document.all(6).contentWindow
	frame.document.getElementByID("cmdSubmit").Click ;submit
	sleep 500
	PageLoading(PartMovePointer)
	while PartMovePointer.busy
		sleep 250
} else {
	msgbox % Part . " is out of stock"
	PartcodeQuantity%i% = out of stock
}
PartMovePointer.quit

I+=1
Part := PartCode%I%
 Quantity := PartcodeQuantity%i%
}

DymoAddIn.Open("Modules\Part Order.label")
DymoAddin.StartPrintJob()
StringReplace,Engineer,Engineer,BK,,
DymoLabel.SetField( "Engineer", Engineer)
I = 1
Part := PartCode%I%
Quantity := PartcodeQuantity%i%
Loop,5 {
	if part {
DymoLabel.SetField( "Part1", Part) 
DymoLabel.SetField( "Quantity1", Quantity) 
iniread,description,%PartsDatabase%,PartDescriptions,%part%
DymoLabel.SetField( "Description1", description) 
OutputDebug % "[T-Enhanced] Label part set to " .  Part
OutputDebug % "[[T-Enhanced] label quantity set to " . Quantity
DymoAddIn.Print( 1, TRUE )
}
I+=1
Part := PartCode%I%
Quantity := PartcodeQuantity%i%
}

DymoAddin.EndPrintJob()

gui,Partdescriptions:destroy
gui,Partselector:destroy
gui,ProductSelector:destroy
Qunatuty:= ""
Listofmovedparts:=""
return
;============================================================

PartselectorGuiEscaper:
ProductSelectorGuiEscape:
EndMovement:
ProductSelectorGuiClose:
EligiblepartsGuiClose:
SaveWinPos("T-Enhanced Part Select")
gui,Partdescriptions:destroy
gui,Partselector:destroy
gui,ProductSelector:destroy
Qunatuty:= ""
Listofmovedparts:=""
return

RequestParts:
run, %A_scriptdir%/modules/Parts-Request.msg
return