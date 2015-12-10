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
global  PartsDatabase:="modules/database/PartsDataBase.ini"
FileInstall,C:/Users/kieran.wynn/Projects/Git/T-Enhanced [ZULU]/InstallMe/PartsDataBase.ini,modules/database/PartsDataBase.ini,1
iniread,EligibleProducts,modules/database/PartsDataBase.ini,EligibleProducts,List
Authorized=406

gui,ProductSelector:add,text,,product
Gui,ProductSelector:add,ComboBox,0x100 vProductcode,%EligibleProducts%|%currentproduct%
gui,ProductSelector:add,button,gsubmitstage1,submit

IniRead,Engineer,%Config%,Engineer,Number
StringReplace,Engineer,Engineer,BK,,
IfInString,Authorized,%Engineer%
gui,ProductSelector:add,button, gCheckProduct,check products


Gui, ProductSelector:- +AlwaysOnTop +ToolWindow
X:=GetWinPosX("T-Enhanced Product Select")
Y:=GetWinPosY("T-Enhanced Product Select")
if (X = "" OR Y = "" OR X= "Error" OR Y="Error"){
Gui, ProductSelector: Show, ,T-Enhanced Product Select
} else {
Gui, ProductSelector: Show, X%x% Y%y%  ,T-Enhanced Product Select
}
return

submitstage1:
SaveWinPos("T-Enhanced Product Select")
gui,ProductSelector:submit
iniread,Eligibleparts,%PartsDatabase%,PartsbyProduct,%ProductCode%
if (EligibleParts="error") {
		MsgBox, %Productcode% has no eligibleparts for this function. `n email your line manager to get some added.
		gosub,EndMovement
		return
}


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

SubmitParts:
SaveWinPos("T-Enhanced Part Select")
gui,Partselector:submit
I = 1
Part := PartCode%I%
Quantity := PartcodeQuantity%i%
while (Part != "") {
; code for the submission
PartMovePointer:=IEVget(Title)
URL=http://hypappbs005/SC5/SC_StockMove/aspx/stockmove_frameset.aspx
	PartMovePointer.Navigate2(URL,2048)
	Loop {
		try {
			PartMovePointer:=IEGetURL("http://hypappbs005/SC5/SC_StockMove/aspx/stockmove_frameset.aspx")
			If (PartMovePointer.document.GetElementById("MainTitle").innertext = "This page can't be displayed") {
				PartMovePointer.refresh()
				sleep, 2500
			}
			Frame:=PartMovePointer.document.all(9).contentwindow
			Frame.document.GetElementById("cboPartNum").value := Part
		}
	}Until (Frame.document.GetElementById("cboPartNum").value = Part)

frame := PartMovePointer.document.all(9).contentWindow

frame.document.getelementbyID("cboPartNum").value := Part

iniread,AltlocationList,%PartsDatabase%,STOKGOODS,List
IfInString,AltLocationList,%Part% 
{
MovementSite := "STOKGOODS"

}
else
{
MovementSite := "STOWPARTS"
}
frame.document.getelementbyID("cboSourceSiteNum").value := MovementSite
OutputDebug, [COK]Movement site is set to %MovementSite%
ModalDialogue() 
frame.document.getElementsByTagName("IMG")[2] .click 
WinWaitClose, ahk_class Internet Explorer_TridentDlgFrame
sleep, 500

loop {
	sleep, 250
} until (frame.document.getelementbyID("txtSourceSerialised").value != "" )

if (frame.document.getelementbyID("txtSourceTotalQty").value != "" OR 0) {
	SourceQuantity := frame.document.getelementbyID("txtSourceTotalQty").value
	OutputDebug % "[COK] Quantity in stock " . SourceQuantity
	if (frame.document.getelementbyID("txtSourceTotalQty").value < Qunatity) {
		While (SourceQuantity < Qunatity) {
		SourceQuantity := frame.document.getelementbyID("txtSourceTotalQty").value
		OutputDebug % "[COK] Quantity in stock " SourceQuantity
		Gui +LastFound +OwnDialogs +AlwaysOnTop
		inputbox, QUANTITY, Insufficient Stock, There is insufficient stock in STOWPARTS`n maximum available is %SourceQuantity%`ninput a new amount
	}
	}
	IniRead,Engineer,%Config%,Engineer,Number
	
	frame.document.getelementbyID("cboDestSiteNum").value := Engineer
	ModalDialogue() 
	frame.document.getElementsByTagName("IMG")[6] .click 
	while (frame.document.getelementbyID("txtDestTotalNeed").value = "")
		sleep, 500
	frame.document.getelementbyID("txtMoveTotalQty").value := Quantity
	frame.document.getelementbyID("cboAdjustCode").value := "MV"
	frame.document.getelementbyID("txtReason").value := "Automated movement - TG"
	TP_Show("Stock movement from " .  MovementSite , "Blue", "12", "White", 2000)
	frame.document.getelementbyID("cboSourceSiteNum").value := MovementSite
	ModalDialogue() 
	frame.document.getElementsByTagName("IMG")[2] .click 
	sleep, 500
	frame.document.getelementbyid("chkAllowNewStockFlag").click
	frame := PartMovePointer.document.all(6).contentWindow
	frame.document.getElementByID("cmdSubmit").Click
	PageLoading(PartMovePointer)
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

;Manage Database


CheckProduct:
Gui, CheckProduct:add, text,, Insert Product Code
gui, CheckProduct:add, edit, vProduct,
gui, CheckProduct:add,button, gsubmitProductCheck, Submit
gui, CheckProduct:show
return



submitProductcheck:
Gui,CheckProduct:Submit
gui,ProductSelector:destroy
if CheckEligibleProducts(Product)
	msgbox % Product . " is eligible"
else
	msgbox % Product . " does not exist"
gui,CheckProduct:destroy
return

AddProduct:
Gui, addproduct:add, text,,select a product
Gui,  CheckProduct:add, edit, vproductcodeadd,
gui, addproduct:add, button,gsubmitProductadd,submit
gui, addproduct:show
return

submitProductadd:
gui, addproduct:submit
if CheckEligibleProducts(productcodeadd) {
	MsgBox, %product% already exists
	Gui,CheckProduct:destroy
	Gui,addproduct:destroy
	return
}


