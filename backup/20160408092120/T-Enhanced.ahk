#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%
Onexit,Endit
#singleinstance, force
#Persistent
RunAsAdmin() 
if not A_isadmin {
	OutputDebug, [T-Enhanced] Not running as administrator
	MsgBox, Unable to get privileges. Check your permissions
	ExitApp	
}

showSplashScreen()
#include Modules\Lib\Functions.ahk
#include Modules\Lib\Api.ahk
#include Modules\Lib\Rini.ahk
#include Modules\config.ahk
#include Modules/Logistics.ahk
#Include OOP.ahk
installFiles()
A:=true
B:=3
C:=false
InitializeDymo()
setConfigLocation("\Modules\Config.ini")
global TesseractVersion := ""
setTesseractVersion("5.40.14")
LaunchCustomScripts()
CreateTrayMenu()
destroySplashScreen()


/*
	###########
	Master GUI
	###########
	Launch the Main User Interface
*/


Gui, Master: Font, s8
Gui, Master: Add, Tab2, x0 y0 w265 h150 vTab gTabClick 0x108, Home|Engineer|Logistics
Gui, Master: Tab, Home
Gui, Master: Font, s10 Bold
Gui, Master: Add, text, x40 y30 w150 h50  center ,T-Enhanced `n[ZULU]
Gui, Master: Add, text, x40 y65 w150 h50  center ,By Kieran Wynne
Gui, Master: Add, picture, x195 y30 w50 h50,icon.png
Gui, Master: Font, s8 norm
Gui, Master: add, Button , x20 y85 w112 h20 gChangelog 0x8000, Changelog
Gui, Master: Add, Button, vConfigButton x20 y110 w225 h35  gConfig 0x8000, Configuration
Gui, Master: Add, Button, x5 y25 w35 h35 gEndit 0x8000, Quit
Gui, Master: Tab, Engineer
Gui, Master: Add, Button, x5 y30 w80 h45 gCreate vCreate 0x8000, Create Job
Gui, Master: Add, Button, x92 y85 w80 h45 gPanic 0x8000, Reload T-Enhanced
Gui, Master: Add, Button, x5 y85 w80 h45 gPrintFunction vPrint 0x8000, Print Labels
Gui, Master: Add, Button, x180 y30 w80 h45 gShip vShipOut 0x8000, Ship Current Job
Gui, Master: Add, Button, x92 y30 w80 h45 gReport vReport 0x8000, Service Report
Gui, Master: Add, Button, x180 y85 w80 h45 gLetsMoveSomeShit 0x8000, Move Parts
Gui, Master: Tab, Logistics
Gui, Master: Add, Button, x5 y30 w80 h45 gAssets vAssets 0x8000, Book In
Gui, Master: Add, Button, x92 y30 w80 h45 gLogShipout 0x8000, Ship Out
Gui, Master: Add, Button, x180 y30 w80 h45 0x8000 gImacLines, Add Lines
Gui, Master: +AlwaysOnTop +ToolWindow +OwnDialogs -DPIScale 
X:=GetWinPosX("T-Enhanced Master Window")
Y:=GetWinPosY("T-Enhanced Master Window")
if (X = "ERROR" || X= "" OR Y = "ERROR" || Y=""){
	Gui, Master: Show, ,T-Enhanced Master Window
} else {
	Gui, Master: Show, X%x% Y%y%  ,T-Enhanced Master Window
}
if (settings.Engineer = "ERROR" or settings.Engineer = "" Or settings.WorkshopSite= "Error" or settings.WorkshopSite= ""){
	OutputDebug,[T-Enhanced]  Failed to find settings
	gosub, config
}
OutputDebug,[T-Enhanced]  Master Gui loaded
WinGet,MasterWindow,ID,T-Enhanced Master Window
return

MasterGuiClose:
MasterGuiEscape:
SaveWinPos("T-Enhanced Master Window")
gosub,Hide
return

/*
	###########
	Simple Quit
	###########
	Quits the master Gui
*/
Endit:
try
	pwb := ""
SaveWinPos("T-Enhanced Master Window")
OutputDebug,[T-Enhanced]  Force quit
Exitapp
return


/*
	###########
	Configuration Interface
	###########
	Opens up the Configuration Interface
*/
config:
Gui, Master: Tab, Home
GuiControl,Master:, Tab, |Home
Gui, Master: Add, text, ym xm+267 center, Insert Engineer Number
Gui, Master: Add, Edit, xm+267 yp+17 vEng1,
Gui, Master:Add, Text, xm+267 yp+20, Workshop Site?
Gui, Master:add, DDL, xm+267 yp+17 vMySite, NSC|Cumbernauld
Gui, Master:Add, Text, xm+267 yp+20 BackgroundTrans, Username
Gui, Master:Add, Edit, xm+267 yp+17 vUserNameIn,
Gui, Master:Add, Text, xm+267 yp+20, Password
Gui, Master:Add, Edit, xm+267 yp+17 vPasswordIn Password,
GuiControl, Master:hide, configButton
Gui, Master: Add, Button, x20 y110 w225 h35  gDone 0x8000, Submit
Gui, Master: show, autosize
return

Done:
gui,Master:submit,nohide
settings.save(Eng1,mySite,UserNameIn,PasswordIn)
reload
return

Changelog:
run, https://github.com/k33k00/T-Enhanced--ZULU-/commits/master
return

Create:
Create:= new TEnhanced.Create(settings)
Create:= ""
return

Report:
#Include Modules/ServicePlease.ahk
return

Ship:
Shipout:= new TEnhanced.Shipout(settings)
ShipOut:= ""
return

Panic:
SaveWinPos("T-Enhanced Master Window")
OutputDebug,[T-Enhanced]  Force Reload
Reload
return

Hide:
SaveWinPos("T-Enhanced Master Window")
Gui,Master: Hide
return

Show:
Gui, Master:Show
return



PrintFunction:
#include Modules/ManualPrint.ahk
return


;{ ----On tab click
TabClick:
if (settings.Engineer = "" OR settings.Engineer = "Error") {
	GuiControl,Master:, Tab, |Home
	return
}
gui,Master:show, autosize

return
;}

;{ ----AutoLogin
MasterGuiContextMenu:
gui,Master:submit, noHide
Login := new TEnhanced.AutoLogin(settings,tab)
return
;}



LetsMoveSomeShit:
#Include Modules/Move.ahk
return

Assets:
Bookin := new Logistics.Bookin()
Bookin:= ""
return

LogShipout:
BookOut := new Logistics.BookOut()
BookOut := ""
return

ImacLines:
#Include Modules\IMAC.ahk
return

#if settings.Engineer = "406"
#include Modules\406.ahk



showSplashScreen(){
	gui, splash: add, picture, w128 h-1 BackgroundTrans,icon.png
	Gui, splash: Font, s12
	gui, splash: add, Text,w128 center cblue BackgroundTrans, T-Enhanced
	Gui, splash: Font, s10
	gui, splash: add, Text,w128 center cblue BackgroundTrans, Created by`nKieran Wynne
	gui, splash: -border -caption +alwaysontop +LastFound +ToolWindow
	Gui, splash:Color, EEAA99
	WinSet, TransColor, EEAA99
	gui, splash:show, autosize, T-Enhanced
}
destroySplashScreen(){
	gui, splash:destroy
}
installFiles(){
	
	FileInstall, InstallMe/icon.png,icon.png, 1
	FileInstall, InstallMe/BerForm.docx,Modules/BerForm.docx,1
	FileInstall, InstallMe/Part Order.label,Modules/Part Order.label,1
	FileInstall, InstallMe/Workshop Codes Large.label,Modules/Workshop Codes Large.label,1
	FileInstall, InstallMe/Workshop Codes.label,Modules/Workshop Codes.label,1
	FileInstall, InstallMe/Zulu-book-in.label,Modules/Zulu-book-in.label,1
	FileCreateDir, Modules/Lib
	FileInstall, InstallMe/AutoHotkeyMini.dll,Modules/Lib/AutoHotkeyMini.dll,1
}
InitializeDymo(){
	try {
		Global DymoLabel := ComObjCreate("DYMO.DymoLabels")
		OutputDebug, [T-Enhanced] Enabled DYMO.DymoLabels 1/3 
		global DymoAddIn := ComObjCreate("DYMO.DymoAddIn")
		OutputDebug, [T-Enhanced] Enabled DYMO.DymoAddIn 2/3
		global DymoEngine := ComObjCreate("DYMO.LabelEngine")
		OutputDebug, [T-Enhanced] Enabled DYMO.LabelEngine 3/3
	} catch {
		msgbox, Unable to activate Dymo.`nPlease check you have the latest Dymo software installed.
	}
}
setConfigLocation(path){
	global Config:=A_ScriptDir . path
}
setTesseractVersion(version){
	TesseractVersion:= version
}
LaunchCustomScripts(){
	Loop %A_ScriptDir%\Custom Scripts\*.ahk
		Run %A_LoopFileFullPath%
}
CreateTrayMenu(){
	if (A_IsCompiled){
		Menu,tray,Nostandard
	}
	Menu, Home, add, Changelog,Changelog
	Menu, Workshop, add,Create Job,Create
	Menu, Workshop, add,Service Report,Report
	Menu, Workshop, add,Ship Out,Ship
	Menu, Workshop, add,Print,PrintFunction
	Menu, tray, add, Home, :Home
	Menu, tray, add, Workshop, :Workshop
	Menu,Tray,add,Show
	Menu,Tray,add,Reload,panic
	Menu,Tray,add,Quit,Endit
}