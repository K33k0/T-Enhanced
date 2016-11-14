/*
	Owner - Kieran Wynne
	Module - Functions for Tesseract
	Version - 1.0
	
	----- Info -----
	The new home to all my functions
*/
AhkDllPath := A_ScriptDir . "/Modules/Lib/AutoHotkeyMini.dll"
hModule := DllCall("LoadLibrary","Str",AhkDllPath)


SaveWinPos(title=""){
	WinGetPos, Xpos, Ypos,,,%Title%
	OutputDebug, %Title% - Y position = %Ypos% : X position = %Xpos%
	IniWrite, %Xpos%, % Settings.iniPath,%Title%,Xpos
	IniWrite, %Ypos%, % Settings.iniPath,%Title%,Ypos
	return
}

GetWinPosX(title=""){
	IniRead, PosX,% Settings.iniPath,%Title%,Xpos
	Return PosX
}

GetWinPosY(title=""){
	IniRead, PosY, % Settings.iniPath,%Title%,Ypos
	Return PosY
}

GetProductCode() {
	try{
		If Pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " settings.Tesseract " / Page Ver: " settings.Tesseract) {
			frame := Pwb.document.all(10).contentWindow
			If ProductCode:=frame.document.getElementById("cboJobPartNum").value ;Job Query
				return ProductCode
			If ProductCode:=frame.document.getElementById("txtStockPartNo").value ;Serialized Stock Query
				return ProductCode
			If ProductCode:=frame.document.getElementById("cboSerProdNum").value  ;Serialized Customer Assets
				return ProductCode	
			
			frame := Pwb.document.all(11).contentWindow
			If ProductCode:=frame.document.getElementById("cboFSRPartNum").value ;Service Report
				return ProductCode
			
			frame := Pwb.document.all(9).contentWindow
			If ProductCode:=frame.document.getElementById("cboPartNum").value  ;Stock Movement
				return ProductCode	
			
			If Pwb := IETitle("Repair Shipping Wizard - " settings.Tesseract) { ;Job Shipout
				If ProductCode:=Pwb.document.getElementById("cbaListCallSerNumLineArray").value 
					return ProductCode
			}
			
		}
	}
}

IETitle(name="")  {
	for WB in ComObjCreate("Shell.Application").Windows
		if (wb.LocationName ~= Name) and InStr(wb.FullName, "iexplore.exe")
			return wb
}

PageLoading(Handle=""){
	While (Handle.document.Readystate != "Complete") {
		While (Handle.busy) {
			Sleep 100
		}
	}
	
	return
}

IEGetUrl(url="") {
	IfEqual, Name,, WinGetTitle, Name, ahk_class IEFrame
		Name := ( Name="New Tab - Windows Internet Explorer" ) ? "about:Tabs"
	: RegExReplace( Name, " - (Windows|Microsoft) Internet Explorer" )
	For Pwb in ComObjCreate( "Shell.Application" ).Windows
		If ( Pwb.Locationurl = url ) && InStr( Pwb.FullName, "iexplore.exe" )
			Return Pwb
}

ModalDialogue() {
	global AhkDllPath
	OutputDebug % AhkDllPath
	global hModule
	OutputDebug % hModule
	
	Text = 
	(
#SingleInstance, Force
OutputDebug, [TZ] ------ Begin Modal Dialogue -------
Critical
Running := 1
Loop {
	IfWinExist,Popup List -- Webpage Dialog
	{
		OutputDebug, [TZ]Modal Window Found
		while not Pdoc
			pdoc := Get_MODAL_DOCUMENT()
		while (pdoc.getElementsByTagName("TD")[0].Innertext = "")
			sleep, 10
		if (pdoc.getElementByID("grdDropdown:_ctl4:_ctl0").Innertext = "")
		{
			pdoc.getElementByID("grdDropdown:_ctl3:_ctl0").click
			OutputDebug, [TZ] AutoConfirmed Modal Dialog
		} else {
			OutputDebug,[TZ] Multiple Records in Modal
		}
		OutputDebug, [TZ] -------Exited ModalDialogue-------
		exitapp
	}
}
return

Get_MODAL_DOCUMENT()
{
	static msg := DllCall("RegisterWindowMessage", "str", "WM_HTML_GETOBJECT")
	SendMessage msg, 0, 0, Internet Explorer_Server1, ahk_class Internet Explorer_TridentDlgFrame
	if ErrorLevel = FAIL
		return
	lResult := ErrorLevel
	DllCall("oleacc\ObjectFromLresult", "ptr", lResult
	, "ptr", GUID(IID_IHTMLDocument2,"{332C4425-26CB-11D0-B483-00C04FD90119}")
	, "ptr", 0, "ptr*", pdoc)
	static VT_DISPATCH := 9, F_OWNVALUE := 1
	return ComObject(VT_DISPATCH, pdoc, F_OWNVALUE)
}

GUID(ByRef GUID, sGUID)
{
	VarSetCapacity(GUID, 16, 0)
	return DllCall("ole32\CLSIDFromString", "wstr", sGUID, "ptr", &GUID) >= 0 ? &GUID : ""
}
	
	)
	
	;~ FileAppend,%TEXT%,TempThread.ahk
	
	
	
	;~ AhkDllPath := A_ScriptDir "\AutoHotkey.dll"
	;~ hModule := DllCall("LoadLibrary","Str",AhkDllPath)
	;~ Sleep 500
	DllCall(AhkDllPath "\ahktextdll","Str",TEXT,"Str","","Str","","Cdecl UPTR")
	
	;~ MsgBox, End main thread
	;~ DllCall("FreeLibrary","PTR",hModule)
	
	return
}

PageAlert() {
	global AhkDllPath
	OutputDebug % AhkDllPath
	global hModule
	OutputDebug % hModule
	Text = 
(
#NoTrayIcon
#SingleInstance, Force
#Persistent
Loop {
	IfWinExist,ahk_class #32770
	{
		OutputDebug, [TE] Well the window exists...
		WinClose, ahk_class #32770
		OutputDebug, [TE] And The window is closed
		ExitApp
	}
}
return
)
	
	DllCall(AhkDllPath "\ahktextdll","Str",Text,"Str","","Str","","Cdecl UPTR")
	return true
}

IEGet(Name="")
{
	IfEqual, Name,, WinGetTitle, Name, ahk_class IEFrame
		Name := ( Name="New Tab - Windows Internet Explorer" ) ? "about:Tabs"
	: RegExReplace( Name, " - (Windows|Microsoft) Internet Explorer" )
	For Pwb in ComObjCreate( "Shell.Application" ).Windows
		If ( Pwb.LocationName = Name ) && InStr( Pwb.FullName, "iexplore.exe" )
			Return Pwb
}

IELoad(Pwb)
{
	If !Pwb
		Return False
	Loop
		Sleep,100
	Until (Pwb.busy)
	Loop
		Sleep,100
	Until (!Pwb.busy)
	Loop
		Sleep,100
	Until (Pwb.document.Readystate = "Complete")
	Return True
}

IELoad1(Pwb)
{
	If !Pwb
		Return False
	try{
		Loop, 50
			Sleep,100
		Until (Pwb.document.Readystate != "Complete")
	}
	Try{
		Loop
			Sleep,100
		Until (Pwb.document.Readystate = "Complete")
	}
	Return True
}

IEvGet(Name="")
{
	IfEqual, Name,, WinGetTitle, Name, ahk_class IEFrame
		Name := ( Name="New Tab - Windows Internet Explorer" ) ? "about:Tabs"
	: RegExReplace( Name, " - (Windows|Microsoft) Internet Explorer.*$" )
	For Pwb in ComObjCreate( "Shell.Application" ).Windows
		If ( Pwb.LocationName = Name ) && InStr( Pwb.FullName, "iexplore.exe" )
			Return Pwb
}

;{ ----IniRead()
IniRead(_IniFile="", _Options="")
{
	Local _Reading, _Prepend, _Entries, _nSec, _nKey, _@, _@1, _@2, _@3, _@4, _@5 := 1, _Literal := """", _Commands := "sa|ka|sl|sr|p|d|r|e|t|c|b|f"
	, _sa := "",              _sa_user := "Sections*"
	, _ka := "",              _ka_user := "Keys*"
	, _sl := "",              _sl_user := "*"
	, _sr := "",              _sr_user := ""
	, _p := "",               _p_user := "*_"
	, _d := "",               _d_user := "|"
	, _r := "",               _r_user := "_"
	, _e := "fso",            _e_user := ""
	, _t := True,             _t_user := False
	, _c := True,             _c_user := False
	, _b := True,             _b_user := False
	, _f := True,             _f_user := False
	, _UserConfig_Foo := "x12 y34"
	, _UserConfig_Bar := "cWhite -a"
	While (_@5 := RegExMatch(_Options, "i)(?:^|\s)(?:!(\w+)|(\+|-)?(" _Commands ")(" _Literal "(?:[^" _Literal "]|" _Literal _Literal ")*" _Literal "(?= |$)|[^ ]*))", _@, _@5 + StrLen(_@)))
		If (_@1 <> "")
			_Options := SubStr(_Options, 1, _@5 + StrLen(_@)) _UserConfig_%_@1% SubStr(_Options, _@5 + StrLen(_@))
	Else If (_@4 <> "") {
		If (InStr(_@4, _Literal) = 1) and (_@4 <> _Literal) and (SubStr(_@4, 0, 1) = _Literal) and (_@4 := SubStr(_@4, 2, -1))
			StringReplace, _@4, _@4, %_Literal%%_Literal%, %_Literal%, All
		_%_@3% := _@4
	} Else
		_%_@3% := _@2 = "+" ? True : _@2 = "-" ? False : _%_@3%_user
	If (_IniFile = "") {
		If !FileExist(_IniFile := SubStr(A_ScriptFullPath, 1, InStr(A_ScriptFullPath, ".", 0, 0)) "ini") {
			If InStr(_e, "f")
				MsgBox, 262160, %A_ScriptName% - %A_ThisFunc%(): Error, The IniFile parameter was omitted or blank, which the function interprets as an ini file with the same name as the script and in the same dir, i.e.:`n`n%_IniFile%`n`nThis file does not exist.
			Return
		}
	} Else If (IniFile = "*") {
		Loop, *.ini
		{
			_IniFile := A_LoopFileFullPath
			Break
		}
		If (_IniFile = "*") {
			If InStr(_e, "f")
				MsgBox, 262160, %A_ScriptName% - %A_ThisFunc%(): Error, No .ini file found in working directory.`n`not avoid this error, specify an explicit .ini file path in the first parameter of the function.
			Return
		}
	} Else If !FileExist(_IniFile) {
		If InStr(_e, "f")
			MsgBox, 262160, %A_ScriptName% - %A_ThisFunc%(): Error, File "%_IniFile%" not found or does not exist.
		Return
	}
	If RegExMatch(_r, "[^\w#@$?]") or RegExMatch(_p, "[^\w*#@$?]") {
		MsgBox, 262160, %A_ScriptName% - %A_ThisFunc%(): Error, Neither the p nor r options may contain characters that are not alloewd in AutoHotkey variable names.
		Return
	}
	_Entries := _d = "" ? 0 : _d
	If !InStr(_p, "*")
		_Prepend := _p
	If (_sl <> "") {
		If (_sr <> "") {
			MsgBox, 262160, %A_ScriptName% - %A_ThisFunc%(): Error, Please enter either a sl (Section - Literal) or sr (Section - RegEx) value, not both.
			Return
		}
	} Else If (_sr = "")
		_Reading := True
	If (_sa <> "") {
		If RegExMatch(_sa, "[^\w#@$?*]")
			_sa := RegExReplace(_sa, "[^\w#@$?*]")
		If !InStr(_sa, "*")
			_sa .= "*"
	}
	If (_ka <> "") {
		If RegExMatch(_ka, "[^\w#@$?*]")
			_ka := RegExReplace(_ka, "[^\w#@$?*]")
		If !InStr(_ka, "*")
			_ka .= "*"
		If (_ka = _sa) {
			MsgBox, 262160, %A_ScriptName% - %A_ThisFunc%(): Error, The sa (Section Output Array) and ka (Key Output Array) options cannot be the same.
			Return
		}
	}
	Loop, Read, %_IniFile%
		If RegExMatch(A_LoopReadLine, "^[^=]*\[\K[^\]]+(?=\])", _@) {
			If _t
				_@ = %_@%
			If (_sr <> "")
				_Reading := RegExMatch(_@, _sr) ? True : False
			Else If (_sl <> "")
				If _Reading
					Break
			Else If (_@ = _sl) or (_sl = "*")
				_Reading := True
			If !_Reading
				Continue
			If InStr(_p, "*") {
				StringReplace, _Prepend, _p, *, % RegExReplace(_@, "[^\w#_t$?]+", _r, _@2), All
				If _@2 and InStr(_e, "r") {
					MsgBox, 262420, %A_ScriptName% - %A_ThisFunc%(): Error, The section "%_@%" contains characters not allowed in AutoHotkey variable names. Replace these characters with "%_r%"?`n`nTo change the replacement character, use the r option.
					IfMsgBox NO
						Return
				}
			}
			If _sa {
				_nSec += 1
				StringReplace, _@1, _sa, *, %_nSec%, All
				%_@1% := _@
			}
		} Else If _Reading and InStr(A_LoopReadLine, "=") {
		_@ := SubStr(A_LoopReadLine, 1, InStr(A_LoopReadLine, "=") - 1), _@2 := SubStr(A_LoopReadLine, InStr(A_LoopReadLine, "=") + 1)
		If _t {
			_@ = %_@%
			_@2 = %_@2%
		}
		If _c
			_@2 := RegExReplace(_@2, "(?:\s+|^);.*")
		_@1 := RegExReplace(_@, "[^\w#_t$?]+", _r)
		If (_@1 <> _@) and InStr(_e, "r") {
			MsgBox, 262420, %A_ScriptName% - %A_ThisFunc%(): Error, The key name "%_@%" contains characters not allowed in AutoHotkey variable names. Replace these characters with "%_r%"?`n`nTo change the replacement character, use the r option.
			IfMsgBox NO
				Return
		}
		If (%_Prepend%%_@1% <> "") and InStr(_e, "o") {
			MsgBox, 262420, %A_ScriptName% - %A_ThisFunc%(): Error, The variable "%_Prepend%%_@1%" has already been assigned, either by %A_ThisFunc%() or elsewhere in the script. Overwrite it with "%_@2%" (the value from the .ini file)?`n`nTo avoid this error, try using the p or s options to make output variable names more unique.
			IfMsgBox NO
				Return
		}
		If _f
			Transform, _@2, Deref, %_@2%
		%_Prepend%%_@1% := !_b ? _@2 : _@2 = "True" or _@2 = "Yes" ? True : _@2 = "False" or _@2 = "No" ? False : _@2
		If (_d <> "") {
			StringReplace, _Entries, _Entries, %_d%%_Prepend%%_@1%%_d%, %_d%, All
			_Entries .= _Prepend _@1 _d
		} Else
			_Entries += 1
		If _ka {
			_nKey += 1
			StringReplace, _@, _ka, *, %_nKey%, All
			%_@% := _Prepend _@1
		}
	}
	If (_sl <> "") and !_Reading and InStr(_e, "s")
		MsgBox, 262160, %A_ScriptName% - %A_ThisFunc%(): Error, Section "%_sl%" was not found in ini file "%_IniFile%", therefore no variables were assigned.`n`nTo avoid this error, use the sr (Section Name - RegEx) option instead of sl (Section Name - Literal), or omit both options.
	If _sa {
		StringReplace, _@, _sa, *, 0, All
		%_@% := _nSec
	}
	If _ka {
		StringReplace, _@, _ka, *, 0, All
		%_@% := _nKey
	}
	Return _d = "" ? _Entries : SubStr(_Entries, 2, -1)
}
;}

;{ ----AniGif
AniGif_CreateControl(_guiHwnd, _x, _y, _w, _h, _style="")
{
	local hAniGif, agHwnd
	local msg, style
	static $bFirstCall := true
	If ($bFirstCall)
	{
		$bFirstCall := false
		hAniGif := DllCall("LoadLibrary", "Str","Modules\lib\AniGif.dll")
	}
	style := 0
	If (_style != "")
	{
		If _style contains autosize
			style |= 1
		If _style contains center
			style |= 2
		If _style contains hyperlink
			style |= 4
	}
	style := 0x50000000 | style
	agHwnd := DLLCall("CreateWindowEx"
	, "UInt", 0
	, "Str", "AniGIF"
	, "Str", "AnimatedGif"
	, "UInt",  style
	, "Int", _x
	, "Int", _y
	, "Int", _w
	, "Int", _h
	, "UInt", _guiHwnd
	, "UInt", 0
	, "UInt", 0
	, "UInt", 0)
	If (ErrorLevel != 0 or agHwnd = 0)
	{
		msg = %msg% Cannot create AniGif control (%ErrorLevel%/%A_LastError%)
		Gosub AniGif_CreateControl_CleanUp
		Return msg
	}
	Return agHwnd
	AniGif_CreateControl_CleanUp:
	Return
}
AniGif_DestroyControl(_agHwnd)
{
	If (_agHwnd != 0)
	{
		AniGif_UnloadGif(_agHwnd)
		DllCall("DestroyWindow", "UInt", _agHwnd)
	}
}
AniGif_LoadGifFromFile(_agHwnd, _gifFile)
{
	VarSetCapacity(var, StrPut(_gifFile, "cp0") * 2)
	StrPut(_gifFile, &var, "cp0")
	SendMessage, 2024, 0, &var, , ahk_id %_agHwnd%
}
AniGif_UnloadGif(_agHwnd)
{
	SendMessage 2026, 0, 0, , ahk_id %_agHwnd%
}
AniGif_SetHyperlink(_agHwnd, _url)
{
	SendMessage 2027, 0, &_url, , ahk_id %_agHwnd%
}
AniGif_Zoom(_agHwnd, _bZoomIn)
{
	PostMessage 2028, 0, _bZoomIn, , ahk_id %_agHwnd%
}
AniGif_SetBkColor(_agHwnd, _backColor)
{
	PostMessage 2029, 0, _backColor, , ahk_id %_agHwnd%
}
;}

/*
	
	---PartsDB--- 
	
*/

CheckEligibleProducts(Product="") {
	try {
		iniread,Eligible,\\pratechvf\TECH_CTS_Branch_Gen\Tesseract Zoanthropy\PartsDatabase.ini,EligibleProducts,List
		IfInString,Eligible, %Product%
			Return True
		else
			return false
	}
}

RunAsAdmin() {
	Loop, %0%
	{
		param := %A_Index%
		params .= A_Space . param
	}
	ShellExecute := A_IsUnicode ? "shell32\ShellExecute":"shell32\ShellExecuteA"
	if not A_IsAdmin
	{
		If A_IsCompiled
			DllCall(ShellExecute, uint, 0, str, "RunAs", str, A_ScriptFullPath, str, params , str, A_WorkingDir, int, 1)
		Else
			DllCall(ShellExecute, uint, 0, str, "RunAs", str, A_AhkPath, str, """" . A_ScriptFullPath . """" . A_Space . params, str, A_WorkingDir, int, 1)
		ExitApp
	}
}

/*
	toaster popups 
*/

TP_Show(TP_Message="Hello, World", TP_FontColor="Blue", TP_FontSize="12", TP_BGColor="White", TP_Lifespan=0)
{
	Global TP_GUI_ID 
	DetectHiddenWindows, On
	SysGet, Workspace, MonitorWorkArea
	Gui, 89:-Caption +ToolWindow +LastFound +AlwaysOnTop +Border
	Gui, 89:Color, %TP_BGColor%
	Gui, 89:Font, s%TP_FontSize% c%TP_FontColor%
	Gui, 89:Add, Text, , %TP_Message%
	Gui, 89:Show, Hide
	TP_GUI_ID := WinExist()
	WinGetPos, GUIX, GUIY, GUIWidth, GUIHeight, ahk_id %TP_GUI_ID%
	NewX := WorkSpaceRight-GUIWidth-5
	NewY := WorkspaceBottom-GUIHeight-5
	Gui, 89:Show, Hide x%NewX% y%NewY%
	
	DllCall("AnimateWindow","UInt",TP_GUI_ID,"Int",500,"UInt","0x00040008") ; TOAST!
	DllCall("AnimateWindow","UInt",TP_GUI_ID,"Int",2500,"UInt","0x90000") ; Fade out when clicked
	Gui, 89:Destroy
	Return
}

;{ ----TF
TF_CountLines(Text)
{
	TF_GetData(OW, Text, FileName)
	StringReplace, Text, Text, `n, `n, UseErrorLevel
	Return ErrorLevel + 1
}
TF_ReadLines(Text, StartLine = 1, EndLine = 0, Trailing = 0)
{
	TF_GetData(OW, Text, FileName)
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
			OutPut .= A_LoopField "`n"
		Else if (A_Index => EndLine)
			Break
	}
	OW = 2
	Return TF_ReturnOutPut(OW, OutPut, FileName, Trailing)
}
TF_ReplaceInLines(Text, StartLine = 1, EndLine = 0, SearchText = "", ReplaceText = "")
{
	TF_GetData(OW, Text, FileName)
	IfNotInString, Text, %SearchText%
		Return Text
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
		{
			StringReplace, LoopField, A_LoopField, %SearchText%, %ReplaceText%, All
			OutPut .= LoopField "`n"
		}
		Else
			OutPut .= A_LoopField "`n"
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_Replace(Text, SearchText, ReplaceText="")
{
	TF_GetData(OW, Text, FileName)
	IfNotInString, Text, %SearchText%
		Return Text
	Loop
	{
		StringReplace, Text, Text, %SearchText%, %ReplaceText%, All
		if (ErrorLevel = 0)
			break
	}
	Return TF_ReturnOutPut(OW, Text, FileName, 0)
}
TF_RegExReplaceInLines(Text, StartLine = 1, EndLine = 0, NeedleRegEx = "", Replacement = "")
{
	TF_GetData(OW, Text, FileName)
	If (RegExMatch(Text, NeedleRegEx) < 1)
		Return Text
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
		{
			LoopField := RegExReplace(A_LoopField, NeedleRegEx, Replacement)
			OutPut .= LoopField "`n"
		}
		Else
			OutPut .= A_LoopField "`n"
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_RegExReplace(Text, NeedleRegEx = "", Replacement = "")
{
	TF_GetData(OW, Text, FileName)
	If (RegExMatch(Text, NeedleRegEx) < 1)
		Return Text
	Text := RegExReplace(Text, NeedleRegEx, Replacement)
	Return TF_ReturnOutPut(OW, Text, FileName, 0)
}
TF_RemoveLines(Text, StartLine = 1, EndLine = 0)
{
	TF_GetData(OW, Text, FileName)
	If (StartLine < 0)
	{
		StartLine:=TF_CountLines(Text) + StartLine
		EndLine=0
	}
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
			Continue
		Else
			OutPut .= A_LoopField "`n"
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_RemoveBlankLines(Text, StartLine = 1, EndLine = 0)
{
	TF_GetData(OW, Text, FileName)
	If (RegExMatch(Text, "[\S]+?\r?\n?") < 1)
		Return Text
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
			OutPut .= (RegExMatch(A_LoopField,"[\S]+?\r?\n?")) ? A_LoopField "`n" :
		Else
			OutPut .= A_LoopField "`n"
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_RemoveDuplicateLines(Text, StartLine = 1, Endline = 0, Consecutive = 0, CaseSensitive = false)
{
	TF_GetData(OW, Text, FileName)
	If (StartLine = "")
		StartLine = 1
	If (Endline = 0 OR Endline = "")
		EndLine := TF_Count(Text, "`n") + 1
	Loop, Parse, Text, `n, `r
	{
		If (A_Index < StartLine)
			Section1 .= A_LoopField "`n"
		If A_Index between %StartLine% and %Endline%
		{
			If (Consecutive = 1)
			{
				If (A_LoopField <> PreviousLine)
					Section2 .= A_LoopField "`n"
				PreviousLine:=A_LoopField
			}
			Else
			{
				If !(InStr(SearchForSection2,"__bol__" . A_LoopField . "__eol__",CaseSensitive))
				{
					SearchForSection2 .= "__bol__" A_LoopField "__eol__"
					Section2 .= A_LoopField "`n"
				}
			}
		}
		If (A_Index > EndLine)
			Section3 .= A_LoopField "`n"
	}
	Output .= Section1 Section2 Section3
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_InsertLine(Text, StartLine = 1, Endline = 0, InsertText = "")
{
	TF_GetData(OW, Text, FileName)
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
			Output .= InsertText "`n" A_LoopField "`n"
		Else
			Output .= A_LoopField "`n"
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_ReplaceLine(Text, StartLine = 1, Endline = 0, ReplaceText = "")
{
	TF_GetData(OW, Text, FileName)
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
			Output .= ReplaceText "`n"
		Else
			Output .= A_LoopField "`n"
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_InsertPrefix(Text, StartLine = 1, EndLine = 0, InsertText = "")
{
	TF_GetData(OW, Text, FileName)
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
			OutPut .= InsertText A_LoopField "`n"
		Else
			OutPut .= A_LoopField "`n"
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_InsertSuffix(Text, StartLine = 1, EndLine = 0 , InsertText = "")
{
	TF_GetData(OW, Text, FileName)
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
			OutPut .= A_LoopField InsertText "`n"
		Else
			OutPut .= A_LoopField "`n"
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_TrimLeft(Text, StartLine = 1, EndLine = 0, Count = 1)
{
	TF_GetData(OW, Text, FileName)
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
		{
			StringTrimLeft, StrOutPut, A_LoopField, %Count%
			OutPut .= StrOutPut "`n"
		}
		Else
			OutPut .= A_LoopField "`n"
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_TrimRight(Text, StartLine = 1, EndLine = 0, Count = 1)
{
	TF_GetData(OW, Text, FileName)
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
		{
			StringTrimRight, StrOutPut, A_LoopField, %Count%
			OutPut .= StrOutPut "`n"
		}
		Else
			OutPut .= A_LoopField "`n"
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_AlignLeft(Text, StartLine = 1, EndLine = 0, Columns = 80, Padding = 0)
{
	Trim:=A_AutoTrim
	AutoTrim, On
	TF_GetData(OW, Text, FileName)
	If (Endline = 0 OR Endline = "")
		EndLine := TF_Count(Text, "`n") + 1
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
		{
			LoopField = %A_LoopField%
			SpaceNum := Columns-StrLen(LoopField)-1
			If (SpaceNum > 0) and (Padding = 1)
			{
				Left:=TF_SetWidth(LoopField,Columns, 0)
				OutPut .= Left "`n"
			}
			Else
				OutPut .= LoopField "`n"
		}
		Else
			OutPut .= A_LoopField "`n"
	}
	AutoTrim, %Trim%
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_AlignCenter(Text, StartLine = 1, EndLine = 0, Columns = 80, Padding = 0)
{
	Trim:=A_AutoTrim
	AutoTrim, On
	TF_GetData(OW, Text, FileName)
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
		{
			LoopField = %A_LoopField%
			SpaceNum := (Columns-StrLen(LoopField)-1)/2
			If (Padding = 1) and (LoopField = "")
			{
				OutPut .= "`n"
				Continue
			}
			If (StrLen(LoopField) >= Columns)
			{
				OutPut .= LoopField "`n"
				Continue
			}
			Centered:=TF_SetWidth(LoopField,Columns, 1)
			OutPut .= Centered "`n"
		}
		Else
			OutPut .= A_LoopField "`n"
	}
	AutoTrim, %Trim%
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_AlignRight(Text, StartLine = 1, EndLine = 0, Columns = 80, Skip = 0)
{
	Trim:=A_AutoTrim
	AutoTrim, On
	TF_GetData(OW, Text, FileName)
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
		{
			LoopField = %A_LoopField%
			If (Skip = 1) and (LoopField = "")
			{
				OutPut .= "`n"
				Continue
			}
			If (StrLen(LoopField) >= Columns)
			{
				OutPut .= LoopField "`n"
				Continue
			}
			Right:=TF_SetWidth(LoopField,Columns, 2)
			OutPut .= Right "`n"
		}
		Else
			OutPut .= A_LoopField "`n"
	}
	AutoTrim, %Trim%
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_ConCat(FirstTextFile, SecondTextFile, OutputFile = "", Blanks = 0, FirstPadMargin = 0, SecondPadMargin = 0)
{
	If (Blanks > 0)
		Loop, %Blanks%
			InsertBlanks .= A_Space
	If (FirstPadMargin > 0)
		Loop, %FirstPadMargin%
			PaddingFile1 .= A_Space
	If (SecondPadMargin > 0)
		Loop, %SecondPadMargin%
			PaddingFile2 .= A_Space
	Text:=FirstTextFile
	TF_GetData(OW, Text, FileName)
	StringSplit, Str1Lines, Text, `n, `r
	Text:=SecondTextFile
	TF_GetData(OW, Text, FileName)
	StringSplit, Str2Lines, Text, `n, `r
	Text=
	If (Str1Lines0 > Str2Lines0)
		MaxLoop:=Str1Lines0
	Else
		MaxLoop:=Str2Lines0
	Loop, %MaxLoop%
	{
		Section1:=Str1Lines%A_Index%
		Section2:=Str2Lines%A_Index%
		OutPut .=  Section1 PaddingFile1 InsertBlanks Section2 PaddingFile2 "`n"
		Section1=
		Section2=
	}
	OW=1
	If (OutPutFile = "")
		OW=2
	Return TF_ReturnOutPut(OW, OutPut, OutputFile, 1, 1)
}
TF_LineNumber(Text, Leading = 0, Restart = 0, Char = 0)
{
	global t
	TF_GetData(OW, Text, FileName)
	Lines:=TF_Count(Text, "`n") + 1
	Padding:=StrLen(Lines)
	If (Leading = 0) and (Char = 0)
		Char := A_Space
	Loop, %Padding%
		PadLines .= Char
	Loop, Parse, Text, `n
	{
		If Restart = 0
			MaxNo = %A_Index%
		Else
		{
			MaxNo++
			If MaxNo > %Restart%
				MaxNo = 1
		}
		LineNumber:= MaxNo
		If (Leading = 1)
		{
			LineNumber := Padlines LineNumber
			StringRight, LineNumber, LineNumber, StrLen(Lines)
		}
		If (Leading = 0)
		{
			LineNumber := LineNumber Padlines
			StringLeft, LineNumber, LineNumber, StrLen(Lines)
		}
		OutPut .= LineNumber A_Space A_LoopField "`n"
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_ColGet(Text, StartLine = 1, EndLine = 0, StartColumn = 1, EndColumn = 1, Skip = 0)
{
	TF_GetData(OW, Text, FileName)
	EndColumn:=(EndColumn+1)-StartColumn
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
		{
			StringMid, Section, A_LoopField, StartColumn, EndColumn
			If (Skip = 1) and (StrLen(A_LoopField) < StartColumn)
				Continue
			OutPut .= Section "`n"
		}
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_ColPut(Text, Startline = 1, EndLine = 0, StartColumn = 1, InsertText = "", Skip = 0)
{
	TF_GetData(OW, Text, FileName)
	StartColumn--
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
		{
			StringLeft, Section1, A_LoopField, StartColumn
			StringMid, Section2, A_LoopField, StartColumn+1
			If (Skip = 1) and (StrLen(A_LoopField) < StartColumn)
				OutPut .= Section1 Section2 "`n"
			Else
				OutPut .= Section1 InsertText Section2 "`n"
		}
		Else
			OutPut .= A_LoopField "`n"
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_ColCut(Text, StartLine = 1, EndLine = 0, StartColumn = 1, EndColumn = 1)
{
	StartColumn--
	EndColumn++
	TF_GetData(OW, Text, FileName)
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
		{
			StringLeft, Section1, A_LoopField, StartColumn
			StringMid, Section2, A_LoopField, EndColumn
			OutPut .= Section1 Section2 "`n"
		}
		Else
			OutPut .= A_LoopField "`n"
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_ReverseLines(Text, StartLine = 1, EndLine = 0)
{
	TF_GetData(OW, Text, FileName)
	StringSplit, Line, Text, `n, `r
	If (EndLine = 0 OR EndLine = "")
		EndLine:=Line0
	If (EndLine > Line0)
		EndLine:=Line0
	CountDown:=EndLine+1
	Loop, Parse, Text, `n, `r
	{
		If (A_Index < StartLine)
			Output1 .= A_LoopField "`n"
		If A_Index between %StartLine% and %Endline%
		{
			CountDown--
			Output2 .= Line%CountDown% "`n" section2
		}
		If (A_Index > EndLine)
			Output3 .= A_LoopField "`n"
	}
	OutPut.= Output1 Output2 Output3
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_SplitFileByLines(Text, SplitAt, Prefix = "file", Extension = "txt", InFile = 1)
{
	LineCounter=1
	FileCounter=1
	Where:=SplitAt
	Method=1
	TF_GetData(OW, Text, FileName)
	IfInString, SplitAt, `-
	{
		StringSplit, Split, SplitAt, `-
		Part=1
		Where:=Split%Part%
		Method=2
	}
	IfInString, SplitAt, `,
	{
		StringSplit, Split, SplitAt, `,
		Part=1
		Where:=Split%Part%
		Method=3
	}
	Loop, Parse, Text, `n, `r
	{
		OutPut .= A_LoopField "`n"
		If (LineCounter = Where)
		{
			If (InFile = 0)
			{
				StringReplace, CheckOutput, PreviousOutput, `n, , All
				StringReplace, CheckOutput, CheckOutput, `r, , All
				If (CheckOutput <> "") and (OW <> 2)
					TF_ReturnOutPut(1, PreviousOutput, Prefix FileCounter "." Extension, 0, 1)
				If (CheckOutput <> "") and (OW = 2)
					TF_SetGlobal(Prefix FileCounter,PreviousOutput)
				Output:=
			}
			If (InFile = 1)
			{
				StringReplace, CheckOutput, Output, `n, , All
				StringReplace, CheckOutput, CheckOutput, `r, , All
				If (CheckOutput <> "") and (OW <> 2)
					TF_ReturnOutPut(1, Output, Prefix FileCounter "." Extension, 0, 1)
				If (CheckOutput <> "") and (OW = 2)
					TF_SetGlobal(Prefix FileCounter,Output)
				Output:=
			}
			If (InFile = 2)
			{
				OutPut := PreviousOutput
				StringReplace, CheckOutput, Output, `n, , All
				StringReplace, CheckOutput, CheckOutput, `r, , All
				If (CheckOutput <> "") and (OW <> 2)
					TF_ReturnOutPut(1, Output, Prefix FileCounter "." Extension, 0, 1)
				If (CheckOutput <> "") and (OW = 2)
					TF_SetGlobal(Prefix FileCounter,Output)
				OutPut := A_LoopField "`n"
			}
			If (Method <> 3)
				LineCounter=0
			FileCounter++
			Part++
			If (Method = 2)
			{
				If (Part > Split0)
				{
					Part=1
				}
				Where:=Split%Part%
			}
			If (Method = 3)
			{
				If (Part > Split0)
					Where:=Split%Split0%
				Else
					Where:=Split%Part%
			}
		}
		LineCounter++
		PreviousOutput:=Output
		PreviousLine:=A_LoopField
	}
	StringReplace, CheckOutput, Output, `n, , All
	StringReplace, CheckOutput, CheckOutput, `r, , All
	If (CheckOutPut <> "") and (OW <> 2)
		TF_ReturnOutPut(1, Output, Prefix FileCounter "." Extension, 0, 1)
	If (CheckOutput <> "") and (OW = 2)
	{
		TF_SetGlobal(Prefix FileCounter,Output)
		TF_SetGlobal(Prefix . "0" , FileCounter)
	}
}
TF_SplitFileByText(Text, SplitAt, Prefix = "file", Extension = "txt",  InFile = 1)
{
	LineCounter=1
	FileCounter=1
	TF_GetData(OW, Text, FileName)
	SplitPath, TextFile,, Dir
	Loop, Parse, Text, `n, `r
	{
		OutPut .= A_LoopField "`n"
		FoundPos:=RegExMatch(A_LoopField, SplitAt)
		If (FoundPos > 0)
		{
			If (InFile = 0)
			{
				StringReplace, CheckOutput, PreviousOutput, `n, , All
				StringReplace, CheckOutput, CheckOutput, `r, , All
				If (CheckOutput <> "") and (OW <> 2)
					TF_ReturnOutPut(1, PreviousOutput, Prefix FileCounter "." Extension, 0, 1)
				If (CheckOutput <> "") and (OW = 2)
					TF_SetGlobal(Prefix FileCounter,PreviousOutput)
				Output:=
			}
			If (InFile = 1)
			{
				StringReplace, CheckOutput, Output, `n, , All
				StringReplace, CheckOutput, CheckOutput, `r, , All
				If (CheckOutput <> "") and (OW <> 2)
					TF_ReturnOutPut(1, Output, Prefix FileCounter "." Extension, 0, 1)
				If (CheckOutput <> "") and (OW = 2)
					TF_SetGlobal(Prefix FileCounter,Output)
				Output:=
			}
			If (InFile = 2)
			{
				OutPut := PreviousOutput
				StringReplace, CheckOutput, Output, `n, , All
				StringReplace, CheckOutput, CheckOutput, `r, , All
				If (CheckOutput <> "") and (OW <> 2)
					TF_ReturnOutPut(1, Output, Prefix FileCounter "." Extension, 0, 1)
				If (CheckOutput <> "") and (OW = 2)
					TF_SetGlobal(Prefix FileCounter,Output)
				OutPut := A_LoopField "`n"
			}
			LineCounter=0
			FileCounter++
		}
		LineCounter++
		PreviousOutput:=Output
		PreviousLine:=A_LoopField
	}
	StringReplace, CheckOutput, Output, `n, , All
	StringReplace, CheckOutput, CheckOutput, `r, , All
	If (CheckOutPut <> "") and (OW <> 2)
		TF_ReturnOutPut(1, Output, Prefix FileCounter "." Extension, 0, 1)
	If (CheckOutput <> "") and (OW = 2)
	{
		TF_SetGlobal(Prefix FileCounter,Output)
		TF_SetGlobal(Prefix . "0" , FileCounter)
	}
}
TF_Find(Text, StartLine = 1, EndLine = 0, SearchText = "", ReturnFirst = 1, ReturnText = 0)
{
	TF_GetData(OW, Text, FileName)
	If (RegExMatch(Text, SearchText) < 1)
		Return "0"
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Loop, Parse, Text, `n
	{
		If A_Index in %TF_MatchList%
		{
			If (RegExMatch(A_LoopField, SearchText) > 0)
			{
				If (ReturnText = 0)
					Lines .= A_Index ","
				Else If (ReturnText = 1)
					Lines .= A_LoopField "`n"
				Else If (ReturnText = 2)
					Lines .= A_Index ": " A_LoopField "`n"
				If (ReturnFirst = 1)
					Break
			}
		}
	}
	If (Lines <> "")
		StringTrimRight, Lines, Lines, 1
	Else
		Lines = 0
	Return Lines
}
TF_FindLines(Text, StartLine = 1, EndLine = 0, SearchText = "", CaseSensitive = false)
{
	Return TF_Find(Text, StartLine, EndLine, SearchText, 0)
}
TF_Prepend(File1, File2)
{
	FileList=
(
%File1%
%File2%
)
	TF_Merge(FileList,"`n", "!" . File2)
	Return
}
TF_Append(File1, File2)
{
	FileList=
(
%File2%
%File1%
)
	TF_Merge(FileList,"`n", "!" . File2)
	Return
}
TF_Merge(FileList, Separator = "`n", FileName = "merged.txt")
{
	OW=0
	Loop, Parse, FileList, `n, `r
	{
		Append2File=
		IfExist, %A_LoopField%
		{
			FileRead, Append2File, %A_LoopField%
			If not ErrorLevel
				Output .= Append2File Separator
		}
	}
	If (SubStr(FileName,1,1)="!")
	{
		FileName:=SubStr(FileName,2)
		OW=1
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName, 0, 1)
}
TF_Wrap(Text, Columns = 80, AllowBreak = 0, StartLine = 1, EndLine = 0)
{
	TF_GetData(OW, Text, FileName)
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	If (AllowBreak = 1)
		Break=
	Else
		Break=[ \r?\n]
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
		{
			If (StrLen(A_LoopField) > Columns)
			{
				LoopField := A_LoopField " "
				OutPut .= RegExReplace(LoopField, "(.{1," . Columns . "})" . Break , "$1`n")
			}
			Else
				OutPut .= A_LoopField "`n"
		}
		Else
			OutPut .= A_LoopField "`n"
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_WhiteSpace(Text, RemoveLeading = 1, RemoveTrailing = 1, StartLine = 1, EndLine = 0) {
	TF_GetData(OW, Text, FileName)
	TF_MatchList:=_MakeMatchList(Text, StartLine, EndLine)
	Trim:=A_AutoTrim
	AutoTrim, On
	Loop, Parse, Text, `n, `r
	{
		If A_Index in %TF_MatchList%
		{
			If (RemoveLeading = 1) AND (RemoveTrailing = 1)
			{
				LoopField = %A_LoopField%
				Output .= LoopField "`n"
				Continue
			}
			If (RemoveLeading = 1) AND (RemoveTrailing = 0)
			{
				LoopField := A_LoopField . "."
				LoopField = %LoopField%
				StringTrimRight, LoopField, LoopField, 1
				Output .=  LoopField "`n"
				Continue
			}
			If (RemoveLeading = 0) AND (RemoveTrailing = 1)
			{
				LoopField := "." A_LoopField
				LoopField = %LoopField%
				StringTrimLeft, LoopField, LoopField, 1
				Output .= LoopField "`n"
				Continue
			}
			If (RemoveLeading = 0) AND (RemoveTrailing = 0)
			{
				Output .= A_LoopField "`n"
				Continue
			}
		}
		Else
			Output .= A_LoopField "`n"
	}
	AutoTrim, %Trim%
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_Substract(File1, File2, PartialMatch = 0) {
	Text:=File1
	TF_GetData(OW, Text, FileName)
	Str1:=Text
	Text:=File2
	TF_GetData(OW, Text, FileName)
	OutPut:=Text
	If (OW = 2)
		File1=
	OutPut .= "`n"
	If (PartialMatch = 1)
	{
		Loop, Parse, Str1, `n, `r
			StringReplace, Output, Output, %A_LoopField%, , All
	}
	Else
	{
		search:="m)^(.*)$"
		replace=__bol__$1__eol__
		Output:=RegExReplace(Output, search, replace)
		StringReplace, Output, Output, `n__eol__,__eol__ , All
		Loop, Parse, Str1, `n, `r
			StringReplace, Output, Output, __bol__%A_LoopField%__eol__, , All
	}
	If (PartialMatch = 0)
	{
		StringReplace, Output, Output, __bol__, , All
		StringReplace, Output, Output, __eol__, , All
	}
	Loop
	{
		StringReplace, Output, Output, `r`n`r`n, `r`n, UseErrorLevel
		if (ErrorLevel = 0) or (ErrorLevel = 1)
			break
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName, 0)
}
TF_RangeReplace(Text, SearchTextBegin, SearchTextEnd, ReplaceText = "", CaseSensitive = "False", KeepBegin = 0, KeepEnd = 0)
{
	TF_GetData(OW, Text, FileName)
	IfNotInString, Text, %SearchText%
		Return Text
	Start = 0
	End = 0
	If (KeepBegin = 1)
		KeepBegin:=SearchTextBegin
	Else
		KeepBegin=
	If (KeepEnd = 1)
		KeepEnd:= SearchTextEnd
	Else
		KeepEnd=
	If (SearchTextBegin = "")
		Start=1
	If (SearchTextEnd = "")
		End=2
	Loop, Parse, Text, `n, `r
	{
		If (End = 1)
		{
			Output .= A_LoopField "`n"
			Continue
		}
		If (Start = 0)
		{
			If (InStr(A_LoopField,SearchTextBegin,CaseSensitive))
			{
				Start = 1
				KeepSection := SubStr(A_LoopField, 1, InStr(A_LoopField, SearchTextBegin)-1)
				EndSection := SubStr(A_LoopField, InStr(A_LoopField, SearchTextBegin)-1)
				If (InStr(EndSection,SearchTextEnd,CaseSensitive))
				{
					EndSection := ReplaceText KeepEnd SubStr(EndSection, InStr(EndSection, SearchTextEnd) + StrLen(SearchTextEnd) ) "`n"
					If (End <> 2)
						End=1
					If (End = 2)
						EndSection=
				}
				Else
					EndSection=
				Output .= KeepSection KeepBegin EndSection
				Continue
			}
			Else
				Output .= A_LoopField "`n"
		}
		If (Start = 1) and (End <> 2)
		{
			If (InStr(A_LoopField,SearchTextEnd,CaseSensitive))
			{
				End = 1
				Output .= ReplaceText KeepEnd SubStr(A_LoopField, InStr(A_LoopField, SearchTextEnd) + StrLen(SearchTextEnd) ) "`n"
			}
		}
	}
	If (End = 2)
		Output .= ReplaceText
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_MakeFile(Text, Lines = 1, Columns = 1, Fill = " ")
{
	OW=1
	If (Text = "")
		OW=2
	Loop, % Columns
		Cols .= Fill
	Loop, % Lines
		Output .= Cols "`n"
	Return TF_ReturnOutPut(OW, OutPut, Text, 1, 1)
}
TF_Tab2Spaces(Text, TabStop = 4, StartLine = 1, EndLine =0)
{
	Loop, % TabStop
		Replace .= A_Space
	Return TF_ReplaceInLines(Text, StartLine, EndLine, A_Tab, Replace)
}
TF_Spaces2Tab(Text, TabStop = 4, StartLine = 1, EndLine =0)
{
	Loop, % TabStop
		Replace .= A_Space
	Return TF_ReplaceInLines(Text, StartLine, EndLine, Replace, A_Tab)
}
TF_Sort(Text, SortOptions = "", StartLine = 1, EndLine = 0)
{
	TF_GetData(OW, Text, FileName)
	If StartLine contains -,+,`,
		Return
	If (StartLine = 1) and (Endline = 0)
	{
		Output:=Text
		Sort, Output, %SortOptions%
	}
	Else
	{
		Output := TF_ReadLines(Text, 1, StartLine-1)
		ToSort := TF_ReadLines(Text, StartLine, EndLine)
		Sort, ToSort, %SortOptions%
		OutPut .= ToSort
		OutPut .= TF_ReadLines(Text, EndLine+1)
	}
	Return TF_ReturnOutPut(OW, OutPut, FileName)
}
TF_Tail(Text, Lines = 1, RemoveTrailing = 0, ReturnEmpty = 1)
{
	TF_GetData(OW, Text, FileName)
	Neg = 0
	If (Lines < 0)
	{
		Neg=1
		Lines:= Lines * -1
	}
	If (ReturnEmpty = 0)
	{
		Loop, Parse, Text, `n, `r
			OutPut .= (RegExMatch(A_LoopField,"[\S]+?\r?\n?")) ? A_LoopField "`n" :
		StringTrimRight, OutPut, OutPut, 1
		Text:=OutPut
		OutPut=
	}
	If (Neg = 1)
	{
		Lines++
		Output:=Text
		StringGetPos, Pos, Output, `n, R%Lines%
		StringTrimLeft, Output, Output, % ++Pos
		StringGetPos, Pos, Output, `n
		StringLeft, Output, Output, % Pos
		Output .= "`n"
	}
	Else
	{
		Output:=Text
		StringGetPos, Pos, Output, `n, R%Lines%
		StringTrimLeft, Output, Output, % ++Pos
		Output .= "`n"
	}
	OW = 2
	Return TF_ReturnOutPut(OW, OutPut, FileName, RemoveTrailing)
}
TF_Count(String, Char)
{
	StringReplace, String, String, %Char%,, UseErrorLevel
	Return ErrorLevel
}
TF_Save(Text, FileName, OverWrite = 1) {
	Return TF_ReturnOutPut(OverWrite, Text, FileName, 0, 1)
}
TF(TextFile, CreateGlobalVar = "T") {
	global
	FileRead, %CreateGlobalVar%, %TextFile%
	Return, (%CreateGlobalVar%)
}
TF_SetGlobal(var, content = "")
{
	global
	%var% := content
}
TF_GetData(byref OW, byref Text, byref FileName)
{
	OW=0
	IfNotInString, Text, `n
	{
		If (SubStr(Text,1,1)="!")
		{
			Text:=SubStr(Text,2)
			OW=1
		}
		IfNotExist, %Text%
		{
			If (OW=1)
				Text:= "!" . Text
			OW=2
		}
	}
	Else
	{
		OW=2
	}
	If (OW = 0) or (OW = 1)
	{
		Text := (SubStr(Text,1,1)="!") ? (SubStr(Text,2)) : Text
		FileName=%Text%
		FileRead, Text, %Text%
		If (ErrorLevel > 0)
		{
			MsgBox, 48, TF Lib Error, % "Can not read " FileName
			ExitApp
		}
	}
	Return
}
TF_SetWidth(Text,Width,AlignText)
{
	If (AlignText!=0 and AlignText!=1 and AlignText!=2)
		AlignText=0
	If AlignText=0
	{
		RetStr= % (Text)TF_Space(Width)
		StringLeft, RetText, RetText, %Width%
	}
	If AlignText=1
	{
		Spaces:=(Width-(StrLen(Text)))
		RetStr= % TF_Space(Round(Spaces/2))(Text)TF_Space(Spaces-(Round(Spaces/2)))
	}
	If AlignText=2
	{
		RetStr= % TF_Space(Width)(Text)
		StringRight, RetStr, RetStr, %Width%
	}
	Return RetStr
}
TF_Space(Width)
{
	Loop,%Width%
		Space=% Space Chr(32)
	Return Space
}
TF_ReturnOutPut(OW, Text, FileName, TrimTrailing = 1, CreateNewFile = 0) {
	If (OW = 0)
	{
		IfNotExist, % FileName
		{
			If (CreateNewFile = 1)
			{
				OW = 1
				Goto CreateNewFile
			}
			Else
				Return
		}
		If (TrimTrailing = 1)
			StringTrimRight, Text, Text, 1
		SplitPath, FileName,, Dir, Ext, Name
		If (Dir = "")
			Dir := A_ScriptDir
		IfExist, % Dir "\backup"
			FileCopy, % Dir "\" Name "_copy." Ext, % Dir "\backup\" Name "_copy.bak", 1
		FileDelete, % Dir "\" Name "_copy." Ext
		FileAppend, %Text%, % Dir "\" Name "_copy." Ext
		Return Errorlevel ? False : True
	}
	CreateNewFile:
	If (OW = 1)
	{
		IfNotExist, % FileName
		{
			If (CreateNewFile = 0)
				Return
		}
		If (TrimTrailing = 1)
			StringTrimRight, Text, Text, 1
		SplitPath, FileName,, Dir, Ext, Name
		If (Dir = "")
			Dir := A_ScriptDir
		IfExist, % Dir "\backup"
			FileCopy, % Dir "\" Name "." Ext, % Dir "\backup\" Name ".bak", 1
		FileDelete, % Dir "\" Name "." Ext
		FileAppend, %Text%, % Dir "\" Name "." Ext
		Return Errorlevel ? False : True
	}
	If (OW = 2)
	{
		If (TrimTrailing = 1)
			StringTrimRight, Text, Text, 1
		Return Text
	}
}
_MakeMatchList(Text, Start = 1, End = 0)
{
	ErrorList=
	 (join|
	 Error 01: Invalid StartLine parameter (non numerical character)
	 Error 02: Invalid EndLine parameter (non numerical character)
	 Error 03: Invalid StartLine parameter (only one + allowed)
)
	StringSplit, ErrorMessage, ErrorList, |
	Error = 0
	TF_MatchList=
	If (Start = 0 or Start = "")
		Start = 1
	If (RegExReplace(Start, "[ 0-9+\-\,]", "") <> "")
		Error = 1
	If (RegExReplace(End, "[0-9 ]", "") <> "")
		Error = 2
	If (TF_Count(Start,"+") > 1)
		Error = 3
	If (Error > 0 )
	{
		MsgBox, 48, TF Lib Error, % ErrorMessage%Error%
		ExitApp
	}
	IfInString, Start, `+
	{
		If (End = 0 or End = "")
			End:= TF_Count(Text, "`n") + 1
		StringSplit, Section, Start, `,
		Loop, %Section0%
		{
			StringSplit, SectionLines, Section%A_Index%, `+
			LoopSection:=End + 1 - SectionLines1
			Counter=0
			TF_MatchList .= SectionLines1 ","
			Loop, %LoopSection%
			{
				If (A_Index >= End)
					Break
				If (Counter = (SectionLines2-1))
				{
					TF_MatchList .= (SectionLines1 + A_Index) ","
					Counter=0
				}
				Else
					Counter++
			}
		}
		StringTrimRight, TF_MatchList, TF_MatchList, 1
		Return TF_MatchList
	}
	IfInString, Start, `-
	{
		StringSplit, Section, Start, `,
		Loop, %Section0%
		{
			StringSplit, SectionLines, Section%A_Index%, `-
			LoopSection:=SectionLines2 + 1 - SectionLines1
			Loop, %LoopSection%
			{
				TF_MatchList .= (SectionLines1 - 1 + A_Index) ","
			}
		}
		StringTrimRight, TF_MatchList, TF_MatchList, 1
		Return TF_MatchList
	}
	IfInString, Start, `,
	{
		TF_MatchList:=Start
		Return TF_MatchList
	}
	If (End = 0 or End = "")
		End:= TF_Count(Text, "`n") + 1
	LoopTimes:=End-Start
	Loop, %LoopTimes%
	{
		TF_MatchList .= (Start - 1 + A_Index) ","
	}
	TF_MatchList .= End ","
	StringTrimRight, TF_MatchList, TF_MatchList, 1
	Return TF_MatchList
}
;}

Timeout(){
	timeout:= IEVGET(title)
	Loop{
		sleep,100
		Source1:=timeout.document.getElementsByTagName("HTML")[0].outerhtml
	}until Source1 != ""
	sleep,1000
	Loop{
		sleep,100
		Source2:=timeout.document.getElementsByTagName("HTML")[0].outerhtml
	}until Source1 != Source2
	sleep, 1000
	If title:= IEGET("Service Centre 5 Login"){
		Timeout = true
	}
	title:=""
	return Timeout
}

WinMove() {
	PostMessage, 0xA1, 2
	WinGetActiveTitle,WinTitle
	while getkeystate("LButton")
	{
		Traytip, %WinTitle%, Moving...
	}
	Traytip
	WingetPos, Xpos, Ypos, W, H, %WinTitle%
	TrayTip,%wintitle%, Moved to %Xpos% x %Ypos%
	sleep, 500
	IniWrite,%Ypos%,Modules\Config.ini,%WinTitle% Position,GuiY
	Iniwrite,%Xpos%,Modules\Config.ini,%WinTitle% Position,GuiX
	TrayTip,%wintitle%, Moved to %Xpos% x %Ypos%`nSaved Position
	sleep, 3000
	TrayTip
	return
}

Taskbar(GuiHeight) {
	ABM_GETSTATE := 0x00000004
	NumPut(VarSetCapacity(ABD, 36, 0), ABD)
	Taskbar := DllCall("Shell32.dll\SHAppBarMessage", UInt, ABM_GETSTATE, UInt, &ABD)
	If (Taskbar = 3) {
		Ypos := A_ScreenHeight - GuiHeight
	} else {
		WinGetPos,,,,TrayHeight,ahk_class Shell_TrayWnd,,,
		Ypos := A_ScreenHeight-GuiHeight-TrayHeight
	}
	return Ypos
}

