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
	IniWrite, %Xpos%, %Config%,%Title%,Xpos
	IniWrite, %Ypos%, %Config%,%Title%,Ypos
	return
}

GetWinPosX(title=""){
	IniRead, PosX,%Config%,%Title%,Xpos
	Return PosX
}

GetWinPosY(title=""){
	IniRead, PosY,%Config%,%Title%,Ypos
		Return PosY
}

GetProductCode() {
	try{
	If Pwb := IETitle("ESOLBRANCH LIVE DB / \w+ / DLL Ver: " TesseractVersion " / Page Ver: " TesseractVersion) {
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
		
		If Pwb := IETitle("Repair Shipping Wizard - " TesseractVersion) { ;Job Shipout
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
		tooltip, closed
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

;{ ----Encryption
	class Crypt
	{
		class Encrypt
		{
			static StrEncoding := "UTF-16"
			static PassEncoding := "UTF-16"
			StrDecryptToFile(EncryptedHash,pFileOut,password,CryptAlg = 1, HashAlg = 1)
			{
				if !EncryptedHash
					return ""
				if !len := b64Decode( EncryptedHash, encr_Buf )
					return ""
				temp_file := "crypt.temp"
				f := FileOpen(temp_file,"w","CP0")
				if !IsObject(f)
					return ""
				if !f.RawWrite(encr_Buf,len)
					return ""
				f.close()
				bytes := this._Encrypt( p, pp, password, 0, temp_file, pFileOut, CryptAlg, HashAlg )
				FileDelete,% temp_file
				return bytes
			}
			FileEncryptToStr(pFileIn,password,CryptAlg = 1, HashAlg = 1)
			{
				temp_file := "crypt.temp"
				if !this._Encrypt( p, pp, password, 1, pFileIn, temp_file, CryptAlg, HashAlg )
					return ""
				f := FileOpen(temp_file,"r","CP0")
				if !IsObject(f)
				{
					FileDelete,% temp_file
					return ""
				}
				f.Pos := 0
				fLen := f.Length
				VarSetCapacity(tembBuf,fLen,0)
				if !f.RawRead(tembBuf,fLen)
				{
					Free(tembBuf)
					return ""
				}
				f.Close()
				FileDelete,% temp_file
				return b64Encode( tembBuf, fLen )
			}
			FileEncrypt(pFileIn,pFileOut,password,CryptAlg = 1, HashAlg = 1)
			{
				return this._Encrypt( p, pp, password, 1, pFileIn, pFileOut, CryptAlg, HashAlg )
			}
			FileDecrypt(pFileIn,pFileOut,password,CryptAlg = 1, HashAlg = 1)
			{
				return this._Encrypt( p, pp, password, 0, pFileIn, pFileOut, CryptAlg, HashAlg )
			}
			StrEncrypt(string,password,CryptAlg = 1, HashAlg = 1)
			{
				len := StrPutVar(string, str_buf,100,this.StrEncoding)
				if this._Encrypt(str_buf,len, password, 1,0,0,CryptAlg,HashAlg)
					return b64Encode( str_buf, len )
				else
					return ""
			}
			StrDecrypt(EncryptedHash,password,CryptAlg = 1, HashAlg = 1)
			{
				if !EncryptedHash
					return ""
				if !len := b64Decode( EncryptedHash, encr_Buf )
					return 0
				if sLen := this._Encrypt(encr_Buf,len, password, 0,0,0,CryptAlg,HashAlg)
				{
					if ( this.StrEncoding = "utf-16" || this.StrEncoding = "cp1200" )
						sLen /= 2
					return strget(&encr_Buf,sLen,this.StrEncoding)
				}
				else
					return ""
			}
			_Encrypt(ByRef encr_Buf,ByRef Buf_Len, password, mode, pFileIn=0, pFileOut=0, CryptAlg = 1,HashAlg = 1)
			{
				c := CryptConst
				CUR_PWD_HASH_ALG := HashAlg == 1 || HashAlg = "MD5" ?c.CALG_MD5
				:HashAlg==2 || HashAlg = "MD2" 	?c.CALG_MD2
				:HashAlg==3 || HashAlg = "SHA"	?c.CALG_SHA
				:HashAlg==4 || HashAlg = "SHA256" ?c.CALG_SHA_256
				:HashAlg==5 || HashAlg = "SHA384" ?c.CALG_SHA_384
				:HashAlg==6 || HashAlg = "SHA512" ?c.CALG_SHA_512
				:0
				CUR_ENC_ALG 	:= CryptAlg==1 || CryptAlg = "RC4" 			? ( c.CALG_RC4, KEY_LENGHT:=0x80 )
				:CryptAlg==2 || CryptAlg = "RC2" 		? ( c.CALG_RC2, KEY_LENGHT:=0x80 )
				:CryptAlg==3 || CryptAlg = "3DES" 		? ( c.CALG_3DES, KEY_LENGHT:=0xC0 )
				:CryptAlg==4 || CryptAlg = "3DES112" ? ( c.CALG_3DES_112, KEY_LENGHT:=0x80 )
				:CryptAlg==5 || CryptAlg = "AES128" 	? ( c.CALG_AES_128, KEY_LENGHT:=0x80 )
				:CryptAlg==6 || CryptAlg = "AES192" 	? ( c.CALG_AES_192, KEY_LENGHT:=0xC0 )
				:CryptAlg==7 || CryptAlg = "AES256" 	? ( c.CALG_AES_256, KEY_LENGHT:=0x100 )
				:0
				KEY_LENGHT <<= 16
				if (CUR_PWD_HASH_ALG = 0 || CUR_ENC_ALG = 0)
					return 0
				if !dllCall("Advapi32\CryptAcquireContextW","Ptr*",hCryptProv,"Uint",0,"Uint",0,"Uint",c.PROV_RSA_AES,"UInt",c.CRYPT_VERIFYCONTEXT)
					{foo := "CryptAcquireContextW", err := GetLastError(), err2 := ErrorLevel
				GoTO FINITA_LA_COMEDIA
			}
			if !dllCall("Advapi32\CryptCreateHash","Ptr",hCryptProv,"Uint",CUR_PWD_HASH_ALG,"Uint",0,"Uint",0,"Ptr*",hHash )
				{foo := "CryptCreateHash", err := GetLastError(), err2 := ErrorLevel
			GoTO FINITA_LA_COMEDIA
		}
		passLen := StrPutVar(password, passBuf,0,this.PassEncoding)
		if !dllCall("Advapi32\CryptHashData","Ptr",hHash,"Ptr",&passBuf,"Uint",passLen,"Uint",0 )
			{foo := "CryptHashData", err := GetLastError(), err2 := ErrorLevel
		GoTO FINITA_LA_COMEDIA
	}
	if !dllCall("Advapi32\CryptDeriveKey","Ptr",hCryptProv,"Uint",CUR_ENC_ALG,"Ptr",hHash,"Uint",KEY_LENGHT,"Ptr*",hKey )
		{foo := "CryptDeriveKey", err := GetLastError(), err2 := ErrorLevel
	GoTO FINITA_LA_COMEDIA
}
if !dllCall("Advapi32\CryptGetKeyParam","Ptr",hKey,"Uint",c.KP_BLOCKLEN,"Uint*",BlockLen,"Uint*",dwCount := 4,"Uint",0)
	{foo := "CryptGetKeyParam", err := GetLastError(), err2 := ErrorLevel
GoTO FINITA_LA_COMEDIA
}
BlockLen /= 8
if (mode == 1)
{
if (pFileIn && pFileOut)
{
	ReadBufSize := 10240 - mod(10240,BlockLen==0?1:BlockLen )
	pfin := FileOpen(pFileIn,"r","CP0")
	pfout := FileOpen(pFileOut,"w","CP0")
	if !IsObject(pfin)
		{foo := "File Opening " . pFileIn
	GoTO FINITA_LA_COMEDIA
}
if !IsObject(pfout)
	{foo := "File Opening " . pFileOut
GoTO FINITA_LA_COMEDIA
}
pfin.Pos := 0
VarSetCapacity(ReadBuf,ReadBufSize+BlockLen,0)
isFinal := 0
hModule := DllCall("LoadLibrary", "Str", "Advapi32.dll","UPtr")
CryptEnc := DllCall("GetProcAddress", "Ptr", hModule, "AStr", "CryptEncrypt","UPtr")
while !pfin.AtEOF
{
BytesRead := pfin.RawRead(ReadBuf, ReadBufSize)
if pfin.AtEOF
	isFinal := 1
if !dllCall(CryptEnc
	,"Ptr",hKey
,"Ptr",0
,"Uint",isFinal
,"Uint",0
,"Ptr",&ReadBuf
,"Uint*",BytesRead
,"Uint",ReadBufSize+BlockLen )
{foo := "CryptEncrypt", err := GetLastError(), err2 := ErrorLevel
GoTO FINITA_LA_COMEDIA
}
pfout.RawWrite(ReadBuf,BytesRead)
Buf_Len += BytesRead
}
DllCall("FreeLibrary", "Ptr", hModule)
pfin.Close()
pfout.Close()
}
else
{
if !dllCall("Advapi32\CryptEncrypt"
,"Ptr",hKey
,"Ptr",0
,"Uint",1
,"Uint",0
,"Ptr",&encr_Buf
,"Uint*",Buf_Len
,"Uint",Buf_Len + BlockLen )
{foo := "CryptEncrypt", err := GetLastError(), err2 := ErrorLevel
GoTO FINITA_LA_COMEDIA
}
}
}
else if (mode == 0)
{
if (pFileIn && pFileOut)
{
ReadBufSize := 10240 - mod(10240,BlockLen==0?1:BlockLen )
pfin := FileOpen(pFileIn,"r","CP0")
pfout := FileOpen(pFileOut,"w","CP0")
if !IsObject(pfin)
{foo := "File Opening " . pFileIn
GoTO FINITA_LA_COMEDIA
}
if !IsObject(pfout)
{foo := "File Opening " . pFileOut
GoTO FINITA_LA_COMEDIA
}
pfin.Pos := 0
VarSetCapacity(ReadBuf,ReadBufSize+BlockLen,0)
isFinal := 0
hModule := DllCall("LoadLibrary", "Str", "Advapi32.dll","UPtr")
CryptDec := DllCall("GetProcAddress", "Ptr", hModule, "AStr", "CryptDecrypt","UPtr")
while !pfin.AtEOF
{
BytesRead := pfin.RawRead(ReadBuf, ReadBufSize)
if pfin.AtEOF
isFinal := 1
if !dllCall(CryptDec
,"Ptr",hKey
,"Ptr",0
,"Uint",isFinal
,"Uint",0
,"Ptr",&ReadBuf
,"Uint*",BytesRead )
{foo := "CryptDecrypt", err := GetLastError(), err2 := ErrorLevel
GoTO FINITA_LA_COMEDIA
}
pfout.RawWrite(ReadBuf,BytesRead)
Buf_Len += BytesRead
}
DllCall("FreeLibrary", "Ptr", hModule)
pfin.Close()
pfout.Close()
}
else if !dllCall("Advapi32\CryptDecrypt"
,"Ptr",hKey
,"Ptr",0
,"Uint",1
,"Uint",0
,"Ptr",&encr_Buf
,"Uint*",Buf_Len )
{foo := "CryptDecrypt", err := GetLastError(), err2 := ErrorLevel
GoTO FINITA_LA_COMEDIA
}
}
FINITA_LA_COMEDIA:
dllCall("Advapi32\CryptDestroyKey","Ptr",hKey )
dllCall("Advapi32\CryptDestroyHash","Ptr",hHash)
dllCall("Advapi32\CryptReleaseContext","Ptr",hCryptProv,"UInt",0)
if (A_ThisLabel = "FINITA_LA_COMEDIA")
{
if (A_IsCompiled = 1)
return ""
else
msgbox % foo " call failed with:`nErrorLevel: " err2 "`nLastError: " err "`n" ErrorFormat(err)
return ""
}
return Buf_Len
}
}
class Hash
{
static StrEncoding := "CP0"
static PassEncoding := "UTF-16"
FileHash(pFile,HashAlg = 1,pwd = "",hmac_alg = 1)
{
return this._CalcHash(p,pp,pFile,HashAlg,pwd,hmac_alg)
}
StrHash(string,HashAlg = 1,pwd = "",hmac_alg = 1)
{
buf_len := StrPutVar(string, buf,0,this.StrEncoding)
return this._CalcHash(buf,buf_len,0,HashAlg,pwd,hmac_alg)
}
_CalcHash(ByRef bBuffer,BufferLen,pFile,HashAlg = 1,pwd = "",hmac_alg = 1)
{
c := CryptConst
HASH_ALG := HashAlg==1?c.CALG_MD5
:HashAlg==2?c.CALG_MD2
:HashAlg==3?c.CALG_SHA
:HashAlg==4?c.CALG_SHA_256
:HashAlg==5?c.CALG_SHA_384
:HashAlg==6?c.CALG_SHA_512
:0
HMAC_KEY_ALG 	:= hmac_alg==1?c.CALG_RC4
:hmac_alg==2?c.CALG_RC2
:hmac_alg==3?c.CALG_3DES
:hmac_alg==4?c.CALG_3DES_112
:hmac_alg==5?c.CALG_AES_128
:hmac_alg==6?c.CALG_AES_192
:hmac_alg==7?c.CALG_AES_256
:0
KEY_LENGHT 		:= hmac_alg==1?0x80
:hmac_alg==2?0x80
:hmac_alg==3?0xC0
:hmac_alg==4?0x80
:hmac_alg==5?0x80
:hmac_alg==6?0xC0
:hmac_alg==7?0x100
:0
KEY_LENGHT <<= 16
if (!HASH_ALG || !HMAC_KEY_ALG)
return 0
if !dllCall("Advapi32\CryptAcquireContextW","Ptr*",hCryptProv,"Uint",0,"Uint",0,"Uint",c.PROV_RSA_AES,"UInt",c.CRYPT_VERIFYCONTEXT )
{foo := "CryptAcquireContextW", err := GetLastError(), err2 := ErrorLevel
GoTO FINITA_DA_COMEDIA
}
if !dllCall("Advapi32\CryptCreateHash","Ptr",hCryptProv,"Uint",HASH_ALG,"Uint",0,"Uint",0,"Ptr*",hHash )
{foo := "CryptCreateHash1", err := GetLastError(), err2 := ErrorLevel
GoTO FINITA_DA_COMEDIA
}
if (pwd != "")
{
passLen := StrPutVar(pwd, passBuf,0,this.PassEncoding)
if !dllCall("Advapi32\CryptHashData","Ptr",hHash,"Ptr",&passBuf,"Uint",passLen,"Uint",0 )
{foo := "CryptHashData Pwd", err := GetLastError(), err2 := ErrorLevel
GoTO FINITA_DA_COMEDIA
}
if !dllCall("Advapi32\CryptDeriveKey","Ptr",hCryptProv,"Uint",HMAC_KEY_ALG,"Ptr",hHash,"Uint",KEY_LENGHT,"Ptr*",hKey )
{foo := "CryptDeriveKey Pwd", err := GetLastError(), err2 := ErrorLevel
GoTO FINITA_DA_COMEDIA
}
dllCall("Advapi32\CryptDestroyHash","Ptr",hHash)
if !dllCall("Advapi32\CryptCreateHash","Ptr",hCryptProv,"Uint",c.CALG_HMAC,"Ptr",hKey,"Uint",0,"Ptr*",hHash )
{foo := "CryptCreateHash2", err := GetLastError(), err2 := ErrorLevel
GoTO FINITA_DA_COMEDIA
}
VarSetCapacity(HmacInfoStruct,4*A_PtrSize + 4,0)
NumPut(HASH_ALG,HmacInfoStruct,0,"UInt")
if !dllCall("Advapi32\CryptSetHashParam","Ptr",hHash,"Uint",c.HP_HMAC_INFO,"Ptr",&HmacInfoStruct,"Uint",0)
{foo := "CryptSetHashParam", err := GetLastError(), err2 := ErrorLevel
GoTO FINITA_DA_COMEDIA
}
}
if pFile
{
f := FileOpen(pFile,"r","CP0")
BUFF_SIZE := 1024 * 1024
if !IsObject(f)
{foo := "File Opening"
GoTO FINITA_DA_COMEDIA
}
if !hModule := DllCall( "GetModuleHandleW", "str", "Advapi32.dll", "UPtr" )
hModule := DllCall( "LoadLibraryW", "str", "Advapi32.dll", "UPtr" )
hCryptHashData := DllCall("GetProcAddress", "Ptr", hModule, "AStr", "CryptHashData", "UPtr")
VarSetCapacity(read_buf,BUFF_SIZE,0)
f.Pos := 0
While (cbCount := f.RawRead(read_buf, BUFF_SIZE))
{
if (cbCount = 0)
break
if !dllCall(hCryptHashData
,"Ptr",hHash
,"Ptr",&read_buf
,"Uint",cbCount
,"Uint",0 )
{foo := "CryptHashData", err := GetLastError(), err2 := ErrorLevel
GoTO FINITA_DA_COMEDIA
}
}
f.Close()
}
else
{
if !dllCall("Advapi32\CryptHashData"
,"Ptr",hHash
,"Ptr",&bBuffer
,"Uint",BufferLen
,"Uint",0 )
{foo := "CryptHashData", err := GetLastError(), err2 := ErrorLevel
GoTO FINITA_DA_COMEDIA
}
}
if !dllCall("Advapi32\CryptGetHashParam","Ptr",hHash,"Uint",c.HP_HASHSIZE,"Uint*",HashLen,"Uint*",HashLenSize := 4,"UInt",0 )
{foo := "CryptGetHashParam HP_HASHSIZE", err := GetLastError(), err2 := ErrorLevel
GoTO FINITA_DA_COMEDIA
}
VarSetCapacity(pbHash,HashLen,0)
if !dllCall("Advapi32\CryptGetHashParam","Ptr",hHash,"Uint",c.HP_HASHVAL,"Ptr",&pbHash,"Uint*",HashLen,"UInt",0 )
{foo := "CryptGetHashParam HP_HASHVAL", err := GetLastError(), err2 := ErrorLevel
GoTO FINITA_DA_COMEDIA
}
hashval := b2a_hex( pbHash, HashLen )
FINITA_DA_COMEDIA:
DllCall("FreeLibrary", "Ptr", hModule)
dllCall("Advapi32\CryptDestroyHash","Ptr",hHash)
dllCall("Advapi32\CryptDestroyKey","Ptr",hKey )
dllCall("Advapi32\CryptReleaseContext","Ptr",hCryptProv,"UInt",0)
if (A_ThisLabel = "FINITA_LA_COMEDIA")
{
if (A_IsCompiled = 1)
return ""
else
msgbox % foo " call failed with:`nErrorLevel: " err2 "`nLastError: " err "`n" ErrorFormat(err)
return 0
}
return hashval
}
}
}
GetLastError()
{
return ToHex(A_LastError < 0 ? A_LastError & 0xFFFFFFFF : A_LastError)
}
ToHex(num)
{
if num is not integer
return num
oldFmt := A_FormatInteger
SetFormat, integer, hex
num := num + 0
SetFormat, integer,% oldFmt
return num
}
ErrorFormat(error_id)
{
VarSetCapacity(msg,1000,0)
if !len := DllCall("FormatMessageW"
,"UInt",FORMAT_MESSAGE_FROM_SYSTEM := 0x00001000 | FORMAT_MESSAGE_IGNORE_INSERTS := 0x00000200
,"Ptr",0
,"UInt",error_id
,"UInt",0
,"Ptr",&msg
,"UInt",500)
return
return 	strget(&msg,len)
}
StrPutVar(string, ByRef var, addBufLen = 0,encoding="UTF-16")
{
tlen := ((encoding="utf-16"||encoding="cp1200") ? 2 : 1)
str_len := StrPut(string, encoding) * tlen
VarSetCapacity( var, str_len + addBufLen,0 )
StrPut( string, &var, encoding )
return str_len - tlen
}
SetKeySalt(hKey,hProv)
{
KP_SALT_EX := 10
SALT := "89ABF9C1005EDD40"
VarSetCapacity(st,2*A_PtrSize,0)
NumPut(len,st,0,"UInt")
NumPut(&pb,st,A_PtrSize,"UPtr")
if !dllCall("Advapi32\CryptSetKeyParam"
,"Ptr",hKey
,"Uint",KP_SALT_EX
,"Ptr",&st
,"Uint",0)
msgbox % ErrorFormat(GetLastError())
}
GetKeySalt(hKey)
{
KP_IV := 1
KP_SALT := 2
if !dllCall("Advapi32\CryptGetKeyParam"
,"Ptr",hKey
,"Uint",KP_SALT
,"Uint",0
,"Uint*",dwCount
,"Uint",0)
msgbox % "Fail to get SALT length."
msgbox % "SALT length.`n" dwCount
VarSetCapacity(pb,dwCount,0)
if !dllCall("Advapi32\CryptGetKeyParam"
,"Ptr",hKey
,"Uint",KP_SALT
,"Ptr",&pb
,"Uint*",dwCount
,"Uint",0)
msgbox % "Fail to get SALT"
}
class CryptConst
{
static ALG_CLASS_ANY := (0)
static ALG_CLASS_SIGNATURE := (1 << 13)
static ALG_CLASS_MSG_ENCRYPT := (2 << 13)
static ALG_CLASS_DATA_ENCRYPT := (3 << 13)
static ALG_CLASS_HASH := (4 << 13)
static ALG_CLASS_KEY_EXCHANGE := (5 << 13)
static ALG_CLASS_ALL := (7 << 13)
static ALG_TYPE_ANY := (0)
static ALG_TYPE_DSS := (1 << 9)
static ALG_TYPE_RSA := (2 << 9)
static ALG_TYPE_BLOCK := (3 << 9)
static ALG_TYPE_STREAM := (4 << 9)
static ALG_TYPE_DH := (5 << 9)
static ALG_TYPE_SECURECHANNEL := (6 << 9)
static ALG_SID_ANY := (0)
static ALG_SID_RSA_ANY := 0
static ALG_SID_RSA_PKCS := 1
static ALG_SID_RSA_MSATWORK := 2
static ALG_SID_RSA_ENTRUST := 3
static ALG_SID_RSA_PGP := 4
static ALG_SID_DSS_ANY := 0
static ALG_SID_DSS_PKCS := 1
static ALG_SID_DSS_DMS := 2
static ALG_SID_ECDSA := 3
static ALG_SID_DES := 1
static ALG_SID_3DES := 3
static ALG_SID_DESX := 4
static ALG_SID_IDEA := 5
static ALG_SID_CAST := 6
static ALG_SID_SAFERSK64 := 7
static ALG_SID_SAFERSK128 := 8
static ALG_SID_3DES_112 := 9
static ALG_SID_CYLINK_MEK := 12
static ALG_SID_RC5 := 13
static ALG_SID_AES_128 := 14
static ALG_SID_AES_192 := 15
static ALG_SID_AES_256 := 16
static ALG_SID_AES := 17
static ALG_SID_SKIPJACK := 10
static ALG_SID_TEK := 11
static CRYPT_MODE_CBCI := 6
static CRYPT_MODE_CFBP := 7
static CRYPT_MODE_OFBP := 8
static CRYPT_MODE_CBCOFM := 9
static CRYPT_MODE_CBCOFMI := 10
static ALG_SID_RC2 := 2
static ALG_SID_RC4 := 1
static ALG_SID_SEAL := 2
static ALG_SID_DH_SANDF := 1
static ALG_SID_DH_EPHEM := 2
static ALG_SID_AGREED_KEY_ANY := 3
static ALG_SID_KEA := 4
static ALG_SID_ECDH := 5
static ALG_SID_MD2 := 1
static ALG_SID_MD4 := 2
static ALG_SID_MD5 := 3
static ALG_SID_SHA := 4
static ALG_SID_SHA1 := 4
static ALG_SID_MAC := 5
static ALG_SID_RIPEMD := 6
static ALG_SID_RIPEMD160 := 7
static ALG_SID_SSL3SHAMD5 := 8
static ALG_SID_HMAC := 9
static ALG_SID_TLS1PRF := 10
static ALG_SID_HASH_REPLACE_OWF := 11
static ALG_SID_SHA_256 := 12
static ALG_SID_SHA_384 := 13
static ALG_SID_SHA_512 := 14
static ALG_SID_SSL3_MASTER := 1
static ALG_SID_SCHANNEL_MASTER_HASH := 2
static ALG_SID_SCHANNEL_MAC_KEY := 3
static ALG_SID_PCT1_MASTER := 4
static ALG_SID_SSL2_MASTER := 5
static ALG_SID_TLS1_MASTER := 6
static ALG_SID_SCHANNEL_ENC_KEY := 7
static ALG_SID_ECMQV := 1
static ALG_SID_EXAMPLE := 80
static CALG_MD2 := (CryptConst.ALG_CLASS_HASH | CryptConst.ALG_TYPE_ANY | CryptConst.ALG_SID_MD2)
static CALG_MD4 := (CryptConst.ALG_CLASS_HASH | CryptConst.ALG_TYPE_ANY | CryptConst.ALG_SID_MD4)
static CALG_MD5 := (CryptConst.ALG_CLASS_HASH | CryptConst.ALG_TYPE_ANY | CryptConst.ALG_SID_MD5)
static CALG_SHA := (CryptConst.ALG_CLASS_HASH | CryptConst.ALG_TYPE_ANY | CryptConst.ALG_SID_SHA)
static CALG_SHA1 := (CryptConst.ALG_CLASS_HASH | CryptConst.ALG_TYPE_ANY | CryptConst.ALG_SID_SHA1)
static CALG_MAC := (CryptConst.ALG_CLASS_HASH | CryptConst.ALG_TYPE_ANY | CryptConst.ALG_SID_MAC)
static CALG_RSA_SIGN := (CryptConst.ALG_CLASS_SIGNATURE | CryptConst.ALG_TYPE_RSA | CryptConst.ALG_SID_RSA_ANY)
static CALG_DSS_SIGN := (CryptConst.ALG_CLASS_SIGNATURE | CryptConst.ALG_TYPE_DSS | CryptConst.ALG_SID_DSS_ANY)
static CALG_NO_SIGN := (CryptConst.ALG_CLASS_SIGNATURE | CryptConst.ALG_TYPE_ANY | CryptConst.ALG_SID_ANY)
static CALG_RSA_KEYX := (CryptConst.ALG_CLASS_KEY_EXCHANGE|CryptConst.ALG_TYPE_RSA|CryptConst.ALG_SID_RSA_ANY)
static CALG_DES := (CryptConst.ALG_CLASS_DATA_ENCRYPT|CryptConst.ALG_TYPE_BLOCK|CryptConst.ALG_SID_DES)
static CALG_3DES_112 := (CryptConst.ALG_CLASS_DATA_ENCRYPT|CryptConst.ALG_TYPE_BLOCK|CryptConst.ALG_SID_3DES_112)
static CALG_3DES := (CryptConst.ALG_CLASS_DATA_ENCRYPT|CryptConst.ALG_TYPE_BLOCK|CryptConst.ALG_SID_3DES)
static CALG_DESX := (CryptConst.ALG_CLASS_DATA_ENCRYPT|CryptConst.ALG_TYPE_BLOCK|CryptConst.ALG_SID_DESX)
static CALG_RC2 := (CryptConst.ALG_CLASS_DATA_ENCRYPT|CryptConst.ALG_TYPE_BLOCK|CryptConst.ALG_SID_RC2)
static CALG_RC4 := (CryptConst.ALG_CLASS_DATA_ENCRYPT|CryptConst.ALG_TYPE_STREAM|CryptConst.ALG_SID_RC4)
static CALG_SEAL := (CryptConst.ALG_CLASS_DATA_ENCRYPT|CryptConst.ALG_TYPE_STREAM|CryptConst.ALG_SID_SEA)
static CALG_DH_SF := (CryptConst.ALG_CLASS_KEY_EXCHANGE|CryptConst.ALG_TYPE_DH|CryptConst.ALG_SID_DH_SANDF)
static CALG_DH_EPHEM := (CryptConst.ALG_CLASS_KEY_EXCHANGE|CryptConst.ALG_TYPE_DH|CryptConst.ALG_SID_DH_EPHEM)
static CALG_AGREEDKEY_ANY := (CryptConst.ALG_CLASS_KEY_EXCHANGE|CryptConst.ALG_TYPE_DH|CryptConst.ALG_SID_AGREED_KEY_ANY)
static CALG_KEA_KEYX := (CryptConst.ALG_CLASS_KEY_EXCHANGE|CryptConst.ALG_TYPE_DH|CryptConst.ALG_SID_KEA)
static CALG_HUGHES_MD5 := (CryptConst.ALG_CLASS_KEY_EXCHANGE|CryptConst.ALG_TYPE_ANY|CryptConst.ALG_SID_MD5)
static CALG_SKIPJACK := (CryptConst.ALG_CLASS_DATA_ENCRYPT|CryptConst.ALG_TYPE_BLOCK|CryptConst.ALG_SID_SKIPJACK)
static CALG_TEK := (CryptConst.ALG_CLASS_DATA_ENCRYPT|CryptConst.ALG_TYPE_BLOCK|CryptConst.ALG_SID_TEK)
static CALG_CYLINK_MEK := (CryptConst.ALG_CLASS_DATA_ENCRYPT|CryptConst.ALG_TYPE_BLOCK|CryptConst.ALG_SID_CYLINK_MEK)
static CALG_SSL3_SHAMD5 := (CryptConst.ALG_CLASS_HASH | CryptConst.ALG_TYPE_ANY | CryptConst.ALG_SID_SSL3SHAMD5)
static CALG_SSL3_MASTER := (CryptConst.ALG_CLASS_MSG_ENCRYPT|CryptConst.ALG_TYPE_SECURECHANNEL|CryptConst.ALG_SID_SSL3_MASTER)
static CALG_SCHANNEL_MASTER_HASH := (CryptConst.ALG_CLASS_MSG_ENCRYPT|CryptConst.ALG_TYPE_SECURECHANNEL|CryptConst.ALG_SID_SCHANNEL_MASTER_HASH)
static CALG_SCHANNEL_MAC_KEY := (CryptConst.ALG_CLASS_MSG_ENCRYPT|CryptConst.ALG_TYPE_SECURECHANNEL|CryptConst.ALG_SID_SCHANNEL_MAC_KEY)
static CALG_SCHANNEL_ENC_KEY := (CryptConst.ALG_CLASS_MSG_ENCRYPT|CryptConst.ALG_TYPE_SECURECHANNEL|CryptConst.ALG_SID_SCHANNEL_ENC_KEY)
static CALG_PCT1_MASTER := (CryptConst.ALG_CLASS_MSG_ENCRYPT|CryptConst.ALG_TYPE_SECURECHANNEL|CryptConst.ALG_SID_PCT1_MASTER)
static CALG_SSL2_MASTER := (CryptConst.ALG_CLASS_MSG_ENCRYPT|CryptConst.ALG_TYPE_SECURECHANNEL|CryptConst.ALG_SID_SSL2_MASTER)
static CALG_TLS1_MASTER := (CryptConst.ALG_CLASS_MSG_ENCRYPT|CryptConst.ALG_TYPE_SECURECHANNEL|CryptConst.ALG_SID_TLS1_MASTER)
static CALG_RC5 := (CryptConst.ALG_CLASS_DATA_ENCRYPT|CryptConst.ALG_TYPE_BLOCK|CryptConst.ALG_SID_RC5)
static CALG_HMAC := (CryptConst.ALG_CLASS_HASH | CryptConst.ALG_TYPE_ANY | CryptConst.ALG_SID_HMAC)
static CALG_TLS1PRF := (CryptConst.ALG_CLASS_HASH | CryptConst.ALG_TYPE_ANY | CryptConst.ALG_SID_TLS1PRF)
static CALG_HASH_REPLACE_OWF := (CryptConst.ALG_CLASS_HASH | CryptConst.ALG_TYPE_ANY | CryptConst.ALG_SID_HASH_REPLACE_OWF)
static CALG_AES_128 := (CryptConst.ALG_CLASS_DATA_ENCRYPT|CryptConst.ALG_TYPE_BLOCK|CryptConst.ALG_SID_AES_128)
static CALG_AES_192 := (CryptConst.ALG_CLASS_DATA_ENCRYPT|CryptConst.ALG_TYPE_BLOCK|CryptConst.ALG_SID_AES_192)
static CALG_AES_256 := (CryptConst.ALG_CLASS_DATA_ENCRYPT|CryptConst.ALG_TYPE_BLOCK|CryptConst.ALG_SID_AES_256)
static CALG_AES := (CryptConst.ALG_CLASS_DATA_ENCRYPT|CryptConst.ALG_TYPE_BLOCK|CryptConst.ALG_SID_AES)
static CALG_SHA_256 := (CryptConst.ALG_CLASS_HASH | CryptConst.ALG_TYPE_ANY | CryptConst.ALG_SID_SHA_256)
static CALG_SHA_384 := (CryptConst.ALG_CLASS_HASH | CryptConst.ALG_TYPE_ANY | CryptConst.ALG_SID_SHA_384)
static CALG_SHA_512 := (CryptConst.ALG_CLASS_HASH | CryptConst.ALG_TYPE_ANY | CryptConst.ALG_SID_SHA_512)
static CALG_ECDH := (CryptConst.ALG_CLASS_KEY_EXCHANGE | CryptConst.ALG_TYPE_DH | CryptConst.ALG_SID_ECDH)
static CALG_ECMQV := (CryptConst.ALG_CLASS_KEY_EXCHANGE | CryptConst.ALG_TYPE_ANY | CryptConst.ALG_SID_ECMQV)
static CALG_ECDSA := (CryptConst.ALG_CLASS_SIGNATURE | CryptConst.ALG_TYPE_DSS | CryptConst.ALG_SID_ECDSA)
static CRYPT_VERIFYCONTEXT := 0xF0000000
static CRYPT_NEWKEYSET := 0x00000008
static CRYPT_DELETEKEYSET := 0x00000010
static CRYPT_MACHINE_KEYSET := 0x00000020
static CRYPT_SILENT := 0x00000040
static CRYPT_DEFAULT_CONTAINER_OPTIONAL := 0x00000080
static CRYPT_EXPORTABLE := 0x00000001
static CRYPT_USER_PROTECTED := 0x00000002
static CRYPT_CREATE_SALT := 0x00000004
static CRYPT_UPDATE_KEY := 0x00000008
static CRYPT_NO_SALT := 0x00000010
static CRYPT_PREGEN := 0x00000040
static CRYPT_RECIPIENT := 0x00000010
static CRYPT_INITIATOR := 0x00000040
static CRYPT_ONLINE := 0x00000080
static CRYPT_SF := 0x00000100
static CRYPT_CREATE_IV := 0x00000200
static CRYPT_KEK := 0x00000400
static CRYPT_DATA_KEY := 0x00000800
static CRYPT_VOLATILE := 0x00001000
static CRYPT_SGCKEY := 0x00002000
static CRYPT_ARCHIVABLE := 0x00004000
static CRYPT_FORCE_KEY_PROTECTION_HIGH := 0x00008000
static RSA1024BIT_KEY := 0x04000000
static CRYPT_SERVER := 0x00000400
static KEY_LENGTH_MASK := 0xFFFF0000
static CRYPT_Y_ONLY := 0x00000001
static CRYPT_SSL2_FALLBACK := 0x00000002
static CRYPT_DESTROYKEY := 0x00000004
static CRYPT_OAEP := 0x00000040
static CRYPT_BLOB_VER3 := 0x00000080
static CRYPT_IPSEC_HMAC_KEY := 0x00000100
static CRYPT_DECRYPT_RSA_NO_PADDING_CHECK := 0x00000020
static CRYPT_SECRETDIGEST := 0x00000001
static CRYPT_OWF_REPL_LM_HASH := 0x00000001
static CRYPT_LITTLE_ENDIAN := 0x00000001
static CRYPT_NOHASHOID := 0x00000001
static CRYPT_TYPE2_FORMAT := 0x00000002
static CRYPT_X931_FORMAT := 0x00000004
static CRYPT_MACHINE_DEFAULT := 0x00000001
static CRYPT_USER_DEFAULT := 0x00000002
static CRYPT_DELETE_DEFAULT := 0x00000004
static SIMPLEBLOB := 0x1
static PUBLICKEYBLOB := 0x6
static PRIVATEKEYBLOB := 0x7
static PLAINTEXTKEYBLOB := 0x8
static OPAQUEKEYBLOB := 0x9
static PUBLICKEYBLOBEX := 0xA
static SYMMETRICWRAPKEYBLOB := 0xB
static KEYSTATEBLOB := 0xC
static AT_KEYEXCHANGE := 1
static AT_SIGNATURE := 2
static CRYPT_USERDATA := 1
static KP_IV := 1
static KP_SALT := 2
static KP_PADDING := 3
static KP_MODE := 4
static KP_MODE_BITS := 5
static KP_PERMISSIONS := 6
static KP_ALGID := 7
static KP_BLOCKLEN := 8
static KP_KEYLEN := 9
static KP_SALT_EX := 10
static KP_P := 11
static KP_G := 12
static KP_Q := 13
static KP_X := 14
static KP_Y := 15
static KP_RA := 16
static KP_RB := 17
static KP_INFO := 18
static KP_EFFECTIVE_KEYLEN := 19
static KP_SCHANNEL_ALG := 20
static KP_CLIENT_RANDOM := 21
static KP_SERVER_RANDOM := 22
static KP_RP := 23
static KP_PRECOMP_MD5 := 24
static KP_PRECOMP_SHA := 25
static KP_CERTIFICATE := 26
static KP_CLEAR_KEY := 27
static KP_PUB_EX_LEN := 28
static KP_PUB_EX_VAL := 29
static KP_KEYVAL := 30
static KP_ADMIN_PIN := 31
static KP_KEYEXCHANGE_PIN := 32
static KP_SIGNATURE_PIN := 33
static KP_PREHASH := 34
static KP_ROUNDS := 35
static KP_OAEP_PARAMS := 36
static KP_CMS_KEY_INFO := 37
static KP_CMS_DH_KEY_INFO := 38
static KP_PUB_PARAMS := 39
static KP_VERIFY_PARAMS := 40
static KP_HIGHEST_VERSION := 41
static KP_GET_USE_COUNT := 42
static KP_PIN_ID := 43
static KP_PIN_INFO := 44
static HP_ALGID := 0x0001
static HP_HASHVAL := 0x0002
static HP_HASHSIZE := 0x0004
static HP_HMAC_INFO := 0x0005
static HP_TLS1PRF_LABEL := 0x0006
static HP_TLS1PRF_SEED := 0x0007
static PROV_RSA_FULL := 1
static PROV_RSA_SIG := 2
static PROV_DSS := 3
static PROV_FORTEZZA := 4
static PROV_MS_EXCHANGE := 5
static PROV_SSL := 6
static PROV_RSA_SCHANNEL := 12
static PROV_DSS_DH := 13
static PROV_EC_ECDSA_SIG := 14
static PROV_EC_ECNRA_SIG := 15
static PROV_EC_ECDSA_FULL := 16
static PROV_EC_ECNRA_FULL := 17
static PROV_DH_SCHANNEL := 18
static PROV_SPYRUS_LYNKS := 20
static PROV_RNG := 21
static PROV_INTEL_SEC := 22
static PROV_REPLACE_OWF := 23
static PROV_RSA_AES := 24
static PROV_STT_MER := 7
static PROV_STT_ACQ := 8
static PROV_STT_BRND := 9
static PROV_STT_ROOT := 10
static PROV_STT_ISS := 11
}
b64Encode( ByRef buf, bufLen )
{
DllCall( "crypt32\CryptBinaryToStringA", "ptr", &buf, "UInt", bufLen, "Uint", 1 | 0x40000000, "Ptr", 0, "UInt*", outLen )
VarSetCapacity( outBuf, outLen, 0 )
DllCall( "crypt32\CryptBinaryToStringA", "ptr", &buf, "UInt", bufLen, "Uint", 1 | 0x40000000, "Ptr", &outBuf, "UInt*", outLen )
return strget( &outBuf, outLen, "CP0" )
}
b64Decode( b64str, ByRef outBuf )
{
static CryptStringToBinary := "crypt32\CryptStringToBinary" (A_IsUnicode ? "W" : "A")
DllCall( CryptStringToBinary, "ptr", &b64str, "UInt", 0, "Uint", 1, "Ptr", 0, "UInt*", outLen, "ptr", 0, "ptr", 0 )
VarSetCapacity( outBuf, outLen, 0 )
DllCall( CryptStringToBinary, "ptr", &b64str, "UInt", 0, "Uint", 1, "Ptr", &outBuf, "UInt*", outLen, "ptr", 0, "ptr", 0 )
return outLen
}
b2a_hex( ByRef pbData, dwLen )
{
if (dwLen < 1)
return 0
if pbData is integer
ptr := pbData
else
ptr := &pbData
SetFormat,integer,Hex
loop,%dwLen%
{
num := numget(ptr+0,A_index-1,"UChar")
hash .= substr((num >> 4),0) . substr((num & 0xf),0)
}
SetFormat,integer,D
StringLower,hash,hash
return hash
}
a2b_hex( sHash,ByRef ByteBuf )
{
if (sHash == "" || RegExMatch(sHash,"[^\dABCDEFabcdef]") || mod(StrLen(sHash),2))
return 0
BufLen := StrLen(sHash)/2
VarSetCapacity(ByteBuf,BufLen,0)
loop,%BufLen%
{
num1 := (p := "0x" . SubStr(sHash,(A_Index-1)*2+1,1)) << 4
num2 := "0x" . SubStr(sHash,(A_Index-1)*2+2,1)
num := num1 | num2
NumPut(num,ByteBuf,A_Index-1,"UChar")
}
return BufLen
}
Free(byRef var)
{
VarSetCapacity(var,0)
return
}
;}

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

