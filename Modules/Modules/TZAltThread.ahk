; <COMPILER: v1.1.16.05>
#NoTrayIcon
#SingleInstance, Force
#Persistent
SetTimer,msgboxclose,100
SetTimer,IEINterupt,100
return
msgboxclose:
IfWinExist,Message from webpage
{
Winclose,Message from webpage
ExitApp
}
return
IEINterupt:
IfWinExist,Popup List -- Webpage Dialog
{
pdoc := Get_MODAL_DOCUMENT()
while (pdoc.readystate != "complete")
sleep, 100
if (pdoc.getElementByID("grdDropdown:_ctl4:_ctl0").Innertext = "")
{
pdoc.getElementByID("grdDropdown:_ctl3:_ctl0").click
}
exitapp
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