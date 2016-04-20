
!1::
;{ check stats
if  pwb:=IEgeturl("http://hypappbs005/SC5/SC_RepairJob/aspx/RepairJob_frameset.aspx") {
msgbox, this function is not available on this page
pwb:=""
return
}
FormatTime,Today, YYYYMMDD, dd/MM/yyyy
pwb:=IEVget(Title)
pwb.Navigate2("http://hypappbs005/SC5/SC_RepairJob/aspx/RepairJob_frameset.aspx",2048)
sleep 2500
if not pwb:=IEGeturl("http://hypappbs005/SC5/SC_RepairJob/aspx/RepairJob_frameset.aspx") {
msgbox,Page not accessable
}
Loop {
try {
sleep,250
frame := Pwb.document.all(10).contentWindow
PageLoad:= frame.document.getElementsByTagName("LABEL")[0].innertext
frame := Pwb.document.all(7).contentWindow
PageLoad2:= frame.document.getElementsByTagName("Input")[0].value
}
}until (PageLoad = "Job Details" AND PageLoad2 = "submit")

try {
frame := Pwb.document.all(10).contentWindow
frame.document.getElementByID("datJobCDate").Value:= Today
frame.document.getElementByID("cboCallEmployNum").Value:= "406"
frame := Pwb.document.all(7).contentWindow
frame.document.getElementByID("cmdSubmit").click
} catch {
msgbox , failed to pull %value%
}

Loop {
try {
frame := Pwb.document.all(10).contentWindow
TotalRecords:=frame.document.getElementByID("lblRecordCount").innertext
SingleRecord:=frame.document.getElementByID("cboJobWorkshopSiteNum").value
}
}Until (TotalRecords != "" OR SingleRecord ="STOWS")

If (SingleRecord = "STOWS") {
Jobs := 1
}
else
{
Jobs := TotalRecords
}
pwb.quit
MsgBox % TotalRecords
return
;}

!2::
#include Modules/KitCheck.ahk
