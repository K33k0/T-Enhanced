Gui Add, tab2, w940 h660 vURL, MainTab|BackgroundWorker
gui,Tab,MainTab
Gui Add, ActiveX, xm+10 yp+22 h630 w925 vWB, Shell.Explorer
gui,Tab,BackGroundWorker
Gui Add, ActiveX, xm+10 yp h630 w925 vbWB, Shell.Explorer

ComObjConnect(WB, WB_events)  ; Connect WB's events to the WB_events class object.
ComObjConnect(bWB, WB_events)  ; Connect WB's events to the WB_events class object.
WB.Navigate2("http://hypappbs005/SC5/SC_Login/aspx/login_launch.aspx?SOURCE=ESOLBRANCHLIVE")
Gui Show
; Continue on to load the initial page:
ButtonGo:
Gui Submit, NoHide
return

class WB_events
{
    NavigateComplete2(wb, NewURL)
    {
		msgbox, Navigation Complete
    }
}

GuiClose:
ExitApp