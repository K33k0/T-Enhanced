#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;run, update/git-pull.exe "../../"
url:= "https://api.github.com/repos/k33k00/T-Enhanced--ZULU-/releases/latest"
 
oHTTP := ComObjCreate("WinHttp.WinHttpRequest.5.1")
;Post request
oHTTP.Open("GET", URL , False)

oHTTP.Send(PostData)
;Get received data
Gui, Add, Edit, w800 r30, % oHTTP.ResponseText
Gui, Show
return
 
GuiClose:
ExitApp