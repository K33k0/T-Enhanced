class AutoLogin extends TEnhanced {
	__New(settings,Tab){
		OutputDebug,[T-Enhanced]  Quick login started
		if (A_GuiControl = "Tab"){
			if (Tab = "Engineer") {
			IniRead,UserHash,%Config%,Login,UserName
			IniRead,PassHash,%Config%,Login,Password
			If (UserHash = "" OR UserHash = "Error"){
				return
			}
			if not PWB:= IEGET("Service Centre 5 Login") {
				pwb:=""
				return
			} else {
				pwb.document.getElementById("txtUserName").value := settings.decrypt("username")
				pwb.document.getElementById("txtPassword").value := settings.decrypt("password")
				pwb.document.getElementsByTagName("IMG")[7].click
				pwb:=""
			}
		}
	}
		OutputDebug,[T-Enhanced]  Quick login ended
	}
}