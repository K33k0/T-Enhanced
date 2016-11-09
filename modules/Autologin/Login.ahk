login(config){
	IniRead,User,%config%,Default,Username
	IniRead,Pass,%config%,Default,Password
	if PWB := IEGET("Service Centre 5 Login"){
		pwb.document.getElementById("txtUserName").value := User
		pwb.document.getElementById("txtPassword").value := Pass
		pwb.document.getElementsByTagName("IMG")[7].click
		pwb:=""
	}
}
	

