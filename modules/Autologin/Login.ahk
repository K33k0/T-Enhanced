login(config){
	IniRead,User,%config%,Login,Username
	IniRead,Pass,%config%,Login,Password
	msgbox, %config%
	if PWB := IEGET("Service Centre 5 Login"){
		pwb.document.getElementById("txtUserName").value := User
		pwb.document.getElementById("txtPassword").value := Pass
		pwb.document.getElementsByTagName("IMG")[7].click
		pwb:=""
	}
}
	

