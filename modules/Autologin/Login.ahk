login(){
	if PWB := IEGET("Service Centre 5 Login"){
		pwb.document.getElementById("txtUserName").value := settings.Username
		pwb.document.getElementById("txtPassword").value := settings.Password
		pwb.document.getElementsByTagName("IMG")[7].click
		pwb:=""
	}
}
	

