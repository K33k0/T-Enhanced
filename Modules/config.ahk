#Include Lib/Functions.ahk
class config {
	
	static Engineer
	static WorkshopSite
	static HashedUserName
	static HashedPassword
	static BenchKit
	
	__New(ini) {
		IniRead, Engineer, %ini%, Enginer, Number
		if (Engineer = "ERROR")
			return false
		IniRead, user, %ini%, login, UserName
		IniRead, pass, %ini%, login, Password
		IniRead,wSite, %ini%, Site, location
		if (wSite = "ERROR") {
			return False
		}
		this.HashedUserName := user
		this.HashedPassword := pass
		this.BenchKit := BenchKit . "BK"
		this.Engineer := Engineer
		this.workshopSite := wSite
	}
	
	decrypt(setting) {
		keys := this.BenchKit
		if (setting = "password")	{
			return Crypt.Encrypt.StrDecrypt(this.HashedPassword,keys,5,1)	
		} else if ( setting = "username" )	{
			return Crypt.Encrypt.StrDecrypt(this.HashedUserName,keys,5,1)	
		} 	else {
			return "failed to find setting"
		}
	}
	
	encrypt(setting) {
		keys := this.BenchKit
		if (setting = "password")	{
			return Crypt.Encrypt.StrEncrypt(this.HashedPassword,keys,5,1)	
		} else if ( setting = "username" )	{
			return Crypt.Encrypt.StrEncrypt(this.HashedUserName,keys,5,1)	
		} 	else {
			return "failed to find setting"
		}
	}

	setup() {
		gui, new
		gui, add, text,, Select your workshop site
		gui, add, ddl,vWorkshopSite, NSC
		gui, add, text,, Insert your engineer number
		gui, add, edit, vEngineer,
		gui, add,text,, insert your username
		gui, add,edit, vUsername
		gui, add,text,, insert your password
		gui, add,edit, vPassword Password,
		gui,show
		WinWaitClose
		return
		
		
	}
}

if not config := new config("config.ini") {
	msgbox, Config not set
}


