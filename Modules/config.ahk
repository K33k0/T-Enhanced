#Include Lib/Functions.ahk
class config {
	
	static Engineer
	static WorkshopSite
	static HashedUserName
	static HashedPassword
	static BenchKit
	static Firstrun = False
	
	__New(ini) {
		IniRead, Engineer, %ini%, Enginer, Number
		IniRead, user, %ini%, login, UserName
		IniRead, pass, %ini%, login, Password
		IniRead,wSite, %ini%, Site, location
		this.HashedUserName := user
		this.HashedPassword := pass
		this.BenchKit := BenchKit . "BK"
		this.Engineer := Engineer
		this.workshopSite := wSite
		if (Engineer = "ERROR"){
			this.FirstRun := True
		}
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
		msgbox, requires configuration
		;~ gui, new
		;~ gui, add, text,, Select your workshop site
		;~ gui, add, ddl,vWorkshopSite, NSC
		;~ gui, add, text,, Insert your engineer number
		;~ gui, add, edit, vEngineer,
		;~ gui, add,text,, insert your username
		;~ gui, add,edit, vUsername
		;~ gui, add,text,, insert your password
		;~ gui, add,edit, vPassword Password,
		;~ gui, add, button, hwndDone, Done
		;~ gui,show
	}
}

config := new config("config.ini")
if (config.FirstRun){
	msgbox, test
	config.setup()
}
	
