#Include Lib/Functions.ahk
class config {
	
	static Engineer
	static WorkshopSite
	static HashedUserName
	static HashedPassword
	static BenchKit
	static Firstrun = False
	static configGui
	
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
	
	class Gui
	{
		static GuiID
		
		__New(){
			gui, new, +LastFound +hwndConfigGui
			this.guiID := ConfigGui
			if (config.firstrun){
				this.firstRun()
			}
		}
		
		firstRun(){
			static myDDL
			guiID := this.GuiID
			gui, %GuiID%:add, text,, First Run
			gui, %GuiID%:add, ddl, vmyDDL, one||two
			gui, %GuiID%:show
			this.submitGui()
		}
		
		submitGui(){
			guiID := this.GuiID
			ControlGetText, val,
			msgbox % this.firstrun.myDDL
		}
	}
}

config := new config("config.ini")
if (config.FirstRun){
	configGui := new config.Gui()
}
	return
		gui, config:new, +hwndConfig
		gui, config:add, text,, Select your workshop site
		gui, config:add, ddl,hwndWorkshopSite, NSC
		gui, config:add, text,, Insert your engineer number
		gui, config:add, edit, vEngineer,
		gui, config:add,text,, insert your username
		gui, config:add,edit, vUsername
		gui, config:add,text,, insert your password
		gui, config:add,edit, vPassword Password,
		gui, config:add, button, hwndhDone, Done
		gui,config:show