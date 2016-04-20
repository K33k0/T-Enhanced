class config {
	static ini
	static Engineer
	static WorkshopSite
	static HashedUserName
	static HashedPassword
	static BenchKit
	static Firstrun = False
	global TesseractVersion := "5.40.14"
	
	__New(ini) {
		this.ini := ini
		IniRead, Engineer, %ini%, Engineer, Number
		IniRead, user, %ini%, login, UserName
		IniRead, pass, %ini%, login, Password
		IniRead,wSite, %ini%, Site, location
		this.HashedUserName := user
		this.HashedPassword := pass
		this.BenchKit := Engineer . "BK"
		this.Engineer := Engineer
		this.workshopSite := wSite
	}
	
	decrypt(setting) {
		keys := this.Engineer
		if (setting = "password")	{
			return Crypt.Encrypt.StrDecrypt(this.HashedPassword,keys)	
		} else if ( setting = "username" )	{
			return Crypt.Encrypt.StrDecrypt(this.HashedUserName,keys)	
		} 	else {
			return "failed to find setting"
		}
	}
	
	encrypt(setting, value) {
		keys := this.Engineer
		if (setting = "password")	{
			return Crypt.Encrypt.StrEncrypt(value,keys)	
		} else if ( setting = "username" )	{
			return Crypt.Encrypt.StrEncrypt(value,keys)	
		} 	else {
			return false
		}
		return true
	}
	
	save(Engineer,WorkshopSite,UserName="",Password="") {
		IniWrite, %Engineer%, % this.ini, Engineer, Number
		this.engineer := Engineer
		IniWrite, %WorkshopSite%, % this.ini, Site, location
		if (userName) {
			IniWrite, % this.encrypt("username",UserName), % this.ini, login, UserName
		}
		if (password) {
			IniWrite, % this.encrypt("password",password), % this.ini, login, password 
		}
	}
}

settings := new config(A_ScriptDir "\Modules\Config.ini")
