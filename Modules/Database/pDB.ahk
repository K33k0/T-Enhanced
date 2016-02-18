iniread, Products, PartsDatabase.ini,EligibleProducts,List
FileAppend,IBM:= [,NewDB.txt

StringSplit, ProductArray, Products, "|",
Loop, %ProductArray0%
{
	This_Product:= ProductArray%A_Index%
	msgbox,4,select,Is %This_Product% IBM?
	IfMsgBox,Yes 
	{
	FileAppend,"%This_Product%"`,,NewDB.txt
	}
}

FileAppend,],NewDB.txt