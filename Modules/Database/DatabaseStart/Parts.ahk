PartList(Manufacturer ="", Type="") {
	iniread,ManufacturerTypes,partList.ini,%Manufacturer%
	ManufacturerTypesNew := RegExReplace(ManufacturerTypes, ")\=(.*)")
	MsgBox % ManufacturerTypesNew
}

MsgBox % PartList("IBM")
	