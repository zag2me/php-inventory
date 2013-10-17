On Error Resume Next

Dim colSettings, objWMIService, strComputer
Dim MACAddressSet, NetworkAdaptSet, IPAddressSet, ComputerNameSet, GUIDSet, RAMMemorySet, sIDsET, RAMMemoryBankLabelSet
Dim RAMMemoryCapacitySet, TotalRAMMemoryDeviceLocatorSet, RAMMemoryManufacturerSet, ComputerManufacturerSet
Dim ComputerModelSet, ComputerSerialSet, MotherboardManufacturerSet, MotherboardModelSet, MotherboardProductSet
Dim MotherboardSerialNumberSet, MotherboardVersionSet, HardDiskDriveLetterSet, HardDiskSizeSet, HardDiskFreeSpaceSet
Dim HardDiskVolumeSerialNumberSet, HardDiskManufacturerSet, HardDiskModelSet, HardDiskDeviceIDSet, HardDiskInterfaceTypeSet
Dim HardDiskPartitionsSet, HardDiskTotalCylindersSet, HardDiskTotalHeadsSet, HardDiskTotalSectorsSet, HardDiskTotalTracksSet
Dim HardDiskTracksPerCylinderSet, ProcessorNameSet, ProcessorTypeSet, ProcessorFamilySet, SoundCardManufacturerSet
Dim SoundCardProductNameSet, SoundCardDescriptionSet, SoundCardNameSet, VideoCardNameSet, VideoCardSettingIDSet 
Dim FloppyManufacturerSet, FloppyNameSet, FloppyDeviceIDSet, CDROMCaptionSet, CDROMDescriptionSet, CDROMDeviceIDSet
Dim CDROMDriveSet, CDROMManufacturerSet, CDROMNameSet, BIOSVersionSet, ISDellComputerSet, DellServiceTagSet
Dim filepath, filesys, filetxt, FSO, maccheck, inventorystats, countitem, MacConvertionGUID, OUpart, MacConvert, FullMacConvert
Dim countitemb, readfullline
Dim macread, netcardread, ipread, comnameread, FieldProp

'// Settings
filepath = "\\server\scripts$\inventorylog.csv"
strComputer = "."
ISDellComputerSet = "NOTDELL"
inventorystats = "COMPLETE"
'// End Settings

Set FSO = CreateObject("Scripting.FileSystemObject")
Set filesys = CreateObject("Scripting.FileSystemObject")

MACAddressSet = "N/A"
NetworkAdaptSet = "N/A"
IPAddressSet = "N/A"
ComputerNameSet = "N/A"
GUIDSet = "Unknown"
MacConvertionGUID = "N/A"
TotalRAMMemorySet = "N/A"
SidSet = "N/A"
RAMMemoryBankLabelSet = ""
RAMMemoryCapacitySet = ""
RAMMemoryDeviceLocatorSet = ""
RAMMemoryManufacturerSet = ""
ComputerManufacturerSet = "N/A"
ComputerModelSet = "N/A"
ComputerSerialSet = "N/A"
MotherboardManufacturerSet = "N/A"
MotherboardModelSet = "N/A"
MotherboardProductSet = "N/A"
MotherboardSerialNumberSet = "N/A"
MotherboardVersionSet = "N/A"
HardDiskDriveLetterSet = ""
HardDiskSizeSet = ""
HardDiskFreeSpaceSet = ""
HardDiskVolumeSerialNumberSet = ""
HardDiskManufacturerSet = ""
HardDiskModelSet = ""
HardDiskDeviceIDSet = ""
HardDiskInterfaceTypeSet = ""
HardDiskPartitionsSet = ""
HardDiskTotalCylindersSet = ""
HardDiskTotalHeadsSet = ""
HardDiskTotalSectorsSet = ""
HardDiskTotalTracksSet = ""
HardDiskTracksPerCylinderSet = ""
ProcessorNameSet = ""
ProcessorTypeSet = ""
ProcessorFamilySet = ""
SoundCardManufacturerSet = "N/A"
SoundCardProductNameSet = "N/A"
SoundCardDescriptionSet = "N/A"
SoundCardNameSet = "N/A"
VideoCardNameSet = "N/A"
VideoCardSettingIDSet = "N/A"
FloppyManufacturerSet = "N/A"
FloppyNameSet = "N/A"
FloppyDeviceIDSet = "N/A"
CDROMCaptionSet = ""
CDROMDescriptionSet = ""
CDROMDeviceIDSet = ""
CDROMDriveSet = ""
CDROMManufacturerSet = ""
CDROMNameSet = ""
BIOSVersionSet = "N/A"
DellServiceTagSet = "N/A"

If NOT FSO.FileExists (filepath) Then
	call writefiletitles
End If

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapter")

	For Each objItem in colItems
		If objItem.AdapterType = "Ethernet 802.3" AND objItem.Description <> "Packet Scheduler Miniport" AND objItem.Description <> "Virtual Machine Network Services Driver" AND objItem.Description <> "1394 Net Adapter" AND objItem.Description <> "Bluetooth Device (Personal Area Network)" Then
			MACAddressSet = objItem.MACAddress
			NetworkAdaptSet = objItem.Description
		End If
	Next

NetworkAdaptSet = CheckCommas(NetworkAdaptSet)

Set colItems = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
For Each IPConfig in colItems
            IPAddressSet = IPConfig.IPAddress(i)
Next

IPAddressSet = CheckCommas(IPAddressSet)

Set colItems = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
	For Each objComputer in colItems
		ComputerNameSet = objComputer.Name
		GUIDSet = GetUUID
		TotalRAMMemorySet = objComputer.TotalPhysicalMemory
		SidSet = GetSid
	Next

call checkthemac

OUpart = "a"
FullMacConvert = ""
countitem = 0
MacConvert = MACAddressSet & ":"
Do until OUpart = ""
	OUpart = split(MacConvert, ":")
		IF OUpart(countitem) = ""  THEN Exit Do
		OUpart = "a"
		countitem = countitem + 1
		Loop
	IF countitem >=1 THEN
		countitemb = 0

	Do until countitemb = countitem

	IF countitemb < countitem AND countitemb <> countitem-1 THEN  FullMacConvert = FullMacConvert + OUpart(countitemb)
	IF countitemb = countitem-1 THEN FullMacConvert = FullMacConvert + OUpart(countitemb)

	countitemb = countitemb + 1

Loop
	END IF

MacConvertionGUID = "00000000-0000-0000-0000-" & FullMacConvert

ComputerNameSet = CheckCommas(ComputerNameSet)
GUIDSet = CheckCommas(GUIDSet)
TotalRAMMemorySet = CheckCommas(TotalRAMMemorySet)
SidSet = CheckCommas(SidSet)
MacConvertionGUID = CheckCommas(MacConvertionGUID)

countitem = 0
Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")
	For Each objItem in colItems
		countitem = countitem + 1
		RAMMemoryBankLabelSet = RAMMemoryBankLabelSet & countitem & ":" & objItem.BankLabel & " "
		RAMMemoryCapacitySet = RAMMemoryCapacitySet & countitem & ":" & objItem.Capacity & " "
		RAMMemoryDeviceLocatorSet = RAMMemoryDeviceLocatorSet & countitem & ":" & objItem.DeviceLocator & " "
		RAMMemoryManufacturerSet = RAMMemoryManufacturerSet & countitem & ":" & objItem.Manufacturer & " "
	Next

RAMMemoryBankLabelSet = CheckCommas(RAMMemoryBankLabelSet)
RAMMemoryCapacitySet = CheckCommas(RAMMemoryCapacitySet)
RAMMemoryDeviceLocatorSet = CheckCommas(RAMMemoryDeviceLocatorSet)
RAMMemoryManufacturerSet = CheckCommas(RAMMemoryManufacturerSet)

Set colItems = objWMIService.ExecQuery ("Select * from Win32_ComputerSystemProduct")
	For Each objItem In colItems
		ComputerManufacturerSet = objItem.Vendor
		ComputerModelSet = objItem.Name
		ComputerSerialSet = objItem.IdentifyingNumber
		ISDellComputerSet = uCASE(Left(objItem.Vendor, 4))
	Next

ComputerManufacturerSet = CheckCommas(ComputerManufacturerSet)
ComputerModelSet = CheckCommas(ComputerModelSet)
ComputerSerialSet = CheckCommas(ComputerSerialSet)

Set colItems = objWMIService.ExecQuery ("Select * from Win32_BaseBoard") 

	For Each objItem in colItems 
		MotherboardManufacturerSet = objItem.Manufacturer 
		MotherboardModelSet = objItem.Model 
		MotherboardProductSet = objItem.Product 
		MotherboardSerialNumberSet = objItem.SerialNumber 
		MotherboardVersionSet = objItem.Version 
	Next

MotherboardManufacturerSet = CheckCommas(MotherboardManufacturerSet)
MotherboardModelSet = CheckCommas(MotherboardModelSet)
MotherboardProductSet = CheckCommas(MotherboardProductSet)
MotherboardSerialNumberSet = CheckCommas(MotherboardSerialNumberSet)
MotherboardVersionSet = CheckCommas(MotherboardVersionSet)

countitem = 0
Set colItems = objWMIService.ExecQuery ("Select * from Win32_LogicalDisk")
	For Each objItem in colItems
		If objItem.Name = "C:" Then
			countitem = countitem + 1
			HardDiskDriveLetterSet =  HardDiskDriveLetterSet & countitem & ":" & objItem.Name & " "
			HardDiskSizeSet =  HardDiskSizeSet & countitem & ":" & Int(objItem.Size /1073741824) & " GB" & " "
			HardDiskFreeSpaceSet =  HardDiskFreeSpaceSet & countitem & ":" & Int(objItem.FreeSpace /1073741824) & " GB" & " "
			HardDiskVolumeSerialNumberSet =  HardDiskVolumeSerialNumberSet & countitem & ":" & objItem.VolumeSerialNumber & " "
		End If
	Next

HardDiskDriveLetterSet = CheckCommas(HardDiskDriveLetterSet)
HardDiskSizeSet = CheckCommas(HardDiskSizeSet)
HardDiskFreeSpaceSet = CheckCommas(HardDiskFreeSpaceSet)
HardDiskVolumeSerialNumberSet = CheckCommas(HardDiskVolumeSerialNumberSet)

countitem = 0
Set colItems = objWMIService.ExecQuery ("Select * from Win32_DiskDrive")
	For each objDiskDrive in colItems
			countitem = countitem + 1
			HardDiskManufacturerSet = HardDiskManufacturerSet & countitem & ":" & objDiskDrive.Manufacturer & " "
			HardDiskModelSet = HardDiskModelSet & countitem & ":" & objDiskDrive.Model & " "
			HardDiskDeviceIDSet = HardDiskDeviceIDSet & countitem & ":" & objDiskDrive.DeviceID & " "
			HardDiskInterfaceTypeSet = HardDiskInterfaceTypeSet & countitem & ": " & objDiskDrive.InterfaceType & " "
			HardDiskPartitionsSet = HardDiskPartitionsSet & countitem & ":" & objDiskDrive.Partitions & " "
			HardDiskTotalCylindersSet = HardDiskTotalCylindersSet & countitem & ":" & objDiskDrive.TotalCylinders & " "
			HardDiskTotalHeadsSet = HardDiskTotalHeadsSet & countitem & ":" & objDiskDrive.TotalHeads & " "
			HardDiskTotalSectorsSet = HardDiskTotalSectorsSet & countitem & ":" & objDiskDrive.TotalSectors & " "
			HardDiskTotalTracksSet = HardDiskTotalTracksSet & countitem & ":" & objDiskDrive.TotalTracks & " "
			HardDiskTracksPerCylinderSet = HardDiskTracksPerCylinderSet & countitem & ":" & objDiskDrive.TracksPerCylinder & " "
	Next

HardDiskManufacturerSet = CheckCommas(HardDiskManufacturerSet)
HardDiskModelSet = CheckCommas(HardDiskModelSet)
HardDiskDeviceIDSet = CheckCommas(HardDiskDeviceIDSet)
HardDiskInterfaceTypeSet = CheckCommas(HardDiskInterfaceTypeSet)
HardDiskPartitionsSet = CheckCommas(HardDiskPartitionsSet)
HardDiskTotalCylindersSet = CheckCommas(HardDiskTotalCylindersSet)
HardDiskTotalHeadsSet = CheckCommas(HardDiskTotalHeadsSet)
HardDiskTotalSectorsSet = CheckCommas(HardDiskTotalSectorsSet)
HardDiskTotalTracksSet = CheckCommas(HardDiskTotalTracksSet)
HardDiskTracksPerCylinderSet = CheckCommas(HardDiskTracksPerCylinderSet)


countitem = 0
Set colItems = objWMIService.ExecQuery ("Select * from Win32_Processor")
	For Each objItem in colItems
		countitem = countitem + 1
		ProcessorNameSet = ProcessorNameSet & countitem & ":" & objItem.Name & " "
	Next

ProcessorNameSet = CheckCommas(ProcessorNameSet)

intProc = 1
countitem = 0
For Each objProcessor in colItems
	countitem = countitem + 1
	ProcessorTypeSet = ProcessorTypeSet & countitem & ":" & objProcessor.Architecture & " "
	ProcessorFamilySet = ProcessorFamilySet & countitem & ":" & objProcessor.Description & " "
	intProc = intProc + 1
Next

ProcessorTypeSet = CheckCommas(ProcessorTypeSet)
ProcessorFamilySet = CheckCommas(ProcessorFamilySet)

Set colItems = objWMIService.ExecQuery ("Select * from Win32_SoundDevice")
	For Each objItem in colItems
		SoundCardManufacturerSet = objItem.Manufacturer
		SoundCardProductNameSet = objItem.ProductName
		SoundCardDescriptionSet = objItem.Description
		SoundCardNameSet = objItem.Name
	Next

SoundCardManufacturerSet = CheckCommas(SoundCardManufacturerSet)
SoundCardProductNameSet = CheckCommas(SoundCardProductNameSet)
SoundCardDescriptionSet = CheckCommas(SoundCardDescriptionSet)
SoundCardNameSet = CheckCommas(SoundCardNameSet)

Set colItems = objWMIService.ExecQuery ("Select * from Win32_DisplayControllerConfiguration")
	For Each objItem in colItems
		VideoCardNameSet = objItem.Name
		VideoCardSettingIDSet = objItem.SettingID
	Next

VideoCardNameSet = CheckCommas(VideoCardNameSet)
VideoCardSettingIDSet = CheckCommas(VideoCardSettingIDSet)

Set colItems = objWMIService.ExecQuery ("Select * from Win32_FloppyDrive")
	For Each objItem in colItems
		FloppyManufacturerSet = objItem.Manufacturer
		FloppyNameSet = objItem.Name
		FloppyDeviceIDSet = objItem.DeviceID
	Next

FloppyManufacturerSet = CheckCommas(FloppyManufacturerSet)
FloppyNameSet = CheckCommas(FloppyNameSet)
FloppyDeviceIDSet = CheckCommas(FloppyDeviceIDSet)

countitem = 0
Set colItems = objWMIService.ExecQuery("Select * from Win32_CDROMDrive")
	For Each objItem in colItems
		If objItem.Caption <> "Avantis Open CDMVD001 SCSI CdRom Device" Then
			CDROMCaptionSet = CDROMCaptionSet & countitem & ":" & objItem.Caption & " "
			CDROMDescriptionSet = CDROMDescriptionSet & countitem & ":" & objItem.Description & " "
			CDROMDeviceIDSet = CDROMDeviceIDSet & countitem & ":" & objItem.DeviceID & " "
			CDROMDriveSet = CDROMDriveSet & countitem & ":" & objItem.Drive & " "
			CDROMManufacturerSet = CDROMManufacturerSet & countitem & ":" & objItem.Manufacturer & " "
			CDROMNameSet = CDROMNameSet & countitem & ":" & objItem.Name & " "
		End If
	Next

CDROMCaptionSet = CheckCommas(CDROMCaptionSet)
CDROMDescriptionSet = CheckCommas(CDROMDescriptionSet)
CDROMDeviceIDSet = CheckCommas(CDROMDeviceIDSet)
CDROMDriveSet = CheckCommas(CDROMDriveSet)
CDROMManufacturerSet = CheckCommas(CDROMManufacturerSet)
CDROMNameSet = CheckCommas(CDROMNameSet)

Set colSettings = objWMIService.ExecQuery ("Select * from Win32_BIOS")
	For Each objBIOS in colSettings 
   		BIOSVersionSet = objBIOS.Version
	Next

BIOSVersionSet = CheckCommas(BIOSVersionSet)

If ISDellComputerSet = "DELL" Then
	Set colItems = objWMIService.ExecQuery ("Select * from Win32_BIOS",,48)
		For Each objItem In colItems
			DellServiceTagSet = objItem.SerialNumber
		Next
End If

DellServiceTagSet = CheckCommas(DellServiceTagSet)

call checkfields

':::::::::::::::::::::::::
':: Write Inventory Log ::
':::::::::::::::::::::::::
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set filetxt = filesys.OpenTextFile(filepath, ForAppending, True)
filetxt.WriteLine MACAddressSet & "," & NetworkAdaptSet & "," & IPAddressSet & "," & ComputerNameSet & "," & GUIDSet & "," & MacConvertionGUID & "," & TotalRAMMemorySet & "," & SidSet & "," & RAMMemoryBankLabelSet & "," & RAMMemoryCapacitySet & "," & RAMMemoryDeviceLocatorSet & "," & RAMMemoryManufacturerSet & "," & ComputerManufacturerSet & "," & ComputerModelSet & "," & ComputerSerialSet & "," & MotherboardManufacturerSet & "," & MotherboardModelSet & "," & MotherboardProductSet & "," & MotherboardSerialNumberSet & "," & MotherboardVersionSet & "," & HardDiskDriveLetterSet & "," & HardDiskSizeSet & "," & HardDiskFreeSpaceSet & "," & HardDiskVolumeSerialNumberSet & "," & HardDiskManufacturerSet & "," & HardDiskModelSet & "," & HardDiskDeviceIDSet & "," & HardDiskInterfaceTypeSet & "," & HardDiskPartitionsSet & "," & HardDiskTotalCylindersSet & "," & HardDiskTotalHeadsSet & "," & HardDiskTotalSectorsSet & "," & HardDiskTotalTracksSet & "," & HardDiskTracksPerCylinderSet & "," & ProcessorNameSet & "," & ProcessorTypeSet & "," & ProcessorFamilySet & "," & SoundCardManufacturerSet & "," & SoundCardProductNameSet & "," & SoundCardDescriptionSet & "," & SoundCardNameSet & "," & VideoCardNameSet & "," & VideoCardSettingIDSet & "," & FloppyManufacturerSet & "," & FloppyNameSet & "," & FloppyDeviceIDSet & "," & CDROMCaptionSet & "," & CDROMDescriptionSet & "," & CDROMDeviceIDSet & "," & CDROMDriveSet & "," & CDROMManufacturerSet & "," & CDROMNameSet & "," & BIOSVersionSet & "," & DellServiceTagSet & "," & Date & "," & Time

filetxt.Close

Wscript.Quit

':::::::::::::::::::::::::::
':: Checkfields the fields::
':::::::::::::::::::::::::::
sub checkfields
	If RAMMemoryBankLabelSet = "" Then
		RAMMemoryBankLabelSet = "N/A"
	End If
	If RAMMemoryCapacitySet = "" Then
		RAMMemoryCapacitySet = "N/A"
	End If
	If RAMMemoryDeviceLocatorSet = "" Then
		RAMMemoryDeviceLocatorSet = "N/A"
	End If
	If RAMMemoryManufacturerSet = "" Then
		RAMMemoryManufacturerSet = "N/A"
	End If
	If HardDiskDriveLetterSet = "" Then
		HardDiskDriveLetterSet = "N/A"
	End If
	If HardDiskSizeSet = "" Then
		HardDiskSizeSet = "N/A"
	End If
	If HardDiskFreeSpaceSet = "" Then
		HardDiskFreeSpaceSet = "N/A"
	End If
	If HardDiskVolumeSerialNumberSet = "" Then
		HardDiskVolumeSerialNumberSet = "N/A"
	End If
	If HardDiskManufacturerSet = "" Then
		HardDiskManufacturerSet = "N/A"
	End If
	If HardDiskModelSet = "" Then
		HardDiskModelSet = "N/A"
	End If
	If HardDiskDeviceIDSet = "" Then
		HardDiskDeviceIDSet = "N/A"
	End If
	If HardDiskInterfaceTypeSet = "" Then
		HardDiskInterfaceTypeSet = "N/A"
	End If
	If HardDiskPartitionsSet = "" Then
		HardDiskPartitionsSet = "N/A"
	End If
	If HardDiskTotalCylindersSet = "" Then
		HardDiskTotalCylindersSet = "N/A"
	End If
	If HardDiskTotalHeadsSet = "" Then
		HardDiskTotalHeadsSet = "N/A"
	End If
	If HardDiskTotalSectorsSet = "" Then
		HardDiskTotalSectorsSet = "N/A"
	End If
	If HardDiskTotalTracksSet = "" Then
		HardDiskTotalTracksSet = "N/A"
	End If
	If HardDiskTracksPerCylinderSet = "" Then
		HardDiskTracksPerCylinderSet = "N/A"
	End If
	If ProcessorNameSet = "" Then
		ProcessorNameSet = "N/A"
	End If
	If ProcessorTypeSet = "" Then
		ProcessorTypeSet = "N/A"
	End If
	If ProcessorFamilySet = "" Then
		ProcessorFamilySet = "N/A"
	End If
	If CDROMCaptionSet = "" Then
		CDROMCaptionSet = "N/A"
	End If
	If CDROMDescriptionSet = "" Then
		CDROMDescriptionSet = "N/A"
	End If
	If CDROMDeviceIDSet = "" Then
		CDROMDeviceIDSet = "N/A"
	End If
	If CDROMDriveSet = "" Then
		CDROMDriveSet = "N/A"
	End If
	If CDROMManufacturerSet = "" Then
		CDROMManufacturerSet = "N/A"
	End If
	If CDROMNameSet = "" Then
		CDROMNameSet = "N/A"
	End If
end sub

'::::::::::::::::::::::::::
':: Get GUID of Computer ::
'::::::::::::::::::::::::::
Function GetUUID
	Dim objWmi, colItems, objItem, strUUID, blnValidUUID    

	Set objWmi = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")    
	Set colItems = objWmi.ExecQuery("Select * from Win32_ComputerSystemProduct")    
	strUUID = ""    
	blnValidUUID = False    
	For Each objItem in colItems        
		strUUID = objItem.UUID        
		If Not IsEmpty(strUUID) OR Not IsNull(strUUID) Then            
			If (strUUID <> "00000000-0000-0000-0000-000000000000") AND _               
			(strUUID <> "FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF") Then                
			blnValidUUID = True                
			Exit For            
			End If        
		End If    
		Next    
	Set objWmi = Nothing    
	If Not blnValidUUID Then        
		Set colItems = GetObject("winmgmts:" & strComputer & "\root\cimv2").InstancesOf("Win32_NetworkAdapter")        
		For Each objItem In colItems            
			If (objItem.AdapterType = "Ethernet 802.3") Then                
				If (objItem.Description <> "Packet Scheduler Miniport") Then                    
				strUUID = "00000000-0000-0000-0000-" & Replace(objItem.MACAddress, ":", "")                    
				Exit For                
				End If            
			End If        
		Next        
		Set NicSet = Nothing    
	End If    
	GetUUID = strUUID
End Function

':::::::::::::::::::::::::
':: Get SID of Computer ::
':::::::::::::::::::::::::
Function GetSID
Dim wshShell, fso, tfolder, tname, TempFile, results, retString, ts 
Const ForReading = 1, TemporaryFolder = 2 

Set wshShell = CreateObject("wscript.shell")
Set fso = CreateObject("Scripting.FileSystemObject") 
Set tfolder = fso.GetSpecialFolder(TemporaryFolder) 

tname = fso.GetTempName
TempFile =fso.buildpath(tfolder, tname)
strCommand=fso.BuildPath(strScriptPath, "psgetsid.exe") 
GetSID ="N/A" 

strBlast="cmd /c " & Chr(34) & strCommand  & Chr(34) & " \\" & strComputer & " >" &  TempFile 

	wshShell.Run strBlast , 0, True

	Set results = fso.GetFile(TempFile)
	Set ts = results.OpenAsTextStream(ForReading)
Do While ts.AtEndOfStream <> True
	retString = ts.ReadLine
	If InStr(retString, "S-") > 0 Then
	GetSID="SID = " & Right(retString,Len(retString)-1) 
		'retString = GetSIDAddress(retString)
Exit Do
	End If
Loop

ts.Close
results.Delete
End Function

'::::::::::::::::::::::
':: Check for commas ::
'::::::::::::::::::::::
Function CheckCommas(Field)
Dim CommaRemove, FieldSet

OUpart = "a"
CommaRemove = ""
countitem = 0
FieldSet = Field & ","

Do until OUpart = ""
	OUpart = split(FieldSet, ",")
		IF OUpart(countitem) = ""  THEN Exit Do
		OUpart = "a"
		countitem = countitem + 1
		Loop
	IF countitem >=1 THEN
		countitemb = 0

	Do until countitemb = countitem

	IF countitemb < countitem AND countitemb <> countitem-1 THEN  CommaRemove = CommaRemove + OUpart(countitemb)
	IF countitemb = countitem-1 THEN CommaRemove = CommaRemove + OUpart(countitemb)

	countitemb = countitemb + 1

Loop
	END IF

CheckCommas = CommaRemove

End Function

':::::::::::::::::::::::
':: Write File Titles ::
':::::::::::::::::::::::
sub writefiletitles
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Set filetxt = filesys.OpenTextFile(filepath, ForAppending, True)
	filetxt.WriteLine "MAC Address,Network Adapter,IP Address,Computer Name,GUID/UUID,GUID/UUID Mac Address Conversion,Total RAM Memory,Sid,RAM Memory Bank Label,RAM Memory Capacity,RAM Memory Device Locator,RAM Memory Manufacturer,Computer Manufacturer,Computer Model,Computer Serial,Motherboard Manufacturer,Motherboard Model,Motherboard Product,Motherboard Serial Number,Motherboard Version,Hard Disk Drive Letter,Hard Disk Size,Hard Disk Free Space,Hard Disk Volume Serial Number,Hard Disk Manufacturer,Hard Disk Model,Hard Disk Device ID,Hard Disk Interface Type,Hard Disk Partitions,Hard Disk Total Cylinders,Hard Disk Total Heads,Hard Disk Total Sectors,Hard Disk Total Tracks,Hard Disk Tracks Per Cylinder,Processor Name,Processor Type,Processor Family,Sound Card Manufacturer,Sound Card Product Name,Sound Card Description,Sound Card Name,Video Card Name,Video Card Setting ID,Floppy Manufacturer,Floppy Name,Floppy Device ID,CDROMCaptionSet,CDROMDescriptionSet,CDROMDeviceIDSet,CDROMDriveSet,CDROMManufacturerSet,CDROMNameSet,BIOS Version,Dell Service Tag,Date of Log,Time of Log"
	filetxt.Close
end sub

'::::::::::::::::::::::::::::::::::::
':: Check if inventory already run ::
'::::::::::::::::::::::::::::::::::::
sub checkthemac
	Set filetxt = filesys.OpenTextFile(filepath)
		Do Until filetxt.AtEndOfStream
 			maccheck = (Left(filetxt.ReadLine,17))
			if MACAddressSet = maccheck then inventorystats = "COMPLETED"
		Loop
	filetxt.Close

If inventorystats = "COMPLETE" Then
	Exit Sub
End If

call ipcomnamechange

If inventorystats = "COMPLETED" then
	Wscript.Quit
End If
end sub

'::::::::::::::::::::::::::::::::::::::
':: Have IP or Computer Name Changed ::
'::::::::::::::::::::::::::::::::::::::
sub ipcomnamechange
Set filetxt = filesys.OpenTextFile(filepath)
	Do Until filetxt.AtEndOfStream
		readfullline = (filetxt.ReadLine)
		maccheck = (Left(readfullline,17))
			if MACAddressSet = maccheck then
				countitem = 0
				Do until countitem = 4
					FieldProp = split(readfullline, ",")
						If countitem = 0 Then
							macread = FieldProp(countitem)
						End If
						If countitem = 1 Then
							netcardread = FieldProp(countitem)
						End If
						If countitem = 2 Then
							ipread = FieldProp(countitem)
						End If
						If countitem = 3 Then
							comnameread = FieldProp(countitem)
						End If
					countitem = countitem + 1
				Loop
			End If
	Loop
filetxt.Close
if NetworkAdaptSet = netcardread AND IPAddressSet = ipread AND ComputerNameSet = comnameread Then
	Exit Sub
End If
if NetworkAdaptSet <> netcardread OR IPAddressSet <> ipread OR ComputerNameSet <> comnameread Then
	inventorystats = "CHANGE"
End If
if inventorystats = "CHANGE" Then
	call deletethelog
End If
end sub

'::::::::::::::::::::
':: Delete the Log ::
'::::::::::::::::::::
sub deletethelog
Dim strContents
Const FOR_READING = 1
Const FOR_WRITING = 2
strFileName = filepath
strCheckForString = Ucase(MACAddressSet)

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(strFileName, FOR_READING)

	strContents = objTS.ReadAll
	objTS.Close

arrLines = Split(strContents, vbNewLine)
	Set objTS = objFS.OpenTextFile(strFileName, FOR_WRITING)
		For Each strLine In arrLines
		   If Not(Left(UCase(LTrim(strLine)),Len(strCheckForString)) = strCheckForString) Then
			If strLine <> "" Then
				objTS.WriteLine strLine
			End If
		   End If
		Next
end sub