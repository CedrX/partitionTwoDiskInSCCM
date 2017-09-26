letter = "X"
outFileOneDisk = letter & ":\DiskpartOneDisk.txt"
outFileSmallerDisk = letter & ":\DiskpartSmallerDisk.txt"
outFileGreaterDisk = letter & ":\DiskpartGreaterDisk.txt"
logOneDiskFile = letter & ":\outputOneDisk.txt"
logSmallerFile = letter & ":\outputSmaller.txt"
logGreaterFile = letter & ":\outputGreater.txt"

set oShell = WScript.CreateObject("WScript.Shell")
strDiskpart = oShell.ExpandEnvironmentStrings("%windir%\system32")

Function getSmallerTwoDisks(disk1,disk2) 
	size1 = Int(disk1.Size /(1073741824)) 
	size2 = Int(disk2.Size /(1073741824)) 
	if size1 < size2 then
		set getSmallerTwoDisks =  disk1
	else
		set getSmallerTwoDisks = disk2
	end if
End Function

Function getGreaterTwoDisks(disk1,disk2) 
	size1 = Int(disk1.Size /(1073741824)) 
	size2 = Int(disk2.Size /(1073741824)) 
	if size1 > size2 then
		set getGreaterTwoDisks =  disk1
	else
		set getGreaterTwoDisks = disk2
	end if
End Function

Function setVarDriveLetter
	set colDriveLetter = objWMIService.ExecQuery("select driveletter from win32_volume where label='Windows'")
	driveletter = colDriveLetter.itemindex(0).Driveletter
	wscript.echo "Drive letter for SYSTEM volume is:" & driveletter
	Set TSEnv = CreateObject("Microsoft.SMS.TSEnvironment") 
	TSEnv("OSDCustomTinDriveLetter") = driveletter
End Function

Function createPartOneDisk(sizeDisk)
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	sizeDiskMega = int(sizeDisk /1048576)
	sizePartC = int(sizeDiskMega * (30/100))
	if sizePartC < 102400 and sizeDiskMega > 102400 then 
		sizePartC = 102400
	End if
	
	'Calculte 1% disk size for WinRE Tools partition
	sizePartRecoveryTools = int(sizeDiskMega * (1/100))
	
	Set objFile = objFSO.CreateTextFile(outFileOneDisk,True)
	objFile.Write "select disk 0" & vbCrLf
	objFile.Write "clean" & vbCrLf
	objFile.Write "convert gpt" & vbCrLf
	
	'create Windows RE tools partition
	'objFile.Write "create partition primary size=300" & vbCrLf
	'objFile.Write "format fs=ntfs quick label=""Windows Re Tools""" & vbCrLf
	'objFile.Write "assign letter=T" & vbCrLf
	'objFile.Write "set id=""de94bba4-06d1-4d40-a16a-bfd50179d6ac""" & vbCrLf
	'objFile.Write "gpt attributes=0x8000000000000001" & vbCrLf
	
	'create System partition 
	objFile.Write "create partition efi size=500" & vbCrLf
	objFile.Write "format fs=fat32 quick label=System" & vbCrLf
	objFile.Write "assign letter=S" & vbCrLf
	
	'create Microsoft Reserved Partition 
	objFile.Write "create partition msr size=128" & vbCrLf
	
	'create Windows partition 
	objFile.Write "create partition primary size=" & sizePartC & vbCrLf
	ObjFile.Write "format fs=ntfs quick label=Windows" & vbCrLf
	objFile.Write "assign" & vbCrLf
	
	'create DATAS Partition
	ObjFile.Write "create partition primary" & vbCrLf
	objFile.Write "shrink minimum=" & sizePartRecoveryTools & vbCrLf
	ObjFile.Write "assign" & vbCrLf
	ObjFile.Write "format fs=ntfs quick label=DATAS" & vbCrLf
	
	'create Recovery Image Partition
	ObjFile.Write "create partition primary" & vbCrLf
	ObjFile.Write "format fs=ntfs quick label=""Windows RE Tools""" & vbCrLf
	objFile.Write "assign letter=T" & vbCrLf
	objFile.Write "set id=""de94bba4-06d1-4d40-a16a-bfd50179d6ac""" & vbCrLf
	objFile.Write "gpt attributes=0x8000000000000001" & vbCrLf

	objFile.Close
End Function

Sub createPartSmallerDisk(index,sizeDisk)
	sizeDiskMega = int(sizeDisk /1048576)
	sizePartRecoveryTools = int(sizeDiskMega * (1/100))
	
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.CreateTextFile(outFileSmallerDisk,True)
	objFile.Write "select disk " & index & vbCrLf
	objFile.Write "clean" & vbCrLf
	objFile.Write "convert gpt" & vbCrLf

	'create Windows RE tools partition
	'objFile.Write "create partition primary size=300" & vbCrLf
	'objFile.Write "format fs=ntfs quick label=""Windows Re Tools""" & vbCrLf
	'objFile.Write "assign letter=T" & vbCrLf
	'objFile.Write "set id=""de94bba4-06d1-4d40-a16a-bfd50179d6ac""" & vbCrLf
	'objFile.Write "gpt attributes=0x8000000000000001" & vbCrLf
	
	'create System partition 
	objFile.Write "create partition efi size=500" & vbCrLf
	objFile.Write "format fs=fat32 quick label=System" & vbCrLf
	objFile.Write "assign letter=S" & vbCrLf
	
	'create Microsoft Reserved Partition 
	objFile.Write "create partition msr size=128" & vbCrLf
	
	'create Windows partition 
	objFile.Write "create partition primary" & vbCrLf
	objFile.Write "shrink minimum=" & sizePartRecoveryTools & vbCrLf
	ObjFile.Write "format fs=ntfs quick label=Windows" & vbCrLf
	objFile.Write "assign" & vbCrLf

	'create Recovery Image Partition
	ObjFile.Write "create partition primary" & vbCrLf
	ObjFile.Write "format fs=ntfs quick label=""WinRE Tools""" & vbCrLf
	objFile.Write "assign letter=T" & vbCrLf
	objFile.Write "set id=""de94bba4-06d1-4d40-a16a-bfd50179d6ac""" & vbCrLf
	objFile.Write "gpt attributes=0x8000000000000001" & vbCrLf

	objFile.Close
End Sub

Function createPartGreaterDisk(index)
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.CreateTextFile(outFileGreaterDisk,True)
	objFile.Write "select disk " & index & vbCrLf
	objFile.Write "clean" & vbCrLf
	objFile.Write "convert gpt" & vbCrLf
	objFile.Write "create partition primary" & vbCrLf
	ObjFile.Write "format fs=ntfs quick label=DATAS" & vbCrLf
	ObjFile.Write "assign" & vbCrLf
	objFile.Close
End Function

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_diskdrive where interfacetype <> 'usb'",,16)
'wscript.echo colItems.Count

'si le nombre de disques durs est égal à 1
if colItems.count = 1 then 
	set objItem = colItems.ItemIndex(0)
	createPartOneDisk(objItem.size)
	'wscript.echo "Disk size : " & objItem.size
	wscript.echo "cmd /C " & strDiskpart & "\diskpart.exe /s " & Chr(34) & outFileOneDisk
	strDispartCmdm = "cmd /C " & strDiskpart & "\diskpart.exe /s " & Chr(34) & outFileOneDisk & chr(34) & " > " & chr(34) & logOneDiskFile & chr(34)
	oShell.Run strDispartCmdm, 1, True
	setVarDriveLetter
end if

'si le nombre de disques durs est égal à 2
if colItems.count = 2 then
	'wscript.echo "On passe dans la fonction de comparaison"
	set smallerDisk = getSmallerTwoDisks(colItems.ItemIndex(0),colItems.ItemIndex(1))
	set greaterDisk = getGreaterTwoDisks(colItems.ItemIndex(0),colItems.ItemIndex(1))
	wscript.echo "On passe dans la condition 2 disques durs"
	createPartSmallerDisk smallerDisk.index, smallerDisk.size
	strDispartCmdm = "cmd /C " & strDiskpart & "\diskpart.exe /s " & Chr(34) & outFileSmallerDisk  & chr(34) & " > " & chr(34) & logSmallerFile & chr(34)
	wscript.echo strDispartCmdm
	oShell.Run strDispartCmdm, 1, True
	'setVarDriveLetter
	createPartGreaterDisk(greaterDisk.index)
	strDispartCmdm = "cmd /C " & strDiskpart & "\diskpart.exe /s " & Chr(34) & outFileGreaterDisk & chr(34) & " > " & chr(34) & logGreaterFile & chr(34)
	wscript.echo strDispartCmdm
	oShell.Run strDispartCmdm, 1, True
	setVarDriveLetter
end if

wscript.quit(0)

'For Each objItem in colItems
'	wscript.echo "Disk index : " & objItem.index
'	wscript.echo "Disk Model : " & objItem.model
'	wscript.echo "Disk size : " & objItem.size
'Next

'wscript.echo "Disk Model : " & smallerDisk.model
'wscript.echo "Disk size : " & smallerDisk.size