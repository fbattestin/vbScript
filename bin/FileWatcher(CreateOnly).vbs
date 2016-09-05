Dim objFile
Set objFSO=CreateObject("Scripting.FileSystemObject")
Const ForReading = 1, ForWriting = 2, ForAppending = 8
 
outFile="c:\temp\DUMPWATCHER.inf"

strComputer = "."

Set objFile = objFSO.OpenTextFile(outFile, ForAppending, True)


Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & _
    strComputer & "\root\cimv2")

Set colMonitoredEvents = objWMIService.ExecNotificationQuery _
    ("Select * From __InstanceCreationEvent Within 5 Where " _
    & "Targetinstance Isa 'CIM_DirectoryContainsFile' and " _
    & "TargetInstance.GroupComponent= " _
    & "'Win32_Directory.Name=""c:\\\\temp\\\\scripts""'")

Do
    Set objLatestEvent = colMonitoredEvents.NextEvent
    'Wscript.Echo objLatestEvent.TargetInstance.PartComponent
	objFile.Write Now & ";" & objLatestEvent.TargetInstance.PartComponent & vbCrLf
	
	'close connection
	'objFile.Close
Loop