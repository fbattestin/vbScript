' VBScript source code
Dim objFile
Set objFSO=CreateObject("Scripting.FileSystemObject")
Const ForReading = 1, ForWriting = 2, ForAppending = 8

outFile="c:\temp\DUMPWATCHER.inf"

strComputer = "."

Set objFile = objFSO.OpenTextFile(outFile, ForAppending, True)

intInterval = "2"
strDrive = "C:" 
strFolder = "\\windows\\installer\\"
strComputer = "." 

' Connect to WMI

Set objWMIService = GetObject( "winmgmts:" &_ 
    "{impersonationLevel=impersonate}!\\" &_ 
    strComputer & "\root\cimv2" )

' The query string

strQuery =  _
    "Select * From __InstanceOperationEvent" _
    & " Within " & intInterval _
    & " Where Targetinstance Isa 'CIM_DataFile'" _
    & " And TargetInstance.Drive='" & strDrive & "'"_
    & " And TargetInstance.Path='" & strFolder & "'"

' Execute the query

Set colEvents = _
    objWMIService. ExecNotificationQuery (strQuery) 

' The loop

Do 
    ' Wait for the next event  
    ' Get SWbemEventSource object
    ' Get SWbemObject for the target instance
    
    Set objEvent = colEvents.NextEvent()
    Set objTargetInst = objEvent.TargetInstance
    
    ' Check the class name for SWbemEventSource
    ' It cane be one of the following:
    ' - __InstanceCreationEvent
    ' - __INstanceDeletionEvent
    ' - __InstanceModificationEvent
    
    Select Case objEvent.Path_.Class 
        
        ' If it is file creation or deletion event
        ' just echo the file name
        
        Case "__InstanceCreationEvent" 
            'WScript.Echo "Created: " & objTargetInst.Name 
			objFile.Write Now & ";" & "Created: " & objTargetInst.Name  & vbCrLf
			
        Case "__InstanceDeletionEvent" 
            'WScript.Echo "Deleted: " & objTargetInst.Name 
			objFile.Write Now & ";" & "Deleted: " & objTargetInst.Name  & vbCrLf
			
        ' If it is file modification event, 
        ' compare property values of the target and previous
        ' instance and echo the properties that have changed
        
        Case "__InstanceModificationEvent" 
        
            Set objPrevInst = objEvent.PreviousInstance
        
            For Each objProperty In objTargetInst.Properties_
                If objProperty.Value <> _
                objPrevInst.Properties_(objProperty.Name) Then
                    objFile.Write Now & ";" & "Changed:" & objTargetInst.Name & vbCrLf
                    objFile.Write Now & ";" & "Property:" & objProperty.Name & vbCrLf
                    objFile.Write Now & ";" & "Previous value:" & objPrevInst.Properties_(objProperty.Name)  & vbCrLf
                    objFile.Write Now & ";" & "New value:" & objProperty.Value  & vbCrLf
                End If            
            Next

    End Select 

Loop

	'close connection
	objFile.Close