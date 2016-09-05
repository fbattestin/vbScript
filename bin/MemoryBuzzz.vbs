'HOW TO USE: <cmd> cscript !MemoryBuzzz.vbs 75 <percent to memory pressure> <how long in seconds>
'INPUT: <percent to memory pressure>, TYPE: integer, VALUES: > 0 to <=100
'		<how long in seconds>,TYPE: integer,VALUES: > 0 
' SILENT MODE:

'links uteis: http://www.visualbasicscript.com/VBScript-memory-leaks-m40532.aspx
'https://msdn.microsoft.com/en-us/library/aa394239(v=vs.85).aspx
'http://www.activexperts.com/admin/scripts/wmi/vbscript/0407/

'each cicle in the MemoryBuzzz Sub routine generate 1.6KB memory allocate
'1 cicle = 1.6kb
'100000 cicles = 160mb

Class MyClass
      Public Foo
  End Class

Dim X,Y,objFile,ceiling,totalMemoryWork,intCont,intSemaphore,TotalVisibleMemorySize,FreePhysicalMemory,Sleep,Silent
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

'WScript.Echo WScript.arguments(0) / 100
'ceiling is the top of memory pressure
ceiling = WScript.arguments(0) / 100
'Sleep is the total time the pressure act.
Sleep = WScript.arguments(1)

'Silent Mode:  0 = Verbose 1 = Silent
Silent = WScript.arguments(2)
If Silent <> 1 Then 
	Silent = 0
Else 
	Silent = 1
End If

strComputer = "."
intCont = 1

'calculating ceiling 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem",,48)

For Each objItem in colItems

	'#highcode totalMemoryWork =  (objItem.TotalVisibleMemorySize - (objItem.TotalVisibleMemorySize - objItem.FreePhysicalMemory)) 
	totalMemoryWork = objItem.TotalVisibleMemorySize * ceiling
	TotalVisibleMemorySize = objItem.TotalVisibleMemorySize
	FreePhysicalMemory = objItem.FreePhysicalMemory
Next

If (TotalVisibleMemorySize - FreePhysicalMemory) > totalMemoryWork Then
	Wscript.Echo "Pressure ceiling requested is equal to or less than the current memory consumption."
	WScript.Echo "Memory already in use      : " & (TotalVisibleMemorySize - FreePhysicalMemory) & " KB" & "( " & Int(((TotalVisibleMemorySize - FreePhysicalMemory) / TotalVisibleMemorySize)*100) & "% )"
	WScript.Quit()
End If

'Wscript.Echo totalMemoryWork & " total free * ceiling"
'Wscript.Echo (TotalVisibleMemorySize - FreePhysicalMemory)& " total em uso"
totalMemoryWork = totalMemoryWork - (TotalVisibleMemorySize - FreePhysicalMemory)
'Wscript.Echo totalMemoryWork & " total free * ceiling - o ja utilizado"

'#highcode totalMemoryWork = totalMemoryWork * ceiling

'Define numero de ciclos para alocacao
intSemaphore = Int(totalMemoryWork / 1.6)

'Start
'Silent Verify

If Silent = 1 Then
	Call MemoryBuzzz()
	
	If Sleep > 1 Then
		WScript.Sleep(Sleep * 1000)
	End If
Else 
	WScript.Echo "Operation System Memory    : " & TotalVisibleMemorySize & " KB"
	WScript.Echo "Memory already in use      : " & (TotalVisibleMemorySize - FreePhysicalMemory) & " KB" & "( " & Int(((TotalVisibleMemorySize - FreePhysicalMemory) / TotalVisibleMemorySize)*100) & "% )"
	WScript.Echo "Memory unit allocated(avg) : " &  "1.6 KB"
	WScript.Echo "Memory will be allocated   : " &  Int(totalMemoryWork) & " KB"
	WScript.Echo "Starting proccess... "
	Call MemoryBuzzz()
	WScript.Echo "Proccess completed. Memory Allocated"

	If Sleep > 1 Then
			WScript.Echo "Sleeping " & Sleep & "seconds"
			WScript.Sleep(Sleep * 1000)
	End If

End If

Sub MemoryBuzzz()
	
	Dim ctrlAlloc,contAlloc,ctrlPct
	ctrlAlloc = Int(intSemaphore/10)
	ctrlAllocUnit = ctrlAlloc
	contAlloc = 1
	ctrlPct = 10
	
	
	Do While intCont < (intSemaphore + 1)
	  Set X = New MyClass
	  Set Y = New MyClass
	  
	  Set X.Foo = Y
	  Set Y.Foo = X
	  
	  Set X = Nothing
	  Set Y = Nothing

	  intCont = intCont + 1
	  
		If intCont = ctrlAlloc and Silent = 0 Then
			Wscript.Echo ctrlPct & "% - Memory Allocation:" & Int(totalMemoryWork*("0."& ctrlPct)) & " KB"" of Total:" & (Int(totalMemoryWork)) & " KB"
			
			ctrlPct = ctrlPct + 10
			ctrlAlloc = ctrlAlloc + ctrlAllocUnit
		End If
	Loop
	  
 End Sub
