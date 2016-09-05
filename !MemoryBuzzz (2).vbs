'HOW TO USE: <cmd> cscript !MemoryBuzzz.vbs 75 <percent to memory pressure> <how long in seconds>
' SILENT MODE:

'links uteis: http://www.visualbasicscript.com/VBScript-memory-leaks-m40532.aspx
'https://msdn.microsoft.com/en-us/library/aa394239(v=vs.85).aspx
'http://www.activexperts.com/admin/scripts/wmi/vbscript/0407/

Class MyClass
      Public Foo
  End Class

Dim X,Y,objFile,ceiling,totalMemoryWork,intCont,intSemaphore,TotalVisibleMemorySize,FreePhysicalMemory
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

'WScript.Echo WScript.arguments(0) / 100
ceiling = WScript.arguments(0) / 100
strComputer = "."

'1 ciclo = 1.6kb
'100000 ciclos = 160mb
intSemaphore = 100001
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

Wscript.Echo totalMemoryWork & " total free * ceiling"
Wscript.Echo (TotalVisibleMemorySize - FreePhysicalMemory)& " total em uso"
totalMemoryWork = totalMemoryWork - (TotalVisibleMemorySize - FreePhysicalMemory)
Wscript.Echo totalMemoryWork & " total free * ceiling - o ja utilizado"

'#highcode totalMemoryWork = totalMemoryWork * ceiling

'Define numero de ciclos para alocacao
intSemaphore = Int(totalMemoryWork / 1.6)

'Inicio
WScript.Echo "Operation System Memory    : " & TotalVisibleMemorySize & " KB"
WScript.Echo "Memory already in use      : " & (TotalVisibleMemorySize - FreePhysicalMemory) & " KB" & "( " & Int(((TotalVisibleMemorySize - FreePhysicalMemory) / TotalVisibleMemorySize)*100) & "% )"
WScript.Echo "Memory unit allocated(avg) : " &  "1.6 KB"
WScript.Echo "Memory will be allocated   : " &  Int(totalMemoryWork) & " KB"
WScript.Echo "Starting proccess... "
Call MemoryBusy()
WScript.Echo "Proccess completed. Memory Allocated"
WScript.Echo "Sleeping"
WScript.Sleep(10000)


Sub CoreMonitor()
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem",,48)
	'Set objFSO=CreateObject("Scripting.FileSystemObject")
	
		For Each objItem in colItems
			 
			'Wscript.Echo "TotalVisibleMemorySize: " & objItem.TotalVisibleMemorySize
			'Wscript.Echo "FreePhysicalMemory: " & objItem.FreePhysicalMemory
			Wscript.Echo objItem.FreePhysicalMemory 
			Wscript.Echo 1 - (objItem.FreePhysicalMemory / objItem.TotalVisibleMemorySize)
		Next
End Sub

Sub MemoryBusy()
	
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
	  
		If intCont = ctrlAlloc  Then
			Wscript.Echo ctrlPct & "% - Memory Allocation:" & Int(totalMemoryWork*("0."& ctrlPct)) & " KB"" of Total:" & (Int(totalMemoryWork)) & " KB"
			
			ctrlPct = ctrlPct + 10
			ctrlAlloc = ctrlAlloc + ctrlAllocUnit
		End If
	Loop
	  
 End Sub
 
 
Sub CoreMonitorResidual()
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem",,48)
	'Set objFSO=CreateObject("Scripting.FileSystemObject")
	
		For Each objItem in colItems
			 
		If (objItem.TotalVisibleMemorySize - objItem.FreePhysicalMemory) < totalMemoryWork Then
			totalMemoryWork = totalMemoryWork - (objItem.TotalVisibleMemorySize - objItem.FreePhysicalMemory)
			intSemaphore = totalMemoryWork / 1.6
			Call MemoryBusy()
		End If
		
		Next
End Sub