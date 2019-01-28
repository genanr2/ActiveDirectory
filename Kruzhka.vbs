Option Explicit  
On Error Resume Next  
Dim WshShell
dim DesktopPath 
dim WorkingPath
dim oShortCut

Dim oShell' as WshShell
Set oShell = WScript.CreateObject("WScript.Shell")
DesktopPath = oShell.SpecialFolders("Desktop")
'WorkingPath = oShell.SpecialFolders("\\192.168.1.209\Garant\Garant-FS")
oShell.LogEvent 1, "Гарант Платформа F1 Эксперт 1"
Set oShortCut = oShell.CreateShortcut(DesktopPath & "\Гарант Платформа F1 Эксперт.lnk")
oShortCut.TargetPath = "\\192.168.1.224\Garant\Garant-FS\garant.exe"
oShortCut.WorkingDirectory = "\\192.168.1.224\Garant\Garant-FS"
oShortCut.Description = "Гарант Платформа F1 Эксперт"
oShortCut.Save
oShell.LogEvent 1, "Гарант Платформа F1 Эксперт 2"

Dim WshNetwork  
Set WshNetwork = WScript.CreateObject("WScript.Network")  
On Error Resume Next  
'WshNetwork.MapNetworkDrive "y:", "\\192.168.1.226\sbisNET", true, "Администратор", "Ghj100rdf"
'oShell.LogEvent 2, "sbisNET"
'WshNetwork.MapNetworkDrive"z:", "\\192.168.1.224\pub", true ', "BuhOff", "Buh2010")
'oShell.LogEvent 2, "pub"
'WshNetwork.MapNetworkDrive"x:", "\\192.168.1.224\1Cdata", true ', "BuhOff", "Buh2010")
'oShell.LogEvent 2, "1Cdata"
'oShell.LogEvent 2, CStr(Err.Number)
'oShell.LogEvent 2, Err.Description

oShell.LogEvent 2, "Кружка."
oShell.LogEvent 2, "Кружка."

'под каким именем на любом компьютере в локальной сети вошли
Dim objAD, objUserName, objComputerName
Dim objFS, objFile
Dim objWMI
dim objNetAdapter
dim strIP
dim strComputer
dim objItem
dim objCollection
dim colItems
dim objWMIService
dim colNetAdapters
dim strAddress

'Const strPath = "\\192.168.1.230\ПользователиДомена\Log\Log.txt" 'Здесь надо задать UNC-путь к доступному для всех пользователей на запись сетевому ресурсу
Const strPath = "\\192.168.1.230\Log\Log.txt" 'Здесь надо задать UNC-путь к доступному для всех пользователей на запись сетевому ресурсу
Const ForAppending = 8
Set objAD = CreateObject("ADSystemInfo")
Set objUserName = GetObject("LDAP://" & objAD.UserName)
'WshNetwork.
Set objComputerName = GetObject("LDAP://" & objAD.ComputerName)
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFile = objFS.OpenTextFile(strPath, ForAppending, True)
objFile.WriteLine(Date & "; " & Time & "; " & objComputerName.cn & "; " & objUserName.cn)
oShell.LogEvent 2, "Кружка Date & Time."

strComputer = "."
strIP = "."
set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}//" & strComputer & "/root/cimv2")
Set colNetAdapters = objWMIService.InstancesOf("Win32_NetworkAdapterConfiguration")' Where IPEnabled = True")
For Each objNetAdapter  in colNetAdapters 
  If IsArray( objNetAdapter.IPAddress ) Then
    objFile.WriteLine("MAC : "  & objNetAdapter.MACAddress)
    If UBound( objNetAdapter.IPAddress ) = 0 Then
      objFile.WriteLine("IP : " & objNetAdapter.IPAddress(0))
    Else
      strIP = "IP : " & Join( objNetAdapter.IPAddress, "," )
      objFile.WriteLine("; MAC42: ")'  & strIP)
    End If
  End If
Next

dim stringx 
dim freef 
dim free
stringx = "Локальные диски"
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
For each i In fso.Drives
  If i.DriveType=1 Then
    If i.DriveLetter<>"A:" Then
      freef = FormatNumber(fso.GetDrive(i.DriveLetter).FreeSpace/1048576, 1)'frit(i)
    End If
  End If
  If i.DriveType=2 Then
    free=FormatNumber(fso.GetDrive(i.DriveLetter).FreeSpace/1048576, 1)'frit(i)'frit2(i)
    stringx= stringx & " На диске " & i & " свободно " & free & " Мб " & vbNewLine
  End If
Next
Const AlertHigh = .9  
dim objSvc 
dim objRet
dim item
dim strMessage
set objSvc = GetObject("winmgmts:{impersonationLevel=impersonate}//" & strComputer & "/root/cimv2")
set objRet = objSvc.InstancesOf("win32_LogicalDisk")
  for each item in objRet
    if item.DriveType = 7 then
    else
      strMessage = strMessage & "Свободно на '" & item.caption & "' = " & FormatNumber((item.FreeSpace/1024000),0) & " Мб из " & FormatNumber((item.size/1024000),0) & " Мб"
      objFile.WriteLine(strMessage)
      strMessage=""
    end if
  next
'    next
set objSvc = Nothing
set objRet = Nothing
Set objCollection = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each objItem In objCollection
  objFile.WriteLine("Версия ОС: " & objItem.Version & " Пакет обновления: " & objItem.ServicePackMajorVersion & "." & objItem.ServicePackMinorVersion & vbNewLine)
Next
dim strTextBody
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration",,48) 
For Each objItem in colItems 
    If isNull(objItem.IPAddress) Then
    Else
      strTextBody = strTextBody + vbCrLf + "Сетевая карта: " & objItem.Caption & ", IP адрес: " & Join(objItem.IPAddress, ",")
    End If
Next
objFile.WriteLine(strTextBody)

dim constmb, constgb, sizegb 
dim compname, temp, compad 
constmb=1048576 
Set colItems = objWMIService.InstancesOf("win32_ComputerSystem")
for each objItem in colItems 
  objFile.WriteLine("Оперативная память " & cstr(round(objItem.totalphysicalmemory/constmb)))
  objFile.WriteLine("Модель компьютера " & objitem.model & vbCrLf)
next
Set colItems = objWMIService.InstancesOf("win32_bios")
for each objItem in colItems 
  objFile.WriteLine("Материнская плата " & objitem.SMBIOSBIOSVersion)
  objFile.WriteLine("BIOS " & objitem.caption & vbCrLf)
next
Set colItems = objWMIService.InstancesOf("win32_processor")
for each objItem in colItems 
  s=s+1 
  objFile.WriteLine("Процессор " & cstr(s) & " " & objitem.name)
  objFile.WriteLine("Процессор " & objitem.caption & vbCrLf & " Частота "+cstr(objitem.CurrentClockSpeed))
next

Set colItems = objWMIService.InstancesOf("win32_videocontroller")
for each objItem in colItems 
  objFile.WriteLine("Видеоконтролер " & objitem.name)
next

'Set colItems = objWMIService.InstancesOf("win32_printer")
'for each objItem in colItems 
'  objFile.WriteLine("Принтер "  & objitem.name)
'  objFile.WriteLine("BIOS " & objitem.caption & vbCrLf & " Частота "+cstr(objitem.CurrentClockSpeed))
'next
Set colItems = objWMIService.InstancesOf("Win32_PrinterConfiguration")
for each objItem in colItems 
  objFile.WriteLine("Name "  & objitem.Name)
  objFile.WriteLine("Caption "  & objitem.Caption)
  objFile.WriteLine("DeviceName "  & objitem.DeviceName)
  objFile.WriteLine("Scale "  & objitem.Scale)
  objFile.WriteLine("PrintQuality "  & objitem.PrintQuality)
'  objFile.WriteLine("Caption "  & objitem.Caption)
next

Set colItems = objWMIService.InstancesOf("Win32_DesktopMonitor")
for each objItem in colItems 
  objFile.WriteLine("Монитор " & objitem.name)
'  objFile.WriteLine("BIOS " & objitem.caption & vbCrLf & " Частота "+cstr(objitem.CurrentClockSpeed))
next

Set colItems = objWMIService.InstancesOf("Win32_UserAccount")
for each objItem in colItems 
    if objitem.LocalAccount = true then
      objFile.WriteLine("Локальный пользователь " & objitem.name)
    end if
next
'objFile.WriteLine(" " & vbCrLf)

Set colItems = objWMIService.InstancesOf("Win32_Product")
for each objItem in colItems 
  objFile.WriteLine("Программа " & objitem.name)
  objFile.WriteLine("ID " & cstr(objitem.ProductID))
  objFile.WriteLine("Version " & objitem.Version)
  objFile.WriteLine("PackageCode " & cstr(objitem.PackageCode))
  objFile.WriteLine("PackageName " & cstr(objitem.PackageName))
next

Set colItems = objWMIService.InstancesOf("Win32_ProcessStartup")
objFile.WriteLine("Процесс " & objitem.Title)
for each objItem in colItems 
  objFile.WriteLine("Процесс " & objitem.Title)
next

objFile.WriteLine("**************************************************************")
objFile.Close
Set objFS = Nothing
Set objFile = Nothing
Set objUserName = Nothing
Set objComputerName = Nothing
Set objAD = Nothing
Set objCollection = Nothing
Set objWMI = Nothing


WScript.Quit()

