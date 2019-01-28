Option Explicit  
On Error Resume Next  
Dim WshShell
dim DesktopPath 
dim WorkingPath
dim oShortCut
'net use y: \\192.168.1.226\sbisNET Ghj100rdf /user:Администратор
'MapDrv "y:", "\\192.168.1.226\sbisNET", "Администратор", "Ghj100rdf"
'net use x: \\192.168.1.224\1Cdata Buh2010 /user:BuhOff
'MapDrv "x:", "\\192.168.1.224\1Cdata", "BuhOff", "Buh2010"
'net use z: \\192.168.1.224\pub Buh2010 /user:BuhOff
'MapDrv "z:", "\\192.168.1.224\pub", "BuhOff", "Buh2010"

' Замена пароля Администратора на всех компьютерах 
'strComputer = "."
'Set objUser = GetObject("WinNT://" & strComputer & "/Администратор,user")
'objUser.SetPassword "gztECPh8"
'objUser.SetInfo

Dim oShell' as WshShell
Set oShell = WScript.CreateObject("WScript.Shell")
DesktopPath = oShell.SpecialFolders("Desktop")
'WorkingPath = oShell.SpecialFolders("\\192.168.1.209\Garant\Garant-FS")

' Настройки истемы Гарант
'oShell.LogEvent 1, "Гарант Платформа F1 Эксперт 1"
'Set oShortCut = oShell.CreateShortcut(DesktopPath & "\Гарант Платформа F1 Эксперт.lnk")
'oShortCut.TargetPath = "\\192.168.1.224\Garant\Garant-FS\garant.exe"
'oShortCut.WorkingDirectory = "\\192.168.1.224\Garant\Garant-FS"
'oShortCut.Description = "Гарант Платформа F1 Эксперт"
'oShortCut.Save
'oShell.LogEvent 1, "Гарант Платформа F1 Эксперт 2"


rem net use x: \\192.168.1.224\1Cdata Buh2010 /user:BuhOff
'oShell.run()

Dim WshNetwork  
'Dim strUserDN, objSysInfo, GroupObj, UserGroups, UserObj 
'UserGroups=""    
'Set WshShell = WScript.CreateObject("WScript.Shell")  
Set WshNetwork = WScript.CreateObject("WScript.Network")  
'WshNetwork.RemoveNetworkDrive "y:"
'MapNetworkDrive "y:", "\\192.168.1.226\sbisNET", "Администратор", "Ghj100rdf"
'oShell.LogEvent 1, "Кружка."
'oShell.LogEvent 1, "Кружка."
'WshNetwork.RemoveNetworkDrive "x"
'WshNetwork.RemoveNetworkDrive "y"
'WshNetwork.RemoveNetworkDrive "z"
'WshNetwork.MapNetworkDrive "y:", "\\192.168.1.226\sbisNET", true, "Администратор", "Ghj100rdf"
'WshNetwork.MapNetworkDrive "z:", "\\192.168.1.224\\chm2010", true ', "data\BuhOff", "Buh2010"
'WshNetwork.MapNetworkDrive "x:", "\\192.168.1.224\1Cdata", true', "data\BuhOff", "Buh2010"

' Set WshNetwork = WScript.CreateObject("WScript.Network")  
' On Error Resume Next  
' WshNetwork.RemoveNetworkDrive DrvLet  

'oShell.LogEvent 2, "Кружка."
oShell.LogEvent 2, "Кружка."
oShell.LogEvent 2, CStr(Err.Number)
oShell.LogEvent 2, Err.Description 
oShell.LogEvent 2, "Кружка2."


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

'Const strPath = "\\192.168.1.230\пользователидомена\log\Log.txt" 'Здесь надо задать UNC-путь к доступному для всех пользователей на запись сетевому ресурсу
Const strPath = "\\192.168.1.230\log\Log.txt" 'Здесь надо задать UNC-путь к доступному для всех пользователей на запись сетевому ресурсу
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
'Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}//" & strComputer & "/root/cimv2")
'Set colNetAdapters = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
Set colNetAdapters = objWMIService.InstancesOf("Win32_NetworkAdapterConfiguration")' Where IPEnabled = True")
For Each objNetAdapter  in colNetAdapters 
'  For Each strAddress in objNetAdapter.IPAddress
'    objFile.WriteLine("; MAC: ")
'    arrOctets = Split(strAddress, ".")
'    objFile.WriteLine(arrOctets)
'    If arrOctets(0) and arrOctets(1) Then
'      strNewAddress = arroctets(0) & "." & arroctets(1) & "." & arrOctets(2) & "." & "211"            
'      arrIPAddress = Array(strNewAddress)
'      strSubnetMask = objNetAdapter.IPSubnet
'      strGateway = objNetAdapter.DefaultIPGateway
'      strGatewayMetric = objNetAdapter.GatewayCostMetric
'      arrDNSServers = objNetAdapter.DNSServerSearchOrder
'      errEnable = objNetAdapter.EnableStatic(arrIPAddress, strSubnetMask)
'      errGateways = objNetAdapter.SetGateways(strGateway, strGatewaymetric)
'      objNetAdapter.SetDNSServerSearchOrder(arrDNSServers)
'    End If
'  Next
'  objFile.WriteLine("; MAC: ")'  & objNetAdapter .MACAddress)
  If IsArray( objNetAdapter.IPAddress ) Then
    objFile.WriteLine("MAC : "  & objNetAdapter.MACAddress)
    If UBound( objNetAdapter.IPAddress ) = 0 Then
      objFile.WriteLine("IP : " & objNetAdapter.IPAddress(0))
    Else
'      objFile.WriteLine("; MAC41: ")'  & strIP)
      strIP = "IP : " & Join( objNetAdapter.IPAddress, "," )
      objFile.WriteLine("; MAC42: ")'  & strIP)
    End If
  End If
'  objFile.WriteLine("; IP: " & strIP)

'  For Each strAddress in objNetAdapter.IPAddress
'    objFile.WriteLine("IP: " & strAddress)
'  Next
Next
'WshNetwork.SetDefaultPrinter=
'Set objCollection = objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
'Set objWMI = GetObject("winmgmts:{impersonationlevel=impersonate}!\\.\root\cimv2")
'strQuery = "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE MACAddress > ''"
'strIP = "dfdsf"
'Set objWMIService2 = GetObject( "winmgmts://./root/CIMV2" )
'Set colItems = objWMIService2.ExecQuery( strQuery, "WQL", 48 )

'For Each objItem In colItems
'  If IsArray( objItem.IPAddress ) Then
'    If UBound( objItem.IPAddress ) = 0 Then
'      strIP = "IP Address: " & objItem.IPAddress(0)
'    Else
'    strIP = "IP Addresses: " & Join( objItem.IPAddress, "," )
'   End If
'  End If
'Next
'objFile.WriteLine("; IP: " & strIP)
'WScript.Echo strIP
dim stringx 
dim freef 
dim free
stringx = "Локальные диски"
'  & vbNewLine  & vbNewLine
'objFile.WriteLine(stringx & vbNewLine)
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
'Set WSHShell = WScript.CreateObject("WScript.Shell")
'Проверяем все драйвы (HDD, FDD, CDD) в системе    
For each i In fso.Drives
  If i.DriveType=1 Then
    If i.DriveLetter<>"A:" Then
      freef = FormatNumber(fso.GetDrive(i.DriveLetter).FreeSpace/1048576, 1)'frit(i)
'      oShell.Popup(freef)
    End If
  End If
  If i.DriveType=2 Then
'    objFile.WriteLine(i)
    free=FormatNumber(fso.GetDrive(i.DriveLetter).FreeSpace/1048576, 1)'frit(i)'frit2(i)
'    oShell.Popup(free)
    stringx= stringx & " На диске " & i & " свободно " & free & " Мб " & vbNewLine
  End If
Next
'stringx = stringx
'objFile.WriteLine(stringx)
'oShell.Popup(stringx)
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
'    end if
'    if item.FreeSpace/item.size <= AlertHigh then
'      strMessage = strMessage & UCase(strComputer) & ": Диск '" & item.caption & "' is low on HDD space!  There are " & FormatNumber((item.FreeSpace/1024000),0) & " MB free <7%" & vbCRLF
      strMessage = strMessage & "Свободно на '" & item.caption & "' = " & FormatNumber((item.FreeSpace/1024000),0) & " Мб из " & FormatNumber((item.size/1024000),0) & " Мб"
'      oShell.Popup(strMessage)
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
'  s=s+1 
  objFile.WriteLine("Видеоконтролер " & objitem.name)
'  objFile.WriteLine("BIOS " & objitem.caption & vbCrLf & " Частота "+cstr(objitem.CurrentClockSpeed))
next

'Set colItems = objWMIService.InstancesOf("win32_printer")
'for each objItem in colItems 
'  objFile.WriteLine("Принтер "  & objitem.name)
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
'  s=s+1 
  objFile.WriteLine("Монитор " & objitem.name)
'  objFile.WriteLine("BIOS " & objitem.caption & vbCrLf & " Частота "+cstr(objitem.CurrentClockSpeed))
next

Set colItems = objWMIService.InstancesOf("Win32_UserAccount")
for each objItem in colItems 
    if objitem.LocalAccount = true then
      objFile.WriteLine("Локальный пользователь " & objitem.name)
    end if
next

Set colItems = objWMIService.InstancesOf("Win32_Product")
for each objItem in colItems 
  objFile.WriteLine("Программа " & objitem.name)
  objFile.WriteLine("ID " & cstr(objitem.ProductID))
  objFile.WriteLine("Version " & objitem.Version)
  objFile.WriteLine("PackageCode " & cstr(objitem.PackageCode))
  objFile.WriteLine("PackageName " & cstr(objitem.PackageName))
'& " ID " & objitem.ProductID & " Version " & objitem.Version & " PackageCode " & objitem.PackageCode & " PackageName " & objitem.PackageName
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

'Msg = "Mapping network drive: " & CStr(Err.Number) & " 0x" & Hex(Err.Number) & vbCrLf & _  
'  "Описание: " & Err.Description & vbCrLf  
'  Msg = Msg & "Domain: " & WshNetwork.UserDomain & vbCrLf  
'  Msg = Msg & "Computer Name: " & WshNetwork.ComputerName & vbCrLf  
'  Msg = Msg & "Пользователь: " & WshNetwork.UserName & vbCrLf & vbCrLf  
'  Msg = Msg & "Диск: " & "y:" & vbCrLf  
'  Msg = Msg & "Маршрут: " & "\\192.168.1.226\sbisNET"
'oShell.LogEvent 1, Msg
'Set objSysInfo = CreateObject("ADSystemInfo")   
'strUserDN = objSysInfo.userName   
'Set UserObj = GetObject(".LDAP://" & strUserDN)   
'For Each GroupObj In UserObj.Groups    
'        UserGroups=UserGroups & "[" & GroupObj.Name & "]"    
'Next    
'MsgBox "Member of "& UserGroups    
'if InGroup("Supports Admins") then    
'        MapDrv "Z:", "\\SRV\SUPPORT$"  
'end if   
'if InGroup("1C Users") then    
'        MapDrv "W:", "\\SRV\Base" 
'end if   
'MapDrv "L:", "\\SRV\Users\" & WshShell.ExpandEnvironmentStrings("%USERNAME%")  
'==========================================================================  
' Function MapDrv(DrvLet, UNCPath)  
' DrvLet -  Буква устройства  
' UNCPath - Сетевой путь  
' COMMENT: Подключение сетевых дисков с записью ошибок в EventLog  
'==========================================================================  
function frit2(gg)
  frit = FormatNumber(fso.GetDrive(gg.DriveLetter).FreeSpace/1048576, 1)
end function
Function MapDrv(DrvLet, UNCPath, sUsername, sPassword)  
    Dim WshNetwork         ' Object variable  
    Dim Msg  
    Set WshNetwork = WScript.CreateObject("WScript.Network")  
    On Error Resume Next  
    WshNetwork.RemoveNetworkDrive DrvLet  
'    WshNetwork.MapNetworkDrive DrvLet, UNCPath  
    MapNetworkDrive DrvLet,UNCPath,sUsername,sPassword
    Msg = "Mapping network drive: " & vbCrLf  ' & CStr(Err.Number) & " 0x" & Hex(Err.Number) & vbCrLf & _  
'      "Error description: " & Err.Description & vbCrLf  
      Msg = Msg & "Domain: " & WshNetwork.UserDomain & vbCrLf  
      Msg = Msg & "Computer Name: " & WshNetwork.ComputerName & vbCrLf  
      Msg = Msg & "User Name: " & WshNetwork.UserName & vbCrLf & vbCrLf  
      Msg = Msg & "Device name: " & DrvLet & vbCrLf  
      Msg = Msg & "Map path: " & UNCPath   
    WshShell.LogEvent 1, Msg, "\\SRV"  
    WshShell.LogEvent 4, "Кружка."
    WshShell.LogEvent 2, "Кружка."
    WshShell.LogEvent 0, "Кружка."
    WshShell.LogEvent 1, "Кружка."

'  0  SUCCESS
'  1  ERROR
'  2  WARNING
'  4  INFORMATION
'  8  AUDIT_SUCCESS
' 16  AUDIT_FAILURE

    Select Case Err.Number  
        Case 0            ' No error  
        Case -2147023694   
            WshNetwork.RemoveNetworkDrive DrvLet  
            WshNetwork.MapNetworkDrive DrvLet, UNCPath  
        Case -2147024811   
            WshNetwork.RemoveNetworkDrive DrvLet  
            WshNetwork.MapNetworkDrive DrvLet, UNCPath  
        Case Else  
            Msg = "Mapping network drive error: " & _   
                   CStr(Err.Number) & " 0x" & Hex(Err.Number) & vbCrLf & _  
                  "Error description: " & Err.Description & vbCrLf  
            Msg = Msg & "Domain: " & WshNetwork.UserDomain & vbCrLf  
            Msg = Msg & "Computer Name: " & WshNetwork.ComputerName & vbCrLf  
            Msg = Msg & "User Name: " & WshNetwork.UserName & vbCrLf & vbCrLf  
            Msg = Msg & "Device name: " & DrvLet & vbCrLf  
            Msg = Msg & "Map path: " & UNCPath   
            WshShell.LogEvent 1, Msg, "\\SRV"  
    End Select  
End Function 
'==========================================================================  
' Function InGroup(strGroup) 
' strGroup - группа, принадлежность к которой проверяем 
' COMMENT: проверка принадлежности пользователя к группе 
'==========================================================================  
Function InGroup(strGroup)    
  InGroup=False    
  If InStr(UserGroups,"[CN=" & strGroup & "]") Then    
    InGroup=True    
  End If    
End Function


Function MapNetworkDrive(sDriveLetter,sNetworkPath,sUsername,sPassword)
 On Error Resume Next				'Will continue even if there is a network error
 Err.Clear  				'Setting Error Value to Zero
 Set GetDrive = CreateObject("WScript.Network")
 If sUsername="" Or sPassword="" Then
  GetDrive.MapNetworkDrive sDriveLetter, sNetworkPath,True
 Else
  GetDrive.MapNetworkDrive sDriveLetter, sNetworkPath,True,sUserName,SPassword
 End If
 MapNetworkDrive = Err.Number 
End Function

Function RemoveNetworkDrive(sDriveLetter)
  On Error Resume Next				'Will continue even if there is a network error
   Err.Clear  				'Setting Error Value to Zero
  Set WshNetwork = CreateObject("WScript.Network")
  WshNetwork.RemoveNetworkDrive sDriveLetter,true,true
  RemoveNetworkDrive = Err.Number
End Function

Function Shell(sCommand)
   Dim oShell, oExec, sLine
   Set oShell = CreateObject("WScript.Shell")
   Set oExec = oShell.Exec(sCommand)
   Do While Not oExec.StdOut.AtEndOfStream
      sLine = oExec.StdOut.ReadLine
      WScript.StdOut.WriteLine "Output: " & sLine
      WScript.Sleep 10
   Loop
   Do While oExec.Status = 0
      WScript.Sleep 100
   Loop
End Function

Function Computers_List(arrTemp)
Dim objAD, objItem, strTemp, strList
Dim strDomain, objWSNet
Const strGroup = "Компьютеры домена"
Set objWSNet = CreateObject("WScript.Network")
strDomain = objWSNet.UserDomain
Set objWSNet = Nothing
Set objAD = GetObject("WinNT://" & strDomain & "/" & strGroup & ",group")
For Each objItem In objAD.Members
    If Not objItem.AccountDisabled Then
        strTemp = objItem.Name
        strList = strList & Left(strTemp, Len(strTemp) - 1) & vbNewLine
    End If
Next
Set objAD = Nothing
arrTemp = Split(strList, vbNewLine)
ReDim Preserve arrTemp(UBound(arrTemp) - 1)
Call Sorting_Array(arrTemp)
End Function
 
'=======
Function Logged_User(strComputer)
Dim objWMI, objCollection, objItem, strTemp
Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objCollection = objWMI.ExecQuery("SELECT UserName FROM Win32_ComputerSystem")
For Each objItem In objCollection
    strTemp = objItem.UserName
    If Not IsNull(strTemp) Then
        strTemp = Mid(strTemp, InstrRev(strTemp, "\") + 1)
    End If
Next
Set objCollection = Nothing
Set objWMI = Nothing
Logged_User = strTemp
End Function
 
'=======
Function Sorting_Array(arrTemp)
Dim blnStopSort, intNumChange, strTemp
blnStopSort = False
i = 1
Do
    intNumChange = 0
    For j = 0 To UBound(arrTemp) - 1
        If arrTemp(j) > arrTemp(j + 1) Then
            strTemp = arrTemp(j)
            arrTemp(j) = arrTemp(j + 1)
            arrTemp(j + 1) = strTemp
            intNumChange = intNumChange + 1
        End If
    Next
    If intNumChange = 0 Then
        blnStopSort = True
    Else
        If i < UBound(arrTemp) Then
            i = i + 1
        Else
            blnStopSort = True
        End If
    End If
Loop While Not blnStopSort
End Function






Function GetUserDN(username)
    Dim objConnection
    Dim objCommand
    Dim objRecordSet
    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Open "Provider=ADsDSOObject;"
    Set objCommand = CreateObject("ADODB.Command")
    objCommand.ActiveConnection = objConnection
    objCommand.CommandText="<LDAP://dc=DOMAINNAME,dc=ru>;(&(objectCategory=User)(samaccountname="&username&"));ADsPath;subtree"
    Set objRecordSet=objCommand.Execute
    If objRecordSet.RecordCount=0 Then
        GetUserDN=""
        Else
        strADsPath=objRecordset.Fields("ADsPath")
        Set objUser=GetObject(strADsPath)
        objUser.GetInfo
        GetUserDN=objUser.Get("distinguishedName")
    End If
    objConnection.Close
End Function
 
 
rem ------------------------------------------------------------
rem GetFreeDrive функция поиска свободной буквы диска по списку.
rem Принимает параметр в формате "A,B,C" где A,B,C буквы дисков.
rem Выполняет последовательный перебор до первой свободной.
rem В случае занятости всех возвращает пустую строку "".
rem ------------------------------------------------------------
function GetFreeDrive(DriveList)
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
found=""
while (len(DriveList)>0 and found="")
   drv=left(DriveList,1)&":"
   Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk where DeviceID='"&drv&"'")
   if (colDisks.Count>0) then
      if len(DriveList)>2 then
         DriveList=mid(DriveList,3)
      else
         DriveList=""
      end if
   else
      found=drv
   end if
wend
GetFreeDrive=found
end function
rem ------------------------------------------------------------
 
 
rem ------------------------------------------------------------
rem GetMountArray Создает массив соответствий диск-путь.
rem Принимает параметром пользовательский DistinguishedName.
rem Возвращает двумерный массив ("Буква диска", "путь").
rem ------------------------------------------------------------
function GetMountArray(strUserDN)
Set objUser=GetObject("LDAP://"&strUserDN)
grps=objUser.GetEx("memberOf")
dim Groups
redim Groups(ubound(grps),2)
size=0
  for each grp in grps
      On Error Resume Next
      Dim objGroup
      Set objGroup=GetObject("LDAP://"&grp)
      Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D
      descr=""
      descr=objGroup.Get("info")
      If err.number <> E_ADS_PROPERTY_NOT_FOUND Then
    if (left(descr,1)="(") then
       size=size+1
       pos=InStr(descr,")")
       letters=mid(descr,2,pos-2)
       path=mid(descr,pos+1)
       drv=GetFreeDrive(letters)
       Groups(size,1)=drv
       Groups(size,2)=path
    end if
      End If
  next
dim Data
redim Data(size,2)
for i=1 to size
  Data(i,1)=Groups(i,1)
  Data(i,2)=Groups(i,2)    
next
 
  
GetMountArray=data
end function
rem ------------------------------------------------------------
 
rem ------------------------------------------------------------
rem ParseVariables функция поиска переменных в строке и подстановки значений.
rem возвращает обработанную строку
rem ------------------------------------------------------------
function ParseVariables(strLine)
set objNet=CreateObject("wscript.network")
str=replace(lcase(strLine),"%username%",objNet.UserName)
str=replace(str,"%computername%",objNet.ComputerName)
ParseVariables=str
end function

' добавляет сетевой диск на основание заметки у групп...
function SetDisk()
  on error resume next
  set objNet=CreateObject("wscript.network")
  UsrName=objNet.Username
  strUsrDN=GetUserDN(usrname)
  dim lst
  lst=GetMountArray(strUsrDN)
  for i=1 to ubound(lst)
    if lst(i,1)="U:" then
     objNet.MapNetworkDrive lst(i,1), lst(i,2)&UsrName
    else
     objNet.MapNetworkDrive lst(i,1), lst(i,2)
    end if
  next
end function

'  изменение ip адреса сетевого адаптера через wmi
function ChangeIP()
  strComputer = "."
  Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
  Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
  strIPAddress = Array("192.168.0.3")
  strSubnetMask = Array("255.255.255.255")
  strGateway = Array("192.168.1.100")
  strGatewayMetric = Array(1)
  For Each objNetAdapter in colNetAdapters
    errEnable = objNetAdapter.EnableStatic(strIPAddress, strSubnetMask)
    errGateways = objNetAdapter.SetGateways(strGateway, strGatewaymetric)
    arrDNSServers = Array("192.168.1.100", "192.168.1.200")
    objNetAdapter.SetDNSServerSearchOrder(arrDNSServers)
  Next
end function

function ChangeIP2()
  strComputer = "."
  Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
  Set colNetAdapters = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE and DHCPEnabled=TRUE")
  strIPAddress = Array("192.168.0.3")
  For Each objNetAdapter in colNetAdapters
    If Left(objNetAdapter.IPAddress,7) = Left(strIPAddress,7) Then
      strSubnetMask = objNetAdapter.IPSubnet
      strGateway = objNetAdapter.DefaultIPGateway
      strGatewayMetric = objNetAdapter.GatewayCostMetric
      arrDNSServers = objNetAdapter.DNSServerSearchOrder
      errEnable = objNetAdapter.EnableStatic(strIPAddress, strSubnetMask)
      errGateways = objNetAdapter.SetGateways(strGateway, strGatewaymetric)
      objNetAdapter.SetDNSServerSearchOrder(arrDNSServers)
    End If
  Next
end function


Function getip()
'  set myobj = getobject("winmgmts:{impersonationlevel=" & "impersonate}") "!//localhost".execquery ("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
'  for each ipadress in myobj
'  if ipaddress.ipaddress(0) <> "0.0.0.0" then
'  localip = ipaddress.ipaddress(0)
'  exit for
'  end if
'  next
'  getip = localip
end function


' выводит свободное дисковое пространтсво)

Function getFreeSpace()
  on Error resume Next
  stringx = "Локальные диски" & vbNewLine  & vbNewLine
  Set fso = WScript.CreateObject("Scripting.FileSystemObject")
  Set WSHShell = WScript.CreateObject("WScript.Shell")
  'Проверяем все драйвы (HDD, FDD, CDD) в системе    
  For each i In fso.Drives
    If i.DriveType=1 Then
      If i<>"A:" Then
        freef = frit(i)
      End If
    End If
    If i.DriveType=2 Then
      free=frit(i)
      stringx= stringx & " На диске " & i & " свободно " & free & " Мб " & vbNewLine
    End If
  Next
  stringx = stringx
  WSHShell.Popup(stringx)
  WScript.Quit()
end function
 
function frit(gg)
  frit = FormatNumber(fso.GetDrive(gg.DriveLetter).FreeSpace/1048576, 1)
End function


' Выводит данные в процентах. 
' При превышении определенного порога % свободного пространства скрипт отсылает предупреждающее 
' (о том что на системном диске 0% свободно) сообщение на почту. 
' Задействовать SimpleMAPI. См. "Microsoft Collaboration Data Objects Programmer's Reference" в MSDN. 
function getFreeSpace2()
  Const AlertHigh = .9                    
  Const emailFrom = "xx@xxx.ru"        
  Const emailTo = "xx@xxx.ru"          
  Const MailServer = "mail.xxx.ru"
  Const WaitTimeInMinutes = 1                
  Dim strMessage
  Dim arrServerList
  arrServerList = array("server name")    
  Do until i = 2
    strMessage = ""
    PollServers(arrServerList)
    if strMessage <> "" then
      EmailAlert(strMessage)
    end if
    WScript.Sleep(WaitTimeInMinutes*60000)
  Loop
End function
 
Sub PollServers(arrServers)
  on error resume next
  for each Server in arrServers
    set objSvc = GetObject("winmgmts:{impersonationLevel=impersonate}//" & Server & "/root/cimv2")
    set objRet = objSvc.InstancesOf("win32_LogicalDisk")
    for each item in objRet
      if item.DriveType = 7 then
      end if
      if item.FreeSpace/item.size <= AlertHigh then
        strMessage = strMessage & UCase(Server) & ": Alert, drive '" & item.caption & "' is low on HDD space!  There are " & FormatNumber((item.FreeSpace/1024000),0) & " MB free <7%" & vbCRLF
      end if
    next
  next
  set objSvc = Nothing
  set objRet = Nothing
End Sub
 
Sub EmailAlert(Message)
  on error resume next
  Set objMessage = CreateObject("CDO.Message")
  with objMessage
    .From = emailFrom
    .To = emailTo
    .Subject = "Low Disk Space Update"
    .TextBody = Message
    .Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MailServer
    .Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
    .Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "xxx"
    .Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "xxx"
    .Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    .Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
    .Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
    .Configuration.Fields.Update
    .Send
  end with
  Set objMessage = Nothing
End Sub