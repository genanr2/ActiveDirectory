Option Explicit  
On Error Resume Next  
Dim WshShell
dim DesktopPath 
dim WorkingPath
dim oShortCut
'net use y: \\192.168.1.226\sbisNET Ghj100rdf /user:�������������
'MapDrv "y:", "\\192.168.1.226\sbisNET", "�������������", "Ghj100rdf"
'net use x: \\192.168.1.224\1Cdata Buh2010 /user:BuhOff
'MapDrv "x:", "\\192.168.1.224\1Cdata", "BuhOff", "Buh2010"
'net use z: \\192.168.1.224\pub Buh2010 /user:BuhOff
'MapDrv "z:", "\\192.168.1.224\pub", "BuhOff", "Buh2010"

' ������ ������ �������������� �� ���� ����������� 
'strComputer = "."
'Set objUser = GetObject("WinNT://" & strComputer & "/�������������,user")
'objUser.SetPassword "gztECPh8"
'objUser.SetInfo

Dim oShell ' as WshShell
Dim objAD, objUserName, objComputerName
Set oShell = WScript.CreateObject("WScript.Shell")
Set objAD = CreateObject("ADSystemInfo")

DesktopPath = oShell.SpecialFolders("Desktop")
'WorkingPath = oShell.SpecialFolders("\\192.168.1.209\Garant\Garant-FS")

' ��������� ������� ������
'oShell.LogEvent 1, "������ ��������� F1 ������� 1"
'Set oShortCut = oShell.CreateShortcut(DesktopPath & "\������ ��������� F1 �������.lnk")
'oShortCut.TargetPath = "\\192.168.1.224\Garant\Garant-FS\garant.exe"
'oShortCut.WorkingDirectory = "\\192.168.1.224\Garant\Garant-FS"
'oShortCut.Description = "������ ��������� F1 �������"
'oShortCut.Save
'oShell.LogEvent 1, "������ ��������� F1 ������� 2"


rem net use x: \\192.168.1.224\1Cdata Buh2010 /user:BuhOff
'oShell.run()

Dim WshNetwork  
Set WshNetwork = WScript.CreateObject("WScript.Network")  
'Dim strUserDN, objSysInfo, GroupObj, UserGroups, UserObj 
'UserGroups=""    
'Set WshShell = WScript.CreateObject("WScript.Shell")  
'WshNetwork.RemoveNetworkDrive "y:"
'MapNetworkDrive "y:", "\\192.168.1.226\sbisNET", "�������������", "Ghj100rdf"
'oShell.LogEvent 1, "������."
'oShell.LogEvent 1, "������."
'WshNetwork.RemoveNetworkDrive "x"
'WshNetwork.RemoveNetworkDrive "y"
'WshNetwork.RemoveNetworkDrive "z"
On Error Resume Next  
WshNetwork.MapNetworkDrive "y:", "\\192.168.1.226\sbisNET", true, "�������������", "Ghj100rdf"
oShell.LogEvent 2, "sbisNET"
WshNetwork.MapNetworkDrive"z:", "\\192.168.1.224\pub", true ', "BuhOff", "Buh2010")
oShell.LogEvent 2, "pub"
WshNetwork.MapNetworkDrive"x:", "\\192.168.1.224\1Cdata", true ', "BuhOff", "Buh2010")
oShell.LogEvent 2, "1Cdata"
oShell.LogEvent 2, CStr(Err.Number)
oShell.LogEvent 2, Err.Description

' Set WshNetwork = WScript.CreateObject("WScript.Network")  
' On Error Resume Next  
' WshNetwork.RemoveNetworkDrive DrvLet  

oShell.LogEvent 2, "������."
oShell.LogEvent 2, "������."
'oShell.LogEvent 2, CStr(Err.Number)
'oShell.LogEvent 2, Err.Description


'��� ����� ������ �� ����� ���������� � ��������� ���� �����
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

'Const strPath = "\\192.168.1.230\������������������\Log\Log.txt" '����� ���� ������ UNC-���� � ���������� ��� ���� ������������� �� ������ �������� �������
Const strPath = "\\192.168.1.230\Log\Log.txt" '����� ���� ������ UNC-���� � ���������� ��� ���� ������������� �� ������ �������� �������
Const ForAppending = 8
Set objUserName = GetObject("LDAP://" & objAD.UserName)
'WshNetwork.
Set objComputerName = GetObject("LDAP://" & objAD.ComputerName)
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFile = objFS.OpenTextFile(strPath, ForAppending, True)
objFile.WriteLine(Date & "; " & Time & "; " & objComputerName.cn & "; " & objUserName.cn)
oShell.LogEvent 2, "������ Date & Time."

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
stringx = "��������� �����"
'  & vbNewLine  & vbNewLine
'objFile.WriteLine(stringx & vbNewLine)
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
'Set WSHShell = WScript.CreateObject("WScript.Shell")
'��������� ��� ������ (HDD, FDD, CDD) � �������    
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
    stringx= stringx & " �� ����� " & i & " �������� " & free & " �� " & vbNewLine
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
'      strMessage = strMessage & UCase(strComputer) & ": ���� '" & item.caption & "' is low on HDD space!  There are " & FormatNumber((item.FreeSpace/1024000),0) & " MB free <7%" & vbCRLF
      strMessage = strMessage & "�������� �� '" & item.caption & "' = " & FormatNumber((item.FreeSpace/1024000),0) & " �� �� " & FormatNumber((item.size/1024000),0) & " ��"
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
  objFile.WriteLine("������ ��: " & objItem.Version & " ����� ����������: " & objItem.ServicePackMajorVersion & "." & objItem.ServicePackMinorVersion & vbNewLine)
Next
dim strTextBody
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration",,48) 
For Each objItem in colItems 
    If isNull(objItem.IPAddress) Then
    Else
      strTextBody = strTextBody + vbCrLf + "������� �����: " & objItem.Caption & ", IP �����: " & Join(objItem.IPAddress, ",")
    End If
Next
objFile.WriteLine(strTextBody)

dim constmb, constgb, sizegb 
dim compname, temp, compad 
constmb=1048576 
Set colItems = objWMIService.InstancesOf("win32_ComputerSystem")
for each objItem in colItems 
  objFile.WriteLine("����������� ������ " & cstr(round(objItem.totalphysicalmemory/constmb)))
  objFile.WriteLine("������ ���������� " & objitem.model & vbCrLf)
next
Set colItems = objWMIService.InstancesOf("win32_bios")
for each objItem in colItems 
  objFile.WriteLine("����������� ����� " & objitem.SMBIOSBIOSVersion)
  objFile.WriteLine("BIOS " & objitem.caption & vbCrLf)
next
Set colItems = objWMIService.InstancesOf("win32_processor")
for each objItem in colItems 
  s=s+1 
  objFile.WriteLine("��������� " & cstr(s) & " " & objitem.name)
  objFile.WriteLine("��������� " & objitem.caption & vbCrLf & " ������� "+cstr(objitem.CurrentClockSpeed))
next

Set colItems = objWMIService.InstancesOf("win32_videocontroller")
for each objItem in colItems 
  objFile.WriteLine("�������������� " & objitem.name)
next

Set colItems = objWMIService.InstancesOf("win32_printer")
for each objItem in colItems 
  objFile.WriteLine("������� "  & objitem.name)
'  objFile.WriteLine("BIOS " & objitem.caption & vbCrLf & " ������� "+cstr(objitem.CurrentClockSpeed))
next
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
  objFile.WriteLine("������� " & objitem.name)
'  objFile.WriteLine("BIOS " & objitem.caption & vbCrLf & " ������� "+cstr(objitem.CurrentClockSpeed))
next

Set colItems = objWMIService.InstancesOf("Win32_UserAccount")
for each objItem in colItems 
    if objitem.LocalAccount = true then
      objFile.WriteLine("��������� ������������ " & objitem.name)
    end if
next
'objFile.WriteLine(" " & vbCrLf)

Set colItems = objWMIService.InstancesOf("Win32_Product")
for each objItem in colItems 
  objFile.WriteLine("��������� " & objitem.name)
  objFile.WriteLine("ID " & cstr(objitem.ProductID))
  objFile.WriteLine("Version " & objitem.Version)
  objFile.WriteLine("PackageCode " & cstr(objitem.PackageCode))
  objFile.WriteLine("PackageName " & cstr(objitem.PackageName))
'& " ID " & objitem.ProductID & " Version " & objitem.Version & " PackageCode " & objitem.PackageCode & " PackageName " & objitem.PackageName
next

Set colItems = objWMIService.InstancesOf("Win32_ProcessStartup")
objFile.WriteLine("������� " & objitem.Title)
for each objItem in colItems 
  objFile.WriteLine("������� " & objitem.Title)
next

'Set colItems = objWMIService.InstancesOf("Win32_Account")
'for each objItem in colItems 
'  objFile.WriteLine("Account " & objitem.name)
'next

'Set colItems = objWMIService.InstancesOf("Win32_SystemAccount")
'for each objItem in colItems 
'  objFile.WriteLine("SystemAccount " & objitem.name)
'next


'objFile.WriteLine("������� " & vbCrLf)

'dim strclass, objAD, obj 
'dim invdate 
'dim constmb, constgb, sizegb 
'dim compname, temp, compad 
'constmb=1048576 
'constgb=1073741824 
'strclass = array( "win32_ComputerSystem", "win32_bios", "win32_processor", "win32_diskdrive", "win32_videocontroller", "win32_NetworkAdapter",_ 
'"win32_sounddevice", "win32_SCSIController", "win32_printer") 

objFile.WriteLine("**************************************************************")

'Set objWMIService = GetObject( "winmgmts://./root/CIMV2" )
strQuery = "SELECT * FROM Win32_Environment WHERE Name='TEMP'"
Set colItems = objWMIService.ExecQuery( strQuery, "WQL", 48 )

For Each objItem In colItems
' 	WScript.Echo "Caption        : " & objItem.Caption
' 	WScript.Echo "Description    : " & objItem.Description
' 	WScript.Echo "Name           : " & objItem.Name
' 	WScript.Echo "Status         : " & objItem.Status
' 	WScript.Echo "SystemVariable : " & objItem.SystemVariable
'	WScript.Echo "UserName       : " & objItem.UserName
'	WScript.Echo "VariableValue  : " & objItem.VariableValue
'	WScript.Echo
    objFile.WriteLine("Caption        : ")
Next
Set wshUserEnv = oShell.Environment( "USER" )
For Each strItem In wshUserEnv
    objFile.WriteLine(strItem)
Next
Set wshUserEnv = Nothing
objFile.WriteLine("User")

d = True
Dim strComputerName
strComputer = "." 
'***Ignore OU Named "Computers" when getting computerlocation from OU structure
IgnoreComputersOU = True       
strRootDSE = GetRootDSE()
'strComputerName = GetComputerName(oShell)
If IsServer() Then 
'    wscript.echo "1"
'    strComputerName = lcase(oShell.ExpandEnvironmentStrings("%CLIENTNAME%"))
    strComputerName = oShell.ExpandEnvironmentStrings("%CLIENTNAME%")
'    wscript.echo "1"
    objFile.WriteLine(strComputerName)
Else 
    wscript.echo "2"
'    GetComputerName = lcase(oShell.ExpandEnvironmentStrings("%COMPUTERNAME%"))
    strComputerName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
    wscript.echo "2"
    objFile.WriteLine(strComputerName)
end if

Set WshShell2 = CreateObject("WScript.Shell")
strComputerName = WshShell2.ExpandEnvironmentStrings("%CLIENTNAME%")
objFile.WriteLine(strComputerName)

dim strOriginalString
dim strExpandedString
strOriginalString = "Windows is installed in %WinDir%. %XYZ% is undefined."
strExpandedString = WshShell2.ExpandEnvironmentStrings(strOriginalString)
'WshShell2.GetLastError 

'objFile.WriteLine(WshShell2.ExpandEnvironmentStrings( "PATH=%PATH%" ))
Set wshSystemEnv = wshShell2.Environment( "SYSTEM" )
objFile.WriteLine(wshSystemEnv( "TEMP" ))


'objFile.WriteLine(strOriginalString)
'objFile.WriteLine(strExpandedString)

'WScript.Echo "Before: " & strOriginalString
'WScript.Echo "After: " & strExpandedString


'objFile.WriteLine(strComputerName)
strDN = GetComputerDN(strComputerName) 
Dim str2
If d = true Then
'    str2 = " "
    if IsServer() = true then
'        str2 = "Runs on server: " & "True"
    else
        str2 = "Runs on server: " & "False"
    end if
    objFile.WriteLine(str2)
'    If IsServer() = true Then
'        str2 = " "
'        str2 = "TS Client Name: " & strComputerName
'        objFile.WriteLine(strComputerName)
'        objFile.WriteLine(str2)
'        str2 = "TS Client DN: " & strDN
'        objFile.WriteLine(str2)
        objFile.WriteLine(strDN)
'    Else
 '       str2 = " "
'        str2 = "Computer Name: " & strComputerName
'        objFile.WriteLine(str2)
'        objFile.WriteLine(strComputerName)
'        str2 = "TS Client DN: " & strDN
'        objFile.WriteLine(str2)
'        objFile.WriteLine(strDN)
'    End If
End If    
'***getting computerlocation from OU structure
'If Not left(strDN,9) = "Could not" Then 
    computerLocation = GetComputerADLocation(strDN)
    parentLocation = GetParentLocation(computerLocation)
'    If d = true Then
''        wscript.echo "Computer Location: " & computerLocation
''        str2 = " "
''        str2 = "Computer Location: " & computerLocation
''        objFile.WriteLine(str2)
'        objFile.WriteLine(computerLocation)
''        str2 = "Computer Parent Location: " & parentLocation
''        objFile.WriteLine(str2)
'        objFile.WriteLine(parentLocation)
'    End If
'    EnumerateLocalPrinters()
'    EnumerateADPrinters(objFile) ', d)
'End If    
EnumerateLocalPrinters(objFile)
EnumerateADPrinters(objFile) ',d)

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
'  "��������: " & Err.Description & vbCrLf  
'  Msg = Msg & "Domain: " & WshNetwork.UserDomain & vbCrLf  
'  Msg = Msg & "Computer Name: " & WshNetwork.ComputerName & vbCrLf  
'  Msg = Msg & "������������: " & WshNetwork.UserName & vbCrLf & vbCrLf  
'  Msg = Msg & "����: " & "y:" & vbCrLf  
'  Msg = Msg & "�������: " & "\\192.168.1.226\sbisNET"
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
' DrvLet -  ����� ����������  
' UNCPath - ������� ����  
' COMMENT: ����������� ������� ������ � ������� ������ � EventLog  
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
    WshShell.LogEvent 4, "������."
    WshShell.LogEvent 2, "������."
    WshShell.LogEvent 0, "������."
    WshShell.LogEvent 1, "������."

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
' strGroup - ������, �������������� � ������� ��������� 
' COMMENT: �������� �������������� ������������ � ������ 
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
Const strGroup = "���������� ������"
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
rem GetFreeDrive ������� ������ ��������� ����� ����� �� ������.
rem ��������� �������� � ������� "A,B,C" ��� A,B,C ����� ������.
rem ��������� ���������������� ������� �� ������ ���������.
rem � ������ ��������� ���� ���������� ������ ������ "".
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
rem GetMountArray ������� ������ ������������ ����-����.
rem ��������� ���������� ���������������� DistinguishedName.
rem ���������� ��������� ������ ("����� �����", "����").
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
rem ParseVariables ������� ������ ���������� � ������ � ����������� ��������.
rem ���������� ������������ ������
rem ------------------------------------------------------------
function ParseVariables(strLine)
set objNet=CreateObject("wscript.network")
str=replace(lcase(strLine),"%username%",objNet.UserName)
str=replace(str,"%computername%",objNet.ComputerName)
ParseVariables=str
end function

' ��������� ������� ���� �� ��������� ������� � �����...
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

'  ��������� ip ������ �������� �������� ����� wmi
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


' ������� ��������� �������� ������������)

Function getFreeSpace()
  on Error resume Next
  stringx = "��������� �����" & vbNewLine  & vbNewLine
  Set fso = WScript.CreateObject("Scripting.FileSystemObject")
  Set WSHShell = WScript.CreateObject("WScript.Shell")
  '��������� ��� ������ (HDD, FDD, CDD) � �������    
  For each i In fso.Drives
    If i.DriveType=1 Then
      If i<>"A:" Then
        freef = frit(i)
      End If
    End If
    If i.DriveType=2 Then
      free=frit(i)
      stringx= stringx & " �� ����� " & i & " �������� " & free & " �� " & vbNewLine
    End If
  Next
  stringx = stringx
  WSHShell.Popup(stringx)
  WScript.Quit()
end function
 
function frit(gg)
  frit = FormatNumber(fso.GetDrive(gg.DriveLetter).FreeSpace/1048576, 1)
End function


' ������� ������ � ���������. 
' ��� ���������� ������������� ������ % ���������� ������������ ������ �������� ��������������� 
' (� ��� ��� �� ��������� ����� 0% ��������) ��������� �� �����. 
' ������������� SimpleMAPI. ��. "Microsoft Collaboration Data Objects Programmer's Reference" � MSDN. 
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
'--------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------- 
Function GetComputerLocation() 
  const HKEY_LOCAL_MACHINE = &H80000002
  Set oReg=GetObject( "winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
  strKeyPath = "SOFTWARE\Policies\Microsoft\Windows NT\Printers"
  strValueName = "PhysicalLocation"
  oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,computerLocation
  GetComputerLocation = trim(lCase(computerLocation)) 
End Function 

Function GetComputerDN (computername) 
  'On Error Resume Next
  Const ADS_SCOPE_SUBTREE = 2
  Set objConnection = CreateObject("ADODB.Connection")
  Set objCommand = CreateObject("ADODB.Command")
  objConnection.Provider = "ADsDSOObject"
  objConnection.Open "Active Directory Provider"
  Set objCommand.ActiveConnection = objConnection
  objCommand.CommandText = "SELECT dnsHostName, distinguishedName FROM " & "'LDAP://" & strRootDSE & "' WHERE objectClass='computer' AND Name='" & computername & "'"
  objCommand.Properties("Page Size") = 1000
  objCommand.Properties("Timeout") = 30
  objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
  objCommand.Properties("Cache Results") = False
  Set objRecordSet = objCommand.Execute
  If Not objRecordSet.EOF Then 
    GetComputerDN = objRecordSet.Fields("distinguishedName").Value 
  Else 
    GetComputerDN = "Could not find Computer '" & ComputerName & "' in AD."
  end if
End Function 

Function GetComputerName (oShell)
  If IsServer() Then 
    GetComputerName = lcase(oShell.ExpandEnvironmentStrings("%clientname%")) 
  Else 
    GetComputerName = lcase(oShell.ExpandEnvironmentStrings("%computername%")) 
  end if
End Function  

Function GetComputerADLocation (strDN) 
  Set objComputerName = Getobject("LDAP://" & strDN)
  Set objOU = GetObject(objComputerName.Parent)
  strOU = replace(objOU.Name,"OU=","") 
  Do While Not left(objOU.Name,3) = "DC="
    If left(objOU.Name,3) = "OU=" Then
      strOU = replace(objOU.Name,"OU=","")
      If IgnoreComputersOU And strOU = "Computers" Then strOU = ""
      If len(ADPath) Then a = "/" Else a = ""
      If strOU <> "" Then  ADPath = strOU & a & ADPath
    End If   
    Set objOU = GetObject(objOU.Parent)
  Loop
  GetComputerADLocation = ADPath 
End Function    

Private Sub EnumerateLocalPrinters(objFile)
  Dim computerLoc, parentLoc, printerLoc
  parentLoc = parentLocation
  computerLoc = computerLocation
  Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & "." & "\root\cimv2")
  Set colInstalledPrinters =  objWMIService.ExecQuery ("Select * from Win32_Printer WHERE Location<>'" & computerLoc & "' AND Location<>'" & parentLoc & "'")
  For Each objPrinter in colInstalledPrinters
    If left(objPrinter.Name,2) = "\\" Then objNetwork.RemovePrinterConnection objPrinter.Name
  Next
End Sub 

Public Function GetRootDSE() 
   Dim objRootDSE
   Set objRootDSE = GetObject("LDAP://rootDSE")
   GetRootDSE= objRootDSE.Get("defaultNamingContext")
   Set objRootDSE = Nothing 
End Function 

Function GetParentLocation (computer) 
    cL = computer
    a = False
    Do Until a
      cL = mid(cL,InStr(cL,"/")+1,len(cL)-InStr(cL,"/")+1)
      if (InStr(cL,"/") = 0) Then a = True
    Loop
    If cL = computer Then GetParentLocation = "abracadabra" Else GetParentLocation = left(computer,len(computer)-len(cL)-1) 
End Function 

'Private Sub EnumerateADPrinters(objFile, d)
Function EnumerateADPrinters (objFile) ', d)
  Dim strStr
  d = true
  computerLoc = computerLocation
  parentLoc = parentLocation
  Const ADS_SCOPE_SUBTREE = 2
  Set objConnection = CreateObject("ADODB.Connection")
  Set objCommand = CreateObject("ADODB.Command")
  objConnection.Provider = "ADsDSOObject"
  objConnection.Open "Active Directory Provider"
  Set objCommand.ActiveConnection = objConnection
  objCommand.CommandText = "SELECT printerName, serverName, Location, UNCName, Description FROM " _
    & "'LDAP://" & strRootDSE & "' WHERE objectClass='printQueue' AND (Location='" & computerLoc & "' OR Location='" & parentLoc & "')"
  objCommand.Properties("Page Size") = 1000
  objCommand.Properties("Timeout") = 30
  objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
  objCommand.Properties("Cache Results") = False
  Set objRecordSet = objCommand.Execute
  If Not objRecordSet.EOF Then objRecordSet.MoveFirst
  Do Until objRecordSet.EOF
    printerLocation = trim(lcase(objRecordSet.Fields("Location").Value))
    printerShare = objRecordSet.Fields("UNCName").Value
    If lcase(objRecordSet.Fields("serverName").Value) <> strComputerDNSName Then
        objNetwork.AddWindowsPrinterConnection printerShare
        If d Then 
'            wscript.echo "Printer Share: " & printerShare & " - True"
            strStr = "Printer Share: " & printerShare & " - True"
            objFile.WriteLine(strStr)
        end if
    Else
        If d Then 
            strStr = "Printer Share: " & printerShare & " - False (Local)"
'            wscript.echo "Printer Share: " & printerShare & " - False (Local)"
            objFile.WriteLine(strStr)
        end if
    End If
    objRecordSet.MoveNext
  Loop 
End Function

Function IsServer 
  IsServer = False
  Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
  Set colOSes = objWMIService.ExecQuery("Select Caption from Win32_OperatingSystem")
  For Each objOS in colOSes
    If InStr(objOS.Caption,"Server") Then 
        IsServer = True
    end if
  Next 
End Function
