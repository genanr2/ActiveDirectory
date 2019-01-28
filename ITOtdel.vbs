Option Explicit  
On Error Resume Next  
Dim WshShell
dim DesktopPath 
dim WorkingPath
dim oShortCut

' ������ ������ �������������� �� ���� ����������� 
'strComputer = "."
'Set objUser = GetObject("WinNT://" & strComputer & "/�������������,user")
'objUser.SetPassword "gztECPh8"
'objUser.SetInfo

Dim oShell' as WshShell
Set oShell = WScript.CreateObject("WScript.Shell")
DesktopPath = oShell.SpecialFolders("Desktop")
'WorkingPath = oShell.SpecialFolders("\\192.168.1.209\Garant\Garant-FS")
oShell.LogEvent 1, "������ ��������� F1 ������� 1"
Set oShortCut = oShell.CreateShortcut(DesktopPath & "\������ ��������� F1 �������.lnk")
oShortCut.TargetPath = "\\192.168.1.224\Garant\Garant-FS\garant.exe"
oShortCut.WorkingDirectory = "\\192.168.1.224\Garant\Garant-FS"
oShortCut.Description = "������ ��������� F1 �������"
oShortCut.Save
oShell.LogEvent 1, "������ ��������� F1 ������� 2"


rem net use x: \\192.168.1.224\1Cdata Buh2010 /user:BuhOff
'oShell.run()

Dim WshNetwork  
'Dim strUserDN, objSysInfo, GroupObj, UserGroups, UserObj 
'UserGroups=""    
'Set WshShell = WScript.CreateObject("WScript.Shell")  
Set WshNetwork = WScript.CreateObject("WScript.Network")  
'WshNetwork.RemoveNetworkDrive "y:"
'MapNetworkDrive "y:", "\\192.168.1.226\sbisNET", "�������������", "Ghj100rdf"
'oShell.LogEvent 1, "������."
'oShell.LogEvent 1, "������."
'WshNetwork.RemoveNetworkDrive "x"
'WshNetwork.RemoveNetworkDrive "y"
'WshNetwork.RemoveNetworkDrive "z"

' Set WshNetwork = WScript.CreateObject("WScript.Network")  
' On Error Resume Next  
' WshNetwork.RemoveNetworkDrive DrvLet  

oShell.LogEvent 2, "�� �����."
oShell.LogEvent 2, "�� �����."
oShell.LogEvent 2, CStr(Err.Number)
oShell.LogEvent 2, Err.Description


'��� ����� ������ �� ����� ���������� � ��������� ���� �����
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

Const strPath = "\\192.168.1.230\������������������\Log.txt" '����� ���� ������ UNC-���� � ���������� ��� ���� ������������� �� ������ �������� �������
Const ForAppending = 8
Set objAD = CreateObject("ADSystemInfo")
Set objUserName = GetObject("LDAP://" & objAD.UserName)
'WshNetwork.
Set objComputerName = GetObject("LDAP://" & objAD.ComputerName)
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFile = objFS.OpenTextFile(strPath, ForAppending, True)
objFile.WriteLine(Date & "; " & Time & "; " & objComputerName.cn & "; " & objUserName.cn)

strComputer = "."
strIP = "."
set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}//" & strComputer & "/root/cimv2")
Set colNetAdapters = objWMIService.InstancesOf("Win32_NetworkAdapterConfiguration")' Where IPEnabled = True")
'Set colNetAdapters = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
For Each objNetAdapter  in colNetAdapters 
  If IsArray( objNetAdapter.IPAddress ) Then
    objFile.WriteLine("MAC : "  & objNetAdapter.MACAddress)
    If UBound( objNetAdapter.IPAddress ) = 0 Then
      objFile.WriteLine("IP : " & objNetAdapter.IPAddress(0))
    Else
      strIP = "IP : " & Join( objNetAdapter.IPAddress, "," )
    End If
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
    strMessage = strMessage & "�������� �� '" & item.caption & "' = " & FormatNumber((item.FreeSpace/1024000),0) & " �� �� " & FormatNumber((item.size/1024000),0) & " ��"
    objFile.WriteLine(strMessage)
    strMessage=""
  end if
next
set objSvc = Nothing
set objRet = Nothing
Set objCollection = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each objItem In objCollection
  objFile.WriteLine("������ ��: " & objItem.Version & " ����� ����������: " & objItem.ServicePackMajorVersion & "." & objItem.ServicePackMinorVersion & vbNewLine)
Next
'BIOS=
'CPU_Freq_in_MHz=
'CPU=
'Memory_in_Mb=
'RetCode = WshShell.Run("d:\psexec.exe \\comp1 -s \\server\enu\windowsXP-KB957097-x86.exe /quiet /norestart", 1, True)
'MsgBox "���������� ���������! ��� �������� - " & RetCode

'Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
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
  objFile.WriteLine("BIOS " & objitem.caption & vbCrLf & " ������� "+cstr(objitem.CurrentClockSpeed))
next

objFile.WriteLine("**************************************************************")
objFile.Close()
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
'    WshShell.LogEvent 2, "������."
'    WshShell.LogEvent 0, "������."
'    WshShell.LogEvent 1, "������."

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
            Msg = "������ ����������� �������� �����: " & _   
                   CStr(Err.Number) & " 0x" & Hex(Err.Number) & vbCrLf & _  
                  "�������� ������: " & Err.Description & vbCrLf  
            Msg = Msg & "�����: " & WshNetwork.UserDomain & vbCrLf  
            Msg = Msg & "��� ������: " & WshNetwork.ComputerName & vbCrLf  
            Msg = Msg & "������������: " & WshNetwork.UserName & vbCrLf & vbCrLf  
            Msg = Msg & "����������: " & DrvLet & vbCrLf  
            Msg = Msg & "���� �����������: " & UNCPath   
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


Function Shell1()
  dim strclass, objAD, obj 
  dim invdate 
  dim constmb, constgb, sizegb 
  dim compname, temp, compad 
  constmb=1048576 
  constgb=1073741824 
  strclass = array( "win32_ComputerSystem", "win32_bios", "win32_processor",_ 
    "win32_diskdrive", "win32_videocontroller", "win32_NetworkAdapter",_ 
    "win32_sounddevice", "win32_SCSIController", "win32_printer") 

  set objAD=getobject("LDAP://CN=Computers,DC=akos-nissan,DC=local") 
  objAD.filter=array("computer") 
  on error resume next 
  for each obj in objAD 
    CompAD=right(obj.name, len(obj.name)-3) 
    invdate = date 
    temp="<html>"+chr(10)+"���� �����: " & invdate & "<table>"+chr(10) 
    compname="" 
    ' on error resume next 
    set objWMIService = GetObject("winmgmts://"&CompAD&"/root\cimv2") 
    i=0 
    s=0 
    d=0 
    q=0 
    ' sizegb="" 
    for a=0 to 8 
      Set colitems = objwmiservice.instancesof(strclass(a)) 
      for each objitem in colitems 
        select case a 
          case 0 
            temp=temp+"<tr><td>" 
            temp=temp+"��� ����������"+"</td><td>"+objitem.name+ "</td>" + chr(10) 
            temp=temp+"</tr>"+chr(10) 
            temp=temp+"<tr><td>" 
            temp=temp+"����������� ������"+"</td><td>"+cstr(round(objitem.totalphysicalmemory/constmb))+ " MB</td>" + chr(10) 
            temp=temp+"</tr>"+chr(10) 
            temp=temp+"<tr><td>" 
            temp=temp+"������ ����������"+"</td><td>"+objitem.model+ "</td>" + chr(10) 
            temp=temp+"</tr>"+chr(10) 
            compname=objitem.name 
          case 1 
            temp=temp+"<tr><td>" 
            temp=temp+"����������� �����"+"</td><td>"+objitem.SMBIOSBIOSVersion+"</td>"+chr(10) 
            temp=temp+"</tr>"+chr(10) 
            temp=temp+"<tr><td>" 
            temp=temp+"BIOS"+"</td><td>" + objitem.caption+"</td>"+chr(10)+"<td>"+chr(10)+"</td>" 
            temp=temp+"</tr>"+chr(10) 
          case 2 
            s=s+1 
            temp=temp+"<tr>"+chr(10)+"<td>" 
            temp=temp+"���������"+cstr(s)+"</td>"+chr(10)+"<td>"+objitem.name+" ������� "+cstr(objitem.CurrentClockSpeed)+chr(10)+"</td>" 
            temp=temp+"</tr>"+chr(10) 
          case 3 
            i=i+1 
            temp=temp+"<tr>"+chr(10)+"<td>" 
            if objitem.size > 0 then ' = nill then 
              sizegb=cstr(round(objitem.size/constgb,2)) 
            else 
              sizegb=cstr(0) 
            end if 
            temp=temp+"������� ���� "+cstr(i)+"</td>"+chr(10)+"<td>"+objitem.model + " " + sizegb + " GB</td>" + chr(10) 
            temp=temp+"</tr>"+chr(10) 
          case 4 
            temp=temp+"<tr>"+chr(10)+"<td>" 
            temp=temp+"��������������"+"</td>"+chr(10)+"<td>"+objitem.caption+chr(10)+"</td>" 
            temp=temp+"</tr>"+chr(10) 
          case 5 
            if objitem.adaptertypeid=0 and objitem.netconnectionstatus=2 then 
              temp=temp+"<tr>"+chr(10)+"<td>" 
              temp=temp+"������� �������"+"</td>"+chr(10) 
              temp=temp+"<td>"+objitem.name+chr(10)+"</td>" 
              temp=temp+"</tr>"+chr(10) 
             end if 
          case 6 
            temp=temp+"<tr>"+chr(10)+"<td>" 
            temp=temp+"�������� �����"+"</td>"+chr(10) 
            temp=temp+"<td>"+objitem.caption+chr(10)+"</td></tr>"+chr(10) 
          case 7 
            temp=temp+"<tr>"+chr(10)+"<td>" 
            temp=temp+"SCSI �������"+"</td>"+chr(10) 
            temp=temp+"<td>"+objitem.manufacturer+" "+objitem.caption+chr(10)+"</td></tr>"+chr(10) 
          case 8 
            d=d+1 
            temp=temp+"<tr>"+chr(10)+"<td>" 
            temp=temp+"������� "+cstr(d)+"</td>"+chr(10)+"<td>"+objitem.name+chr(10)+"</td>" 
            temp=temp+"</tr>"+chr(10) 
        end select 
      next 
    next 
    '�������������� ����� 
    temp=temp+"</table></html>" 
    '������ ����� 
    Dim fso, tf 
    Set fso = CreateObject("Scripting.FileSystemObject") 
    Set tf = fso.CreateTextFile(""&compname&".htm", True) 
    tf.Write (temp) 
    tf.Close 
  next
End Function
