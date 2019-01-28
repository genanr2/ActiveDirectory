Set oShell = WScript.CreateObject("WScript.Shell")
DesktopPath = oShell.SpecialFolders("Desktop")
WorkingPath = oShell.SpecialFolders("\\192.168.1.209\Garant\Garant-FS")
Set oShortCut = oShell.CreateShortcut(DesktopPath & "\Гарант Платформа F1 Эксперт.lnk")
oShortCut.TargetPath = "\\192.168.1.209\Garant\Garant-FS\garant.exe"
oShortCut.WorkingDirectory = "\\192.168.1.209\Garant\Garant-FS"
oShortCut.Description = "Гарант Платформа F1 Эксперт"
oShortCut.Save

rem net use x: \\192.168.1.224\1Cdata Buh2010 /user:BuhOff


'Option Explicit  
'On Error Resume Next  
'Dim WshShell, WshNetwork  
'Dim strUserDN, objSysInfo, GroupObj, UserGroups, UserObj 
'UserGroups=""    
'Set WshShell = WScript.CreateObject("WScript.Shell")  
'Set objSysInfo = CreateObject("ADSystemInfo")   
'strUserDN = objSysInfo.userName   
'Set UserObj = GetObject("LDAP://" & strUserDN)   
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
Function MapDrv(DrvLet, UNCPath)  
    Dim WshNetwork         ' Object variable  
    Dim Msg  
    Set WshNetwork = WScript.CreateObject("WScript.Network")  
    On Error Resume Next  
    WshNetwork.RemoveNetworkDrive DrvLet  
    WshNetwork.MapNetworkDrive DrvLet, UNCPath  
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
