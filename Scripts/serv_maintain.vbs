Option Explicit
Const HKEY_LOCAL_MACHINE = &H80000002
Const HIDDEN_WINDOW = 0
Dim procc_ver:procc_ver=Check_ver
dim srcript_name:srcript_name=WScript.ScriptName
dim scriptfullPath:scriptfullPath = replace(WScript.ScriptFullName,srcript_name,"")
dim t_enviroments:t_enviroments=get_enviroments

wscript.echo Now() & " Enviroments " & t_enviroments
do
  if t_enviroments<>"Access Driver not Installed" Then
    if not Check_MainTASK_IsWork then Launch_TASK
    WScript.Sleep (60*1000)
  Else
    Call Err.Raise(vbObjectError + 10, srcript_name,t_enviroments & "':=> Program terminated")
    Exit do
  end if
loop
Private Function Check_MainTASK_IsWork
  on error Resume Next
  Dim objWMIService ,colProcesses,objitem
  dim strComputer:strComputer = "."
  Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
  Set colProcesses = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_Process WHERE Name = 'cscript.exe'")
    dim cnt:cnt=0
    If colProcesses.Count > 0 Then
    For Each objitem In colProcesses
      if instr(1,objitem.CommandLine,scriptfullPath & "main.vbs")>0 then cnt=cnt+1
    Next
    end if
    If Err.Number <> 0 Then
        WScript.Echo now () & " Blad " & Str(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    End If
    if cnt>0 then
      Check_MainTASK_IsWork=True
    Else
      Check_MainTASK_IsWork=False
    end if
  Set colProcesses=Nothing
  Set objWMIService=Nothing
end Function
Private Function get_enviroments

  'check what version of drivers are installed 64bit or 32 bit'
  dim chk:chk=False
  dim chk1:chk1=False
  dim is64: is64=FolderExists("c:\Windows\SysWOW64")
  Wscript.echo("System 64? "  & cstr(is64)  & "  Process run from folder =>" & procc_ver )
  DIm strComputer:strComputer = "."
  dim arrValueNames(),strValueName,arrValueTypes(),i,strValue
  dim objRegistry: Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
  dim strKeyPath:strKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"
  objRegistry.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes
  Wscript.echo "ODBC drivers"
  For i = 0 to UBound(arrValueNames)
      strValueName = arrValueNames(i)
      objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
      if strValue = "Installed" Then
          Wscript.echo arrValueNames(i)
          If InStr(1, UCase(arrValueNames(i)), "MDB") <> 0  And InStr(1, UCase(arrValueNames(i)), "MICR") <> 0 Then
            chk = True
            Exit For
          End If
      end if
  Next
  if chk<>True Then
      if is64 then
        Wscript.echo "WOW64 drivers"
        strKeyPath = "SOFTWARE\Wow6432Node\ODBC\ODBCINST.INI\ODBC Drivers"
        objRegistry.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes
        For i = 0 to UBound(arrValueNames)
          strValueName = arrValueNames(i)
          objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
          if strValue = "Installed" Then
            Wscript.echo arrValueNames(i)
              If InStr(1, UCase(arrValueNames(i)), "MDB") <> 0  And InStr(1, UCase(arrValueNames(i)), "MICR") <> 0 Then
                chk1 = True
                Exit For
              End If
          end if
        Next
      end if
  end if
  if is64=True Then
    if chk and procc_ver=64  Then get_enviroments="C:\Windows\system32\cscript.exe"
    if chk and procc_ver=32 Then get_enviroments="C:\Windows\SysWOW64\cscript.exe"
    if chk1 and procc_ver=64 Then get_enviroments="C:\Windows\SysWOW64\cscript.exe"
  Else
    if chk Then
        get_enviroments="C:\Windows\system32\cscript.exe"
    Else
        get_enviroments="Access Driver not Installed"
    end if
  end if
end Function
Private sub Launch_TASK
  on error Resume next
  Dim objWMIService,objProcess, intProcessID
  dim strComputer:strComputer = "."
  Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
  Dim objStartup: Set objStartup = objWMIService.Get("Win32_ProcessStartup")
  Dim objConfig:Set objConfig = objStartup.SpawnInstance_
  objConfig.ShowWindow = HIDDEN_WINDOW
  'objConfig.CreateFlags=Create_New_Process_Group
  Set objProcess = objWMIService.Get("Win32_Process")
  wscript.echo now() & " " & (t_enviroments & " //NoLogo " & scriptfullPath & "main.vbs")
  Dim intReturn:intReturn = objProcess.Create (t_enviroments & " " & scriptfullPath & "main.vbs", Null, objConfig, intProcessID)
    wscript.echo Now() & " ProcessID started PID => " & intProcessID
    If Err.Number <> 0 Then
        WScript.Echo now () & " Blad " & Str(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    End If
end sub
Function FolderExists(Path)
  dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FolderExists(Path) Then
    FolderExists=CBool(1)
  Else
   FolderExists=CBool(0)
   End If
 Set fso = Nothing
End Function
Function Check_ver
  dim fso:Set fso = CreateObject("Scripting.FileSystemObject")
  dim wshShell:Set wshShell = CreateObject( "WScript.Shell" )
  If fso.FolderExists(wshShell.ExpandEnvironmentStrings("%windir%") & "\sysnative" ) Then
    Check_ver=32
  Else
    Check_ver=64
  End if
end function
