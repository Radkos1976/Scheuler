 'Sprawdź czy process już nie chodzi
dim srcript_name:srcript_name=WScript.ScriptName
dim scriptfullPath:scriptfullPath = replace(WScript.ScriptFullName,"\" & srcript_name ,"")
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
chk="false"
chk1="false"
Set colProcesses = objWMIService.ExecQuery _
("SELECT * FROM Win32_Process WHERE Name = 'cscript.exe'")

If colProcesses.Count > 0 Then
For Each objitem In colProcesses
    if instr(1,objitem.CommandLine, scriptfullPath & "\serv_maintain.vbs")>0 then
      objitem.terminate
    end if
Next
end if
