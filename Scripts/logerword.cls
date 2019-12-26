VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "loger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Option Explicit
 Private Type GUID_T
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
#If VBA7 Then
    Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
#Else
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If
Private log_path As String
Private srv_pth As String
Private loger_to_file As Boolean
Public Sub Logerr(Err_msg As String)
    Dim objCon: Set objCon = CreateObject("ADODB.Connection")
    objCon.Open ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + srv_pth & "\Log_Helper.mde")
    objCon.Execute ("Update task_procc set iserr=true where pid=" & GetCurrentProcessId)
    objCon.Close
    Set objCon = Nothing
    Log ("Wyst¹pi³ b³¹d powoduj¹cy zatrzymanie wykonania Tasku => " & Err_msg)
End Sub
Public Sub Log(Log_msg As String)
If loger_to_file Then
    Dim fs, File
     Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FileExists(log_path & "") Then 'check if file exists
        Set File = fs.OpenTextFile(log_path, 8, True, -1) ' append, create a file if doesn't exist, Unicode
        File.Write Now() & " Rozpoczynam zapis logów z  => " & App.ActiveDocument.FullName & Chr(13) + Chr(10) 'add a row to the file
        File.Close
    End If
    Set File = fs.OpenTextFile(log_path, 8, True, -1) ' append, create a file if doesn't exist, Unicode
    File.Write Now() & " " & Log_msg & Chr(13) + Chr(10) 'add a row to the file
    File.Close
Else
    Debug.Print (Now() & " " & Log_msg)
End If
End Sub
Private Sub Class_Initialize()
    Set App = Application
    If check_use_file(App.ActiveDocument.FullName) Then
       loger_to_file = True
    Else
       loger_to_file = False
    End If
End Sub
Private Function check_use_file(file_path) As Boolean
  Dim objWMIService As Object
  Dim colProcesses As Variant
  Dim objitem As Object
  Dim objLoc As Object
  Dim strComputer: strComputer = "."
  Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate, (Debug)}!\\" & strComputer & "\root\cimv2")
    objWMIService.Security_.privileges.addasstring "sedebugprivilege", True
    Set colProcesses = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_Process WHERE Name like '%script.exe%'")
    Dim cnt: cnt = 0
    If colProcesses.Count > 0 Then
      For Each objitem In colProcesses
        If InStr(1, objitem.CommandLine, "main.vbs") > 0 Then
          If Filename_EXIST_in_helper(file_path, Get_path_from_CML(objitem.CommandLine)) Then cnt = cnt + 1
        End If
      Next
    End If
    If cnt > 0 Then
      check_use_file = True
    Else
      check_use_file = False
    End If
    Set colProcesses = Nothing
    Set objWMIService = Nothing
End Function
Private Function Filename_EXIST_in_helper(file_path, serv_path) As Boolean
  If FileExists(serv_path & "\Log_Helper.mde") Then
    Dim objCon: Set objCon = CreateObject("ADODB.Connection")
    Dim rs_exist: Set rs_exist = CreateObject("ADODB.Recordset")
    objCon.Open ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + serv_path & "\Log_Helper.mde")
    With rs_exist
      .ActiveConnection = objCon
      .CursorLocation = 3
      .LockType = 4
      .Open "Select * from task_procc where app_path='" & file_path & "' and pid=0 and app_start>(Now()-(1/24/120))"
    End With
      Set rs_exist.ActiveConnection = Nothing
      If Not rs_exist.EOF Then
        rs_exist("pid") = GetCurrentProcessId()
        log_path = serv_path & "\" & rs_exist("task_id") & "_" & rs_exist("task_pid") & "_" & rs_exist("pid") & ".log"
        rs_exist("curr_log") = log_path
        Set rs_exist.ActiveConnection = objCon
        rs_exist.updatebatch
        Set rs_exist.ActiveConnection = Nothing
      End If
      objCon.Close
      Set objCon = Nothing
      If Not rs_exist.EOF Then srv_pth = serv_path
      Filename_EXIST_in_helper = Not rs_exist.EOF
      Set rs_exist = Nothing
  Else
    Filename_EXIST_in_helper = False
  End If
End Function
Private Function Get_path_from_CML(full_path) As String
    Get_path_from_CML = Replace(Mid(full_path, InStr(1, full_path, " ") + 1), Mid(full_path, InStrRev(full_path, "\")), "")
  'Get_path_from_CML = Replace(Replace(Mid(full_path, InStr(1, full_path, " ") + 1), Mid(Mid(full_path, InStr(1, full_path, " ") + 1), InStrRev(full_path, "\") + 1) & "\", ""), "\" & Mid(full_path, InStrRev(full_path, "\") + 1), "")
End Function
Private Function FileExists(FilePath)
   Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(FilePath) Then
      FileExists = CBool(1)
    Else
      FileExists = CBool(0)
    End If
    Set fso = Nothing
 End Function

 
 
