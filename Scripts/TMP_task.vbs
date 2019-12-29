Option Explicit
Const wdDoNotSaveChanges = 0
dim ext_conn,g_CurrProcessId
Dim ThrowErr:ThrowErr=False
dim srcript_name:srcript_name=WScript.ScriptName
dim scriptfullPath:scriptfullPath = replace(WScript.ScriptFullName,srcript_name,"")
dim tsk : set tsk= new Task
Class Task
  private t_db_connection_String,t_db_path_forLogs,t_db_fulpath,self_id,self_subjob,start_job
  private have_paralell_jobs
  private rs_work
  Private Function CurrProcessId
    Dim oShell, sCmd, oWMI, oChldPrcs, oCols, lOut
    lOut = 0
    Set oShell  = CreateObject("WScript.Shell")
    Set oWMI    = GetObject(_
        "winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    sCmd = "/K " & Left(CreateObject("Scriptlet.TypeLib").Guid, 38)
    oShell.Run "%comspec% " & sCmd, 0
    WScript.Sleep 100 'For healthier skin, get some sleep
    Set oChldPrcs = oWMI.ExecQuery(_
        "Select * From Win32_Process Where CommandLine Like '%" & sCmd & "'",,32)
    For Each oCols In oChldPrcs
        lOut = oCols.ParentProcessId 'get parent
        oCols.Terminate 'process terminated
        Exit For
    Next
    CurrProcessId = lOut
  End Function
  Private Function Check_childTASK_IsWork (TaskID)
    Dim objWMIService ,colProcesses,objitem
    dim strComputer:strComputer = "."
    Set objWMIService = GetObject("winmgmts:" _
      & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery _
      ("SELECT * FROM Win32_Process WHERE Name = 'cscript.exe'")
      dim cnt:cnt=0
      If colProcesses.Count > 0 Then
      For Each objitem In colProcesses
        if instr(1,objitem.CommandLine,"TaskID;" & TaskID)>0 then cnt=cnt+1
      Next
      end if
      if cnt>0 then
        Check_MainTASK_IsWork=True
      Else
        Check_MainTASK_IsWork=False
      end if
    Set colProcesses=Nothing
    Set objWMIService=Nothing
  end Function
  Private Function Prepare_childtask_toRun(TaskID,TaskName,start)
      WScript.Echo("Uruchamiam Task:" & TaskID & "  name:" & TaskName & "   "  & now())
      call Report_start_work(TaskID,TaskName,start)
      if not FileExists(scriptfullPath & "TaskID;" & TaskID & ".vbs") then
        Dim fso :Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFile scriptfullPath & "TMP_task.vbs",scriptfullPath & "TaskID;" & TaskID & ".vbs", True
        Set fso=nothing
      end if
      dim PID:PID=Launch_TASK(TaskID,TaskName)
      Prepare_task_toRun=PID
  end function
  Private sub Get_settings
    on error resume next
    dim settingsXML:Set settingsXML = CreateObject("MSXML2.DOMDocument")
    With settingsXML
      .SetProperty "SelectionLanguage", "XPath"
      .SetProperty "ProhibitDTD", False
      .ValidateOnParse =  True
      .Async = False
      .Load scriptfullPath & "settings.xml"
      End With
    if settingsXML.parseError.errorCode<>0 then
      dim myErr: set myErr= settingsXML.parseError
      WScript.Echo(now() & " You have error in XML => " & scriptfullPath & "settings.xml => " + myErr.reason)
      Call Err.Raise(vbObjectError + 10, now() & " Blad w przetwarzaniu pliku ustawien ", "Blad w " & scriptfullPath & "settings.xml => Task przerwany :=> " &  myErr.reason)
    else
      t_db_connection_String=settingsXML.selectSingleNode("scheduler/db_connection_String").text
      t_db_path_forLogs=settingsXML.selectSingleNode("scheduler/db_path_forLogs").text
      t_db_fulpath=settingsXML.selectSingleNode("scheduler/db_fulpath").text
    End if
    WScript.Echo(now() & " Ustawienia poprawnie pobrane => " & scriptfullPath & "settings.xml")
    set settingsXML=nothing
    start_job=now()
    ext_conn=t_db_connection_String
    self_id=My_ID_is
    If Err.Number <> 0 Then
       Wscript.echo String(150,"-")
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
          Report_problem(self_id)
          Wscript.echo Now() & " Porzucam dalsze wykonanie"
          Wscript.echo String(150,"-")
          WScript.Quit
    End If
  end sub
  Private Function My_ID_is
    on error resume next
    Dim pos_id_txt:pos_id_txt=InStr(1,srcript_name,"TaskID;")
    Dim pos_end_txt:pos_end_txt= InStr(1,srcript_name,".vbs")
    if pos_id_txt<>0 then
      if InStr(1,srcript_name,"Job_ID;")>0 then
        self_subjob=mid(srcript_name,InStr(1,srcript_name,"Job_ID;")+7,InStr(1,srcript_name,"TaskID;")-(InStr(1,srcript_name,"Job_ID;")+7))
      Else
        self_subjob=""
      end if
      My_ID_is=mid(srcript_name,pos_id_txt+7,pos_end_txt-(pos_id_txt+7))
    else
      My_ID_is=0
    end if
    If Err.Number <> 0 Then
       Wscript.echo String(150,"-")
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
          Report_problem(self_id)
          Wscript.echo Now() & " Porzucam dalsze wykonanie"
          Wscript.echo String(150,"-")
          WScript.Quit
    End If
  end function
  Private Sub Get_lst_work
    on error resume next
    WScript.echo now() & " Pobranie listy zadan do wykonania z bazy danych"
    Dim objCon:Set objCon = CreateObject("ADODB.Connection")
    objCon.Open (t_db_connection_String)
    Set rs_work= CreateObject("ADODB.Recordset")
    with rs_work
      .ActiveConnection = objCon
      .CursorLocation = 3
      .LockType=1
      .open "SELECT hdr.id as hdr_id,hdr.name as hdr_name,wrk.item_no,app.program,'0' as on_wrk,tp.* from task_hdr as hdr,task_work as wrk,tblWork as tp ,tbltype as app where hdr.id='" & self_id & "' and hdr.id=wrk.id_task and wrk.id_work=tp.id and tp.type=app.id order by wrk.item_no"
    end with
    Set rs_work.ActiveConnection = Nothing
    objCon.close
    Set objCon=Nothing
    If Err.Number <> 0 Then
       Wscript.echo String(150,"-")
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
          Report_problem(self_id)
          Wscript.echo Now() & " Porzucam dalsze wykonanie"
          Wscript.echo String(150,"-")
          WScript.Quit
    End If
  End Sub
  Private Sub queue_work
    on error resume next
    Dim Present_work
   if not rs_work.EOF then
    rs_work.movefirst
    do until rs_work.EOF
      Wscript.echo String(150,"-")
      WScript.echo now() & " Rozpoczynam zadanie numer " & rs_work("item_no") &  " => " & rs_work("name") & " " & rs_work("description")
      Select Case rs_work("program")
        Case "Word.Application"
          Set Present_work = (new Word) (rs_work("Path"),rs_work("SUB_FUNCTION"),rs_work("Param"),self_id)
        Case "Excel.Application"
          Set Present_work = (new Excel) (rs_work("Path"),rs_work("SUB_FUNCTION"),rs_work("Param"),self_id)
        Case "Access.Application"
          Set Present_work = (new Access) (rs_work("Path"),rs_work("SUB_FUNCTION"),rs_work("Param"),self_id)
        Case "WMI_Process"
          Set Present_work = (new Shell) (rs_work("Path"),rs_work("Param"),self_id)
        Case "Cscript x32"
          Set Present_work = (new Vbscriptx32) (rs_work("Path"),rs_work("Param"),self_id)
        Case "Cscript x64"
          Set Present_work = (new Vbscriptx64) (rs_work("Path"),rs_work("Param"),self_id)
      End select
      if ThrowErr then Exit Do
      if not rs_work.eof then rs_work.movenext
    loop
   end if
   If Err.Number <> 0 Then
      Wscript.echo String(150,"-")
       WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
         & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
         Report_problem(self_id)
         Wscript.echo Now() & " Porzucam dalsze wykonanie"
         Wscript.echo String(150,"-")
         WScript.Quit
   End If
   IF not ThrowErr then
    Wscript.echo String(150,"-")
    Wscript.echo (now() & " TASK skonczony poprawnie")
    Dim objCon:Set objCon = CreateObject("ADODB.Connection")
    objCon.Open (t_db_connection_String)
      objCon.Execute("Delete from schedule_history where state=7 and id_task='" & self_id & "' and round(cdbl(start),5) in (Select round(cdbl(start),5) from schedule_history where id_task='" & self_id & "' and state in (5,7) group by id_task,round(cdbl(start),5) having count(id)>1 )")
      objCon.Execute("update schedule_history set state=4,real_end=now() where id_task='" & self_id & "' and state in (5,8)")
    objCon.Close
   set objCon=Nothing
   Set Present_work=Nothing
   Else
    Wscript.echo String(150,"-")
    Wscript.echo (now() & " TASK zakonczony z bledami - przejrzyj zapis z logu")
   end if
  end sub
  Private Sub Class_Initialize
    wscript.echo String(150,"#")
    Wscript.echo now() & " Rozpoczynam prace"
    Get_settings
    Get_lst_work
    g_CurrProcessId=CurrProcessId
    if not rs_work.eof then
      WScript.echo(now() & " Rozpoczynam Zadania zwiazane z Taskiem " & rs_work("hdr_name") )
      queue_work
      if not ThrowErr then WScript.echo now() & " Pozytywny Raport wykonania do bazy"
    Else
      Report_problem(self_id)
      WScript.echo(now() &  " Nie istnieja Zadania dla Tasku o ID: " & self_id & "  ....Koncze prace Tasku ")
    end if
  End Sub
END Class
Class Access
 Private objACC
 Public default function init(AccPath,Sub_func,zmienna,self_id)
  on error resume next
      WScript.Echo(now() & " Otwieram Plik Access : " & AccPath & "  makro : " & Sub_func &" zmienna : " & zmienna & " id:" & self_id)
  if FileExists(AccPath) then
    dim rsult:rsult=true
    dim Inslog:Set Inslog = (new Office_instance) (self_id,AccPath)
    Set objACC = CreateObject("Access.Application")
      objACC.Visible = False
      call objACC.OpenCurrentDatabase (AccPath)
      WScript.sleep(5000)
        objACC.DoCmd.SetWarnings False
        if IsNull(zmienna) or zmienna=""  then
          WScript.Echo(now() & " Makro :" & Sub_func & " uruchomione Bez zmiennej")
          rsult=objACC.Run (Sub_func)
        else
          WScript.Echo(now() & " Makro :" & Sub_func & " uruchomione z wartoscia " & zmienna)
          rsult=objACC.Run (Sub_func,zmienna)
        end if
      Dim rsp:rsp=Inslog.Logs_to_console
      WScript.Echo  now() & " Wynik poprawny? " & rsult
      if rsult<>true and rsult<>"" then
        Call Err.Raise(vbObjectError + 10, now() & " Uruchomienie Access", "Blad w " & AccPath & " makro:" & Sub_func & "  :=> Task terminated :=> " & rsult)
        Report_problem(self_id)
      end if
      WScript.Echo(now() & " Zamykam Access")
      objACC.CloseCurrentDatabase
      WScript.Sleep(1*1000)
      objACC.Quit
      WScript.Sleep(2*1000)
      Set objACC=Nothing
      Set Inslog = Nothing
    Else
      WScript.Echo(now() & " Blad - podany plik nie istnieje")
      Report_problem(self_id)
      Wscript.echo("Porzucam dalsze wykonanie...")
      Wscript.echo String(150,"-")
    end if
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
          Report_problem(self_id)
          Wscript.echo("Porzucam dalsze wykonanie...")
          Wscript.echo String(150,"-")
    End If
    set Init=Me
  end function
  Private sub Class_Terminate
    on error resume next
    WScript.sleep (3*1000)
    if not objACC is Nothing Then
      objACC.CloseCurrentDatabase
      objACC.Quit
      Set objACC = Nothing
    end if
  end Sub
END Class
Class Excel
  Private objExcel,objWorkbook
  Public default function init(ExclPath,Sub_func,zmienna,self_id)
    on error resume next
    WScript.Echo(now() & " Otwieram Plik EXCEL : " & ExclPath & "  makro : " & Sub_func &" zmienna : " & zmienna & " id:" & self_id)
    if FileExists(ExclPath) then
      dim rsult
      dim Inslog:Set Inslog = (new Office_instance) (self_id,ExclPath)
      Set objExcel = CreateObject("Excel.Application")
        objExcel.Visible = False
        objExcel.DisplayAlerts = False
        Set objWorkbook = objExcel.Workbooks.Open (ExclPath)
      'call Chek_autom_hello(self_id,ExclPath)
          WScript.sleep(5000)
        if IsNull(zmienna) or zmienna=""  then
          WScript.Echo(now() & " Makro :" & Sub_func & " uruchomione Bez zmiennej")
          rsult=objExcel.Run (Sub_func)
        else
          WScript.Echo(now() & " Makro :" & Sub_func & " uruchomione z wartoscia " & zmienna)
          rsult=objExcel.Run (Sub_func,zmienna)
        end if
        Dim rsp:rsp=Inslog.Logs_to_console
        'if rsp then rsult=true
        WScript.Echo  now() & " Wynik poprawny? " & rsult
        if rsult<>true and rsult<>""  then
          Call Err.Raise(vbObjectError + 10, now() & " Uruchomienie EXCEL", "Blad w " & ExclPath & " makro:" & Sub_func & "  :=> Task terminated :=> " & rsult)
          Report_problem(self_id)
        end if
        WScript.Echo(now() & " Zamykam EXCEL")
        objWorkbook.Close (False)
        WScript.Sleep(1*1000)
        Set objWorkbook = Nothing
        objExcel.Quit
        WScript.Sleep(2*1000)
        Set objExcel = Nothing
        Set Inslog = Nothing
      Else
        WScript.Echo(now() & " Blad - podany plik nie istnieje")
        Report_problem(self_id)
        Wscript.echo("Porzucam dalsze wykonanie...")
        Wscript.echo String(150,"-")
      end if
      If Err.Number <> 0 Then
          WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
            & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
            Report_problem(self_id)
            Wscript.echo Now() & " Porzucam dalsze wykonanie"
            Wscript.echo String(150,"-")
      End If
      set Init=Me
  end function
  Private sub Class_Terminate
    on error resume next
    if not objExcel is Nothing Then
      objExcel.Workbooks.Close True
      Set objWorkbook = Nothing
      objExcel.Quit
      Set objExcel = Nothing
    end if
  end Sub
END Class
Class Word
  Private objWord
  Public default function init(WrdPath,Sub_func,zmienna,self_id)
    on error resume next
    WScript.Echo(now() & " Otwieram Plik WORD : " & WrdPath & "  makro : " & Sub_func & " zmienna : " & zmienna & " id:" & self_id)
    if FileExists(WrdPath) Then
      dim rsult:rsult=true
      dim Inslog:Set Inslog = (new Office_instance) (self_id,WrdPath)
      Set objWord = CreateObject("Word.Application")
      objWord.Visible = False
      objWord.DisplayAlerts = False
      objWord.Documents.Open(WrdPath)
        WScript.sleep(5000)
      if IsNull(zmienna) or zmienna="" then
        WScript.Echo(now() & " Makro :" & Sub_func & " uruchomione Bez zmiennej")
        rsult=objWord.Run (Sub_func)
      else
        WScript.Echo(now() & " Makro :" & Sub_func & " uruchomione z wartoscia " & zmienna)
        rsult=objWord.Run (Sub_func,zmienna)
      end if
      Dim rsp:rsp=Inslog.Logs_to_console
      WScript.Echo  now() & " Wynik poprawny? " & rsult
      if rsult<>true and rsult<>"" then
        Call Err.Raise(vbObjectError + 10, now() & " Uruchomienie WORD", "Blad w " & ExclPath & " makro:" & Sub_func & "  :=> Task terminated :=> " & rsult)
        Report_problem(self_id)
      end if
      WScript.Echo(now() & " Zamykam WORD")
      objWord.Documents.Close
      WScript.Sleep(1*1000)
      objWord.Quit wdDoNotSaveChanges
      WScript.Sleep(2*1000)
      Set objWord = Nothing
      Set Inslog = Nothing
    Else
      WScript.Echo(now() & " Blad - podany plik nie istnieje")
      Report_problem(self_id)
      Wscript.echo("Porzucam dalsze wykonanie...")
      Wscript.echo String(150,"-")
    end if
    If Err.Number <> 0 Then
        WScript.Echo now () & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
            Report_problem(self_id)
            Wscript.echo Now() & " Porzucam dalsze wykonanie"
            Wscript.echo String(150,"-")
    End If
    set Init=Me
  end function
  Private sub Class_Terminate
    on error resume next
    WScript.sleep (2*1000)
    if not objWord is Nothing Then
      objWord.Documents.Close
      objWord.Quit wdDoNotSaveChanges
      Set objWord = Nothing
    end if
  end Sub
END Class
Class Vbscriptx32
  private r32wShell
  Public default function init(VCscrtPath,zmienna,self_id)
    on error resume next
    WScript.Echo(now() & " Otwieram Vbscript w wersji 32 bit : " & VCscrtPath & " parametry: " & zmienna & " id:" & self_id)
    if FileExists(VCscrtPath) Then
      dim rsult
      Set r32wShell = WScript.CreateObject("WScript.Shell")
      if FileExists("C:\Windows\SysWOW64\cscript.exe") then
          Set rsult=r32wShell.exec("C:\Windows\SysWOW64\cscript.exe //NoLogo " & VCscrtPath & " " & zmienna & " TaskID;" & self_id)
      Else
          Set rsult=r32wShell.exec("C:\Windows\system32\cscript.exe //NoLogo " & VCscrtPath & " " & zmienna & " TaskID;" & self_id)
      end if
      Wscript.Echo now() & " Zadanie uruchomione jako proces : " & rsult.ProcessID
      Dim allInput, tryCount
      Do While True
        Dim input
        input = ReadAllFromAny(rsult)
        If -1 = input Then
            If tryCount > 30 And rsult.Status = 1 Then
                 Exit Do
            End If
            tryCount = tryCount + 1
            WScript.Sleep 100
        Else
            allInput = allInput & input
            tryCount = 0
        End If
      Loop
      if len(allInput) > 0 then
        Wscript.echo String(150,"/")
        WScript.Echo  now() & " Informacje z wykonywanego zadania => " & VCscrtPath & " parametry: " & zmienna & " id:" & self_id
        WScript.Echo  allInput
        Wscript.echo String(150,"/")
      end if
      dim stderr1:stderr1=False
      if instr(1,allInput,"StdERR : ")>0 then stderr1=True
      WScript.Echo  now() & " Wynik poprawny aplikacja? => (0)-tak => " & rsult.ExitCode
      WScript.Echo  now() & " Blad wygenerowany z poziomu konsoli? => " & stderr1
      if stderr1=True then
        Call Err.Raise(vbObjectError + 10, now() & " Uruchomienie Vbscriptx32", "Blad w " & VCscrtPath & " parametry: " & zmienna & "  :=> Task terminated :=> Błędy w StdERR" )
        Report_problem(self_id)
      end if
      if rsult.ExitCode<>0 then
        Call Err.Raise(vbObjectError + 10, now() & " Uruchomienie Vbscriptx32", "Blad w " & VCscrtPath & " parametry: " & zmienna & "  :=> Task terminated :=> " & rsult.ExitCode)
        Report_problem(self_id)
      end if
      Set rsult=Nothing
      Set r32wShell=Nothing
    Else
      WScript.Echo(now() & " Blad - podany plik nie istnieje")
      Report_problem(self_id)
      Wscript.echo("Porzucam dalsze wykonanie...")
      Wscript.echo String(150,"-")
    end if
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
          Wscript.echo Now() & " Porzucam dalsze wykonanie"
          Wscript.echo String(150,"-")
          Report_problem(self_id)
    End If
    set Init=Me
  end function
  Private sub Class_Terminate
    WScript.sleep (3*1000)
    on error resume next
    if not r32wShell is Nothing Then
      Set  r32wShell = Nothing
    end if
  end Sub
End class
Class Vbscriptx64
  private r32wShell
  Public default function init(VCscrtPath,zmienna,self_id)
  on error resume next
  WScript.Echo(now() & " Otwieram Vbscript w wersji 64 bit : " & VCscrtPath & " parametry: " & zmienna & " id:" & self_id)
  if FileExists(VCscrtPath) Then
    dim rsult
      Set r32wShell = WScript.CreateObject("WScript.Shell")
      if FileExists("C:\Windows\SysWOW64\cscript.exe") then
        Set rsult=r32wShell.exec("C:\Windows\system32\cscript.exe //U //NoLogo " & VCscrtPath & " " & zmienna & " TaskID;" & self_id)
      Else
        WScript.Echo now() & "Komputer w wersji 32 bit => uruchamiam jako 32 bitowa wersja Cscript"
        Set rsult=r32wShell.exec("C:\Windows\system32\cscript.exe //U //NoLogo " & VCscrtPath & " " & zmienna & " TaskID;" & self_id)
      end if
      Wscript.Echo now() & "Zadanie uruchomine jako proces : " & rsult.ProcessID
      Dim allInput, tryCount
      Do While True
        Dim input
        input = ReadAllFromAny(rsult)
        If -1 = input Then
            If tryCount > 30 And rsult.Status = 1 Then
                 Exit Do
            End If
            tryCount = tryCount + 1
            WScript.Sleep 100
        Else
            allInput = allInput & input
            tryCount = 0
        End If
      Loop
      if len(allInput) > 0 then
        Wscript.echo String(150,"/")
        WScript.Echo  now() & " Informacje z wykonywanego zadania => " & VCscrtPath & " parametry: " & zmienna & " id:" & self_id
        WScript.Echo  allInput
        Wscript.echo String(150,"/")
      end if
      dim stderr:stderr=False
      if instr(1,allInput,"StdERR : ")>0 then stderr=True
      WScript.Echo  now() & " Wynik poprawny aplikacja? => (0)-tak => " & rsult.ExitCode
      WScript.Echo  now() & " Blad wygenerowany z poziomu konsoli? => " & stderr
      if stderr1=True then
        Call Err.Raise(vbObjectError + 10, now() & " Uruchomienie Vbscriptx32", "Blad w " & VCscrtPath & " parametry: " & zmienna & "  :=> Task terminated :=> Błędy w StdERR" )
        Report_problem(self_id)
      end if
      if rsult.ExitCode<>0 then
        Call Err.Raise(vbObjectError + 10, now() & " Uruchomienie Vbscriptx32", "Blad w " & VCscrtPath & " parametry: " & zmienna & "  :=> Task terminated :=> " & rsult.ExitCode)
        Report_problem(self_id)
      end if
      Set rsult=Nothing
      Set r32wShell=Nothing
    Else
      WScript.Echo(now() & " Blad - podany plik nie istnieje")
      Report_problem(self_id)
      Wscript.echo("Porzucam dalsze wykonanie...")
      Wscript.echo String(150,"-")
    end if
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
          Wscript.echo Now() & " Porzucam dalsze wykonanie"
          Wscript.echo String(150,"-")
          Report_problem(self_id)
    End If
    set Init=Me
  end function
  Private sub Class_Terminate
    WScript.sleep (3*1000)
    on error resume next
    if not r32wShell is Nothing Then
      Set  r32wShell = Nothing
    end if
  end Sub
End class
Class Shell
  private r32wShell
  Public default function init(VCscrtPath,zmienna,self_id)
  on error resume next
  WScript.Echo(now() & " Otwieram Plik z konsoli : " & VCscrtPath & " parametry: " & zmienna & " id:" & self_id)
  if FileExists(VCscrtPath) Then
    dim rsult:rsult=0
      Set r32wShell = WScript.CreateObject("WScript.Shell")
      Set rsult=r32wShell.exec( VCscrtPath & " " & zmienna & " TaskID;" & self_id)
      Wscript.Echo now() & " Zadanie uruchomine jako proces : " & rsult.ProcessID
       Dim allInput, tryCount
       Do While True
         Dim input
         input = ReadAllFromAny(rsult)
         If -1 = input Then
             If tryCount > 30 And rsult.Status = 1 Then
                  Exit Do
             End If
             tryCount = tryCount + 1
             WScript.Sleep 100
         Else
             allInput = allInput & input
             tryCount = 0
         End If
       Loop
       if len(allInput) > 0 then
         Wscript.echo String(150,"/")
         WScript.Echo  now() & " Informacje z wykonywanego zadania => " & VCscrtPath & " parametry: " & zmienna & " id:" & self_id
         WScript.Echo  allInput
         Wscript.echo String(150,"/")
       end if
         dim stderr:stderr=False
         if instr(1,allInput,"StdERR : ")>0 then stderr=True
         WScript.Echo  now() & " Wynik poprawny aplikacja? => (0)-tak => " & rsult.ExitCode
         WScript.Echo  now() & " Blad wygenerowany z poziomu konsoli? => " & stderr
         if stderr1=True then
           Call Err.Raise(vbObjectError + 10, now() & " Uruchomienie Vbscriptx32", "Blad w " & VCscrtPath & " parametry: " & zmienna & "  :=> Task terminated :=> Błędy w StdERR" )
           Report_problem(self_id)
         end if
         if rsult.ExitCode<>0 then
           Call Err.Raise(vbObjectError + 10, now() & " Uruchomienie Vbscriptx32", "Blad w " & VCscrtPath & " parametry: " & zmienna & "  :=> Task terminated :=> " & rsult.ExitCode)
           Report_problem(self_id)
         end if
      Set rsult=Nothing
      Set r32wShell=Nothing
    Else
      WScript.Echo(now() & " Blad - podany plik nie istnieje")
      Report_problem(self_id)
      Wscript.echo("Porzucam dalsze wykonanie...")
      Wscript.echo String(150,"-")
    end if
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
          Wscript.echo Now() & " Porzucam dalsze wykonanie"
          Wscript.echo String(150,"-")
          Report_problem(self_id)
    End If
    set Init=Me
  END FUNCTION
  Private sub Class_Terminate
    on error resume next
    WScript.sleep (3*1000)
    if not r32wShell is Nothing Then
      Set  r32wShell = Nothing
    end if
  end Sub
End Class
Class Office_instance
  Dim t_have_module,t_existErrlog,t_self_id,t_file_path,t_cml,t_task_pid,rs_resp
  public Function IS_response
    'Wscript.Echo now() & " Czy była odpowiedz?"
    Dim objCon:Set objCon = CreateObject("ADODB.Connection")
    objCon.Open ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + scriptfullPath & "Log_Helper.mde")
    Set rs_resp= CreateObject("ADODB.Recordset")
    with rs_resp
      .ActiveConnection = objCon
      .CursorLocation = 3
      .LockType=4
      .open "Select * from task_procc where app_path='" & t_file_path & "' and task_pid=" & t_task_pid
    end with
    Set rs_resp.ActiveConnection = Nothing
    objCon.Close
    Set objCon = Nothing
    if rs_resp("pid")<>0 Then
      IS_response=True
    Else
      IS_response=False
    End if
    WScript.Echo  now() & " Sprawdzam czy jest odpowiedz z aplikacji office ? => " & IS_response
  end Function
  public Function Logs_to_console
    'Wscript.Echo now() & " Sprawdzam odpowiedz instancji office"
    if IS_response Then
      on error resume next
      if FileExists(rs_resp("curr_log")) Then
        Wscript.echo String(150,"/")
        WScript.Echo now() & " Informacje z wykonywanego zadania "
        dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
        dim f:Set f = fso.OpenTextFile(rs_resp("curr_log"),1,False,-1)
          Do Until f.AtEndOfStream
            WScript.Echo f.ReadLine
          Loop
        f.Close
        Set fso=Nothing
        Wscript.echo String(150,"/")
        if err.number<>0 then wscript.echo now() & " Plik logu nie istnieje" : err.clear
      Else
        wscript.echo now() & " Plik logu nie istnieje"
      End if
    Else
        wscript.echo now() & " Plik logu nie istnieje"
    End if
    if rs_resp("iserr") then ThrowErr=True
    Logs_to_console=rs_resp("iserr")
  end Function
  private Function check_is_closed(timemup)
    dim proccexist:proccexist=False
    dim strComputer:strComputer = "."
    dim objWMIService:Set objWMIService = GetObject("winmgmts:" _
      & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
      dim objitem:DIm colProcesses:Set colProcesses = objWMIService.ExecQuery _
        ("SELECT * FROM Win32_Process WHERE ProcessID =" & rs_resp("pid").value)
      If colProcesses.Count > 0 Then
        For Each objitem In colProcesses
          if timemup Then
            WScript.Echo now() & " Wymuszam zamkniecie instancji Office - aplikacja nie zamyka sie samoczynnie => prosze sprawdzic poprawnosc napisanego makra ..."
            objitem.terminate
          Else
            proccexist=True
          end if
        Next
      end if
    check_is_closed=proccexist
  end function
  Private sub Class_Terminate
    if not isnull(rs_resp("pid").value) then
      if rs_resp("pid").value<>0 then
        wscript.Echo now() & " Sprawdzam czy process office nadal istnieje"
        dim tmpchk,bl:dim endloop:endloop=30:bl=False
        do
          wscript.Sleep (3*1000)
          tmpchk=check_is_closed(bl)
          if not tmpchk then exit do
          endloop=endloop-1
          if endloop<0 then bl=True
        loop

        if FileExists(rs_resp("curr_log")) then
            Wscript.Echo now() & " Usuwam logi z instancji office => " & rs_resp("curr_log")
            dim filesys:Set filesys = CreateObject("Scripting.FileSystemObject")
            filesys.DeleteFile rs_resp("curr_log")
            Set filesys= Nothing
        end if
      end if
    end if
    Dim objCon:Set objCon = CreateObject("ADODB.Connection")
    Wscript.Echo now() & " Usuwam zapis instancji w tymczasowej bazie => " & rs_resp("curr_log")
    objCon.Open ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + scriptfullPath & "Log_Helper.mde")
    objCon.Execute("Delete from task_procc where app_path='" & t_file_path & "' and task_pid=" & t_task_pid)
    objCon.Close
    Wscript.Echo now() & " Instancja office zamknieta => " & rs_resp("curr_log")
    Set objCon= Nothing
  End Sub
  Private sub check_HelpersDB
    on error resume next
    if not FileExists(scriptfullPath & "Log_Helper.mde") Then
      dim dbfile:Set dbfile =CreateObject("ADOX.Catalog")
      dbfile.Create("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + scriptfullPath & "Log_Helper.mde")
      WScript.Echo(now() & " Empty Helper Database Created  : => " + scriptfullPath & "Log_Helper.mde")
      Set dbfile=nothing
      Dim oDb1 : Set oDb1 = CreateObject("ADODB.CONNECTION")
      oDb1.Open ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + scriptfullPath & "Log_Helper.mde")
      oDb1.Execute("Create Table task_procc (pid INTEGER,app_start DATETIME,task_pid INTEGER,task_id GUID,app_path VARCHAR(250),curr_log VARCHAR(250),iserr BIT)")
      oDb1.Close
      set oDb1=Nothing
    end if
    If Err.Number <> 0 Then
        WScript.Echo now () & " Blad " & CStr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
        WScript.Quit
    End If
    On Error Goto 0
  end sub
  Public default function init(self_id,file_path)
    t_self_id=self_id
    t_file_path=file_path
    t_task_pid=g_CurrProcessId
    check_HelpersDB
    do while check_use_file(file_path)
      WScript.echo (now() & " Czekam na zwolnienie procesu")
      WScript.Sleep(1*1000)
    loop
    call Report_start_ActiveX(self_id,file_path)
    set Init=Me
  end Function
  Sub Report_start_ActiveX(task_id,file_path)
    WScript.Echo  now() & " Raport rozpoczecia zadania automatyzacji ActiveX => Data Source=" + scriptfullPath & "Log_Helper.mde"
    Dim objCon:Set objCon = CreateObject("ADODB.Connection")
      objCon.Open ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + scriptfullPath & "Log_Helper.mde")
        objCon.Execute("INSERT into task_procc (pid,app_start,task_pid,task_id,app_path,curr_log,iserr) values (0,now()," & t_task_pid & ",'" & task_id & "','"& file_path & "','',0)")
        objCon.Close
      set objCon=Nothing
  end sub
  private function Get_path_from_CML (full_path)
    Get_path_from_CML = Replace(Mid(full_path, InStr(1, full_path, " ") + 1),Mid(full_path,InStrRev(full_path, "\")),"")
  end Function
  private Function Filename_EXIST_in_helper(file_path,serv_path)
    on error resume next
    WScript.Echo  now() & " Sprawdzam uruchomienie pliku => '" & file_path & "' przez serwis => " & serv_path
    if FileExists(serv_path & "\Log_Helper.mde") then
      Dim objCon:Set objCon = CreateObject("ADODB.Connection")
      Dim rs_exist:Set rs_exist= CreateObject("ADODB.Recordset")
      WScript.Echo  now() & " Sprawdzam w tymczasowej bazie : " & serv_path & "\Log_Helper.mde"
      objCon.Open ("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" & serv_path & "\Log_Helper.mde")
      If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
      End If
      with rs_exist
        .ActiveConnection = objCon
        .CursorLocation = 3
        .LockType=4
        .open "Select * from task_procc where app_path='" & file_path & "' and pid=0 and app_start>(Now()-(1/24/120))"
      end with
      If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
      End If
      WScript.Echo  now() & " Pobrano dane => " & serv_path & "\Log_Helper.mde => " & rs_exist.eof
        Set rs_exist.ActiveConnection = Nothing
          objCon.close
        Set objCon=Nothing
      if rs_exist.eof then
        Filename_EXIST_in_helper = False
      else
        Filename_EXIST_in_helper = True
      end if
      set rs_exist=Nothing
    Else
      Filename_EXIST_in_helper=False
    end if
    If Err.Number <> 0 Then
      WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
        & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    End If
  end function
  private Function check_use_file (file_path)
    on error resume next
    'WScript.Echo  now() & " Czy plik " & file_path & " nie zostal odpalony przez inny task ? "
    Dim objWMIService ,colProcesses,objitem
    dim strComputer:strComputer = "."
    Set objWMIService = GetObject("winmgmts:" _
      & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
      Set colProcesses = objWMIService.ExecQuery _
        ("SELECT * FROM Win32_Process WHERE Name = 'cscript.exe'")
        dim cnt:cnt=0
        dim cml,fil_ext
        If colProcesses.Count > 0 Then
        For Each objitem In colProcesses
          if instr(1,objitem.CommandLine,"serv_maintain.vbs" )>0 then
          'if instr(1,objitem.CommandLine,"TaskID" )>0 then
            t_cml=Get_path_from_CML(objitem.CommandLine)
            Wscript.echo(NOw() & " Sprawdzam lokalizacje => " & cml)
            fil_ext=Filename_EXIST_in_helper(file_path,t_cml)
            if fil_ext then cnt=cnt+1
            end if
        Next
        end if
        if cnt>0 then
          WScript.Echo  now() & " Plik " & file_path & " jest odpalony przez inny task => czekam ..."
          check_use_file=True
        Else
          WScript.Echo  now() & " Plik " & file_path & " nie zostal odpalony przez inny task => mozna dzialac ..."
          check_use_file=False
        end if
      Set colProcesses=Nothing
      Set objWMIService=Nothing
      If Err.Number <> 0 Then
          WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
      End If
      on Error goto 0
  end function
end class
Function ReadAllFromAny(oExec)
     If Not oExec.StdOut.AtEndOfStream Then
          ReadAllFromAny = "StdOut : " & oExec.StdOut.ReadAll
          Exit Function
     End If
     If Not oExec.StdErr.AtEndOfStream Then
          ReadAllFromAny ="StdERR : " & oExec.StdErr.ReadAll
          Exit Function
     End If
     ReadAllFromAny = -1
End Function
Function FileExists(FilePath)
   dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(FilePath) Then
      FileExists=CBool(1)
    Else
      FileExists=CBool(0)
    End If
    Wscript.Echo now() & " Plik istnieje " & FilePath & " wynik =>" & FileExists
    Set fso = Nothing
 End Function
Function Num2digit(num)
   If(Len(num)=1) Then
       Num2digit="0"&num
   Else
       Num2digit=num
   End If
 End Function
Function DTE_form(myDate)
   dim d,m,y
   d = Num2digit(Day(myDate))
   m = Num2digit(Month(myDate))
   y = Year(myDate)
   DTE_form= y &  m  & d
 End Function
Sub Report_problem (task_id)
  on error resume next
  ThrowErr=True
  WScript.Echo  now() & " Raport nie poprawnego wykonania tasku do bazy danych..."
  Dim objCon:Set objCon = CreateObject("ADODB.Connection")
    objCon.Open (ext_conn)
      objCon.Execute("update schedule_history set state=6,real_end="  & replace(cdbl(cdate(now())),",",".") &  " where id_task='" & task_id & "' and state=5")
    objCon.Close
  set objCon=Nothing
  If Err.Number <> 0 Then
      WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
        & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
        Wscript.echo Now() & " Porzucam dalsze wykonanie"
        Wscript.echo String(150,"-")
        Report_problem(self_id)
        WScript.Quit
  End If
end sub
Function FolderExists(Path)
  dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FolderExists(Path) Then
    FolderExists=CBool(1)
  Else
   FolderExists=CBool(0)
   End If
 Set fso = Nothing
End Function
