Option Explicit
Const HIDDEN_WINDOW = 0
Const HKEY_LOCAL_MACHINE = &H80000002
Const Create_New_Process_Group =512
dim srcript_name:srcript_name=WScript.ScriptName
dim scriptfullPath:scriptfullPath = replace(WScript.ScriptFullName,srcript_name,"")
dim Sched : set Sched= new Scheduler
Class Settings
  'Time intervals
  Private t_main_loop_interval,t_refr_task_fromDtbase_intrv,t_refr_settings_intrv
  'Database settings'
  Private t_db_connection_String,t_db_fulpath,t_db_connection_crea_String,t_db_path_forLogs
  public Property Get Db_connection_crea_String
      Db_connection_crea_String=t_db_connection_crea_String
    End Property
  Public Property Get Db_connection
     Db_connection = t_db_connection_String
  End Property
  Public Property Get Db_fulpath
     Db_fulpath = t_db_fulpath
  End Property
  Public Property Get Main_loop_interval
    Main_loop_interval = t_main_loop_interval
  End Property
  Public Property Get Db_path_forLogs
    Db_path_forLogs = t_db_path_forLogs
  End Property
  Public Property Get Refr_task_fromDtbase_intrv
    Refr_task_fromDtbase_intrv = t_refr_task_fromDtbase_intrv
  End Property
  Public Property Get Refr_settings_intrv
      Refr_settings_intrv = t_refr_settings_intrv
  End Property
  'get dta from XML
  Private sub regenerate_calendars
    On Error resume Next
    Dim oDb : Set oDb = CreateObject("ADODB.CONNECTION")
    oDb.Open t_db_connection_String
    Dim oRs : Set oRs = CreateObject("ADODB.Recordset")
    with oRs
      .ActiveConnection = oDb
      .CursorLocation = 3
      .LockType=4
      .open "Select * FROM calendar_hdr"
    end with
    Set oRs.ActiveConnection = Nothing
    oDb.close
    dim que
    Do While NOT oRs.eof
      set que = (new Calendar) (t_db_connection_String,oRs("id"))
      que.generate
      oRs.movenext
    Loop
    set que=Nothing
    set oRs=Nothing
    set oDb=Nothing
    If Err.Number <> 0 Then
        WScript.Echo now () & " Blad " & CStr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    Else
      WScript.Echo(now() & " Calendars object created")
    End If
  end sub
  Public sub get_XML
    On Error resume Next
    WScript.Echo(now() & " Get Settings for service")
    'Check is file ?'
    if not FileExists(scriptfullPath & "settings.xml") Then
      Call Err.Raise(vbObjectError + 10, srcript_name, "No 'settings.xml'=> in FilePath:'" & scriptfullPath & "':=> Program terminated")
    Else
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
          Call Err.Raise(vbObjectError + 10,"You have error in XML => " & scriptfullPath & "settings.xml => " + myErr.reason)
        else
          t_db_connection_String=settingsXML.selectSingleNode("scheduler/db_connection_String").text
          t_db_connection_crea_String=settingsXML.selectSingleNode("scheduler/db_create_connection_String").text
          t_refr_task_fromDtbase_intrv=settingsXML.selectSingleNode("scheduler/refr_task_fromDtbase_intrv").text
          t_refr_settings_intrv=settingsXML.selectSingleNode("scheduler/refr_settings_intrv").text
          t_main_loop_interval=settingsXML.selectSingleNode("scheduler/main_loop_interval").text
          t_db_path_forLogs=settingsXML.selectSingleNode("scheduler/db_path_forLogs").text
          t_db_fulpath=settingsXML.selectSingleNode("scheduler/db_fulpath").text
        End if
      WScript.Echo(now() & " Settings succesfuly loaded => " & scriptfullPath & "settings.xml")
      set settingsXML=nothing
      If Err.Number <> 0 Then
          WScript.Echo now () & " Blad " & CStr(Err.Number) & " wygenerowany przez " _
            & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
          WScript.Quit
      End If
    End if
  End Sub
  Private sub check_Database
    On Error resume Next
    WScript.Echo(now() & " Start Checking Database")
    if not FileExists(t_db_fulpath) Then
      WScript.Echo(now() & " Database don't exist => Create NEW : => " + t_db_fulpath)
      dim sPath:sPath = replace(replace(t_db_fulpath,Mid(t_db_fulpath, InStrRev(t_db_fulpath, "\") + 1) & "\",""),"\" & Mid(t_db_fulpath, InStrRev(t_db_fulpath, "\") + 1),"")
      if not FolderExists(sPath) then
        WScript.Echo(now() & " Folder don't exist => Create NEW : => " + sPath)
        Create_folder(sPath)
      end if
      WScript.Echo(now() & " Db_path " & t_db_fulpath)
      dim dbfile:Set dbfile =CreateObject("ADOX.Catalog")
      dbfile.Create("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + t_db_fulpath)
      If Err.Number <> 0 Then
          WScript.Echo now () & " Blad " & CStr(Err.Number) & " wygenerowany przez " _
            & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
          WScript.Quit
      End If
      WScript.Echo(now() & " Empty Database Created  : => " + t_db_fulpath)
      Set dbfile=nothing
    End if
    WScript.Echo(now() & " Check Schema Database for neded tables => " + t_db_fulpath)
    'check for neded tables and fields'
    Const adSchemaTables = 20
    Dim oDb : Set oDb = CreateObject("ADODB.CONNECTION")
    oDb.Open t_db_connection_String
    Dim oRs : Set oRs = oDb.OpenSchema(adSchemaTables)
    WScript.Echo(now() & " Get Settings for service")
    'Check is file ?'
    if not FileExists(scriptfullPath & "schemaDB.xml") Then
        Wscript.Echo "No 'schemaDB.xml'=> in FilePath:'" & scriptfullPath & "':=> Program terminated"
        Call Err.Raise(vbObjectError + 10, srcript_name, "No 'schemaDB.xml'=> in FilePath:'" & scriptfullPath & "':=> Program terminated")
    Else
      dim schemaXML:Set schemaXML = CreateObject("MSXML2.DOMDocument")
        With schemaXML
          .SetProperty "SelectionLanguage", "XPath"
          .SetProperty "ProhibitDTD", False
          .ValidateOnParse =  True
          .Async = False
          .Load scriptfullPath & "schemaDB.xml"
        End With
        if schemaXML.parseError.errorCode<>0 then
          dim myErr: set myErr= schemaXML.parseError
          Wscript.Echo "You have error in XML => " & scriptfullPath & "schemaDB.xml => " + myErr.reason
          Call Err.Raise (vbObjectError + 10,"You have error in XML => " & scriptfullPath & "schemaDB.xml => " + myErr.reason)
        else
          dim i,j,k
          Dim objNodeList:Set objNodeList = schemaXML.SelectNodes("schema/table")
          Dim oDb1 : Set oDb1 = CreateObject("ADODB.CONNECTION")
          oDb1.Open t_db_connection_crea_String
          If Err.Number <> 0 Then
              WScript.Echo now () & " Blad " & CStr(Err.Number) & " wygenerowany przez " _
                & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
              WScript.Quit
          End If
          WScript.Echo now () & " Schema of neded table(s) in file => contains " & objNodeList.length & " objects => Check database structure => " +  t_db_fulpath
          for i=0 to objNodeList.length-1
            oRs.filter = "TABLE_NAME='" + objNodeList(i).selectSingleNode("name").text + "'"
            if oRs.EOF Then
            WScript.Echo now () & " Checking '" & objNodeList(i).selectSingleNode("name").text + "' table exist in DB " + t_db_fulpath + " => " & not(ors.eof)
            'table not in database create new'
              Dim Flds:set Flds = objNodeList(i).SelectNodes("field")
              if Flds.length>0 then
                dim quer:quer="Create Table " & objNodeList(i).selectSingleNode("name").text & " ("
                for j=0 to Flds.length-1
                  quer=quer & Flds(j).selectSingleNode("fname").text & " " & Flds(j).selectSingleNode("type").text & ","
                next
                quer= mid(quer,1,len(quer)-1) & ")"
                'wscript.echo quer
                with oDb1
                  ''.Open t_db_connection_crea_String
                  .Execute(quer)
                  '.close
                End with
                wscript.echo Now() & " " & quer & " Execute result => " & True
                dim tmp: set tmp= objNodeList(i).SelectNodes("index")
                for k=0 to tmp.length-1
                  with oDb1
                    ''.Open t_db_connection_crea_String
                    .Execute(tmp(k).text)
                    '.close
                  End with
                  If Not tmp Is Nothing Then  wscript.echo Now() & " " & tmp(k).text & " Execute result => " & True
                next
                set tmp= objNodeList(i).SelectNodes("insert")
                for k=0 to tmp.length-1
                  with oDb1
                    '.Open t_db_connection_String
                    .Execute(unique_replace(tmp(k).text))
                    'wscript.sleep (1*200)
                    '.close
                  End with
                  If Not tmp Is Nothing Then  wscript.echo Now() & " " & tmp(k).text & " Execute result => " & True
                next
              end if
            Else
            WScript.Echo now () & " Checking '" & objNodeList(i).selectSingleNode("name").text + "' table exist in DB => " & not(ors.eof)
            end if
          next
          If Err.Number <> 0 Then
              WScript.Echo now () & " Blad " & CStr(Err.Number) & " wygenerowany przez " _
                & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
              WScript.Quit
          Else
              WScript.Echo(now() & " Settings succesfuly loaded => " & scriptfullPath & "schemaDB.xml")
          End If

          oRs.filter=0
          oDb.close
          On Error Goto 0
          set oDb1=nothing
          set oRs=Nothing
          set oDb=Nothing
        End if
      set schemaXML=nothing
    End if
    If Err.Number <> 0 Then
        WScript.Echo now () & " Blad " & CStr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
        WScript.Quit
    End If
  End sub
  Private sub check_HelpersDB
    On Error resume Next
    WScript.Echo(now() & " Check is Helper Database Exist => " & scriptfullPath & "Log_Helper.mde")
    if not FileExists(scriptfullPath & "Log_Helper.mde") Then
      WScript.Echo(now() & " Db_path " & t_db_fulpath)
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
  Private Sub Class_Initialize
    get_XML
    check_Database
    check_HelpersDB
    regenerate_calendars
  End Sub
End Class
Class Calendar_Queue
    Private rs_type_queue
    Private rs_len,rs,first_day,last_day
    Private objCon,cal_id
    Public Sub Next_day
      if rs_len>1 then
        if rs_type_queue("period_pos")=rs_len Then
          rs_type_queue.movefirst
        else
          rs_type_queue.movenext
        end if
      End if
    End Sub
    Public Property Get Calendar_id
       Calendar_id =rs_type_queue("calendar_id")
    End Property
    Public Property Get Period_pos
       Period_pos =rs_type_queue("period_pos")
    End Property
    Public Property Get Queue_counter
       Queue_counter =rs_type_queue("queue_counter")
    End Property
    Public Property Get Day_type
       Day_type =rs_type_queue("day_type")
    End Property
    Public default function init(Dbpath,cal_ids)
      On Error resume Next
      WScript.echo Now() & (" Obiekt kolejka Kalendarza : " & cal_ids)
      Dim objCon:Set objCon = CreateObject("ADODB.Connection")
      objCon.Open (Dbpath)
      Set rs_type_queue = CreateObject("ADODB.Recordset")
      with rs_type_queue
        .ActiveConnection = objCon
        .CursorLocation = 3
        .LockType=4
        .open "Select * FROM day_type_queue where calendar_id='" & cal_ids &  "' order by period_pos"
      End with
      Set rs_type_queue.ActiveConnection = Nothing
      if check_integr_data Then
        Set rs_type_queue.ActiveConnection = objCon
        rs_type_queue.updatebatch
        Set rs_type_queue.ActiveConnection = Nothing
      end if
      objCon.close
      If Err.Number <> 0 Then
          WScript.Echo now () & " Blad " & CStr(Err.Number) & " wygenerowany przez " _
            & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
      End If
      set objCon=Nothing
      On Error Goto 0
      set Init=Me
    end Function
    'check data in table'
    Private function check_integr_data
      On Error resume Next
        Wscript.echo Now() & (" Sprawdzenie poprawnosci kolejki")
        rs_len=rs_type_queue.RecordCount
        rs_type_queue.movefirst
        first_day=rs_type_queue("queue_counter")
        Dim tm_day,tm_count,data_changed
        data_changed=false
        tm_day=first_day
        tm_count=1
        Do While NOT rs_type_queue.eof
          if rs_type_queue("period_pos")<>tm_count then
            rs_type_queue("period_pos")=tm_count
            data_changed=true
          end if
          last_day=rs_type_queue("queue_counter")
          rs_type_queue.movenext
          tm_count=tm_count+1
          tm_day=tm_day+1
          if tm_day=8 then tm_day=1
          if tm_count<=rs_len Then
            if rs_type_queue("queue_counter")<>tm_day then
              rs_type_queue("queue_counter")=tm_day
              data_changed=true
            end if
          End if
        Loop
        If Err.Number <> 0 Then
            WScript.Echo now () & " Blad " & CStr(Err.Number) & " wygenerowany przez " _
              & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
        End If
        On Error Goto 0
        check_integr_data=data_changed
    end function
    Private sub Class_Terminate
      set rs_type_queue=nothing
    End Sub
    Public function get_first_calendar_day(start_date)
        On Error resume Next
        dim tmp_day,cnt
        dim day_start: day_start =weekday(cdate(start_date))
        rs_type_queue.filter=("queue_counter=" & day_start)
        if rs_type_queue.eof Then
          if rs_len>1 then
            rs_type_queue.filter=0
            rs_type_queue.movelast
            cnt=rs_len
            tmp_day=rs_type_queue("queue_counter")
            Do while not day_start=tmp_day
              tmp_day=tmp_day-1
              if tmp_day=0 then tmp_day=7
              cnt=cnt-1
              if cnt=0 Then
                rs_type_queue.movelast
                cnt=rs_len
              Else
                rs_type_queue.moveprevious
              end if
            Loop
          Else
            if rs_len=0 then
                'code for generate default day'
            else
                rs_type_queue.Filter = 0
                tmp_day=rs_type_queue("queue_counter")
            end if
          End if
        Else
          tmp_day=rs_type_queue("queue_counter")
          rs_type_queue.filter=0
          rs_type_queue.movefirst
          Do while not rs_type_queue("queue_counter")=tmp_day
            rs_type_queue.movenext
          loop
        end if
        If Err.Number <> 0 Then
            WScript.Echo now () & " Blad " & CStr(Err.Number) & " wygenerowany przez " _
              & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
        End If
        On Error Goto 0
        get_first_calendar_day= tmp_day
    End Function
End Class
Class Scheduler
  Private t_main_loop_interval,t_refr_task_fromDtbase_intrv,t_refr_settings_intrv
  Private t_db_connection_String,t_db_path_forLogs,serv_control
  Private t_count_activeTask,t_enviroments,tmp_wmiRctset
  Private t_work_to_do,t_old_work_to_do,t_db_fulpath
  Private function Check_serv_maintainExist
    On Error resume Next
    Dim objWMIService ,colProcesses,objitem
    dim strComputer:strComputer = "."
    Set objWMIService = GetObject("winmgmts:" _
      & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
      Set colProcesses = objWMIService.ExecQuery _
      ("SELECT * FROM Win32_Process WHERE Name = 'cscript.exe'")
      dim cnt:cnt=0
      If colProcesses.Count > 0 Then
        For Each objitem In colProcesses
          if instr(1,objitem.CommandLine,scriptfullPath & "serv_maintain.vbs")>0 then cnt=cnt+1
        Next
      end if
      If Err.Number <> 0 Then
        WScript.Echo now () & " Blad " & CStr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
      End If
      On Error Goto 0
      if cnt>0 then
        Check_serv_maintainExist=True
      Else
        Check_serv_maintainExist=False
      end if
    Set colProcesses=Nothing
    Set objWMIService=Nothing
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end Function
  Private sub Report_task_Notdoo(Id,rlStart)
    on error resume Next
    WScript.Echo(now() & " Log do bazy nie wykonanie zadania w czasie => " & id & " z godziny => " & rlStart )
    Dim objCon:Set objCon = CreateObject("ADODB.Connection")
    objCon.Open (t_db_connection_String)
      objCon.Execute("INSERT INTO schedule_history (id,id_task,start,state,real_start,real_end) values ('" & genGUID & "','" & id & "'," & replace(cdbl(cdate(rlStart)),",",".") & ",10," & replace(cdbl(cdate(now())),",",".") & ", " & replace(cdbl(cdate(now())),",",".") & " )")
    objCon.Close
    set objCon=Nothing
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end sub
  Private sub Log_end_of_time_forTask
    on error resume next
   dim chk:chk=False
   dim chk1 :chk1=False
   WScript.Echo(now() & " Sprawdzam przeterminowane zadania")
   With t_old_work_to_do
    .filter="status=2"
    'WScript.Echo(now() & " Sprawdzam przeterminowane zadania krok1")
      if not .EOF Then
        'WScript.Echo(now() & " Sprawdzam przeterminowane zadania krok2")
        t_work_to_do.filter="status<3 or status=7"
        t_work_to_do.Sort="start"
        .Sort="start"
        .movefirst
        dim time_off:time_off=cdbl(now()-(60/1440)+(5/1440))
        DO until .EOF or chk1=True
        'WScript.Echo(now() & " Sprawdzam przeterminowane zadania krok3")
          if CDbl(cdate(.fields("start"))) < time_off then
            WScript.Echo(now() & " Task zagrozony nie wykonaniem => " & .fields("id") & " z planowanej daty => " & .fields("start"))
            t_work_to_do.filter="id='" & .fields("id") & "'"
            t_work_to_do.Sort="start"

            if not t_work_to_do.eof then
                WScript.Echo(now() & " Sprawdzam przeterminowane zadania  => " & t_work_to_do("id"))
                chk=False
                t_work_to_do.movefirst
                DO until t_work_to_do.eof or chk=True
                  WScript.Echo(now() & " Sprawdzam przeterminowane zadania krok4 => " & t_work_to_do("start") & " => " & t_work_to_do("id"))
                  if round(cdbl(t_work_to_do("start")),5)=round(cdbl(.fields("start")),5) then
                    WScript.Echo(now() & " Task znaleziony w bieżacych danych")
                    chk=True
                  end if
                  t_work_to_do.movenext
                LOOP
                if chk=False then Call Report_task_Notdoo(.fields("id"),.fields("start"))
            Else
               call Report_task_Notdoo(.fields("id"),.fields("start"))
            end if
          Else
          WScript.Echo(now() & " Koncze sprawdzanie przeterminowanych taskow")
           chk1=True
          end if
          .movenext
        LOOP
      end if
      .filter=0
      t_work_to_do.filter=0
    end with
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end sub
  private sub Get_list_of_schedule
    on error resume next
    dim have_old_tasklst: have_old_tasklst=False
    WScript.Echo(now() & " Pobieram dane z bazy danych")
    Dim objCon:Set objCon = CreateObject("ADODB.Connection")
    objCon.Open (t_db_connection_String)
    if not IsEmpty(t_work_to_do)  then
      Set t_old_work_to_do = t_work_to_do.clone(1)
      have_old_tasklst=True
    end if
    Set t_work_to_do = CreateObject("ADODB.Recordset")
    with t_work_to_do
      .ActiveConnection = objCon
      .CursorLocation = 3
      .LockType=4
      .open "SELECT TOP " & t_count_activeTask+100 & " aq.id, aq.start, aq.real_start, aq.name, aq.description, aq.time_exec, aq.priority, aq.status,c.state as st_enum FROM (select * from (Select a.id, a.start,b.real_start, a.name, a.description,a.priority ,iif(isnull(b.state),a.state,b.state) as status,a.time_exec from (SELECT hdr.id, w.work_day+tim.start_hour AS start, hdr.name, hdr.description, hdr.state,sch.time_exec,max(sch.prioity) as priority FROM schedule AS sch,(select id_shed, start_hour from (SELECT id_shed, start_hour FROM schedule_timer) UNION (SELECT id_shed, start_hour+1 FROM schedule_timer) )  AS tim,task_hdr AS hdr,(SELECT c_w.calendar_id, c_w.cal_counter, c_w.work_day, c_w.work_day+IIf(isnull(tp.start_hour),0,tp.start_hour) AS Start_work,c_w.work_day+IIf(isnull(tp.end_hour),1,tp.end_hour) AS End_work FROM calendar_wrk AS c_w,calendar_hdr AS chdr,calendar_types_day AS tp WHERE chdr.id=c_w.calendar_id and chdr.state<3 and c_w.work_day>= date()-1 and tp.day_id=c_w.day_id ORDER BY cal_counter)  AS w WHERE (((hdr.id)=[sch].[id_task]) AND ((hdr.state)<3) AND ((w.calendar_id)=[sch].[calendar_id]) AND (([w].[work_day]+[tim].[start_hour])>(select Now()-(60/1440) from dual ) And ([w].[work_day]+[tim].[start_hour]) Between [w].[Start_work] And [w].[End_work]) AND ((sch.state)<3) AND ((sch.id_shed)=[tim].[id_shed])) GROUP BY hdr.id, w.work_day+tim.start_hour, hdr.name, hdr.description, hdr.state,sch.time_exec ORDER BY w.work_day+tim.start_hour) as a left join (SELECT * from schedule_history where start>(select Now()-(60/1440) from dual ) ) as b on round(cdbl(b.start),5)=round(cdbl(a.start),5) and a.id=b.id_task where iif(isnull(b.state),a.state,b.state)<>7) UNION (select aql.id, aql.start,aql.real_start, aql.name, aql.description,aql.priority,iif(ab.cnt>1 ,5,aql.state) as status,0  from (Select a.id, b.start,b.real_start, a.name, a.description,100 as priority,b.state from task_hdr AS a,schedule_history as b where b.start>(select Now()-(60/1440) from dual ) and b.state=7 and b.id_task=a.id)  AS aql left join (Select id_task,round(cdbl(start),5) as rdat,count(id) as cnt from schedule_history where state in (5,7) group by id_task,round(cdbl(start),5))  as ab on  aql.id=ab.id_task and round(cdbl(aql.start),5)=ab.rdat) UNION (Select a.id, d.start,d.real_start, a.name, a.description,100 ,d.state as status,0 FROM task_hdr AS a,(select b.id_task,b.start,b.real_start,b.state,c.time_exec from schedule_history as b left join schedule as c on b.id_task=c.id_task) as d where d.state in (5,8) and d.id_task=a.id and d.start<=(select Now()-(60/1440) from dual )))  AS aq, dbtaskenum AS c WHERE (((c.id)=[aq].[status])) ORDER BY aq.start, aq.priority DESC;"
    end with
    Set t_work_to_do.ActiveConnection = Nothing
    objCon.close
    Set objCon=Nothing
    if have_old_tasklst then Log_end_of_time_forTask
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  End Sub
  Private sub Count_activeTasks
    on error resume Next
    Dim strComputer:strComputer = "."
    Dim objWMIService:Set objWMIService = GetObject("winmgmts:" _
      & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Dim colProcesses:Set colProcesses = objWMIService.ExecQuery _
      ("SELECT * FROM Win32_Process WHERE Name = 'cscript.exe'")
    dim cnt:cnt=0
    If colProcesses.Count > 0 Then
      Dim objitem
      For Each objitem In colProcesses
        if instr(1,objitem.CommandLine,"TaskID;")>0 and instr(1,objitem.CommandLine,"Job_ID;")=0 then cnt=cnt+1
      Next
    end if
    t_count_activeTask=cnt
    Set colProcesses=Nothing
    Set objWMIService=Nothing
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end Sub
  private Function Count_Db_activeTasks
    on error resume Next
      t_work_to_do.filter="status=5"
      Count_Db_activeTasks=t_work_to_do.RecordCount
      If Err.Number <> 0 Then
          WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
            & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
      end if
  end Function
  Private sub Delete_unuseful_tasks
    on error resume Next
    IF t_count_activeTask<>Count_Db_activeTasks or t_count_activeTask=0  Then
      dim objFile
      dim objFSO :Set objFSO = CreateObject("Scripting.FileSystemObject")
      dim objFolder: Set objFolder = objFSO.GetFolder(scriptfullPath)
      dim colFiles:Set colFiles = objFolder.Files
      dim chk:chk =False
      For Each objFile in colFiles
          If UCase(objFSO.GetExtensionName(objFile.name)) = "VBS" Then
              if InStr(1,Left(objFile.Name,8),"TaskID;")>0 then
                WScript.Echo (Now() & " Uruchomione Taski w systemie => " & t_count_activeTask & " => Zaczynam czyszczenie po wykonanych pracach... ")
                Dim pos_id_txt:pos_id_txt=InStr(1,objFile.Name,"TaskID;")
                Dim pos_end_txt:pos_end_txt= InStr(1,objFile.Name,".vbs")
                if not Check_MainTASK_IsWork(mid(objFile.Name,pos_id_txt+7,pos_end_txt-(pos_id_txt+7))) then
                  WScript.Echo (now () & " Usuwam plik nie uzywanego tasku : " & scriptfullPath & "\" & objFile.name)
                  objfile.Delete
                  chk=True
                end if
              end if
          End If
      Next
      Set objFSO = Nothing
      Set objFolder = Nothing
      Set colFiles = Nothing
      if chk then checkOLDER_reportedDB_not_existWMI
    end if
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end sub
  Private sub Report_start_work(TaskID,TaskName,start,stat)
    on error resume Next
    Dim objCon:Set objCon = CreateObject("ADODB.Connection")
    objCon.Open (t_db_connection_String)
    if stat=7 Then
      objCon.Execute("Update schedule_history set state=5,real_start="  & replace(cdbl(cdate(now())),",",".") &  " where id_task='" & TaskID & "' and state=7 and cdbl(start)=" & Replace(CDbl(cdate(start)), ",", "."))
    Else
      objCon.Execute("INSERT INTO schedule_history (id,id_task,start,state,real_start,real_end) values ('" & genGUID & "','" & TaskID & "'," & replace(cdbl(cdate(start)),",",".") & ",5," & replace(cdbl(cdate(now())),",",".") & ", null )")
    end if
    objCon.Close
    set objCon=Nothing
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end sub
  Private Function Prepare_task_toRun(TaskID,TaskName,start,stat)
    on error resume Next
      WScript.Echo(now() & " Uruchamiam Task:" & TaskID & "  name:" & TaskName & "   "  & now())
      call Report_start_work(TaskID,TaskName,start,stat)
      if not FileExists(scriptfullPath & "TaskID;" & TaskID & ".vbs") then
        Dim fso :Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFile scriptfullPath & "TMP_task.vbs",scriptfullPath & "TaskID;" & TaskID & ".vbs", True
        Set fso=nothing
      end if
      dim PID:PID=Launch_TASK(TaskID,TaskName)
      Prepare_task_toRun=PID
      If Err.Number <> 0 Then
          WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
            & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
      end if
  end function
  Private Function Check_MainTASK_IsWork (TaskID)
    on error resume next
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
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end Function
  Private Function get_enviroments
    on error resume Next
    'check what version of drivers are installed 64bit or 32 bit'
    Dim procc_ver:procc_ver=Check_ver
    dim chk:chk=False
    dim chk1:chk1=False
    dim is64: is64=FolderExists("c:\Windows\SysWOW64")
    Wscript.echo(now() &" Check running enviroment => System 64? => "  & cstr(is64)  & "  Process run from folder => " & procc_ver )
    DIm strComputer:strComputer = "."
    dim arrValueNames(),strValueName,arrValueTypes(),i,strValue
    dim objRegistry: Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
    dim strKeyPath:strKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"
    objRegistry.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes
    'Wscript.echo "ODBC drivers"
    For i = 0 to UBound(arrValueNames)
        strValueName = arrValueNames(i)
        objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
        if strValue = "Installed" Then
            'Wscript.echo arrValueNames(i)
            If InStr(1, UCase(arrValueNames(i)), "MDB") <> 0  And InStr(1, UCase(arrValueNames(i)), "MICR") <> 0 Then
              chk = True
              Exit For
            End If
        end if
    Next
    if chk<>True Then
        if is64 then
          'Wscript.echo "WOW64 drivers"
          strKeyPath = "SOFTWARE\Wow6432Node\ODBC\ODBCINST.INI\ODBC Drivers"
          objRegistry.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes
          For i = 0 to UBound(arrValueNames)
            strValueName = arrValueNames(i)
            objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
            if strValue = "Installed" Then
              'Wscript.echo arrValueNames(i)
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
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end Function
  Function Check_ver
    on error resume Next
    dim fso:Set fso = CreateObject("Scripting.FileSystemObject")
    dim wshShell:Set wshShell = CreateObject( "WScript.Shell" )
    If fso.FolderExists(wshShell.ExpandEnvironmentStrings("%windir%") & "\sysnative" ) Then
      Check_ver=32
    Else
      Check_ver=64
    End if
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end function
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
  Private Function Launch_TASK(TaskID,TaskName)
    on error resume Next
    Dim objWMIService,objProcess, intProcessID
    dim strComputer:strComputer = "."
    Set objWMIService = GetObject("winmgmts:" _
      & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Dim objStartup: Set objStartup = objWMIService.Get("Win32_ProcessStartup")
    Dim objConfig:Set objConfig = objStartup.SpawnInstance_
    objConfig.ShowWindow = HIDDEN_WINDOW
    'objConfig.CreateFlags=Create_New_Process_Group
    Set objProcess = objWMIService.Get("Win32_Process")
    wscript.echo now() & " " & (t_enviroments & " //NoLogo " & scriptfullPath & "TaskID;" & TaskID & ".vbs >> "  & t_db_path_forLogs  & "\" & DTE_form(date()) & "TaskID;" & TaskID & ".log")
    Dim intReturn:intReturn = objProcess.Create ("cmd.exe /c " & t_enviroments & " //U //NoLogo " & scriptfullPath & "TaskID;" & TaskID & ".vbs >> "  & t_db_path_forLogs & "\" & DTE_form(date()) & "TaskID" & TaskID & ".log", Null, objConfig, intProcessID)
      wscript.echo Now() & " ProcessID started PID => " & intProcessID
    Launch_TASK=intProcessID
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end Function
  Private Sub Get_settings
    On Error resume Next
    WScript.Echo(now() & " Pobieram ustawienia serwisu")
      dim XML_set : set XML_set= new Settings
        t_main_loop_interval=cdbl(XML_set.Main_loop_interval)
        t_refr_task_fromDtbase_intrv=cdbl(XML_set.Refr_task_fromDtbase_intrv)
        t_refr_settings_intrv=cdbl(XML_set.Refr_settings_intrv)
        t_db_connection_String=XML_set.Db_connection
        t_db_path_forLogs=XML_set.Db_path_forLogs
        t_db_fulpath=XML_set.Db_fulpath
        set XML_set= Nothing
    If Err.Number <> 0 Then
        WScript.Echo now () & " Blad " & CStr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  End sub
  Private Sub Check_Var_LogPath
    WScript.Echo(now() & " Sprawdzam ustawienia sciezki do logow")
    call Add_var_envir("Log",t_db_path_forLogs)
  end sub
  private sub settingdbxml
    WScript.Echo (now() & " Sygnal - pobrano ustawienia")
    call Add_var_envir("Get settingsXML",cstr(Now()))
  end sub
  Private sub Service_alive
    WScript.Echo (now() & " Sygnal - raport pracy")
    call Add_var_envir("Alive",cstr(Now()))
  end sub
  Private sub Add_var_envir(setting,setval)
    on error resume next
    Dim objCon:Set objCon = CreateObject("ADODB.Connection")
    objCon.Open (t_db_connection_String)
    Dim logs_inf:Set logs_inf = CreateObject("ADODB.Recordset")
    with logs_inf
      .ActiveConnection = objCon
      .CursorLocation = 3
      .LockType=4
      .open "SELECT * from Variables where setting='" & setting & "'"
    end with
    set logs_inf.ActiveConnection=Nothing
    objCon.Close
    if logs_inf.eof Then
      objCon.Open (t_db_connection_String)
      objCon.Execute("INSERT into Variables (setting,Val_variab) values ('" & setting & "','" & setval & "')")
      objCon.Close
    Else
      if logs_inf("Val_variab")<>setval Then
        logs_inf("Val_variab")=setval
        objCon.Open (t_db_connection_String)
        set logs_inf.ActiveConnection=objCon
        logs_inf.updatebatch
        set logs_inf.ActiveConnection=Nothing
        objCon.Close
      end if
    end if
    set logs_inf= Nothing
    set objCon=Nothing
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end sub
  Private Sub Report_serv
    WScript.Echo (now() & " Sygnal istnienia do bazy")
    call Add_var_envir("service",scriptfullPath)
  end sub
  Private Sub Class_Initialize
    Get_settings
    t_enviroments=get_enviroments
    Wscript.Echo now() & " Enviroments of main set to => " & t_enviroments
    if not FolderExists(t_db_path_forLogs) then Create_folder(t_db_path_forLogs)
    settingdbxml
    Archive_logs(t_db_path_forLogs)
    Check_Var_LogPath
    Report_serv
    checkOLDER_reportedDB_not_existWMI
    Get_list_of_schedule
    check_reportedDB_not_existWMI
    check_to_loongWrk_terminate
    serv_control=Check_serv_maintainExist
    Service_alive
    Main_Loop
  end Sub
  Private Sub Main_Loop
    on error resume Next
    Dim dbcle,set_cle
    dbcle=0
    set_cle=0
    dim old_task:old_task=0
    DO
      if serv_control then
        if not Check_serv_maintainExist then Exit Do
      end if
      WScript.sleep (t_main_loop_interval*1000)
      dbcle=dbcle+1
      set_cle=set_cle+1
      old_task=old_task+1
      if old_task>=600/t_main_loop_interval Then
        checkOLDER_reportedDB_not_existWMI
        old_task=0
      end if
      IF dbcLe >=t_refr_task_fromDtbase_intrv then
        Get_list_of_schedule
        check_reportedDB_not_existWMI
        check_to_loongWrk_terminate
        Service_alive
        dbcLe=0
      end if
      If set_cle>= t_refr_settings_intrv then
        Get_settings
        settingdbxml
        Archive_logs(t_db_path_forLogs)
        set_cle=0
      end if
      Count_activeTasks
      Delete_unuseful_tasks
      check_Work
      If Err.Number <> 0 Then
          WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
            & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
            err.clear
      end if
    LOOP
  End sub
  Private sub list_selected_Tree_processes(pid)
    on error resume Next
    WScript.Echo(now() & " Pobieram liste procesow powiazanych z PID = > " & pid)
    dim r_proc(),curr_fld_item:curr_fld_item=0
    dim R_FLDS:R_FLDS=0
    redim PRESERVE r_proc(0,R_FLDS)
    r_proc(0,R_FLDS)=pid
    with tmp_wmiRctset
      .filter="PID=" & pid
      .fields("Delete")="Yes"
      do
        .filter="parentPID=" & r_proc(0,curr_fld_item)
        WScript.Echo (now() & " Szukam powiazan do = > " & r_proc(0,curr_fld_item))
        if not .eof Then
            .movefirst
            do until .EOF
              WScript.Echo(now() & " Powiazania PID => " & r_proc(0,R_FLDS) & " = > " & .fields("PID"))
              .fields("Delete")="Yes"
              R_FLDS=R_FLDS+1
              redim PRESERVE r_proc(0,R_FLDS)
              r_proc(0,R_FLDS)=.fields("PID")
              .movenext
            loop
            curr_fld_item=curr_fld_item+1
        Else
          WScript.Echo(now() & " Brak powiazan do = > " & r_proc(0,curr_fld_item))
          WScript.Echo(now() & " check = > " & curr_fld_item & " => " & R_FLDS)
        end if
      loop until curr_fld_item => R_FLDS-1
      .filter=0
    end With
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end sub
  private function procc_to_del(Pid)
    on error resume next
    tmp_wmiRctset.filter="PID=" & pid & " and Delete='Yes'"
    procc_to_del= not tmp_wmiRctset.eof
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end function
  private sub Terminate_processes (Pid)
    on error resume next
    list_selected_Tree_processes(pid)
    Dim objWMIService ,colProcesses,objitem
    dim strComputer:strComputer = "."
    Set objWMIService = GetObject("winmgmts:" _
      & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
      Set colProcesses = objWMIService.ExecQuery _
        ("Select ProcessID FROM Win32_Process WHERE ProcessID =" & Pid)
        If colProcesses.Count > 0 Then
          For Each objitem In colProcesses
             objitem.Terminate()
          Next
        end if
      Set colProcesses=Nothing
      Set objWMIService=Nothing
      If Err.Number <> 0 Then
          WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
            & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
      end if
  end sub
  private sub get_childProcesses
    on error resume next
    set tmp_wmiRctset =  CreateObject("ADODB.Recordset")
    With tmp_wmiRctset
      Set .ActiveConnection = Nothing
      .CursorLocation = 3
      .LockType = 4
        .Fields.Append "command", 8, 200
        .Fields.Append "PID",4
        .Fields.Append "parentPID",4
        .Fields.Append "Delete" ,8 ,3
      .Open
    End With
    WScript.Echo(now() & " Utworzono zestaw rekordow")
    Dim objWMIService ,colProcesses,objitem
    dim strComputer:strComputer = "."
    Set objWMIService = GetObject("winmgmts:" _
      & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery _
      ("Select ProcessID,ParentProcessID,commandline FROM Win32_Process")
      dim cnt:cnt=0
      If colProcesses.Count > 0 Then
      dim tmp_str
      For Each objitem In colProcesses
        with tmp_wmiRctset
          .addnew
            tmp_str=objitem.CommandLine
            if tmp_str<>"" then
              .fields("command")=objitem.CommandLine
            end if
            if objitem.ProcessID=6372 then wscript.Echo ("Linia komend do " & objitem.CommandLine)
            .fields("PID")=objitem.ProcessID
            .fields("parentPID")=objitem.ParentProcessID
            .fields("Delete")="No"
          .update
        end with
      Next
      WScript.Echo(now() & " Dane procesow pobrane")
      end if
    Set colProcesses=Nothing
    Set objWMIService=Nothing
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end sub
  Private sub checkOLDER_reportedDB_not_existWMI
    on error resume next
    WScript.Echo(now() & " Pobieram stare nie zakonczone taski z bazy danych")
    Dim objCon:Set objCon = CreateObject("ADODB.Connection")
    objCon.Open (t_db_connection_String)
    dim old_sweat: Set old_sweat = CreateObject("ADODB.Recordset")
    with old_sweat
      .ActiveConnection = objCon
      .CursorLocation = 3
      .LockType=4
      .open "Select * from schedule_history where state=5 "
    end with
    set old_sweat.ActiveConnection=Nothing
    objCon.Close
    WScript.Echo(now() & " Baza pobrana")
    if not old_sweat.eof Then
      WScript.Echo(now() & " Znaleziono potencjalne taski do usuniecia")
      old_sweat.movefirst
      objCon.Open (t_db_connection_String)
      do until old_sweat.EOF
        if not Check_MainTASK_IsWork(old_sweat("id_task")) then
          old_sweat("real_end")=now()
          old_sweat("state")=6
          WScript.Echo  now() & " Raport do bazy dla popsutych taskow... : => " & old_sweat("id_task")
        end if
        old_sweat.movenext
      Loop
      set old_sweat.ActiveConnection=objCon
      old_sweat.updatebatch
      set old_sweat.ActiveConnection=Nothing
      objCon.Close
    end if
      Set old_sweat=Nothing
      objCon.Open (t_db_connection_String)
      WScript.Echo(now() & " Zbieram statystyki wykonania zadan")
      dim r_schedul: Set r_schedul = CreateObject("ADODB.Recordset")
      with r_schedul
        .ActiveConnection = objCon
        .CursorLocation = 3
        .LockType=4
        .open "Select * from schedule"
      end with
      set r_schedul.ActiveConnection=Nothing
      dim dtachang:dtachang=False
      dim r_com_schedul: Set r_com_schedul = CreateObject("ADODB.Recordset")
      with r_com_schedul
        .ActiveConnection = objCon
        .CursorLocation = 3
        .LockType=4
        .open "SELECT id_task,round(avg(real_end-real_start)*24*60,4) as exec_time from schedule_history where state=4 and real_start>date()-31 group by id_task"
      end with
      set r_com_schedul.ActiveConnection=Nothing
      objCon.Close
      If not r_schedul.eof Then
        r_schedul.movefirst
        do until r_schedul.EOF
          WScript.Echo(now() & " Statystyki wykonania zadan => " & r_schedul("id_task"))
          r_com_schedul.filter="id_task='" & r_schedul("id_task") & "'"
          if not r_com_schedul.eof then
            if r_schedul("time_exec")<>r_com_schedul("exec_time") or isnull(r_schedul("time_exec")) then
              WScript.Echo(now() & " Update statystk dla => " & r_schedul("id_task"))
              r_schedul("time_exec")=r_com_schedul("exec_time")
              dtachang=True
            Else
              WScript.Echo(now() & " Bez zmian dla => " & r_schedul("id_task"))
            end if
          end if
          r_schedul.movenext
        loop
        if dtachang Then
          objCon.Open (t_db_connection_String)
          set r_schedul.ActiveConnection=objCon
          r_schedul.updatebatch
          set r_schedul.ActiveConnection=Nothing
          objCon.Close
        end if
      end if
    set r_schedul=Nothing
    set r_com_schedul=Nothing
    Set objCon=Nothing
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end sub
  Private sub check_reportedDB_not_existWMI
    On Error resume Next
    'Check if some tesks are ended without no report in DB remove bug
    t_work_to_do.filter="status=5"
    if not t_work_to_do.eof Then
      t_work_to_do.movefirst
      do until t_work_to_do.EOF
        if not Check_MainTASK_IsWork(t_work_to_do("id")) then
          WScript.Echo  now() & " Raport do bazy dla popsutych taskow... : => " & t_work_to_do("id")
          Dim objCon:Set objCon = CreateObject("ADODB.Connection")
            objCon.Open (t_db_connection_String)
              objCon.Execute("update schedule_history set state=6,real_end="  & replace(cdbl(cdate(now())),",",".") &  " where id_task='" & t_work_to_do("id") & "' and state=5")
            objCon.Close
          set objCon=Nothing
          if not err.Number <> 0 then t_work_to_do("status")=6
        end if
        t_work_to_do.movenext
      Loop
    end if
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & CStr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    End If
  end sub
  Private sub check_to_loongWrk_terminate
    on error resume Next
    dim rst_chang:rst_chang=False
    WScript.Echo(now() & " Pobieram taski usuniecia z procesow")
    dim r_proc(),cnt,i
    Dim objCon:Set objCon = CreateObject("ADODB.Connection")
    objCon.Open (t_db_connection_String)
    dim old_sweat: Set old_sweat = CreateObject("ADODB.Recordset")
    with old_sweat
      .ActiveConnection = objCon
      .CursorLocation = 3
      .LockType=4
      .open "Select * from schedule_history where state=8"
      end with
      set old_sweat.ActiveConnection=Nothing
    objCon.Close
    WScript.Echo(now() & " Pobrano liste taskow z bazy")
    if not old_sweat.eof Then
      WScript.Echo(now() & " Lista zawiera taski do usuniecia")
      old_sweat.movefirst
      WScript.Echo(now() & " Pobieram mape aktywnych procesow")
      get_childProcesses
      dim chk:chk=False
      do until old_sweat.EOF
        with tmp_wmiRctset
          .filter="command<>''"
          .movefirst
          WScript.Echo(now() & " Szukam pid tasku => " & old_sweat("id_task"))
          do until .eof
            if instr(1,.Fields("command"),"TaskID;" & old_sweat("id_task"))>0 then
              chk=True
              cnt=cnt+1
              WScript.Echo(now() & " Pid tasku => " & old_sweat("id_task") & " => " & .fields("PID") & " item_no => " & cnt)
              REDIM PRESERVE  r_proc(cnt)
              r_proc(cnt)=.fields("PID")
              old_sweat("state")=9
              old_sweat("real_end")=now()
              rst_chang=true
              call Terminate_office(.fields("PID"),old_sweat("id_task"))
            end if
            .movenext
          loop
        end With
        if chk=false then
          old_sweat("state")=6
          old_sweat("real_end")=now()
          rst_chang=true
        end if
        old_sweat.movenext
      loop
      for i=1 to cnt
        WScript.Echo Now() & " Element nr " & i & " => " & r_proc(i)
      next
        for i=1 to cnt
          WScript.Echo Now() & " Element nr " & i & " Zamykam process PID : " & r_proc(i)
          Terminate_processes (r_proc(i))
        next
        end if
        if rst_chang=true then
        objCon.Open (t_db_connection_String)
          Set old_sweat.ActiveConnection = objCon
          old_sweat.updatebatch
          Set old_sweat.ActiveConnection = Nothing
        objCon.close
        end if
        Set old_sweat=Nothing
        Set objCon=Nothing
        If Err.Number <> 0 Then
            WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
              & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
        end if
  end sub
  Private Sub check_terminate
    on error resume Next
    dim r_proc(),cnt,i
    t_work_to_do.filter="status=8"
    if not t_work_to_do.eof Then
        check_to_loongWrk_terminate
    end if
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end sub
  Private Sub check_Work
    on error resume next
    dim chk:chk=false
    t_work_to_do.filter="status<3 or status=7"
    t_work_to_do.Sort="start"
    if not t_work_to_do.eof then
      t_work_to_do.movefirst
      do until chk or t_work_to_do.eof
        if t_work_to_do("start")<=now() then
          if not Check_MainTASK_IsWork(t_work_to_do("id")) then
              dim pid:pid= Prepare_task_toRun(t_work_to_do("id"),t_work_to_do("name"),t_work_to_do("start"),t_work_to_do("status"))
              t_work_to_do("status")=5
              chk=true
          end if
        else
          chk=true
        end if
        t_work_to_do.movenext
      loop
    end if
    t_work_to_do.filter=0
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end Sub
  private Sub Archive_logs(fFolder)
    on error resume Next
    dim objFile
    dim objFSO :Set objFSO = CreateObject("Scripting.FileSystemObject")
    dim objFolder: Set objFolder = objFSO.GetFolder(fFolder)
    dim colFiles:Set colFiles = objFolder.Files
    WScript.Echo now () & " Sprawdzam Logi do archiwizacji w folderze " & fFolder
    For Each objFile in colFiles
      WScript.Echo now () & " Plik =>" & UCase(objFSO.GetExtensionName(objFile.name))
        If UCase(objFSO.GetExtensionName(objFile.name)) = "LOG" Then
            if Left(objFile.Name,8)<>DTE_form(date()) and Left(objFile.Name,8)<>DTE_form(now()-(3/24)) then
              WScript.Echo (now () & " Przenosze stare logi : " & fFolder & "\" & objFile.name & " => " & fFolder & "\Old_logs.zip")
               call WindowsZip(fFolder & "\" & objFile.name ,fFolder & "\Old_logs.zip")
              if Check_FileinZip_Exist(fFolder & "\" & objFile.name ,fFolder & "\Old_logs.zip") then objfile.Delete
            end if
        End If
    Next
    Set objFSO = Nothing
    Set objFolder = Nothing
    Set colFiles = Nothing
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end Sub
  Private Function Check_FileinZip_Exist (sFile, sZipFile)
    on error resume Next
    Dim oZipApp:Set oZipApp = CreateObject("Shell.Application")
    Dim aFileName:aFileName = Split(sFile, "\")
    Dim sFileName:sFileName = (aFileName(Ubound(aFileName)))
    Dim sDupe : sDupe = False
    Dim sFileNameInZip
    For Each sFileNameInZip In oZipApp.NameSpace(sZipFile).items
      If LCase(sFileName) = LCase(sFileNameInZip) Then
        sDupe = True
        Exit For
      End If
    Next
    Check_FileinZip_Exist=sDupe
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  End Function
  Private Sub WindowsZip(sFile, sZipFile)
    on error resume Next
    'This script is provided under the Creative Commons license located
    'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
    'be used for commercial purposes with out the expressed written consent
    'of NateRice.com
    Dim oZipShell:Set oZipShell = CreateObject("WScript.Shell")
    Dim oZipFSO: Set oZipFSO = CreateObject("Scripting.FileSystemObject")
    If Not oZipFSO.FileExists(sZipFile) Then
      NewZip(sZipFile)
    End If
    Dim oZipApp:Set oZipApp = CreateObject("Shell.Application")
    Dim sZipFileCount: sZipFileCount = oZipApp.NameSpace(sZipFile).items.Count
    Dim aFileName:aFileName = Split(sFile, "\")
    Dim sFileName:sFileName = (aFileName(Ubound(aFileName)))
    'listfiles
    Dim sDupe : sDupe = False
    Dim sFileNameInZip
    For Each sFileNameInZip In oZipApp.NameSpace(sZipFile).items
      If LCase(sFileName) = LCase(sFileNameInZip) Then
        sDupe = True
        Exit For
      End If
    Next
      If Not sDupe Then
      oZipApp.NameSpace(sZipFile).Copyhere sFile
      'Keep script waiting until Compressing is done
      On Error resume Next
        sLoop = 0
        Do Until sZipFileCount < oZipApp.NameSpace(sZipFile).Items.Count
          Wscript.Sleep(100)
          sLoop = sLoop + 1
          Loop
          On Error GoTo 0
      End If
      Set oZipApp=Nothing
      Set oZipFSO=Nothing
      Set oZipShell=Nothing
      If Err.Number <> 0 Then
          WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
            & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
      end if
  End Sub
  Private Sub NewZip(sNewZip)
    on error resume next
    'This script is provided under the Creative Commons license located
    'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
    'be used for commercial purposes with out the expressed written consent
    'of NateRice.com
    Dim oNewZipFSO: Set oNewZipFSO = CreateObject("Scripting.FileSystemObject")
    Dim oNewZipFile: Set oNewZipFile = oNewZipFSO.CreateTextFile(sNewZip)
      oNewZipFile.Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)
      oNewZipFile.Close
    Set oNewZipFile=Nothing
    Set oNewZipFSO = Nothing
    Wscript.Sleep(500)
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  End Sub
  Private sub Terminate_office(Pid_offtask,Task_nam_ID)
    on error resume next
    wscript.Echo now() & " Sprawdzam czy process office istnieje do PID => " & Pid_offtask
    Dim objCon:Set objCon = CreateObject("ADODB.Connection")
    objCon.Open ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + scriptfullPath & "Log_Helper.mde")
    dim rs_resp
    Set rs_resp= CreateObject("ADODB.Recordset")
    with rs_resp
      .ActiveConnection = objCon
      .CursorLocation = 3
      .LockType=4
      .open "Select * from task_procc where task_pid=" & Pid_offtask
    end with
    Set rs_resp.ActiveConnection = Nothing
    objCon.Close
    Set objCon = nothing
     if not rs_resp.eof then
      if not isnull(rs_resp("pid").value) then
        if rs_resp("pid").value<>0 then
          wscript.Echo now() & " Sprawdzam czy process office istnieje PID => " & rs_resp("pid").value
          dim tmx:tmx = check_is_closed(rs_resp("pid").value)
          dim tmpchk,bl:bl=False
          if FileExists(rs_resp("curr_log")) then
              Call Add_last_log(rs_resp("curr_log"),Task_nam_ID)
              Wscript.Echo now() & " Usuwam logi z instancji office => " & rs_resp("curr_log")
              dim filesys:Set filesys = CreateObject("Scripting.FileSystemObject")
              filesys.DeleteFile rs_resp("curr_log")
              Set filesys= Nothing
          end if
        end if
      end if
      Set objCon = CreateObject("ADODB.Connection")
      Wscript.Echo now() & " Usuwam zapis instancji w tymczasowej bazie => " & rs_resp("curr_log")
      objCon.Open ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + scriptfullPath & "Log_Helper.mde")
      objCon.Execute("Delete from task_procc where task_pid=" & Pid_offtask)
      objCon.Close
      Wscript.Echo now() & " Instancja office zamknieta => " & rs_resp("curr_log")
      Set objCon= Nothing
      set rs_resp = Nothing
     End if
     If Err.Number <> 0 Then
         WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
           & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
     end if
  End Sub
  Private sub Add_last_log(curr_log,Task_nam_ID)

  end sub
  private Function check_is_closed(Pid)
    on Error resume Next
    dim proccexist:proccexist=False
    dim strComputer:strComputer = "."
    dim objWMIService:Set objWMIService = GetObject("winmgmts:" _
      & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
      dim objitem:DIm colProcesses:Set colProcesses = objWMIService.ExecQuery _
        ("SELECT * FROM Win32_Process WHERE ProcessID =" & Pid)
      If colProcesses.Count > 0 Then
        For Each objitem In colProcesses
            WScript.Echo now() & " Wymuszam zamkniecie instancji Office..."
            objitem.terminate
        Next
      end if
    check_is_closed=proccexist
    If Err.Number <> 0 Then
        WScript.Echo now() & " Blad " & Cstr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    end if
  end function
End Class
Class Calendar
  private rs_exeptions,rs_type_day,rs_calendar_wrk
  private t_Dbpath,t_cal_ids
  private Valid_from,Valid_to,state,calendar_types_day
  private function Is_exeption_day(Valid_from)
    rs_exeptions.filter="Exeption_day='" & cdate(Valid_from )& "'"
    Is_exeption_day=not rs_exeptions.eof
  end Function
  private sub filter_rs_type(num_day)
    rs_type_day.filter="day_id=" & num_day
  end sub
  private function IS_type_work(num_day)
    filter_rs_type(num_day)
    if not rs_type_day.eof then
      if rs_type_day("time_day")>0 then
          IS_type_work=true
        Else
          IS_type_work=false
        end if
    Else
      IS_type_work=false
    end if
  end function
  public sub Generate
    On Error resume Next
    Wscript.echo now() & (" Sprawdzam kalendarz : " & t_cal_ids)
    dim que:set que = (new Calendar_Queue) (t_Dbpath,t_cal_ids)
    que.get_first_calendar_day(Valid_from)
    dim first,last,i,j
    dim cal_counter,work_day,day_id,flds(),Isadd,Isupdt
    dim R_FLDS:R_FLDS=0
    redim flds(4,R_FLDS)
    cal_counter=1
    Isupdt=False
    Isadd=False
    if not rs_calendar_wrk.eof then rs_calendar_wrk.movefirst
    for i=Valid_from to Valid_to
      if not Is_exeption_day(i) then
        day_id=que.Day_type
      Else
        day_id=rs_exeptions("day_type")
      End if
      if IS_type_work(day_id) Then
          work_day=i
          if not rs_calendar_wrk.eof Then
            if rs_calendar_wrk("cal_counter")<>cal_counter or rs_calendar_wrk("work_day")<>work_day or rs_calendar_wrk("day_id")<>day_id Then
              Isupdt=True
              rs_calendar_wrk("cal_counter")=cal_counter
              rs_calendar_wrk("work_day")=work_day
              rs_calendar_wrk("day_id")=day_id
            end if
            rs_calendar_wrk.movenext
          Else
            Isadd=true
            flds(0,R_FLDS)=genGUID
            flds(1,R_FLDS)=t_cal_ids
            flds(2,R_FLDS)=cal_counter
            flds(3,R_FLDS)=work_day
            flds(4,R_FLDS)=day_id
            R_FLDS=R_FLDS+1
            redim Preserve flds(4,R_FLDS)
          End if
          que.Next_day
          cal_counter=cal_counter+1
      end if
    next
    If Err.Number <> 0 Then
        WScript.Echo now () & " Blad " & CStr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    End If
   if Isadd then
      for i=0 to R_FLDS-1
        rs_calendar_wrk.addnew
        for j=0 to 4
		      rs_calendar_wrk(j)=flds(j,i)
	      next
        rs_calendar_wrk.update
      next
    end if
    if Isadd or Isupdt Then
      dim objConnection:Set objConnection = CreateObject("ADODB.Connection")
      objConnection.Open (t_Dbpath)
      set rs_calendar_wrk.ActiveConnection=objConnection
      rs_calendar_wrk.updatebatch
      set rs_calendar_wrk.ActiveConnection=Nothing
      objConnection.close
      Set objConnection=Nothing
    end if
    If Err.Number <> 0 Then
        WScript.Echo now () & " Blad " & CStr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    End If
    On Error Goto 0
  end Sub
  Public default function init(Dbpath,cal_ids)
    On Error resume Next
    WScript.Echo(now() & " Tworze obiekt kalendarza : " & cal_ids)
    t_Dbpath=Dbpath
    t_cal_ids=cal_ids
    Dim objCon:Set objCon = CreateObject("ADODB.Connection")
    objCon.Open (Dbpath)
    Dim cal_hdrRs:Set cal_hdrRs = CreateObject("ADODB.Recordset")
    with cal_hdrRs
      .ActiveConnection = objCon
      .CursorLocation = 3
      .LockType=4
      .open "Select * FROM calendar_hdr where id='" & cal_ids & "'"
    end with
    Set cal_hdrRs.ActiveConnection = Nothing
    Valid_from=cal_hdrRs("Valid_from")
    Valid_to=cal_hdrRs("Valid_to")
    state=cal_hdrRs("state")
    calendar_types_day=cal_hdrRs("calendar_types_day")
    Set cal_hdrRs=Nothing
    Set rs_exeptions = CreateObject("ADODB.Recordset")
    with rs_exeptions
      .ActiveConnection = objCon
      .CursorLocation = 3
      .LockType=4
      .open "Select * FROM calendar_day_exeptions where calendar_id='" & cal_ids & "'"
    end with
    Set rs_exeptions.ActiveConnection = Nothing
    Set rs_type_day = CreateObject("ADODB.Recordset")
    with rs_type_day
      .ActiveConnection = objCon
      .CursorLocation = 3
      .LockType=4
      .open "Select * FROM calendar_types_day where set_id='" & calendar_types_day & "'"
    end with
    Set rs_type_day.ActiveConnection = Nothing
    Set rs_calendar_wrk = CreateObject("ADODB.Recordset")
    with rs_calendar_wrk
      .ActiveConnection = objCon
      .CursorLocation = 3
      .LockType=4
      .open "Select * FROM calendar_wrk where calendar_id='" & cal_ids  & "'"
    end with
    Set rs_calendar_wrk.ActiveConnection = Nothing
    objCon.close
    If Err.Number <> 0 Then
        WScript.Echo now () & " Blad " & CStr(Err.Number) & " wygenerowany przez " _
          & Err.Source & Chr(13) & Err.Description  & Chr(13) & Err.Helpfile & chr(13) & Err.HelpContext
    End If
    Set objCon=Nothing
    On Error Goto 0
    set Init=Me
  end Function
  Private sub Class_Terminate
    set rs_exeptions=nothing
    set rs_type_day=nothing
    set rs_calendar_wrk=nothing
  End Sub
End Class
Sub Create_folder(path)
  dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder path
  Set fso = Nothing
end Sub
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
Function FileExists(FilePath)
   dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(FilePath) Then
      FileExists=CBool(1)
    Else
      FileExists=CBool(0)
    End If
    WScript.Echo now () & " Check FileExists " & FilePath & " result =>" & FileExists
    Set fso = Nothing
 End Function
function unique_replace(text)
   do while not instr(text,"GenGUID()")=0
     text=replace(text,"GenGUID()","'" & genGUID & "'",1,1)
   Loop
   unique_replace=text
 End function
function genGUID
   genGUID = Left(CreateObject("Scriptlet.TypeLib").Guid,38)
 end function
