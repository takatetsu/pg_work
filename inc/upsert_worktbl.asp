<%
' ---------------------------------------------------------------------
' 個人勤務表入力による更新
' ---------------------------------------------------------------------
' ---------------------------------------------------------
' worktbl に対しての更新処理
' ---------------------------------------------------------
' #######################################################
' worktbl 変数設定 ここから
' #######################################################
wk_id               = Request.Form("worktbl_id")(i)
wk_updatetime       = Request.Form("worktbl_updatetime" )(i)
wk_personalcode     = Session("MM_Username")
wk_workingdate      = Request.Form("ymd")(i)
wk_morningwork      = Request.Form("morningwork")(i)
wk_afternoonwork    = Request.Form("afternoonwork")(i)
wk_morningholiday   = Request.Form("morningholiday")(i)
wk_afternoonholiday = Request.Form("afternoonholiday")(i)
wk_summons          = Request.Form("summons")(i)
wk_overtime_begin   = Left(editTime(Request.Form("overtime_begin" )(i)), 2) & _
                      Right(editTime(Request.Form("overtime_begin")(i)), 2)
wk_overtime_end     = Left(editTime(Request.Form("overtime_end"   )(i)), 2) & _
                      Right(editTime(Request.Form("overtime_end"  )(i)), 2)
wk_rest_begin       = Left(editTime(Request.Form("rest_begin"     )(i)), 2) & _
                      Right(editTime(Request.Form("rest_begin"    )(i)), 2)
wk_rest_end         = Left(editTime(Request.Form("rest_end"       )(i)), 2) & _
                      Right(editTime(Request.Form("rest_end"      )(i)), 2)
' ---------------------------------------------------------
' 時間該等算出
' ---------------------------------------------------------
v_overtime_begin    = editTime(Request.Form("overtime_begin")(i))
v_overtime_end      = editTime(Request.Form("overtime_end"  )(i))
v_rest_begin        = editTime(Request.Form("rest_begin"    )(i))
v_rest_end          = editTime(Request.Form("rest_end"      )(i))
v_morningholiday    = Request.Form("morningholiday"  )(i)
v_afternoonholiday  = Request.Form("afternoonholiday")(i)
v_morningwork       = Request.Form("morningwork"     )(i)
v_afternoonwork     = Request.Form("afternoonwork"   )(i)
compOverTime()
wk_overtime         = Left(v_overtime, 2) & Right(v_overtime, 2)
wk_overtimelate     = Left(v_overtimelate, 2) & Right(v_overtimelate, 2)
wk_holidayshift     = Left(v_holidayshift, 2) & Right(v_holidayshift, 2)
wk_holidayshiftovertime = Left(v_holidayshiftovertime, 2) & _
                          Right(v_holidayshiftovertime, 2)
wk_holidayshiftlate = Left(v_holidayshiftlate, 2) & Right(v_holidayshiftlate, 2)
wk_holidayshiftovertimelate = Left(v_holidayshiftovertimelate, 2) & _
                              Right(v_holidayshiftovertimelate, 2)

' 時間単位代休
If workshift = "9" Then
    ' フレックス勤務者は時間単位代休の入力は無い
    temp       = "0000"
    temp_begin = ""
    temp_end   = ""
Else
    wk_requestmin = minDif(editTime(Request.Form("requesttime_begin")(i)),    _
                           editTime(Request.Form("requesttime_end"  )(i))) -  _
                    checkLunchTime(editTime(Request.Form("requesttime_begin")(i)), _
                                   editTime(Request.Form("requesttime_end"  )(i)))
    temp       = min2Time(wk_requestmin)
    temp_begin = Left (editTime(Request.Form("requesttime_begin")(i)), 2) & _
                 Right(editTime(Request.Form("requesttime_begin")(i)), 2)
    temp_end   = Left (editTime(Request.Form("requesttime_end"  )(i)), 2) & _
                 Right(editTime(Request.Form("requesttime_end"  )(i)), 2) 
End If
wk_requesttime       = Left(temp, 2) & Right(temp, 2)
wk_requesttime_begin = temp_begin
wk_requesttime_end   = temp_end

' 深夜割増
temp       = min2Time(minDif(editTime(Request.Form("latetime_begin")(i)), _
                             editTime(Request.Form("latetime_end"  )(i))))
temp_begin = Left (editTime(Request.Form("latetime_begin")(i)), 2) & _
             Right(editTime(Request.Form("latetime_begin")(i)), 2)
temp_end   = Left (editTime(Request.Form("latetime_end"  )(i)), 2) & _
             Right(editTime(Request.Form("latetime_end"  )(i)), 2)
' 交替勤務乙番で深夜割増未入力のときの自動設定
If is_operator Then
    ' オペレータ
    If (Request.Form("operator")(i) = "2"  Or _
        Request.Form("operator")(i) = "4"  Or _
        Request.Form("operator")(i) = "6") Then
        ' 交替勤務が乙番、生産会議乙、見習（乙）
        If (Request.Form("morningwork"  )(i) > "0"  And _
            Request.Form("afternoonwork")(i) > "0") Then
            ' 午前午後とも勤務
            If Len(Trim(Request.Form("latetime_begin")(i))) = 0 And _
               Request.Form("worktbl_id")(i) = "" Then
                ' 深夜割増に入力がなく、新規入力のとき 22:00～05:00で7時間を自動設定する。
                temp       = "0700"
                temp_begin = "2200"
                temp_end   = "0500" 
            End If
        End If
    End If
End If
wk_latetime         = Left(temp, 2) & Right(temp, 2)
wk_latetime_begin   = temp_begin
wk_latetime_end     = temp_end

' 週超過時間
If Not workshift = "9" Then
    wk_weekovertime = Left (editTime(Request.Form("weekovertime")(i)), 2) & _
                      Right(editTime(Request.Form("weekovertime")(i)), 2)
Else
    wk_weekovertime = ""
End If

wk_is_approval      = "0" ' 上長承認
wk_nightduty        = Request.Form("nightduty")(i)
wk_dayduty          = Request.Form("dayduty")(i)
If is_operator Then
    wk_operator  = Request.Form("operator")(i)
    ' 前日交替勤務を設定する
    wk_operator2 = setPreOp(Session("MM_Username"), Request.Form("ymd")(i))
Else
    wk_operator  = "0"
    wk_operator2 = "0" ' 前日交替勤務を設定する
End If

' 時間有給
wk_vacationmin = minDif(editTime(Request.Form("vacationtime_begin")(i)),   _
                        editTime(Request.Form("vacationtime_end"  )(i))) - _
         checkLunchTime(editTime(Request.Form("vacationtime_begin")(i)),   _
                        editTime(Request.Form("vacationtime_end"  )(i)))
temp = min2Time(wk_vacationmin)
wk_vacationtime       = Left (temp, 2) & Right(temp, 2)
wk_vacationtime_begin = Left(editTime(Request.Form("vacationtime_begin")(i)), 2) & _
                        Right(editTime(Request.Form("vacationtime_begin")(i)), 2)
wk_vacationtime_end   = Left(editTime(Request.Form("vacationtime_end" )(i)), 2) & _
                        Right(editTime(Request.Form("vacationtime_end")(i)), 2)

wk_memo             = Request.Form("memo" )(i)
wk_memo2            = Request.Form("memo2")(i)

' 労働時間適正化エラーフラグ
If gradecode >= "033" Or gradecode = "000" Then
    ' 課長以上または等級コード00,01はエラーチェック不要
    temp = "0"
Else
    x = Right(Request.Form("YMD")(i), 2) * 1
    temp = workTimeCheck(wk_operator, _
        Request.Form("morningwork" )(i), Request.Form("afternoonwork" )(i), _
        Request.Form("cometime"    )(x), Request.Form("leavetime"     )(x), _
        Request.Form("pc_ontime"   )(x), Request.Form("pc_offtime"    )(x), _
        Request.Form("dayduty"     )(i), Request.Form("nightduty"     )(i), _
        Request.Form("nightduty2"  )(x), Request.Form("overtime_begin")(i), _
        Request.Form("overtime_end")(i), Request.Form("memo2"         )(i), _
        Session("MM_opentime"), Session("MM_closetime"), _
        Session("MM_is_unionexecutive"), wk_operator2)
End If
wk_is_error = temp

If workshift <> "9" Then
    wk_work_begin     = ""
    wk_work_end       = ""
    wk_break_begin1   = ""
    wk_break_end1     = ""
    wk_break_begin2   = ""
    wk_break_end2     = ""
    wk_workmin        = 0
Else
    ' フレックス勤務
    temp_work_begin   = Trim(Request.Form("work_begin"  )(i))
    temp_work_end     = Trim(Request.Form("work_end"    )(i))
    temp_break_begin1 = Trim(Request.Form("break_begin1")(i))
    temp_break_end1   = Trim(Request.Form("break_end1"  )(i))
    temp_break_begin2 = Trim(Request.Form("break_begin2")(i))
    temp_break_end2   = Trim(Request.Form("break_end2"  )(i))
    ' 勤務時間の入力が無い場合、標準時刻を設定
    If workshift = "9" And _
       Len(temp_work_begin  ) = 0 And Len(temp_work_end  ) = 0 And _
       Len(temp_break_begin1) = 0 And Len(temp_break_end1) = 0 And _
       Len(temp_break_begin2) = 0 And Len(temp_break_end2) = 0 Then
        If Request.Form("morningwork")(i) > "0" And _
           Request.Form("morningwork")(i) <> "3" Then
            If Request.Form("afternoonwork")(i) > "0" And _
               Request.Form("afternoonwork")(i) <> "3" Then
                ' 午前午後出勤
                If holidaytype = "2" Then
                  ' ピポット勤務者
                  temp_work_begin   = "09:30"
                  temp_work_end     = "18:10"
                  temp_break_begin1 = "13:00"
                  temp_break_end1   = "14:00"
                Else
                  temp_work_begin   = "08:30"
                  temp_work_end     = "17:10"
                  temp_break_begin1 = "12:00"
                  temp_break_end1   = "13:00"
                End If
            Else
                ' 午前のみ出勤
                If holidaytype = "2" Then
                  ' ピポット勤務者
                  temp_work_begin   = "09:30"
                  temp_work_end     = "13:00"
                  temp_break_begin1 = ""
                  temp_break_end1   = ""
                Else
                  temp_work_begin   = "08:30"
                  temp_work_end     = "12:00"
                  temp_break_begin1 = ""
                  temp_break_end1   = ""
                End If
            End If
        Else
            If Request.Form("afternoonwork")(i) > "0" And _
               Request.Form("afternoonwork")(i) <> "3" Then
                ' 午後のみ出勤
                If holidaytype = "2" Then
                  ' ピポット勤務者
                  temp_work_begin   = "14:00"
                  temp_work_end     = "18:10"
                  temp_break_begin1 = ""
                  temp_break_end1   = ""
                Else
                  temp_work_begin   = "13:00"
                  temp_work_end     = "17:10"
                  temp_break_begin1 = ""
                  temp_break_end1   = ""
                End If
            End If
        End If
    End If
End If
wk_work_begin   = Left(editTime(temp_work_begin  ), 2) & Right(editTime(temp_work_begin  ), 2)
wk_work_end     = Left(editTime(temp_work_end    ), 2) & Right(editTime(temp_work_end    ), 2)
wk_break_begin1 = Left(editTime(temp_break_begin1), 2) & Right(editTime(temp_break_begin1), 2)
wk_break_end1   = Left(editTime(temp_break_end1  ), 2) & Right(editTime(temp_break_end1  ), 2)
wk_break_begin2 = Left(editTime(temp_break_begin2), 2) & Right(editTime(temp_break_begin2), 2)
wk_break_end2   = Left(editTime(temp_break_end2  ), 2) & Right(editTime(temp_break_end2  ), 2)

wk_workmin = 0
If workshift = "9" Then ' フレックス勤務者
    ' #########################################
    ' 労働時間
    ' #########################################
    If Request.Form("morningholiday")(i) = "3" Or _
       Request.Form("morningholiday")(i) = "5" Or _
       Request.Form("morningholiday")(i) = "6" Then
        ' 午前有給休暇、特別休暇、保存休暇
        wk_workmin = wk_workmin + 210
    ElseIf Request.Form("morningholiday")(i) = "9" Then
        ' 午前コアタイム有休
        wk_workmin = wk_workmin + 120
    End If
    If Request.Form("afternoonholiday")(i) = "3" Or _
       Request.Form("afternoonholiday")(i) = "5" Or _
       Request.Form("afternoonholiday")(i) = "6" Then
        ' 午後有給休暇、特別休暇、保存休暇
        wk_workmin = wk_workmin + 250
    ElseIf Request.Form("afternoonholiday")(i) = "9" Then
        ' 午後コアタイム有休
        wk_workmin = wk_workmin + 120
    End If
    ' 勤務時間計算
    wk_workmin = wk_workmin + minDif(editTime(temp_work_begin), _
                                     editTime(temp_work_end))
    ' 休憩時間1計算
    breaktime1 = minDif(editTime(temp_break_begin1), _
                        editTime(temp_break_end1))
    ' 休憩時間2計算
    breaktime2 = minDif(editTime(temp_break_begin2), _
                        editTime(temp_break_end2))
    ' 実勤務時間計算
    wk_workmin = wk_workmin - breaktime1 - breaktime2 '+ time2min(v_overtime)

    ' フレックス勤務のとき時間外入力も労働時間に加算
    tempWork = minDif(editTime(Request.Form("overtime_begin")(i)), _
                      editTime(Request.Form("overtime_end")(i)))
    tempRest = minDif(editTime(Request.Form("rest_begin")(i)), _
                      editTime(Request.Form("rest_end")(i)))
    wk_workmin = wk_workmin + (tempWork - tempRest)
Else ' フレックス勤務者以外
    tempWork = 0
    ' 午前勤務時間加算
    If wk_morningwork = "1" Or wk_morningwork = "4" Or _
       wk_morningwork = "5" Or wk_morningwork = "9" Then
        tempWork = tempWork + base_am_workmin
    End If
    ' 午後勤務時間加算
    If wk_afternoonwork = "1" Or wk_afternoonwork = "4" Or _
       wk_afternoonwork = "5" Or wk_afternoonwork = "9" Then
        tempWork = tempWork + base_pm_workmin
    End If
    ' 時間代休、時間有給減算
    wk_workmin = tempWork - wk_requestmin - wk_vacationmin
End If

' #######################################################
' worktbl 変数設定 ここまで
' #######################################################

Set MM_editCmd = Server.CreateObject ("ADODB.Command")
MM_editCmd.ActiveConnection = MM_workdbms_STRING
If (Request.Form("worktbl_id")(i) = "") Then
    ' 出勤区分(午前) または、休日区分(午前) どちらかに入力があれば、 INSERT 処理を行う。
    If (Request.Form("morningwork"   )(i) <> "0"  Or _
        Request.Form("morningholiday")(i) <> "0") Then
        Set MM_editCmd = Server.CreateObject ("ADODB.Command")
        MM_editCmd.ActiveConnection = MM_workdbms_STRING
        ' -----------------------------------------------------
        ' INSERT worktbl 処理
        ' -----------------------------------------------------
        MM_editCmd.CommandText = "INSERT INTO dbo.worktbl VALUES(DEFAULT, " & _
            "?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, " & _
            "?, ?, ?, ?, ?, ?, ?, ?, '0', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        MM_editCmd.Prepared = true
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,5, wk_personalcode)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,8, wk_workingdate)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,1, wk_morningwork)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,1, wk_afternoonwork)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,1, wk_morningholiday)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,1, wk_afternoonholiday)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,1, wk_summons)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_overtime_begin)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_overtime_end)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_rest_begin)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_rest_end)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_overtime)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_overtimelate)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_holidayshift)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_holidayshiftovertime)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_holidayshiftlate)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_holidayshiftovertimelate)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_requesttime)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_requesttime_begin)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_requesttime_end)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_latetime)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_latetime_begin)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_latetime_end)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,1, wk_nightduty)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,1, wk_dayduty)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,1, wk_operator)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_vacationtime)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_vacationtime_begin)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_vacationtime_end)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,100, wk_memo)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,2, wk_memo2)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,129,,1, wk_is_error)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_work_begin)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_work_end)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_break_begin1)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_break_end1)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_break_begin2)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_break_end2)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,,4, wk_workmin)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_weekovertime)
        MM_editCmd.Execute
        MM_editCmd.ActiveConnection.Close
    End If
Else
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_workdbms_STRING
    ' ---------------------------------------------------------
    ' UPDATE worktbl 処理
    ' ---------------------------------------------------------
    MM_editCmd.CommandText = "UPDATE dbo.worktbl SET " & _
             "morningwork              = ?," & _
             "afternoonwork            = ?," & _
             "morningholiday           = ?," & _
             "afternoonholiday         = ?," & _
             "summons                  = ?," & _
             "overtime_begin           = ?," & _
             "overtime_end             = ?," & _
             "rest_begin               = ?," & _
             "rest_end                 = ?," & _
             "overtime                 = ?," & _
             "overtimelate             = ?," & _
             "holidayshift             = ?," & _
             "holidayshiftovertime     = ?," & _
             "holidayshiftlate         = ?," & _
             "holidayshiftovertimelate = ?," & _
             "requesttime              = ?," & _
             "requesttime_begin        = ?," & _
             "requesttime_end          = ?," & _
             "latetime                 = ?," & _
             "latetime_begin           = ?," & _
             "latetime_end             = ?," & _
             "nightduty                = ?," & _
             "dayduty                  = ?," & _
             "operator                 = ?," & _
             "vacationtime             = ?," & _
             "vacationtime_begin       = ?," & _
             "vacationtime_end         = ?," & _
             "memo                     = ?," & _
             "memo2                    = ?," & _
             "is_error                 = ?," & _
             "work_begin               = ?," & _
             "work_end                 = ?," & _
             "break_begin1             = ?," & _
             "break_end1               = ?," & _
             "break_begin2             = ?," & _
             "break_end2               = ?," & _
             "workmin                  = ?," & _
             "weekovertime             = ? " & _
             "WHERE id                 = ? " & _
             "AND CONVERT(int,updatetime) = ?"
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,1, wk_morningwork)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,1, wk_afternoonwork)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,1, wk_morningholiday)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,1, wk_afternoonholiday)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,1, wk_summons)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_overtime_begin)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_overtime_end)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_rest_begin)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_rest_end)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_overtime)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_overtimelate)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_holidayshift)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_holidayshiftovertime)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_holidayshiftlate)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_holidayshiftovertimelate)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_requesttime)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_requesttime_begin)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_requesttime_end)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_latetime)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_latetime_begin)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_latetime_end)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,1, wk_nightduty)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,1, wk_dayduty)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,1, wk_operator)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_vacationtime)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_vacationtime_begin)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_vacationtime_end)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,100, wk_memo)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,2, wk_memo2)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,129,,1, wk_is_error)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_work_begin)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_work_end)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_break_begin1)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_break_end1)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_break_begin2)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_break_end2)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,5  ,,4, wk_workmin)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,4, wk_weekovertime)
    
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,5,, -1, wk_id)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,5,, -1, wk_updatetime)
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
End If
%>
