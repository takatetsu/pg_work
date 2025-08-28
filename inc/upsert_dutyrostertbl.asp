<%
' ------------------------------------------------------------------------------
' dutyrostertbl に対しての更新処理
' 登録された対象社員対象月の worktbl を全件読み集計処理を行う。
' ------------------------------------------------------------------------------
Dim Rs_worktbl_sum
Dim Rs_worktbl_sum_cmd
Set Rs_worktbl_sum_cmd = Server.CreateObject ("ADODB.Command")
Rs_worktbl_sum_cmd.ActiveConnection = MM_workdbms_STRING
Rs_worktbl_sum_cmd.CommandText = "SELECT * FROM worktbl " & _
    "WHERE personalcode = ? AND workingdate LIKE ? ORDER BY workingdate ASC"
Rs_worktbl_sum_cmd.Prepared = true
Rs_worktbl_sum_cmd.Parameters.Append Rs_worktbl_sum_cmd.CreateParameter _
    ("param1", 200, 1, 5, target_personalcode)
Rs_worktbl_sum_cmd.Parameters.Append Rs_worktbl_sum_cmd.CreateParameter _
    ("param2", 200, 1, 7, ymb & "%")
Set Rs_worktbl_sum = Rs_worktbl_sum_cmd.Execute
total_workdays                  = lastDay ' 可出勤日数    workdays
total_workholidays              = 0       ' 代休日数      workholidays
total_absencedays               = 0       ' 欠勤日数      absencedays
total_paidvacations             = 0       ' 有休日数      paidvacations
total_preservevacations         = 0       ' 保存休日数    preservevacations
total_specialvacations          = 0       ' 特休日数      specialvacations
total_holidayshifts             = 0       ' 休出日数      holidayshifts
total_realworkdays              = 0       ' 実出勤日数    realworkdays
total_shortdays                 = 0       ' 遅早回数      shortdays
total_nightduty_a               = 0       ' 宿直Ａ回数    nightduty_a
total_nightduty_b               = 0       ' 宿直Ｂ回数    nightduty_b
total_nightduty_c               = 0       ' 宿直Ｃ回数    nightduty_c
total_nightduty_d               = 0       ' 宿直Ｄ回数    nightduty_d
total_holidaypremium            = 0       ' 休日割増      holidaypremium
total_dayduty                   = 0       ' 日直          dayduty
total_shiftwork_kou             = 0       ' 交替勤務甲番  shiftwork_kou
total_shiftwork_otsu            = 0       ' 交替勤務乙番  shiftwork_otsu
total_shiftwork_hei             = 0       ' 交替勤務丙番  shiftwork_hei
total_shiftwork_a               = 0       ' 交替勤務A番   shiftwork_a
total_shiftwork_b               = 0       ' 交替勤務B番   shiftwork_b
total_summons                   = 0       ' 呼出通常回数  summons
total_summonslate               = 0       ' 呼出深夜回数  summonslate
total_yearend1230               = 0       ' 年末年始1230  yearend1230
total_yearend1231               = 0       ' 年末年始1231  yearend1231
total_workholidaytime           = 0       ' 時間代休      workholidaytime
' 時間有休          vacationtime
total_vacationtime              = Request.Form("vacationtimeHidden")
total_latepremium               = 0       ' 深夜割増      latepremium
total_overtime                  = 0       ' 時間外        overtime
total_holidayshifttime          = 0       ' 休日出勤時間  holidayshifttime
total_holidayshiftovertime      = 0       ' 休出時間外    holidayshiftovertime
total_holidayshiftlate          = 0       ' 休出深夜業    holidayshiftlate
total_overtimelate              = 0       ' 時間外深夜業  overtimelate
total_holidayshiftovertimelate  = 0       ' 休出時間外深  holidayshiftovertimelate
' 当月末有給休暇残  vacationnumber
total_vacationnumber            = Request.Form("vacationnumberHidden")
' 当月末振替休日残  holidaynumber
total_holidaynumber             = Request.Form("holidaynumberHidden" )

' 前月必要勤務時間  lastmonth_currentworkmi
lastmonth_currentworkmin        = CInt(Request.Form("lastmonth_currentworkminHidden"))
' 前月実労働時間  lastmonth_workingmins
lastmonth_workingmins           = CInt(Request.Form("lastmonth_workingminsHidden"))
' 前月労働時間不足分を今月時間外集計しないためのエリア(分)
poor_mins                       = 0       ' 前月不足労働時間(分)
If workshift = "9" Then
    If lastmonth_workingmins < lastmonth_currentworkmin Then
        poor_mins               = lastmonth_currentworkmin - lastmonth_workingmins
    End If
End If

total_saturdayworkmin           = 0     ' 土曜日勤務時間
total_weekdaysworkmin           = 0     ' 平日勤務時間
realworkmin                     = 0     ' 勤務実績
baseworkmin                     = Request.Form("thismonth_baseminHidden") ' 当月基準労働時間
currentworkmin                  = baseworkmin   ' 当月労働時間
legalholiday_extra_min          = 0     ' 法定休日割増時間(分)
weekovertime                    = 0     ' 週超過労働時間

While (NOT Rs_worktbl_sum.EOF)
    operatorAddDay = 0  ' 交替勤務時の午前への加算日数
    If (Rs_worktbl_sum.Fields.Item("operator").value = "1"  Or _
        Rs_worktbl_sum.Fields.Item("operator").value = "2"  Or _
        Rs_worktbl_sum.Fields.Item("operator").value = "3"  Or _
        Rs_worktbl_sum.Fields.Item("operator").value = "5"  Or _
        Rs_worktbl_sum.Fields.Item("operator").value = "6") Then
        operatorAddDay = 0.5
    End If
    If (Rs_worktbl_sum.Fields.Item("operator").value = "4") Then
        operatorAddDay = 1.0
    End If
    ' 交替勤務時の午前への加算日数を可出勤日数に集計
    If (Rs_worktbl_sum.Fields.Item("morningwork"   ).value <> "0"   Or  _
       (Rs_worktbl_sum.Fields.Item("morningholiday").value <> "1"   And _
        Rs_worktbl_sum.Fields.Item("morningholiday").value <> "2"   And _
        Rs_worktbl_sum.Fields.Item("morningholiday").value <> "A")) Then
        total_workdays     = total_workdays + operatorAddDay
    Else
        If (Rs_worktbl_sum.Fields.Item("morningholiday").value = "1"  Or _
            Rs_worktbl_sum.Fields.Item("morningholiday").value = "2"  Or _
            Rs_worktbl_sum.Fields.Item("morningholiday").value = "A") Then
            total_workdays = total_workdays - operatorAddDay
        End If
    End If
    ' 実出勤日数
    If (Rs_worktbl_sum.Fields.Item("morningwork"  ).value = "1"  Or _
        Rs_worktbl_sum.Fields.Item("morningwork"  ).value = "4"  Or _
        Rs_worktbl_sum.Fields.Item("morningwork"  ).value = "5"  Or _
        Rs_worktbl_sum.Fields.Item("morningwork"  ).value = "9") Then
        total_realworkdays = total_realworkdays + 0.5 + operatorAddDay
    End If
    If (Rs_worktbl_sum.Fields.Item("afternoonwork").value = "1"  Or _
        Rs_worktbl_sum.Fields.Item("afternoonwork").value = "4"  Or _
        Rs_worktbl_sum.Fields.Item("afternoonwork").value = "5"  Or _
        Rs_worktbl_sum.Fields.Item("afternoonwork").value = "9") Then
        total_realworkdays = total_realworkdays + 0.5
    End If
    ' 交替勤務
    Select Case Rs_worktbl_sum.Fields.Item("operator").value
        Case "1":
            If (Rs_worktbl_sum.Fields.Item("morningwork"  ).value <> "0"  Or _
                Rs_worktbl_sum.Fields.Item("afternoonwork").value <> "0") Then
                total_shiftwork_kou  = total_shiftwork_kou  + 1
            End If
        Case "2":
            If (Rs_worktbl_sum.Fields.Item("morningwork"  ).value <> "0"  Or _
                Rs_worktbl_sum.Fields.Item("afternoonwork").value <> "0") Then
                total_shiftwork_otsu = total_shiftwork_otsu + 1
            End If
        Case "3":
            If (Rs_worktbl_sum.Fields.Item("morningwork"  ).value <> "0"  Or _
                Rs_worktbl_sum.Fields.Item("afternoonwork").value <> "0") Then
                total_shiftwork_kou  = total_shiftwork_kou  + 1
            End If
        Case "4"
            If (Rs_worktbl_sum.Fields.Item("morningwork"  ).value <> "0"  Or _
                Rs_worktbl_sum.Fields.Item("afternoonwork").value <> "0") Then
                total_shiftwork_otsu = total_shiftwork_otsu + 1
            End If
        Case "7"
            If (Rs_worktbl_sum.Fields.Item("morningwork"  ).value <> "0"  Or _
                Rs_worktbl_sum.Fields.Item("afternoonwork").value <> "0") Then
                total_shiftwork_a = total_shiftwork_a + 1
            End If
         Case "8"
            If (Rs_worktbl_sum.Fields.Item("morningwork"  ).value <> "0"  Or _
                Rs_worktbl_sum.Fields.Item("afternoonwork").value <> "0") Then
                total_shiftwork_b = total_shiftwork_b + 1
            End If
       Case Else
    End Select

    If (Rs_worktbl_sum.Fields.Item("morningwork").value = "1"  Or _
        Rs_worktbl_sum.Fields.Item("morningwork").value = "5") Then
        ' 振替出勤のとき
        ' 可出勤日数
        total_workdays      = total_workdays      + 0.5 + operatorAddDay
        ' 振替残日数 = 前月末振替残日数
        '            + (振替出勤(午前)* 0.5)
        '            + (振替出勤(午後)* 0.5)
        '            - (振替休日(午前)* 0.5)
        '            - (振替休日(午後)* 0.5)
        total_holidaynumber = total_holidaynumber + 0.5 + operatorAddDay
        If workshift = "9" Then
            ' フレックス勤務者が振替出勤したとき、当月労働時間を加算
            currentworkmin = currentworkmin + 210
        End If
    End If
    If (Rs_worktbl_sum.Fields.Item("afternoonwork").value = "1"  Or _
        Rs_worktbl_sum.Fields.Item("afternoonwork").value = "5") Then
        ' 振替出勤のとき
        ' 可出勤日数
        total_workdays      = total_workdays      + 0.5
        ' 振替残日数 = 前月末振替残日数
        '            + (振替出勤(午前)* 0.5)
        '            + (振替出勤(午後)* 0.5)
        '            - (振替休日(午前)* 0.5)
        '            - (振替休日(午後)* 0.5)
        total_holidaynumber = total_holidaynumber + 0.5
        If workshift = "9" Then
            ' フレックス勤務者が振替出勤したとき、当月労働時間を加算
            currentworkmin = currentworkmin + 250
        End If
    End If
    Select Case Rs_worktbl_sum.Fields.Item("morningholiday").value
        Case "1"    '公休日
            ' 可出勤日数
            total_workdays          = total_workdays          - 0.5 - operatorAddDay
        Case "2"    '振替休日
            ' 可出勤日数
            total_workdays          = total_workdays          - 0.5 - operatorAddDay
            ' 振替残日数 = 前月末振替残日数
            '            + (振替出勤(午前)* 0.5)
            '            + (振替出勤(午後)* 0.5)
            '            - (振替休日(午前)* 0.5)
            '            - (振替休日(午後)* 0.5)
            total_holidaynumber     = total_holidaynumber     - 0.5 - operatorAddDay
            If workshift = "9" Then
                ' フレックス勤務者が振替休暇のとき、当月労働時間を減算
                currentworkmin = currentworkmin - 210
            End If
        Case "3"    '有給休暇
            ' 有休日数
            total_paidvacations     = total_paidvacations     + 0.5 + operatorAddDay
            ' 有給休暇日数
            total_vacationnumber    = total_vacationnumber    - 0.5 - operatorAddDay
        Case "4"    '代替休暇
            ' 代休日数
            total_workholidays      = total_workholidays      + 0.5 + operatorAddDay
        Case "5"    '特別休暇
            ' 特休日数
            total_specialvacations  = total_specialvacations  + 0.5 + operatorAddDay
        Case "6"    '保存休暇
            ' 保存休日数
            total_preservevacations = total_preservevacations + 0.5 + operatorAddDay
        Case "7"    '半日欠勤
            ' 欠勤日数
            total_absencedays       = total_absencedays       + 0.5 + operatorAddDay
        Case "9"    'コアタイム有休
            ' 有休日数
            total_paidvacations     = total_paidvacations     + 0.25
            ' 有給休暇日数
            total_vacationnumber    = total_vacationnumber    - 0.25
         Case "A"    '法定休日
            ' 可出勤日数
            total_workdays          = total_workdays          - 0.5
        Case "B"    '育児休業
            If workshift = "9" Then
                currentworkmin = currentworkmin - 210
            End If
            total_absencedays       = total_absencedays       + 0.5 + operatorAddDay
    End Select
    Select Case Rs_worktbl_sum.Fields.Item("afternoonholiday").value
        Case "1"    '公休日
            ' 可出勤日数
            total_workdays          = total_workdays          - 0.5
        Case "2"    '振替休日
            ' 可出勤日数
            total_workdays          = total_workdays          - 0.5
            ' 振替残日数 = 前月末振替残日数
            '            + (振替出勤(午前)* 0.5)
            '            + (振替出勤(午後)* 0.5)
            '            - (振替休日(午前)* 0.5)
            '            - (振替休日(午後)* 0.5)
            total_holidaynumber     = total_holidaynumber     - 0.5
            If workshift = "9" Then
                ' フレックス勤務者が振替休暇のとき、当月労働時間を減算
                currentworkmin = currentworkmin - 250
            End If
        Case "3"    '有給休暇
            ' 有休日数
            total_paidvacations     = total_paidvacations     + 0.5
            ' 有給休暇日数
            total_vacationnumber    = total_vacationnumber    - 0.5
        Case "4"    '代替休暇
            ' 代休日数
            total_workholidays      = total_workholidays      + 0.5
        Case "5"    '特別休暇
            ' 特休日数
            total_specialvacations  = total_specialvacations  + 0.5
        Case "6"    '保存休暇
            ' 保存休日数
            total_preservevacations = total_preservevacations + 0.5
        Case "7"    '半日欠勤
            ' 欠勤日数
            total_absenceDays       = total_absenceDays       + 0.5
        Case "9"    'コアタイム有休
            ' 有休日数
            total_paidvacations     = total_paidvacations     + 0.25
            ' 有給休暇日数
            total_vacationnumber    = total_vacationnumber    - 0.25
        Case "A"    '法定休日
            ' 可出勤日数
            total_workdays          = total_workdays          - 0.5
        Case "B"    '育児休業
            If workshift = "9" Then
                currentworkmin = currentworkmin - 250
            End If
            total_absenceDays       = total_absenceDays       + 0.5
    End Select

    Select Case Rs_worktbl_sum.Fields.Item("summons").value
        Case "1"    '呼出通常
            total_summons     = total_summons     + 1     ' 呼出回数通常
        Case "2"    '呼出深夜
            total_summonslate = total_summonslate + 1     ' 呼出回数深夜
    End Select
    ' 時間外
    If (Len(Rs_worktbl_sum.Fields.Item("overtime"                ).value) > 0) Then
        total_overtime = total_overtime _
                       + time2Min(Rs_worktbl_sum.Fields.Item("overtime").value)
    End If
    
    If workshift = "9" Then
        If ((Rs_worktbl_sum.Fields.Item("morningholiday"  ).value <> "A" And _
             Rs_worktbl_sum.Fields.Item("afternoonholiday").value <> "A") And _
            (Rs_worktbl_sum.Fields.Item("morningwork"  ).value = "2" Or _
             Rs_worktbl_sum.Fields.Item("morningwork"  ).value = "6" Or _
             Rs_worktbl_sum.Fields.Item("afternoonwork").value = "2" Or _
             Rs_worktbl_sum.Fields.Item("afternoonwork").value = "6")) Then
           ' 休日出勤 フレックス勤務者は休出時の勤務時間を集計する。法定休日は別で集計されている。
            total_holidayshifttime = total_holidayshifttime + Rs_worktbl_sum.Fields.Item("workmin").Value
        End If
    Else
        ' 休日出勤
        If (Len(Rs_worktbl_sum.Fields.Item("holidayshift"            ).value) > 0) Then
            total_holidayshifttime          = total_holidayshifttime _
                                            + time2Min(Rs_worktbl_sum.Fields.Item(_
                                                "holidayshift").value)
        End If
        ' 休出時間外
        If (Len(Rs_worktbl_sum.Fields.Item("holidayshiftovertime"    ).value) > 0) Then
            total_holidayshiftovertime      = total_holidayshiftovertime _
                                            + time2Min(Rs_worktbl_sum.Fields.Item(_
                                                "holidayshiftovertime").value)
        End If
        ' 休出深夜
        If (Len(Rs_worktbl_sum.Fields.Item("holidayshiftlate"        ).value) > 0) Then
            total_holidayshiftlate          = total_holidayshiftlate _
                                            + time2Min(Rs_worktbl_sum.Fields.Item(_
                                                "holidayshiftlate").value)
        End If
        ' 時間外深夜
        If (Len(Rs_worktbl_sum.Fields.Item("overtimelate"            ).value) > 0) Then
            total_overtimelate              = total_overtimelate _
                                            + time2Min(Rs_worktbl_sum.Fields.Item(_
                                                "overtimelate").value)
        End If
        ' 休出時間外深夜
        If (Len(Rs_worktbl_sum.Fields.Item("holidayshiftovertimelate").value) > 0) Then
            total_holidayshiftovertimelate  = total_holidayshiftovertimelate _
                                            + time2Min(Rs_worktbl_sum.Fields.Item(_
                                                "holidayshiftovertimelate").value)
        End If
    End If
    ' 時間代休
    If (Len(Rs_worktbl_sum.Fields.Item("requesttime"             ).value) > 0) Then
        total_workholidaytime           = total_workholidaytime _
                                        + time2Min(Rs_worktbl_sum.Fields.Item(_
                                            "requesttime").value)
    End If
    ' 時間有休
    If (Len(Rs_worktbl_sum.Fields.Item("vacationtime"            ).value) > 0) Then
        total_vacationtime              = total_vacationtime _
                                        + time2Min(Rs_worktbl_sum.Fields.Item(_
                                            "vacationtime").value)
    End If
    ' 深夜割増
    If (Len(Rs_worktbl_sum.Fields.Item("latetime"                ).value) > 0) Then
        total_latepremium               = total_latepremium _
                                        + time2Min(Rs_worktbl_sum.Fields.Item(_
                                            "latetime").value)
    End If
    ' 宿直件数
    If (Rs_worktbl_sum.Fields.Item("nightduty").value = "1") Then
        total_nightduty_a               = total_nightduty_a     + 1
    End If
    If (Rs_worktbl_sum.Fields.Item("nightduty").value = "2") Then
        total_nightduty_b               = total_nightduty_b     + 1
    End If
    ' 休日割増
    If (Rs_worktbl_sum.Fields.Item("morningholiday"  ).value =  "1"  Or  _
        Rs_worktbl_sum.Fields.Item("morningholiday"  ).value =  "2"  Or  _
        Rs_worktbl_sum.Fields.Item("morningholiday"  ).value =  "4"  Or  _
        Rs_worktbl_sum.Fields.Item("morningholiday"  ).value =  "5"  Or  _
        Rs_worktbl_sum.Fields.Item("morningholiday"  ).value =  "A") And _
       (Rs_worktbl_sum.Fields.Item("afternoonholiday").value =  "1"  Or  _
        Rs_worktbl_sum.Fields.Item("afternoonholiday").value =  "2"  Or  _
        Rs_worktbl_sum.Fields.Item("afternoonholiday").value =  "4"  Or  _
        Rs_worktbl_sum.Fields.Item("afternoonholiday").value =  "5"  Or  _
        Rs_worktbl_sum.Fields.Item("afternoonholiday").value =  "A") And _
       (Rs_worktbl_sum.Fields.Item("morningwork"     ).value <> "1"  And _
        Rs_worktbl_sum.Fields.Item("morningwork"     ).value <> "4"  And _
        Rs_worktbl_sum.Fields.Item("morningwork"     ).value <> "5"  And _
        Rs_worktbl_sum.Fields.Item("morningwork"     ).value <> "9") And _
       (Rs_worktbl_sum.Fields.Item("afternoonwork"   ).value <> "1"  And _
        Rs_worktbl_sum.Fields.Item("afternoonwork"   ).value <> "4"  And _
        Rs_worktbl_sum.Fields.Item("afternoonwork"   ).value <> "5"  And _
        Rs_worktbl_sum.Fields.Item("afternoonwork"   ).value <> "9") Then
        If Rs_worktbl_sum.Fields.Item("nightduty").value <> "0" Then
            total_holidaypremium = total_holidaypremium  + 1
        End If
    End If
    ' 日直件数
    If (Rs_worktbl_sum.Fields.Item("dayduty"  ).value = "1") Then
        total_dayduty = total_dayduty + 1
    End If

'                ' 実出勤日数 = 可出勤日数 - 代休日数   - 欠勤日数
'                '            - 有休日数   - 保存休日数 - 特休日数
'                total_realworkdays = total_workdays             _
'                                   - total_workholidays         _
'                                   - total_absencedays          _
'                                   - total_paidvacations        _
'                                   - total_preservevacations    _
'                                   - total_specialvacations
    If workshift = "1" Or workshift = "2" Or workshift = "3" Then
        tempMin = 0
        If (Rs_worktbl_sum.Fields.Item("morningwork"  ).value <> "0"  Or _
            Rs_worktbl_sum.Fields.Item("afternoonwork").value <> "0") Then
            If workshift = "1" Then
                ' 全日勤務 08:50-17:30 460分
                tempMin = 460
            ElseIf workshift = "2" Then
                ' 午前勤務 08:50-13:00 250分
                tempMin = 250
            ElseIf MM_workshift = "3" Then
                ' 午後勤務 13:00-17:30 270分
                tempMin = 270
            End If
            ' 時間外加算
            If (Len(Rs_worktbl_sum.Fields.Item("overtime").value) > 0) Then
                tempMin = tempMin + + time2Min(Rs_worktbl_sum.Fields.Item("overtime").value)
            End If
        End If
        tempDay = Left(Rs_worktbl_sum.Fields.Item("workingdate").value,4) & "/" & _
                  Mid(Rs_worktbl_sum.Fields.Item("workingdate").Value,5,2) & "/" & _
                  Right(Rs_worktbl_sum.Fields.Item("workingdate").Value,2)
        tempWeek = Weekday(tempDay)
        If (tempWeek = "7") Then
            ' コミュニケータ土曜日勤務時間
            total_saturdayworkmin = total_saturdayworkmin + tempMin
        Else
            ' コミュニケータ平日勤務時間
            total_weekdaysworkmin = total_weekdaysworkmin + tempMin
        End If
    End If
    
    If workshift = "9" Then
        If Len(Rs_worktbl_sum.Fields.Item("workmin").value) > 0 Then
            realworkmin = realworkmin + Rs_worktbl_sum.Fields.Item("workmin").value
        End If
        ' 法定休日かつ振替出勤のとき、割増時間算出 (法定休日の時間外のみ集計する)
        If (Rs_worktbl_sum.Fields.Item("morningholiday"  ).value = "A" And _
            Rs_worktbl_sum.Fields.Item("afternoonholiday").value = "A" And _
           (Rs_worktbl_sum.Fields.Item("morningwork"     ).value = "1" Or _
            Rs_worktbl_sum.Fields.Item("morningwork"     ).value = "2" Or _
            Rs_worktbl_sum.Fields.Item("morningwork"     ).value = "3" Or _
            Rs_worktbl_sum.Fields.Item("morningwork"     ).value = "5" Or _
            Rs_worktbl_sum.Fields.Item("morningwork"     ).value = "6" Or _
            Rs_worktbl_sum.Fields.Item("afternoonwork"   ).value = "1" Or _
            Rs_worktbl_sum.Fields.Item("afternoonwork"   ).value = "2" Or _
            Rs_worktbl_sum.Fields.Item("afternoonwork"   ).value = "3" Or _
            Rs_worktbl_sum.Fields.Item("afternoonwork"   ).value = "5" Or _
            Rs_worktbl_sum.Fields.Item("afternoonwork"   ).value = "6")) Then
            legalholiday_extra_min = legalholiday_extra_min _
                                   + Rs_worktbl_sum.Fields.Item("workmin").value
        End If
    End If

    If Not workshift = "9" Then
        weekovertime = weekovertime + time2Min(editTime(Rs_worktbl_sum.Fields.Item("weekovertime").value))
    End If

    Rs_worktbl_sum.MoveNext()
Wend
Rs_worktbl_sum.Close()
Set Rs_worktbl_sum = Nothing

If (Request.Form("dutyrostertbl_id") = "") Then
    ' INSERT
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_workdbms_STRING
    MM_editCmd.CommandText = "INSERT INTO dutyrostertbl VALUES(DEFAULT, " & _
                             "?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, " & _
                             "?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, " & _
                             "?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,5, target_personalcode            ) ' 個人コード
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,6, Left(Request.Form("ymd")(1), 6)) ' 年月分
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,,6, total_workdays                 ) ' 可出勤日数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,,6, total_workholidays             ) ' 代休日数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,,6, total_absencedays              ) ' 欠勤日数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,,6, total_paidvacations            ) ' 有休日数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,,6, total_preservevacations        ) ' 保存休日数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,,6, total_specialvacations         ) ' 特休日数
    ' 休出日数の小数点計算
    temp_holidayshifts = mm2FloatDay(total_holidayshifttime + total_holidayshiftlate)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, temp_holidayshifts             ) ' 休出日数
    If workshift = "9" Then
        ' フレックス勤務者の場合のみ実出勤日数に休出日数を加算
        total_realworkdays = total_realworkdays + temp_holidayshifts
    End If
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_realworkdays             ) ' 実出勤日数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_shortdays                ) ' 遅早回数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_nightduty_a              ) ' 宿直Ａ回数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_nightduty_b              ) ' 宿直Ｂ回数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_nightduty_c              ) ' 宿直Ｃ回数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_nightduty_d              ) ' 宿直Ｄ回数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_holidaypremium           ) ' 休日割増
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_dayduty                  ) ' 日直
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_shiftwork_kou            ) ' 交替勤務甲番
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_shiftwork_otsu           ) ' 交替勤務乙番
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_shiftwork_hei            ) ' 交替勤務丙番
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_summons                  ) ' 呼出通常回数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_summonslate              ) ' 呼出深夜回数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_yearend1230              ) ' 年末年始1230
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_yearend1231              ) ' 年末年始1231
    If workshift = "9" Then
        ' フレックス勤務者のとき、時間代休に法定休日割増時間を設定する。(HiPer-BTでの給与計算のための措置)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(legalholiday_extra_min)) ' 時間代休
    Else
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_workholidaytime))  ' 時間代休
    End If
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_latepremium)) ' 深夜割増
    If workshift = "0" Or workshift = "9" Then
        ' 一般社員(お客さまセンターオペレータ以外)のとき)
        If workshift = "9" Then
            ' フレックス勤務者のとき
            If (realworkmin - currentworkmin - poor_mins) > 0 Then
                ' 実勤務時間が基準勤務時間より大きいとき、差分を時間外に設定
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(realworkmin - currentworkmin - poor_mins)) ' 時間外
            Else
                ' 実勤務時間が基準勤務時間以下のとき、時間外は0
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0)                                      ' 時間外
            End If
            ' 休日出勤時間外に法定休日割増時間を設定する。(HiPer-BTでの給与計算のための措置)
            MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(legalholiday_extra_min + total_holidayshifttime)) ' 休日出勤時間
        Else
            MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_overtime            )) ' 時間外
            MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_holidayshifttime    )) ' 休日出勤時間
        End If
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_holidayshiftovertime    )) ' 休出時間外
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_holidayshiftlate        )) ' 休出深夜業
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_overtimelate            )) ' 時間外深夜業
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_holidayshiftovertimelate)) ' 休出時間外深夜
    Else
        ' お客さまセンターオペレータのとき
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0) ' 時間外
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0) ' 休日出勤時間
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0) ' 休出時間外
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0) ' 休出深夜業
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0) ' 時間外深夜業
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0) ' 休出時間外深夜
    End If
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_vacationnumber) ' 当月末有給休暇残
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_holidaynumber)  ' 当月末振替休日残
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_vacationtime)   ' 時間有休
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_shiftwork_a)    ' 交替勤務A番
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_shiftwork_b)    ' 交替勤務B番
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_saturdayworkmin)) ' 土曜日勤務時間(コミュニケータ)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_weekdaysworkmin)) ' 平日勤務時間(コミュニケータ)
    If workshift = "9" Then
        ' フレックス勤務者
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, realworkmin)           ' 労働時間
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, currentworkmin)        ' 当月労働時間
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, legalholiday_extra_min)' 法定休日割増時間
    Else
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0)                     ' 労働時間
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0)                     ' 当月労働時間
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0)                     ' 法定休日割増時間
    End If
    If Not workshift = "9" Then
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(weekovertime))' 週超過労働時間
    Else
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0)  ' 週超過労働時間
    End If
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
Else
    ' UPDATE
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_workdbms_STRING
    MM_editCmd.CommandText = "UPDATE dutyrostertbl SET "       & _
                                "workdays                   = ?, " & _
                                "workholidays               = ?, " & _
                                "absencedays                = ?, " & _
                                "paidvacations              = ?, " & _
                                "preservevacations          = ?, " & _
                                "specialvacations           = ?, " & _
                                "holidayshifts              = ?, " & _
                                "realworkdays               = ?, " & _
                                "shortdays                  = ?, " & _
                                "nightduty_a                = ?, " & _
                                "nightduty_b                = ?, " & _
                                "nightduty_c                = ?, " & _
                                "nightduty_d                = ?, " & _
                                "holidaypremium             = ?, " & _
                                "dayduty                    = ?, " & _
                                "shiftwork_kou              = ?, " & _
                                "shiftwork_otsu             = ?, " & _
                                "shiftwork_hei              = ?, " & _
                                "summons                    = ?, " & _
                                "summonslate                = ?, " & _
                                "yearend1230                = ?, " & _
                                "yearend1231                = ?, " & _
                                "workholidaytime            = ?, " & _
                                "latepremium                = ?, " & _
                                "overtime                   = ?, " & _
                                "holidayshifttime           = ?, " & _
                                "holidayshiftovertime       = ?, " & _
                                "holidayshiftlate           = ?, " & _
                                "overtimelate               = ?, " & _
                                "holidayshiftovertimelate   = ?, " & _
                                "vacationnumber             = ?, " & _
                                "holidaynumber              = ?, " & _
                                "vacationtime               = ?, " & _
                                "shiftwork_a                = ?, " & _
                                "shiftwork_b                = ?, " & _
                                "saturday_workmin           = ?, " & _
                                "weekdays_workmin           = ?, " & _
                                "workingmins                = ?, " & _
                                "currentworkmin             = ?, " & _
                                "legalholiday_extra_min     = ?, " & _
                                "weekovertime               = ?  " & _
                                "WHERE id                   = ?"
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_workdays)         ' 可出勤日数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_workholidays)     ' 代休日数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_absencedays)      ' 欠勤日数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_paidvacations)    ' 有休日数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_preservevacations)' 保存休日数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_specialvacations) ' 特休日数
    ' 休出日数の小数点計算
    temp_holidayshifts = mm2FloatDay(total_holidayshifttime + total_holidayshiftlate)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, temp_holidayshifts) ' 休出日数
    If workshift = "9" Then
        ' フレックス勤務者の場合のみ実出勤日数に休出日数を加算
        total_realworkdays = total_realworkdays + temp_holidayshifts
    End If
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_realworkdays) ' 実出勤日数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_shortdays)    ' 遅早回数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_nightduty_a)  ' 宿直Ａ回数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_nightduty_b)  ' 宿直Ｂ回数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_nightduty_c)  ' 宿直Ｃ回数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_nightduty_d)  ' 宿直Ｄ回数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_holidaypremium) ' 休日割増
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_dayduty)      ' 日直
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_shiftwork_kou)' 交替勤務甲番
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_shiftwork_otsu) ' 交替勤務乙番
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_shiftwork_hei)' 交替勤務丙番
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_summons)      ' 呼出通常回数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_summonslate)  ' 呼出深夜回数
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_yearend1230)  ' 年末年始1230
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_yearend1231)  ' 年末年始1231
    If workshift = "9" Then
        ' フレックス勤務者のとき、時間代休に法定休日割増時間を設定する。(HiPer-BTでの給与計算のための措置)
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(legalholiday_extra_min)) ' 時間代休
    Else
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_workholidaytime))  ' 時間代休
    End If
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_latepremium)) ' 深夜割増
    If workshift = "0" Or workshift = "9" Then
        ' 一般社員(お客さまセンターオペレータ以外)のとき)
        If workshift = "9" Then
            ' フレックス勤務者のとき
            If (realworkmin - currentworkmin - poor_mins) > 0 Then
                ' 実勤務時間が基準勤務時間より大きいとき、差分を時間外に設定
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(realworkmin - currentworkmin - poor_mins)) ' 時間外
            Else
                ' 実勤務時間が基準勤務時間以下のとき、時間外は0
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0)                                      ' 時間外
            End If
            ' フレックス勤務者のとき、休日出勤時間外に法定休日割増時間を設定する。(HiPer-BTでの給与計算のための措置)
            MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(legalholiday_extra_min + total_holidayshifttime)) ' 休日出勤時間
        Else
            MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_overtime            )) ' 時間外
            MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_holidayshifttime    )) ' 休日出勤時間
        End If
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_holidayshiftovertime    )) ' 休出時間外
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_holidayshiftlate        )) ' 休出深夜業
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_overtimelate            )) ' 時間外深夜業
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_holidayshiftovertimelate)) ' 休出時間外深夜
    Else
        ' お客さまセンターオペレータのとき
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0) ' 時間外
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0) ' 休日出勤時間
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0) ' 休出時間外
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0) ' 休出深夜業
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0) ' 時間外深夜業
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0) ' 休出時間外深夜
    End If
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_vacationnumber) ' 当月末有給休暇残
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_holidaynumber)  ' 当月末振替休日残
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_vacationtime)   ' 時間有休
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_shiftwork_a)    ' 交替勤務A番
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, total_shiftwork_b)    ' 交替勤務B番
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_saturdayworkmin)) ' 土曜日勤務時間(コミュニケータ)
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(total_weekdaysworkmin)) ' 平日勤務時間(コミュニケータ)
    If workshift = "9" Then
        ' フレックス勤務者
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, realworkmin)           ' 労働時間
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, currentworkmin)        ' 当月労働時間
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, legalholiday_extra_min)' 法定休日割増時間
    Else
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0)                     ' 労働時間
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0)                     ' 当月労働時間
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0)                     ' 法定休日割増時間
    End If
    If Not workshift = "9" Then
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, mm2Float(weekovertime))' 週超過労働時間
    Else
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,,6, 0)  ' 週超過労働時間
    End If
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(, 5,, 6, Request.Form("dutyrostertbl_id")) ' id
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
End If
%>