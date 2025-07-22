<%
' 勤怠情報初期化
v_worktbl_id                  = ""
v_updatetime                  = ""
v_morningwork                 = ""
v_afternoonwork               = ""
v_morningholiday              = ""
v_afternoonholiday            = ""
v_workmin                     = ""
v_work_begin                  = ""
v_work_end                    = ""
v_break_begin1                = ""
v_break_end1                  = ""
v_break_begin2                = ""
v_break_end2                  = ""
v_summons                     = ""
v_overtime_begin              = ""
v_overtime_end                = ""
v_rest_begin                  = ""
v_rest_end                    = ""
v_overtime                    = ""
v_overtimelate                = ""
v_holidayshift                = ""
v_holidayshiftovertime        = ""
v_holidayshiftlate            = ""
v_holidayshiftovertimelate    = ""
v_requesttime                 = ""
v_requesttime_begin           = ""
v_requesttime_end             = ""
v_vacationtime                = ""
v_vacationtime_begin          = ""
v_vacationtime_end            = ""
v_latetime                    = ""
v_latetime_begin              = ""
v_latetime_end                = ""
v_weekovertime                = ""
v_is_approval                 = ""
v_nightduty                   = ""
v_nightduty2                  = ""  ' 前日宿直区分(労働時間適正化で追加)
v_dayduty                     = ""
v_operator                    = ""
v_operator2                   = ""  ' 前日交替勤務(労働時間適正化で追加)
v_memo                        = ""
v_memo2                       = ""
v_is_error                    = ""
v_overwork                    = ""  ' 週労働時間が40時間を超えた日には"warning"がセットされる

' =================================================================
' 勤怠テーブル処理
' =================================================================
If screen = "0" Then
    text_disabled = ""
Else
    text_disabled = "disabled"
End If

text_approval_disabled = "disabled" ' 上長チェックをデフォルトで無効化(下記で解除)
text_timecard_disabled = "disabled" ' タイムカードをデフォルトで無効化(下記で解除)
If Not Rs_worktbl.EOF Then
    If (Rs_worktbl.Fields.Item("workingdate").Value = (ymb & Right("0"&i, 2))) Then
        ' 勤怠データ有り
        v_worktbl_id               = Trim(Rs_worktbl.Fields.Item("id"                          ).Value) ' ID
        v_updatetime               = Trim(Rs_worktbl.Fields.Item("inttime"                     ).Value) ' UPDATETIME を int 変換したもの
        v_morningwork              = Trim(Rs_worktbl.Fields.Item("morningwork"                 ).Value) ' 出勤区分(午前)
        v_afternoonwork            = Trim(Rs_worktbl.Fields.Item("afternoonwork"               ).Value) ' 出勤区分(午後)
        v_morningholiday           = Trim(Rs_worktbl.Fields.Item("morningholiday"              ).Value) ' 休日区分(午前)
        v_afternoonholiday         = Trim(Rs_worktbl.Fields.Item("afternoonholiday"            ).Value) ' 休日区分(午後)
        v_workmin                  = Rs_worktbl.Fields.Item("workmin").Value                            ' 勤務時間
        v_work_begin               = editTime(Rs_worktbl.Fields.Item("work_begin"              ).Value) ' フレックス勤務開始時間
        v_work_end                 = editTime(Rs_worktbl.Fields.Item("work_end"                ).Value) ' フレックス勤務終了時間
        v_break_begin1             = editTime(Rs_worktbl.Fields.Item("break_begin1"            ).Value) ' フレックス勤務休憩開始時間
        v_break_end1               = editTime(Rs_worktbl.Fields.Item("break_end1"              ).Value) ' フレックス勤務休憩終了時間
        v_break_begin2             = editTime(Rs_worktbl.Fields.Item("break_begin2"            ).Value) ' フレックス勤務中抜開始時間
        v_break_end2               = editTime(Rs_worktbl.Fields.Item("break_end2"              ).Value) ' フレックス勤務中抜終了時間
        v_summons                  = Trim(Rs_worktbl.Fields.Item("summons"                     ).Value) ' 呼出
        v_overtime_begin           = editTime(Rs_worktbl.Fields.Item("overtime_begin"          ).Value) ' 時間外(休日出勤)申請分 開始
        v_overtime_end             = editTime(Rs_worktbl.Fields.Item("overtime_end"            ).Value) ' 時間外(休日出勤)申請分 終了
        v_rest_begin               = editTime(Rs_worktbl.Fields.Item("rest_begin"              ).Value) ' 時間外(休日出勤)申請分 休憩開始
        v_rest_end                 = editTime(Rs_worktbl.Fields.Item("rest_end"                ).Value) ' 時間外(休日出勤)申請分 休憩終了
        v_overtime                 = editTime(Rs_worktbl.Fields.Item("overtime"                ).Value) ' 時間外
        v_overtimelate             = editTime(Rs_worktbl.Fields.Item("overtimelate"            ).Value) ' 時間外深夜業
        If workshift = "9" Then
            If v_morningwork = "2" Or v_morningwork = "3" Or v_morningwork = "6" Or v_afternoonwork = "2" Or v_afternoonwork = "3" Or v_afternoonwork = "6" Then
                ' フレックス勤務者は休出時の勤務時間を集計する
                'temp = mm2FloatDay(Rs_worktbl.Fields.Item("workmin").Value)
                'If temp > 1 Then
                '    temp = 1
                'End If
                'sumFlex_holidayshift = sumFlex_holidayshift + temp
                sumFlex_holidayshift = sumFlex_holidayshift + Rs_worktbl.Fields.Item("workmin").Value
            End If
        Else
            v_holidayshift         = editTime(Rs_worktbl.Fields.Item("holidayshift"            ).Value) ' 休日出勤
        End If
        v_holidayshiftovertime     = editTime(Rs_worktbl.Fields.Item("holidayshiftovertime"    ).Value) ' 休出時間外
        v_holidayshiftlate         = editTime(Rs_worktbl.Fields.Item("holidayshiftlate"        ).Value) ' 休出深夜業
        v_holidayshiftovertimelate = editTime(Rs_worktbl.Fields.Item("holidayshiftovertimelate").Value) ' 休出時間外深夜業
        v_requesttime              = editTime(Rs_worktbl.Fields.Item("requesttime"             ).Value) ' 時間代休申請時間数
        v_requesttime_begin        = editTime(Rs_worktbl.Fields.Item("requesttime_begin"       ).Value) ' 時間代休申請開始時刻
        v_requesttime_end          = editTime(Rs_worktbl.Fields.Item("requesttime_end"         ).Value) ' 時間代休申請終了時刻
        v_vacationtime             = editTime(Rs_worktbl.Fields.Item("vacationtime"            ).Value) ' 時間有休申請時間数
        v_vacationtime_begin       = editTime(Rs_worktbl.Fields.Item("vacationtime_begin"      ).Value) ' 時間有休申請開始時刻
        v_vacationtime_end         = editTime(Rs_worktbl.Fields.Item("vacationtime_end"        ).Value) ' 時間有休申請終了時刻
        v_latetime                 = editTime(Rs_worktbl.Fields.Item("latetime"                ).Value) ' 深夜割増時間数
        v_latetime_begin           = editTime(Rs_worktbl.Fields.Item("latetime_begin"          ).Value) ' 深夜割増開始時刻
        v_latetime_end             = editTime(Rs_worktbl.Fields.Item("latetime_end"            ).Value) ' 深夜割増終了時刻
        v_is_approval              = Trim(Rs_worktbl.Fields.Item("is_approval"                 ).Value) ' 上長承認
        v_nightduty                = Trim(Rs_worktbl.Fields.Item("nightduty"                   ).Value) ' 宿直
        v_nightduty2               = Trim(Rs_worktbl.Fields.Item("nightduty2"                  ).Value) ' 前日宿直
        v_dayduty                  = Trim(Rs_worktbl.Fields.Item("dayduty"                     ).Value) ' 日直
        v_operator                 = Trim(Rs_worktbl.Fields.Item("operator"                    ).Value) ' 交替勤務
        v_operator2                = Trim(Rs_worktbl.Fields.Item("operator2"                   ).Value) ' 前日交替勤務
        v_memo                     = Trim(Rs_worktbl.Fields.Item("memo"                        ).Value) ' 備考欄
        v_memo2                    = Trim(Rs_worktbl.Fields.Item("memo2"                       ).Value) ' 備考欄
        v_is_error                 = Trim(Rs_worktbl.Fields.Item("is_error"                    ).Value) ' 労働時間適正化エラーフラグ
        v_overwork                 = Trim(Rs_worktbl.Fields.Item("overwork"                    ).Value) ' 週労働時間が40時間を超えた日には"warning"がセットされる
        v_weekovertime             = editTime(Rs_worktbl.Fields.Item("weekovertime"            ).Value) ' 週超過時間数

        ' 入力画面で上長未承認のとき、入力可設定
        If screen <> "0" Or (Rs_worktbl.Fields.Item("is_approval").Value = "1")Then
            text_disabled = "disabled"
        End If

        ' =========================================================
        ' 集計処理
        ' =========================================================
        ' -----------------------------------------------------
        ' 宿直回数
        ' -----------------------------------------------------
        If v_nightduty <> "0" Then
            sumNightdutyCount = sumNightdutyCount + 1
        End If
        ' -----------------------------------------------------
        ' 日直回数
        ' -----------------------------------------------------
        If v_dayduty = "1" Or v_dayduty = "2" Then
            sumDaydutyCount = sumDaydutyCount + 1
        End If
        ' -----------------------------------------------------
        ' 交替勤務
        ' -----------------------------------------------------
        ' 午前午後共に出勤区分が入力されていなければ、交替勤務が入力されていても集計しません。
        ' 午前午後どちらかに入力されていれば集計します。
        If (v_morningwork > "0" Or v_afternoonwork > "0") Then
            Select Case v_operator
                Case "1":   ' 甲番
                    sumOperatorKou  = sumOperatorKou  + 1
                Case "2":   ' 乙番
                    sumOperatorOtsu = sumOperatorOtsu + 1
                Case "3":   ' 日勤甲
                    sumOperatorKou  = sumOperatorKou  + 1
                Case "4":   ' 生産会議乙
                    sumOperatorOtsu = sumOperatorOtsu + 1
            End Select
        End If
        ' -----------------------------------------------------
        ' 可出勤日数
        ' -----------------------------------------------------
        If is_operator Then
            ' オペレータ時の可出勤日数集計
            sumWorkDays = sumWorkDays + operatorWorkDay(v_morningholiday, v_afternoonholiday, v_morningwork, v_afternoonwork, v_operator)
        Else
            ' オペレータでないときの可出勤日数集計
            If (v_morningholiday   = "1"  Or _
                v_morningholiday   = "2"  Or _
                v_morningholiday   = "A") Then
                sumWorkDays        = sumWorkDays - 0.5
            End if
            If (v_afternoonholiday = "1"  Or _
                v_afternoonholiday = "2"  Or _
                v_afternoonholiday = "A") Then
                sumWorkDays        = sumWorkDays - 0.5
            End if
            If (v_morningwork      = "1"  Or _
                v_morningwork      = "5") Then
                sumWorkDays        = sumWorkDays + 0.5
            End If
            If (v_afternoonwork    = "1"  Or _
                v_afternoonwork    = "5") Then
                sumWorkDays        = sumWorkDays + 0.5
            End If
        End If
        ' -----------------------------------------------------
        ' 実出勤日数
        ' -----------------------------------------------------
        If (v_morningwork = "1"  Or _
            v_morningwork = "4"  Or _
            v_morningwork = "5"  Or _
            v_morningwork = "9") Then
            ' 実出勤日数
            sumRealworkdays = sumRealworkdays + 0.5 + operatorAddDays(v_operator)
        End If
        If (v_afternoonwork = "1"  Or _
            v_afternoonwork = "4"  Or _
            v_afternoonwork = "5"  Or _
            v_afternoonwork = "9") Then
            ' 実出勤日数
            sumRealworkdays = sumRealworkdays + 0.5
        End If
        ' -----------------------------------------------------
        ' 振替残日数
        ' -----------------------------------------------------
        If (v_morningwork = "1"  Or _
            v_morningwork = "5") Then
            ' 振替出勤
            sumHolidaynumber = sumHolidaynumber + 0.5 + operatorAddDays(v_operator)
            If (v_is_approval    = "1") Then
                sumHolidaynumberHidden = sumHolidaynumberHidden + 0.5 + operatorAddDays(v_operator)
            End If
            If workshift = "9" Then
                ' フレックス勤務者が振替出勤したとき、当月労働時間を加算
                currentworkmin = currentworkmin + 210
            End If
        End If
        If (v_afternoonwork = "1"  Or _
            v_afternoonwork = "5") Then
            ' 振替出勤
            sumHolidaynumber     = sumHolidaynumber + 0.5
            If (v_is_approval    = "1") Then
                sumHolidaynumberHidden = sumHolidaynumberHidden + 0.5
            End If
            If workshift = "9" Then
                ' フレックス勤務者が振替出勤したとき、当月労働時間を加算
                currentworkmin = currentworkmin + 250
            End If
        End If
        ' -----------------------------------------------------
        ' 休日区分
        ' -----------------------------------------------------
        Select Case v_morningholiday
            Case "2"    '振替休日
                sumHolidaynumber     = sumHolidaynumber - 0.5 - operatorAddDays(v_operator)
                If (v_is_approval    = "1") Then
                    sumHolidaynumberHidden = sumHolidaynumberHidden - 0.5 - operatorAddDays(v_operator)
                End If
                If workshift = "9" Then
                    ' フレックス勤務者が振替休暇のとき、当月労働時間を減算
                    currentworkmin = currentworkmin - 210
                End If
            Case "B"    '育児休業
                If workshift = "9" Then
                    currentworkmin = currentworkmin - 210
                End If
                sumAbsenceDays       = sumAbsenceDays + 0.5 + operatorAddDays(v_operator)
            Case "3"    '有給休暇
                sumPaidvacations     = sumPaidvacations + 0.5 + operatorAddDays(v_operator)
                If (v_is_approval    = "1") Then
                    sumVacationnumberHidden = sumVacationnumberHidden + 0.5 + operatorAddDays(v_operator)
                End If
            Case "4"    '代替休暇
                sumWorkholidays      = sumWorkholidays + 0.5 + operatorAddDays(v_operator)
            Case "5"    '特別休暇
                sumSpecialvacations  = sumSpecialvacations + 0.5 + operatorAddDays(v_operator)
            Case "6"    '保存休暇
                sumPreservevacations = sumPreservevacations + 0.5 + operatorAddDays(v_operator)
            Case "7"    '半日欠勤
                sumAbsenceDays       = sumAbsenceDays + 0.5 + operatorAddDays(v_operator)
            Case "8"    'コアタイム振休
                sumHolidaynumber     = sumHolidaynumber - 0.25 - operatorAddDays(v_operator)
                If (v_is_approval    = "1") Then
                    sumHolidaynumberHidden = sumHolidaynumberHidden - 0.25 - operatorAddDays(v_operator)
                End If
                If workshift = "9" Then
                    ' フレックス勤務者がコアタイム振休のとき、当月労働時間を減算
                    currentworkmin = currentworkmin - 120
                End If
            Case "9"    'コアタイム有休
                sumPaidvacations     = sumPaidvacations + 0.25 + operatorAddDays(v_operator)
                If (v_is_approval    = "1") Then
                    sumVacationnumberHidden = sumVacationnumberHidden + 0.25 + operatorAddDays(v_operator)
                End If
        End Select
        Select Case v_afternoonholiday
            Case "2"    '振替休日
                sumHolidaynumber     = sumHolidaynumber - 0.5
                If (v_is_approval    = "1") Then
                    sumHolidaynumberHidden = sumHolidaynumberHidden - 0.5
                End If
                If workshift = "9" Then
                    ' フレックス勤務者が振替休暇のとき、当月労働時間を減算
                    currentworkmin = currentworkmin - 250
                End If
            Case "B"    '育児休業
                If workshift = "9" Then
                    currentworkmin = currentworkmin - 250
                End If
                sumAbsenceDays       = sumAbsenceDays + 0.5
            Case "3"    '有給休暇
                sumPaidvacations     = sumPaidvacations + 0.5
                If (v_is_approval    = "1") Then
                    sumVacationnumberHidden = sumVacationnumberHidden + 0.5
                End If
            Case "4"    '代替休暇
                sumWorkholidays      = sumWorkholidays + 0.5
            Case "5"    '特別休暇
                sumSpecialvacations  = sumSpecialvacations + 0.5
            Case "6"    '保存休暇
                sumPreservevacations = sumPreservevacations + 0.5
            Case "7"    '半日欠勤
                sumAbsenceDays       = sumAbsenceDays + 0.5
            Case "8"    'コアタイム振休
                sumHolidaynumber     = sumHolidaynumber - 0.25
                If (v_is_approval    = "1") Then
                    sumHolidaynumberHidden = sumHolidaynumberHidden - 0.25
                End If
                If workshift = "9" Then
                    ' フレックス勤務者がコアタイム振休のとき、当月労働時間を減算
                    currentworkmin = currentworkmin - 120
                End If
            Case "9"    'コアタイム有休
                sumPaidvacations     = sumPaidvacations + 0.25
                If (v_is_approval    = "1") Then
                    sumVacationnumberHidden = sumVacationnumberHidden + 0.25
                End If
        End Select
        ' -----------------------------------------------------
        ' 呼出(通常、深夜)
        ' -----------------------------------------------------
        Select Case v_summons
            Case "1"    '呼出通常
                sumSummons     = sumSummons     + 1     ' 呼出回数通常
            Case "2"    '呼出深夜
                sumSummonslate = sumSummonslate + 1     ' 呼出回数深夜
        End Select
        ' -----------------------------------------------------
        ' 時間外
        ' -----------------------------------------------------
        If (Len(v_overtime                ) > 0) Then
            sumOvertime                 = sumOvertime + time2Min(v_overtime)
        End If
        ' -----------------------------------------------------
        ' 休日出勤
        ' -----------------------------------------------------
        If (Len(v_holidayshift            ) > 0) Then
            sumHolidayshifttime         = sumHolidayshifttime + time2Min(v_holidayshift)
        End If
        ' -----------------------------------------------------
        ' 休出時間外
        ' -----------------------------------------------------
        If (Len(v_holidayshiftovertime    ) > 0) Then
            sumHolidayshiftovertime     = sumHolidayshiftovertime + time2Min(v_holidayshiftovertime)
        End If
        ' -----------------------------------------------------
        ' 休出深夜
        ' -----------------------------------------------------
        If (Len(v_holidayshiftlate        ) > 0) Then
            sumHolidayshiftlate         = sumHolidayshiftlate + time2Min(v_holidayshiftlate)
        End If
        ' -----------------------------------------------------
        ' 時間外深夜
        ' -----------------------------------------------------
        If (Len(v_overtimelate            ) > 0) Then
            sumOvertimelate             = sumOvertimelate + time2Min(v_overtimelate)
        End If
        ' -----------------------------------------------------
        ' 休出時間外深夜
        ' -----------------------------------------------------
        If (Len(v_holidayshiftovertimelate) > 0) Then
            sumHolidayshiftovertimelate = sumHolidayshiftovertimelate + time2Min(v_holidayshiftovertimelate)
        End If
        ' -----------------------------------------------------
        ' 時間代休
        ' -----------------------------------------------------
        If (Len(v_requesttime             ) > 0) Then
            sumWorkholidaytime          = sumWorkholidaytime + time2Min(v_requesttime)
        End If
        ' -----------------------------------------------------
        ' 時間有給
        ' -----------------------------------------------------
        If (Len(v_vacationtime            ) > 0) Then
            sumVacationtime             = sumVacationtime + time2Min(v_vacationtime)
        End If
        ' -----------------------------------------------------
        ' 深夜割増
        ' -----------------------------------------------------
        If (Len(v_latetime                ) > 0) Then
            sumLatepremium              = sumLatepremium + time2Min(v_latetime)
        End If

        ' -----------------------------------------------------------------
        ' 休出回数をカウント
        ' -----------------------------------------------------------------
        If v_morningwork   = "2" Or _
           v_morningwork   = "3" Or _
           v_morningwork   = "6" Or _
           v_afternoonwork = "2" Or _
           v_afternoonwork = "3" Or _
           v_afternoonwork = "6" Then
            sumHolidayWork = sumHolidayWork + 1
        End If

        ' 当日時間外14時間超チェック
        If (v_is_approval   = "1"  Or _
            v_morningwork   = "2"  Or _
            v_morningwork   = "3"  Or _
            v_morningwork   = "6"  Or _
            v_afternoonwork = "2"  Or _
            v_afternoonwork = "3"  Or _
            v_afternoonwork = "6") Then
            ' 上長チェック済み、または休日出勤の時チェックしない
        Else
            ' 時間外(休出)入力時の休憩時間チェック
            ' 時間外(休出)分算出
            overtimeMin = 0
            If (Len(Trim(v_overtime_begin)) > 0  And Len(Trim(v_overtime_end  )) > 0) Then
                If (legalTime(v_overtime_begin)  And legalTime(v_overtime_end  )) Then
                    ' 時間外算出
                    overtimeMin = minDif(editTime(v_overtime_begin), editTime(v_overtime_end  ))
                End If
            End If
            ' 時間外(休出)休憩時間分算出
            restMin     = 0
            If (Len(Trim(v_rest_begin)) > 0  And Len(Trim(v_rest_end  )) > 0) Then
                If (legalTime(v_rest_begin)  And legalTime(v_rest_end  )) Then
                    ' 休憩時間算出
                    restMin = minDif(editTime(v_rest_begin), editTime(v_rest_end  ))
                End If
            End If
            overtimeRealMin = overtimeMin    - restMin          ' 時間外実時間算出
            overtime_count  = overtime_count + overtimeRealMin  ' 時間代休チェック用時間外時間集計

            If overtimeRealMin > 840 Then
                warn_time14over = "1"
            End If
        End If

        ' ---------------------------------------------------------------
        ' コミュニケータ用
        ' 勤務時間
        ' ---------------------------------------------------------------
        If workshift <> "0" Then
            tempMin = 0
            If (v_morningwork <> "0" Or v_afternoonwork <> "0") Then
                If workshift = "1" Then
                    ' 全日勤務 08:50-17:30 460分
                    tempMin = 460
                ElseIf workshift = "2" Then
                    ' 午前勤務 08:50-13:00 250分
                    tempMin = 250
                ElseIf workshift = "3" Then
                    ' 午後勤務 13:00-17:30 270分
                    tempMin = 270
                End If
                ' 時間外加算
                If (Len(v_overtime) > 0) Then
                    tempMin = tempMin + time2Min(v_overtime)
                End If
                ' 時間有給減算
                'If (Len(v_vacationtime) > 0) Then
                '    tempMin = tempMin - time2Min(v_vacationtime)
                'End If
            End If
            tempDay = Left(Rs_worktbl.Fields.Item("workingdate").Value,4) & "/" & Mid(Rs_worktbl.Fields.Item("workingdate").Value,5,2) & "/" & Right(Rs_worktbl.Fields.Item("workingdate").Value,2)
            tempWeek = Weekday(tempDay)
            If (tempWeek = "7") Then
                ' 土曜日勤務時間集計
                sumSaturdayWorkMin = sumSaturdayWorkMin + tempMin
            Else
                ' 平日勤務時間集計
                sumWeekdaysWorkMin = sumWeekdaysWorkMin + tempMin
            End If
        End If

        Rs_worktbl.MoveNext()

        ' 上長チェック入力可否判定処理
        ' 上長チェック画面 かつ 空データでなく、給与担当者の処理済年月以後 かつ 当月以前 は上長チェックを可とする
        If screen = "2" And (v_morningholiday <> "0"  Or v_morningwork <> "0") And _
           ymb > proceseed_ymb And ymb <= inputLimitYmb Then
            text_approval_disabled = ""
        End If
    End If
End If

v_timetbl_id = ""
v_cometime   = ""
v_leavetime  = ""
' =================================================================
' タイムテーブル処理
' =================================================================
If Not Rs_timetbl.EOF Then
    If (Rs_timetbl.Fields.Item("workingdate").Value = (ymb & Right("0"&i, 2))) Then
        ' タイムテーブル有り
        v_timetbl_id = Rs_timetbl.Fields.Item("id").Value
        v_cometime   = editTime(Rs_timetbl.Fields.Item("cometime").Value)   ' 出社時刻
        v_leavetime  = editTime(Rs_timetbl.Fields.Item("leavetime").Value)  ' 退社時刻
        Rs_timetbl.MoveNext()
    End If
End If

' =================================================================
' 公休日テーブル処理
' =================================================================
If Not is_operator And Not Rs_holidaytbl.EOF And Not screen = 2 Then
      ' ピポット職員以外
      If (Rs_holidaytbl.Fields.Item("holidaydate").Value = (ymb & Right("0"&i, 2))) Then
          If v_worktbl_id = "" Then   ' 勤務表データが存在しないとき
              ' 公休日テーブル有り
              If Rs_holidaytbl.Fields.Item("holidaytype").Value = holidaytype Then
                  If workshift = "9" And _
                     Weekday(DateSerial(left(Rs_holidaytbl.Fields.Item("holidaydate").Value, 4), _
                                         Mid(Rs_holidaytbl.Fields.Item("holidaydate").Value, 5, 2), _
                                       Right(Rs_holidaytbl.Fields.Item("holidaydate").Value, 2))) = 1  Then
                      ' フレックス勤務で日曜日のとき、初期表示に法定休日を設定
                      v_morningholiday    = "A"   ' 休日区分(午前)
                      v_afternoonholiday  = "A"   ' 休日区分(午後)
                  Else
                      v_morningholiday    = "1"   ' 休日区分(午前)
                      v_afternoonholiday  = "1"   ' 休日区分(午後)
                  End If
              End If
          End If
          Rs_holidaytbl.MoveNext()
      End If
End If
If holidaytype = "2" And workshift = "9" Then
  ' ピポット職員の時
  pipotdate = ymb & Right("0"&i, 2)
  If v_worktbl_id = "" Then
    ' 勤務表データが存在しない
    If Weekday(DateSerial(left(pipotdate, 4), Mid(pipotdate, 5, 2), Right(pipotdate, 2))) = 4  Then
      ' pipot勤務で水曜日のとき、初期表示に法定休日を設定
      v_morningholiday    = "A"   ' 休日区分(午前)
      v_afternoonholiday  = "A"   ' 休日区分(午後)
    End If
  End If
End If
' =================================================================
' 前回入力情報セット
' 入力エラーの時、前回の入力情報を設定する。
' 上のテーブル情報をセットする中で、入力フォームの有効、無効などを
' 判定しているので、上記テーブル読込み処理を行った後処理を行う。
' 集計処理はエラーの時は行っていない。
' =================================================================
If (errorMsg <> "") Then
    For j = 1 To Request.Form("ymd").count Step 1
        If (ymb & Right("0"&i, 2) = Request.Form("ymd")(j)) Then
            If (Request.QueryString("p"))="" Then
                ' 勤務表入力画面
                v_morningwork        = Request.Form("morningwork"                )(j)  ' 出勤区分(午前)
                v_afternoonwork      = Request.Form("afternoonwork"              )(j)  ' 出勤区分(午後)
                v_morningholiday     = Request.Form("morningholiday"             )(j)  ' 休日区分(午前)
                v_afternoonholiday   = Request.Form("afternoonholiday"           )(j)  ' 休日区分(午後)
                If workshift = "9" Then
                    v_work_begin     = Request.Form("work_begin"                 )(j)  ' フレックス勤務開始時間
                    v_work_end       = Request.Form("work_end"                   )(j)  ' フレックス勤務終了時間
                    v_break_begin1   = Request.Form("break_begin1"               )(j)  ' フレックス勤務休憩開始時間
                    v_break_end1     = Request.Form("break_end1"                 )(j)  ' フレックス勤務休憩終了時間
                    v_break_begin2   = Request.Form("break_begin2"               )(j)  ' フレックス勤務中抜開始時間
                    v_break_end2     = Request.Form("break_end2"                 )(j)  ' フレックス勤務中抜終了時間
                Else
                    v_work_begin     = ""  ' フレックス勤務開始時間
                    v_work_end       = ""  ' フレックス勤務終了時間
                    v_break_begin1   = ""  ' フレックス勤務休憩開始時間
                    v_break_end1     = ""  ' フレックス勤務休憩終了時間
                    v_break_begin2   = ""  ' フレックス勤務中抜開始時間
                    v_break_end2     = ""  ' フレックス勤務中抜終了時間
                End If
                v_summons            = Request.Form("summons"                    )(j)  ' 呼出
                v_overtime_begin     = editTime(Request.Form("overtime_begin"    )(j)) ' 時間外(休日出勤)申請分 開始
                v_overtime_end       = editTime(Request.Form("overtime_end"      )(j)) ' 時間外(休日出勤)申請分 終了
                v_rest_begin         = editTime(Request.Form("rest_begin"        )(j)) ' 時間外(休日出勤)申請分 休憩開始
                v_rest_end           = editTime(Request.Form("rest_end"          )(j)) ' 時間外(休日出勤)申請分 休憩終了
                If workshift = "9" Then
                    v_requesttime_begin = ""    ' 時間代休申請開始時刻
                    v_requesttime_end   = ""    ' 時間代休申請終了時刻
                Else
                    v_requesttime_begin = editTime(Request.Form("requesttime_begin")(j)) ' 時間代休申請開始時刻
                    v_requesttime_end   = editTime(Request.Form("requesttime_end"  )(j)) ' 時間代休申請終了時刻
                End If
                v_vacationtime_begin = editTime(Request.Form("vacationtime_begin")(j)) ' 時間有休申請開始時刻
                v_vacationtime_end   = editTime(Request.Form("vacationtime_end"  )(j)) ' 時間有休申請終了時刻
                v_latetime_begin     = editTime(Request.Form("latetime_begin"    )(j)) ' 深夜割増開始時刻
                v_latetime_end       = editTime(Request.Form("latetime_end"      )(j)) ' 深夜割増終了時刻
                If (workshift = "9" Or is_operator) Then
                    v_weekovertime   = "" ' 週超過時間
                Else
                    v_weekovertime   = editTime(Request.Form("weekovertime"      )(j)) ' 週超過時間
                End If
                v_nightduty          = Request.Form("nightduty"                  )(j)  ' 宿直
                v_dayduty            = Request.Form("dayduty"                    )(j)  ' 日直
                If is_operator Then                                                    ' オペレータ
                    v_operator       = Request.Form("operator"                   )(j)
                Else
                    v_operator       = "0"
                End If
                v_memo               = Request.Form("memo"                       )(j)  ' 備考欄
                v_memo2              = Request.Form("memo2"                      )(j)  ' 備考欄
            Else
                ' 上長チェック画面
                v_cometime           = editTime(Request.Form("beginTime" & ymb & Right("0"&i, 2)))  ' タイムカード出社
                v_leavetime          = editTime(Request.Form("endTime"   & ymb & Right("0"&i, 2)))  ' タイムカード退社
                If (Request.Form("is_approval" & ymb & Right("0"&i, 2))="on") Then  ' 上長承認
                    v_is_approval    = "1"
                Else
                    v_is_approval    = "0"
                End If
            End If

            Exit For
        End If
    Next
End If
' 電源オン時間
pc_ontime = ""
If Not Rs_pctimetbl_on.EOF Then
    If (Rs_pctimetbl_on.Fields.Item("pcdate").Value = (ymb & Right("0"&i, 2))) Then
        ' 電源オン時間テーブル有り
        pc_ontime  = editTime(Trim(Rs_pctimetbl_on.Fields.Item("pctime").Value))
        Rs_pctimetbl_on.MoveNext()
    End If
End If

' 電源オフ時間
pc_offtime = ""
If Not Rs_pctimetbl_off.EOF Then
    If (Rs_pctimetbl_off.Fields.Item("pcdate").Value = (ymb & Right("0"&i, 2))) Then
        ' 電源オン時間テーブル有り
        pc_offtime  = editTime(Trim(Rs_pctimetbl_off.Fields.Item("pctime").Value))
        Rs_pctimetbl_off.MoveNext()
    End If
End If

If screen = "2" And (v_timetbl_id <> "" Or v_worktbl_id <> "") Then
    ' 上長チェック画面 And タイムテーブルあり、またはworktblデータありのとき、タイムカード項目を有効化
    text_timecard_disabled = ""
End If
%>
