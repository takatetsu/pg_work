<%
' -------------------------------------------------------------------------
' 労働時間適正化チェック
' 引数：v_operator      オペレータ
'       v_morningwork   出勤区分(午前)
'       v_afternoonwork 出勤区分(午後)
'       cometime        出勤時刻
'       leavetime       退勤時刻
'       pc_ontime       PC起動時刻
'       pc_offtime      PC終了時刻
'       dayduty         日直
'       nightduty       宿直
'       nightduty2      前日宿直
'       overtime_begin  時間外開始時刻
'       overtime_end    時間外終了時刻
'       memo2           メモ２
'       opentime        始業時刻
'       closetime       終業時刻
'       is_unionexecutive 組合執行役員
'       v_operator      前日交替勤務
'       workshift       勤務体系
'       wk_work_begin   出社申請時刻
'       wk_work_end     退社申請時刻
' 戻値：0     エラーなし
'       0以外 エラー
' -------------------------------------------------------------------------
Function workTimeCheck(v_operator, v_morningwork, v_afternoonwork, cometime, _
                       leavetime, pc_ontime, pc_offtime, dayduty, nightduty, _
                       nightduty2, overtime_begin, overtime_end, memo2, _
                       opentime, closetime, is_unionexecutive, v_operator2, _
                       workshift, wk_work_begin, wk_work_end)
    ' Function内での数値変更に備え、別変数で処理を行う
    c_operator          = v_operator
    c_morningwork       = v_morningwork
    c_afternoonwork     = v_afternoonwork
    c_cometime          = cometime
    c_leavetime         = leavetime
    c_pc_ontime         = pc_ontime
    c_pc_offtime        = pc_offtime
    c_dayduty           = dayduty
    c_nightduty         = nightduty
    c_nightduty2        = nightduty2
    c_overtime_begin    = editTime(overtime_begin)
    c_overtime_end      = editTime(overtime_end)
    c_memo2             = memo2
    c_is_unionexecutive = is_unionexecutive
    c_workshift         = workshift
    c_wk_work_begin     = editTime(wk_work_begin)
    c_wk_work_end       = editTime(wk_work_end)
    ' 【出勤状況に応じたチェック用の基準時刻を設定し、必要なチェック項目のフラグを1にする】
    workTimeCheck  = "0"        ' 判定戻値
    ref_starttime  = opentime   ' 基準就業開始時刻
    ref_endtime    = closetime  ' 基準就業終了時刻
    res_timeDev    = "0"        ' 出退勤時刻、PC起動停止時刻乖離チェック結果
    res_comeCheck  = "0"        ' 開始時刻チェック結果
    res_outCheck   = "0"        ' 終了時刻チェック結果
    flg_checkStart = "0"        ' 開始時間チェックするかのフラグ 0:チェックしない 1:チェックする
    flg_checkEnd   = "0"        ' 終了時間チェックするかのフラグ 0:チェックしない 1:チェックする
    dif_startTime  = 30         ' 出社時刻チェック猶予時間(分)
    dif_endTime    = 30         ' 退社時刻チェック猶予時間(分)

    ' メモ2判定
    If c_memo2 = "1" Or c_memo2 = "2" Then
        ' 通勤渋滞回避、電車時間都合のとき出勤時刻チェックはチェック対象外とする
        c_cometime = ""
    End If
    If c_memo2 = "2" Or c_memo2 = "4" Then
        ' 電車時間都合、懇親会時間待のとき退社時刻はチェック対象外とする
        c_leavetime  = ""
    End If
    If c_memo2 = "3" Then
        ' 組合活動のとき退社時刻はチェック対象外とする
        c_leavetime  = ""
        If c_is_unionexecutive = "1" Then
            ' 組合執行部のときPC終了時刻もチェック対象外とする
            c_pc_offtime = ""
        End If
    End If
    If c_memo2 = "5" Then
        ' PC消し忘れのときPC終了時刻はチェック対象外とする
        c_pc_offtime = ""
    End If

    dec_starttime  = setTime(c_cometime,  c_pc_ontime,  "0") ' 判定用開始時刻
    dec_endtime    = setTime(c_leavetime, c_pc_offtime, "1") ' 判定用終了時刻

    If c_operator = "0" Then
        ' 交代勤務以外(一般、フレックス勤務)
        If c_morningwork = "1" Or _
           c_morningwork = "4" Or _
           c_morningwork = "5" Or _
           c_morningwork = "9" Then
            ' 午前出勤(振替出勤、出張(出勤)、出張(振替出勤)、出勤)
            If c_afternoonwork = "1" Or _
               c_afternoonwork = "4" Or _
               c_afternoonwork = "5" Or _
               c_afternoonwork = "9" Then
                ' 午前出勤・午後出勤(振替出勤、出張(出勤)、出張(振替出勤)、出勤)
                ref_starttime  = opentime
                ref_endtime    = closetime
                flg_checkStart = "1"
                flg_checkEnd   = "1"
            Else
                If (v_afterwork = "2" Or v_afterwork = "3" Or v_afterwork = "6") Then
                    ' 午前出勤・午後休出、休出(半日未満)、出張(休出)
                    ref_endtime    = c_overtime_end
                Else
                    ' 午前出勤・午後出勤せず
                    ref_endtime    = "12:00"
                End If
                ref_starttime  = opentime
                flg_checkStart = "1"
                flg_checkEnd   = "1"
            End If
        Else
             If (c_morningwork = "2" Or c_morningwork = "3" or c_morningwork = "6") Then
                ' 午前休出、休出(半日未満)、出張(休出)
                If c_afternoonwork = "1" Or _
                   c_afternoonwork = "4" Or _
                   c_afternoonwork = "5" Or _
                   c_afternoonwork = "9" Then
                    ' 午前休出、休出(半日未満)、出張(休出)・午後出勤(振替出勤、出張(出勤)、出張(振替出勤)、出勤)
                    ref_starttime  = c_overtime_begin
                    ref_endtime    = closetime
                    flg_checkStart = "1"
                    flg_checkEnd   = "1"
                Else
                    ' 午前休出、休出(半日未満)、出張(休出)・午後休出、休出(半日未満)、出張(休出) もしくは 午後出勤せず のとき
                    ref_starttime  = c_overtime_begin
                    ref_endtime    = c_overtime_end
                    flg_checkStart = "1"
                    flg_checkEnd   = "1"
                End If
             Else
                ' 午前出勤せず
                If c_afternoonwork = "1" Or _
                   c_afternoonwork = "4" Or _
                   c_afternoonwork = "5" Or _
                   c_afternoonwork = "9" Then
                    ' 午前出勤せず・午後出勤(振替出勤、出張(出勤)、出張(振替出勤)、出勤)
                    ref_starttime  = "13:00"
                    ref_endtime    = closetime
                    flg_checkStart = "1"
                    flg_checkEnd   = "1"
                Else
                    If (v_afterwork = "2" Or v_afterwork = "3" Or v_afterwork = "6") Then
                        ' 午前出勤せず・午後休出、休出(半日未満)、出張(休出)
                        ref_starttime  = c_overtime_begin
                        ref_endtime    = c_overtime_end
                        flg_checkStart = "1"
                        flg_checkEnd   = "1"
                    Else
                        ' 午前出勤せず・午後出勤せず
                        If c_cometime <> "" Or c_pc_ontime <> "" Then
                            If c_morningwork = "0" And c_afternoonwork = "0" And c_dayduty = "0" And c_nightduty = "0" Then
                                ' 午前午後とも出勤区分の入力が無く、日直宿直でもないときエラー
                                res_comeCheck = "1"
                                res_outCheck  = "1"
                            Else
                                If c_dayduty <> "0" And c_nightduty =  "0" Then
                                    ' 日直で宿直でないとき
                                    flg_checkStart = "1"
                                    flg_checkEnd   = "1"
                                End If
                                If c_dayduty =  "0" And c_nightduty <> "0" Then
                                    ' 日直でなく宿直のとき
                                    flg_checkStart = "1"
                                    ref_starttime  = "17:10"
                                End If
                                If c_dayduty <> "0" And c_nightduty <> "0" Then
                                    ' 日直宿直のとき
                                    flg_checkStart = "1"
                                End If
                            End If
                        End If
                        If c_leavetime <> "" Or c_pc_offtime <> "" Then
                            If c_morningwork = "0" And c_afternoonwork = "0" And c_dayduty = "0" And c_nightduty = "0" And c_nightduty2 = "0" Then
                                ' 午前午後とも出勤区分の入力が無く、日直宿直でなく、前日宿直でもないときエラー
                                res_comeCheck = "1"
                                res_outCheck  = "1"
                            Else
                                If c_dayduty <> "0" And c_nightduty = "0" Then
                                    ' 日直で宿直でないとき
                                    flg_checkEnd   = "1"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        ' 交代勤務に該当
        If c_operator = "1" or _
           c_operator = "3" or _
           c_operator = "5" Then
           ' 甲番勤務(甲番、日勤甲、見習(甲))
            ref_starttime  = opentime
            ref_endtime    = "20:30"
            flg_checkStart = "1"
            flg_checkEnd   = "1"
        End If
        If c_operator = "2" or _
           c_operator = "6" Then
           ' 乙番勤務(乙番、見習(乙))
            ref_starttime  = "20:30"
            flg_checkStart = "1"
        End If
    End If

    If c_workshift = "9" Then
      ref_starttime = wk_work_begin
      ref_endtime   = wk_work_end
    End If

    ref_starttime  = setTime(ref_starttime,  c_overtime_begin,  "0") ' 判定用開始時刻
    ref_endtime    = setTime(ref_endtime,    c_overtime_end,    "1") ' 判定用終了時刻
    
    ' 宿直判定
    If (c_nightduty2 = "1" Or c_nightduty2 = "2") Then ' 前日宿直
        ' 前日宿直のとき開始時刻チェックは無視する
        res_comeCheck  = "0"
        flg_checkStart = "0"
        If c_morningwork = "0" And c_afternoonwork = "0" And c_dayduty = "0" And c_nightduty = "0" Then
            ' 前日宿直で当日出勤なし、日直宿直なしのとき、終了時刻チェックを基準時08:30、猶予30分で行う
            flg_checkEnd   = "1"
            ref_endtime    = "08:30"
        Else
            ' 前日宿直で当日出勤ありのとき、PC起動時刻をクリア
            c_pc_ontime = ""
        End If
    End If
    If (c_nightduty  = "1" Or c_nightduty  = "2") Then ' 当日宿直
        flg_checkEnd   = "0"
    End If
    
    ' 交替勤務判定
    If (v_operator2 = "2" Or v_operator2 = "4" Or v_operator2 = "6") Then ' 前日乙番
        ' 前日乙番のとき開始時刻チェックは無視する
        res_comeCheck  = "0"
        flg_checkStart = "0"
        If c_morningwork = "0" And c_afternoonwork = "0" Then
            ' 前日乙番で当日出勤なしのとき、終了時刻チェックを基準時08:30、猶予30分で行う
            flg_checkEnd   = "1"
            ref_endtime    = "08:30"
        Else
            ' 前日乙番で当日出勤ありのとき、PC起動時刻をクリア
            c_pc_ontime = ""
        End If
    End If
    
    ' 【時刻チェック】
    If flg_checkStart = "1" And dec_starttime <> "" Then
        ' 開始時刻チェック
        res_comeCheck = checkTimeInterval(dec_starttime, ref_starttime, dif_startTime)
'response.write("<br />START/kijyun:" & ref_starttime & " hantei:" & dec_starttime & " res_comeCheck=" & res_comeCheck & " res_outCheck=" & res_outCheck & " c_wk_work_begin=" & c_wk_work_begin)
    End If
    If flg_checkEnd   = "1" And dec_endtime   <> "" Then
        ' 終了時刻チェック
        res_outCheck  = checkTimeInterval(ref_endtime,   dec_endtime,   dif_endTime)
'response.write("<br />E N D/kijyun:" & ref_endtime & " hantei:" & dec_endtime & " yuuyo:" & dif_endTime & " res_outCheck=" & res_outCheck & " res_comeCheck=" & res_comeCheck & " c_wk_work_end=" & c_wk_work_end & "<br />")
    End If

    ' 結果コード設定
    If res_comeCheck <> "0" Or res_outCheck <> "0" Then
        workTimeCheck = "1"
    End If
End Function

' -----------------------------------------------------------------------------
' 時刻間隔チェック
' 引数：t1 (チェック開始時刻)
'       t2 (チェック終了時刻),
'       m  (許容分数)
' 戻値：0 チェックOK
'       1 t1とt2がm分よりも間隔が開いているためエラー
' -----------------------------------------------------------------------------
Function checkTimeInterval(t1, t2, m)
    checkTimeInterval = "1"
    If t1 <> "" And t2 <> "" Then
        If t1 <= t2 Then
            If minDifIV(editTime(t1), editTime(t2)) <= m Then
                checkTimeInterval = "0"
            End If
        Else
            checkTimeInterval = "0"
        End If
    End If
End Function

' -----------------------------------------------------------------------------
' チェック用時刻設定
' 引数：t1  (チェック対象時刻1)
'       t2  (チェック対象時刻2),
'       j   (判定フラグ 0:早い方を取得、0以外:遅い方を取得),
' 戻値：t   t1とt2のうちjで設定された判定の結果を返す
'           どちらかに値が無いときは値がある方を返す
' -----------------------------------------------------------------------------
Function setTime(t1, t2, j)
    setTime = ""
    If t1 <> "" Then
        If t2 <> "" Then
            If t1 <= t2 Then
                setSmallTime = t1
                setBigTime   = t2
            Else
                setSmallTime = t2
                setBigTime   = t1
            End If
        Else
            setSmallTime = t1
            setBigTime   = t1
        End If
    Else
        If t2 <> "" then
            setSmallTime = t2
            setBigTime   = t2
        End If
    End If
    If j = "0" Then
        setTime = setSmallTime
    Else
        setTime = setBigTime
    End If
End Function

' -----------------------------------------------------------------------------
' 前日交替勤務を設定
' 引数：personalcode 個人コード
'       ymb yyyymmddのフォーマットで日付を設定
' 戻値：t   引数で渡された日付の前日の交替勤務を返す
'           どちらかに値が無いときは値がある方を返す
' -----------------------------------------------------------------------------
Function setPreOp(personalcode, ymb)
    setPreOp = ""
    ' 前日日付算出
    predate = DateAdd("d", -1, CDate(Left(ymb,4) & "/" & Mid(ymb,5,2) & "/" & Right(ymb,2)))
    predate = Left(predate,4) & Mid(predate,6,2) & Right(predate,2)
    ' 前日交替勤務を読み込み設定する
    Dim Rs_previous_worktbl
    Dim Rs_previous_worktbl_cmd
    Dim Rs_previous_worktbl_numRows
    Set Rs_previous_worktbl_cmd = Server.CreateObject ("ADODB.Command")
    Rs_previous_worktbl_cmd.ActiveConnection = MM_workdbms_STRING
    Rs_previous_worktbl_cmd.CommandText = "SELECT operator FROM dbo.worktbl WHERE personalcode = ? AND workingdate = ?"
    Rs_previous_worktbl_cmd.Prepared = true
    Rs_previous_worktbl_cmd.Parameters.Append Rs_previous_worktbl_cmd.CreateParameter("param1", 200, 1, 5, personalcode)
    Rs_previous_worktbl_cmd.Parameters.Append Rs_previous_worktbl_cmd.CreateParameter("param2", 200, 1, 8, predate)
    Set Rs_previous_worktbl = Rs_previous_worktbl_cmd.Execute
    Rs_previous_worktbl_numRows = 0
    If Rs_previous_worktbl.EOF And Rs_previous_worktbl.BOF Then
        setPreOp = ""
    Else
        setPreOp = Rs_previous_worktbl.Fields.Item("operator").Value
    End If
    Rs_previous_worktbl.Close()
    Set Rs_previous_worktbl = Nothing
End Function

%>
