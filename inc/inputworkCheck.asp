<%
' -------------------------------------------------------------------------
' 入力チェック
' -------------------------------------------------------------------------
i = 1
For x = 1 To Request.Form("everyday").count Step 1
    ' 上長チェック済みでも一部の集計処理を行わなければならないので、
    ' 全日付と上長チェックがされていない日付を比較しながら入力チェックと集計処理を行う
    If i <= Request.Form("ymd").count Then
        If Request.Form("everyday")(x) = Request.Form("ymd")(i) Then
            ' 上長チェックされていないとき
            j = Right(Request.Form("ymd")(i), 2)
            ' 休出回数をカウント
            If Request.Form("morningwork"  )(i) = "2" Or _
               Request.Form("morningwork"  )(i) = "3" Or _
               Request.Form("morningwork"  )(i) = "6" Or _
               Request.Form("afternoonwork")(i) = "2" Or _
               Request.Form("afternoonwork")(i) = "3" Or _
               Request.Form("afternoonwork")(i) = "6" Then
                holidaywork_count = holidaywork_count + 1
            End If
            ' -----------------------------------------------------------------
            ' 時刻チェック
            ' -----------------------------------------------------------------
            If workshift = "9" Then
                ' フレックス勤務 自
                If Not (legalTime(Request.Form("work_begin")(i))) Then
                    err_work_begin          = 1
                    style_work_begin(j)     = "errorcolor"
                    dayErrorFlag(j)         = "error"
                End If
                ' フレックス勤務 至
                If Not (legalTime(Request.Form("work_end")(i))) Then
                    err_work_end            = 1
                    style_work_end(j)       = "errorcolor"
                    dayErrorFlag(j)         = "error"
                End If
                ' フレックス休憩 自
                If Not (legalTime(Request.Form("break_begin1")(i))) Then
                    err_break_begin1        = 1
                    style_break_begin1(j)   = "errorcolor"
                    dayErrorFlag(j)         = "error"
                End If
                ' フレックス休憩 至
                If Not (legalTime(Request.Form("break_end1")(i))) Then
                    err_break_end1          = 1
                    style_break_end1(j)     = "errorcolor"
                    dayErrorFlag(j)         = "error"
                End If
                ' フレックス中抜 自
                If Not (legalTime(Request.Form("break_begin2")(i))) Then
                    err_break_begin2        = 1
                    style_break_begin2(j)   = "errorcolor"
                    dayErrorFlag(j)         = "error"
                End If
                ' フレックス中抜 至
                If Not (legalTime(Request.Form("break_end2")(i))) Then
                    err_break_end2          = 1
                    style_break_end2(j)     = "errorcolor"
                    dayErrorFlag(j)         = "error"
                End If
            End If
            ' 時間外(休出)申請分 自
            If Not (legalTime(Request.Form("overtime_begin")(i))) Then
                err_overtime_begin          = 1
                style_overtime_begin(j)     = "errorcolor"
                dayErrorFlag(j)             = "error"
            End If
            ' 時間外(休出)申請分 至
            If Not (legalTime(Request.Form("overtime_end")(i))) Then
                err_overtime_end            = 1
                style_overtime_end(j)       = "errorcolor"
                dayErrorFlag(j)             = "error"
            End If
            ' 時間外(休出)申請分 休憩自
            If Not (legalTime(Request.Form("rest_begin")(i))) Then
                err_rest_begin              = 1
                style_rest_begin(j)         = "errorcolor"
                dayErrorFlag(j)             = "error"
            End If
            ' 時間外(休出)申請分 休憩至
            If Not (legalTime(Request.Form("rest_end")(i))) Then
                err_rest_end                = 1
                style_rest_end(j)           = "errorcolor"
                dayErrorFlag(j)             = "error"
            End If
            If Not workshift = "9" Then ' フレックス勤務者以外のとき
                ' 時間代休申請分 自
                If Not (legalTime(Request.Form("requesttime_begin")(i))) Then
                    err_requesttime_begin       = 1
                    style_requesttime_begin(j)  = "errorcolor"
                    dayErrorFlag(j)             = "error"
                End If
                ' 時間代休申請分 至
                If Not (legalTime(Request.Form("requesttime_end")(i))) Then
                    err_requesttime_end         = 1
                    style_requesttime_end(j)    = "errorcolor"
                    dayErrorFlag(j)             = "error"
                End If
            End If
            ' 時間有給 自
            If Not (legalTime(Request.Form("vacationtime_begin")(i))) Then
                err_vacationtime_begin      = 1
                style_vacationtime_begin(j) = "errorcolor"
                dayErrorFlag(j)             = "error"
            End If
            ' 時間有給 至
            If Not (legalTime(Request.Form("vacationtime_end")(i))) Then
                err_vacationtime_end        = 1
                style_vacationtime_end(j)   = "errorcolor"
                dayErrorFlag(j)             = "error"
            End If
            ' 深夜割増 自
            If Not (legalTime(Request.Form("latetime_begin")(i))) Then
                err_latetime_begin          = 1
                style_latetime_begin(j)     = "errorcolor"
                dayErrorFlag(j)             = "error"
            End If
            ' 深夜割増 至
            If Not (legalTime(Request.Form("latetime_end")(i))) Then
                err_latetime_end            = 1
                style_latetime_end(j)       = "errorcolor"
                dayErrorFlag(j)             = "error"
            End If
            ' 週超過時間
            If Not workshift = "9" And Not is_operator Then
                If Not (legalTime(Request.Form("weekovertime")(i))) Then
                    err_weekovertime            = 1
                    style_weekovertime(j)       = "errorcolor"
                    dayErrorFlag(j)             = "error"
                End If
            End If
            ' -----------------------------------------------------------------
            ' 関連チェック
            ' -----------------------------------------------------------------
            ' 交替勤務で日勤甲と生産会議乙のとき休日区分、出勤区分の入力チェックを
            ' 一部実行させないためのフラグを設定
            operatorNoCheck = ""
            If is_operator Then
                If ((Request.Form("operator")(i) = "3")  Or _
                    (Request.Form("operator")(i) = "4")) Then
                    operatorNoCheck = "y"
                End If
            End If
            ' 交替勤務時の勤務日数算出に使用します。
            If is_operator Then
                v_operator = Request.Form("operator")(i)
            Else
            End If
            ' 休日区分公休日チェック
            ' 休日区分午前=1公休日のとき、休日区分午後<>1公休日でエラー
            ' 休日区分午前<>1公休日のとき、休日区分午後=1公休日でエラー
            ' 休日区分午前=A法定休日のとき、休日区分午後<>A法定休日でエラー
            ' 休日区分午前<>A法定休日のとき、休日区分午後=A法定休日でエラー
            ' ただし、オペレータ, お客さまセンターコミュニケータの場合はエラーとしない
            If Not is_operator And (workshift = "0" Or workshift = "9") Then
                If ((Request.Form("morningholiday"  )(i) =  "1"   And _
                     Request.Form("afternoonholiday")(i) <> "1")  Or  _
                    (Request.Form("morningholiday"  )(i) <> "1"   And _
                     Request.Form("afternoonholiday")(i) =  "1")  Or  _
                    (Request.Form("morningholiday"  )(i) =  "A"   And _
                     Request.Form("afternoonholiday")(i) <> "A")  Or _
                    (Request.Form("morningholiday"  )(i) <> "A"   And _
                     Request.Form("afternoonholiday")(i) =  "A")) Then
                    ' 出勤区分午前=1振替出勤、出勤区分午後=1振替出勤で休日区分午後=5特別休暇または7欠勤のときエラーとしない
                    ' 出勤区分午後=1振替出勤、出勤区分午前=1振替出勤で休日区分午前=5特別休暇または7欠勤のときエラーとしない
                    If ((Request.Form("morningwork"     )(i) = "1"   And _
                         Request.Form("afternoonwork"   )(i) = "1"   And _
                        (Request.Form("afternoonholiday")(i) = "5"   Or  _
                         Request.Form("afternoonholiday")(i) = "7")) Or  _
                        (Request.Form("morningwork"     )(i) = "1"   And _
                        (Request.Form("morningholiday"  )(i) = "5"   Or _
                         Request.Form("morningholiday"  )(i) = "7")  And _
                         Request.Form("afternoonwork"   )(i) = "1")) Then
                         ' エラーとしない
                    Else
                        err_relation_01                 = 1
                        style_morningholiday  (j)       = "errorcolor"
                        style_afternoonholiday(j)       = "errorcolor"
                        dayErrorFlag(j)                 = "error"
                    End If
                End If
            End If
            ' 休日区分・出勤区分チェック
            ' 休日区分午前=1公休日のとき、出勤区分午前=9出勤でエラー
            ' 休日区分午後=1公休日のとき、出勤区分午後=9出勤でエラー
            ' ただしフレックス勤務者は除く
            If (Request.Form("morningholiday")(i) =  "1"  And _
               (Request.Form("morningwork"   )(i) =  "4"  Or  _
                Request.Form("morningwork"   )(i) =  "9") And _
                workshift <> "9") Then
                If is_operator Then
                    If (Request.Form("operator")(i) = "0") Then
                        err_relation_02                 = 1
                        style_morningholiday(j)         = "errorcolor"
                        style_morningwork   (j)         = "errorcolor"
                        dayErrorFlag(j)                 = "error"
                    End If
                Else
                    err_relation_02                 = 1
                    style_morningholiday(j)         = "errorcolor"
                    style_morningwork   (j)         = "errorcolor"
                    dayErrorFlag(j)                 = "error"
                End If
            End If
            If (Request.Form("afternoonholiday")(i) =  "1"  And _
               (Request.Form("afternoonwork"   )(i) =  "4"  Or  _
                Request.Form("afternoonwork"   )(i) =  "9") And _
                workshift <> "9") Then
                If is_operator Then
                    If (Request.Form("operator")(i) = "0") Then
                        err_relation_02                 = 1
                        style_afternoonholiday(j)       = "errorcolor"
                        style_afternoonwork   (j)       = "errorcolor"
                        dayErrorFlag(j)                 = "error"
                    End If
                Else
                    err_relation_02                 = 1
                    style_afternoonholiday(j)       = "errorcolor"
                    style_afternoonwork   (j)       = "errorcolor"
                    dayErrorFlag(j)                 = "error"
                End If
            End If
            ' フレックス勤務者の公休日における休出チェック
            If workshift = "9" Then
                If (Request.Form("morningholiday")(i) = "1" And _
                   (Request.Form("morningwork")(i) = "2" Or _
                    Request.Form("morningwork")(i) = "3" Or _
                    Request.Form("morningwork")(i) = "6")) Then
                    err_relation_54 = 1
                    style_morningholiday(j) = "errorcolor"
                    style_morningwork(j) = "errorcolor"
                    dayErrorFlag(j) = "error"
                End If
                If (Request.Form("afternoonholiday")(i) = "1" And _
                   (Request.Form("afternoonwork")(i) = "2" Or _
                    Request.Form("afternoonwork")(i) = "3" Or _
                    Request.Form("afternoonwork")(i) = "6")) Then
                    err_relation_54 = 1
                    style_afternoonholiday(j) = "errorcolor"
                    style_afternoonwork(j) = "errorcolor"
                    dayErrorFlag(j) = "error"
                End If
            End If
            ' 休日区分午前=A法定休日のとき、出勤区分午前=9出勤でエラー
            ' 休日区分午後=A法定休日のとき、出勤区分午後=9出勤でエラー
            If (Request.Form("morningholiday")(i) = "A"   And _
               (Request.Form("morningwork"   )(i) = "4"   Or  _
                Request.Form("morningwork"   )(i) = "9")) Then
                err_relation_50           = 1
                style_morningholiday(j)   = "errorcolor"
                style_morningwork   (j)   = "errorcolor"
                dayErrorFlag(j)           = "error"
            End If
            If (Request.Form("afternoonholiday")(i) = "A"   And _
               (Request.Form("afternoonwork"   )(i) = "4"   Or  _
                Request.Form("afternoonwork"   )(i) = "9")) Then
                err_relation_50           = 1
                style_afternoonholiday(j) = "errorcolor"
                style_afternoonwork   (j) = "errorcolor"
                dayErrorFlag(j)           = "error"
            End If
            ' 休日区分午前=3有給休暇,5特別休暇,6保存休暇,7半日欠勤のとき、
            ' 出勤区分午前=1振替出勤,2休日出勤,3休日出勤(半日未満),4出張のときエラー
            If ((Request.Form("morningholiday")(i) =  "3"   Or  _
                 Request.Form("morningholiday")(i) =  "5"   Or  _
                 Request.Form("morningholiday")(i) =  "6"   Or  _
                 Request.Form("morningholiday")(i) =  "7")  And _
                (Request.Form("morningwork"   )(i) =  "1"   Or  _
                 Request.Form("morningwork"   )(i) =  "2"   Or  _
                 Request.Form("morningwork"   )(i) =  "3"   Or  _
                 Request.Form("morningwork"   )(i) =  "4"   Or  _
                 Request.Form("morningwork"   )(i) =  "5"   Or  _
                 Request.Form("morningwork"   )(i) =  "6")) Then
                ' 出勤区分午前=1振替出勤、出勤区分午後=1振替出勤で休日区分午後=5特別休暇または7欠勤のときエラーとしない
                ' 出勤区分午後=1振替出勤、出勤区分午前=1振替出勤で休日区分午前=5特別休暇または7欠勤のときエラーとしない
                If ((Request.Form("morningwork"     )(i) = "1"   And _
                     Request.Form("afternoonwork"   )(i) = "1"   And _
                    (Request.Form("afternoonholiday")(i) = "5"   Or  _
                     Request.Form("afternoonholiday")(i) = "7")) Or  _
                    (Request.Form("morningwork"     )(i) = "1"   And _
                    (Request.Form("morningholiday"  )(i) = "5"   Or _
                     Request.Form("morningholiday"  )(i) = "7")  And _
                     Request.Form("afternoonwork"   )(i) = "1")) Then
                     ' エラーとしない
                Else
                    err_relation_03                 = 1
                    style_morningholiday(j)         = "errorcolor"
                    style_morningwork   (j)         = "errorcolor"
                    dayErrorFlag(j)                 = "error"
                End If
            End If
            ' 休日区分午後=3有給休暇,5特別休暇,6保存休暇,7半日欠勤のとき、
            ' 出勤区分午後=1振替出勤,2休日出勤,3休日出勤(半日未満),4出張のときエラー
            If ((Request.Form("afternoonholiday")(i) =  "3"   Or  _
                 Request.Form("afternoonholiday")(i) =  "5"   Or  _
                 Request.Form("afternoonholiday")(i) =  "6"   Or  _
                 Request.Form("afternoonholiday")(i) =  "7")  And _
                (Request.Form("afternoonwork"   )(i) =  "1"   Or  _
                 Request.Form("afternoonwork"   )(i) =  "2"   Or  _
                 Request.Form("afternoonwork"   )(i) =  "3"   Or  _
                 Request.Form("afternoonwork"   )(i) =  "4"   Or  _
                 Request.Form("afternoonwork"   )(i) =  "5"   Or  _
                 Request.Form("afternoonwork"   )(i) =  "6")) Then
                ' 出勤区分午前=1振替出勤、出勤区分午後=1振替出勤で休日区分午後=5特別休暇または7欠勤のときエラーとしない
                ' 出勤区分午後=1振替出勤、出勤区分午前=1振替出勤で休日区分午前=5特別休暇または7欠勤のときエラーとしない
                If ((Request.Form("morningwork"     )(i) = "1"   And _
                     Request.Form("afternoonwork"   )(i) = "1"   And _
                    (Request.Form("afternoonholiday")(i) = "5"   Or  _
                     Request.Form("afternoonholiday")(i) = "7")) Or  _
                    (Request.Form("morningwork"     )(i) = "1"   And _
                    (Request.Form("morningholiday"  )(i) = "5"   Or _
                     Request.Form("morningholiday"  )(i) = "7")  And _
                     Request.Form("afternoonwork"   )(i) = "1")) Then
                     ' エラーとしない
                Else
                    err_relation_03                 = 1
                    style_afternoonholiday(j)       = "errorcolor"
                    style_afternoonwork   (j)       = "errorcolor"
                    dayErrorFlag(j)                 = "error"
                End If
            End If
            If Not workshift = "9" Then ' フレックス勤務者以外のとき
                ' 休日区分=4代替休暇のとき、時間単位代休の入力チェックを行う。
                If (Request.Form("morningholiday")(i) = "4") Then
                    If (Request.Form("afternoonholiday")(i) = "4") Then
                        ' 午前・午後ともに代替休暇のとき、時間単位代休は8:30-17:10でなければエラー
                        If (Len(Trim(Request.Form("requesttime_begin")(i))) > 0) Then
                            If (legalTime(Request.Form("requesttime_begin")(i))) Then
                                If (editTime(Request.Form("requesttime_begin")(i)) <> "08:30") Then
                                    err_relation_36             = 1
                                    style_morningholiday    (j) = "errorcolor"
                                    style_requesttime_begin (j) = "errorcolor"
                                    dayErrorFlag            (j) = "error"
                                End If
                            End if
                        Else
                            err_relation_36             = 1
                            style_morningholiday    (j) = "errorcolor"
                            style_requesttime_begin (j) = "errorcolor"
                            dayErrorFlag            (j) = "error"
                        End If
                        If (Len(Trim(Request.Form("requesttime_end")(i))) > 0) Then
                            If (legalTime(Request.Form("requesttime_end")(i))) Then
                                If (editTime(Request.Form("requesttime_end")(i)) <> "17:10") Then
                                    err_relation_36            = 1
                                    style_afternoonholiday (j) = "errorcolor"
                                    style_requesttime_end  (j) = "errorcolor"
                                    dayErrorFlag           (j) = "error"
                                End If
                            End if
                        Else
                            err_relation_36            = 1
                            style_afternoonholiday (j) = "errorcolor"
                            style_requesttime_end  (j) = "errorcolor"
                            dayErrorFlag           (j) = "error"
                        End If
                    Else
                        ' 午前のみ代替休暇のとき、時間単位代休は8:30-12:00でなければエラー
                        If (Len(Trim(Request.Form("requesttime_begin")(i))) > 0) Then
                            If (legalTime(Request.Form("requesttime_begin")(i))) Then
                                If (editTime(Request.Form("requesttime_begin")(i)) <> "08:30") Then
                                    err_relation_36             = 1
                                    style_morningholiday    (j) = "errorcolor"
                                    style_requesttime_begin (j) = "errorcolor"
                                    dayErrorFlag            (j) = "error"
                                End If
                            End if
                        Else
                            err_relation_36             = 1
                            style_morningholiday    (j) = "errorcolor"
                            style_requesttime_begin (j) = "errorcolor"
                            dayErrorFlag            (j) = "error"
                        End If
                        If (Len(Trim(Request.Form("requesttime_end")(i))) > 0) Then
                            If (legalTime(Request.Form("requesttime_end")(i))) Then
                                If (editTime(Request.Form("requesttime_end")(i)) <> "12:00") Then
                                    err_relation_36           = 1
                                    style_morningholiday  (j) = "errorcolor"
                                    style_requesttime_end (j) = "errorcolor"
                                    dayErrorFlag          (j) = "error"
                                End If
                            End if
                        Else
                            err_relation_36           = 1
                            style_morningholiday  (j) = "errorcolor"
                            style_requesttime_end (j) = "errorcolor"
                            dayErrorFlag          (j) = "error"
                        End If
                    End If
                Else
                    If (Request.Form("afternoonholiday")(i) = "4") Then
                        ' 午後のみ代替休暇のとき、時間単位代休は13:00-17:10でなければエラー
                        If (Len(Trim(Request.Form("requesttime_begin")(i))) > 0) Then
                            If (legalTime(Request.Form("requesttime_begin")(i))) Then
                                If (editTime(Request.Form("requesttime_begin")(i)) <> "13:00") Then
                                    err_relation_36             = 1
                                    style_afternoonholiday  (j) = "errorcolor"
                                    style_requesttime_begin (j) = "errorcolor"
                                    dayErrorFlag            (j) = "error"
                                End If
                            End if
                        Else
                            err_relation_36             = 1
                            style_afternoonholiday  (j) = "errorcolor"
                            style_requesttime_begin (j) = "errorcolor"
                            dayErrorFlag            (j) = "error"
                        End If
                        If (Len(Trim(Request.Form("requesttime_end")(i))) > 0) Then
                            If (legalTime(Request.Form("requesttime_end")(i))) Then
                                If (editTime(Request.Form("requesttime_end")(i)) <> "17:10") Then
                                    err_relation_36            = 1
                                    style_afternoonholiday (j) = "errorcolor"
                                    style_requesttime_end  (j) = "errorcolor"
                                    dayErrorFlag           (j) = "error"
                                End If
                            End if
                        Else
                            err_relation_36           = 1
                            style_afternoonholiday(j) = "errorcolor"
                            style_requesttime_end (j) = "errorcolor"
                            dayErrorFlag          (j) = "error"
                        End If
                    End If
                End If
            End If
            ' コアタイム有休チェック
            ' コアタイム有休入力時は午前、午後ともコアタイム有休でなければエラー
            If (Request.Form("morningholiday"  )(i) =  "9"  Or  _
                Request.Form("afternoonholiday")(i) =  "9") And _
               (Request.Form("morningholiday"  )(i) <> "9"  Or  _
                Request.Form("afternoonholiday")(i) <> "9") Then
                err_relation_51           = 1
                style_morningholiday  (j) = "errorcolor"
                style_afternoonholiday(j) = "errorcolor"
                dayErrorFlag(j)           = "error"
            End If
            ' 振替出勤チェック
            ' 出勤区分午前=1振替出勤のとき、出勤区分午後=2休日出勤, 3休日出勤(半日未満)でエラー
            If ((Request.Form("morningwork"      )(i) = "1"   Or  _
                 Request.Form("morningwork"      )(i) = "5")  And _
                (Request.Form("afternoonwork"    )(i) = "2"   Or  _
                 Request.Form("afternoonwork"    )(i) = "3"   Or  _
                 Request.Form("afternoonwork"    )(i) = "6")) Then
                err_relation_04                 = 1
                style_morningwork  (j)          = "errorcolor"
                style_afternoonwork(j)          = "errorcolor"
                dayErrorFlag(j)                 = "error"
            End If
            ' 出勤区分午後=1振替出勤, 4出張のとき、出勤区分午前=2休日出勤, 3休日出勤(半日未満)でエラー
            If ((Request.Form("afternoonwork")(i) = "1"   Or  _
                 Request.Form("afternoonwork")(i) = "5")  And _
                (Request.Form("morningwork"  )(i) = "2"   Or  _
                 Request.Form("morningwork"  )(i) = "3"   Or  _
                 Request.Form("morningwork"  )(i) = "6")) Then
                err_relation_04                 = 1
                style_afternoonwork  (j)        = "errorcolor"
                style_morningwork    (j)        = "errorcolor"
                dayErrorFlag(j)                 = "error"
            End If
            ' 出勤区分午前=1振替出勤のとき、休日区分午前=1公休日, 2振替休日, A法定休日 以外でエラー
            If ((Request.Form("morningwork"      )(i) =  "1"   Or  _
                 Request.Form("morningwork"      )(i) =  "5")  And _
                (Request.Form("morningholiday"   )(i) <> "1"   And _
                 Request.Form("morningholiday"   )(i) <> "2"   And _
                 Request.Form("morningholiday"   )(i) <> "A")) Then
                ' 出勤区分午前=1振替出勤、出勤区分午後=1振替出勤で休日区分午後=5特別休暇または7欠勤のときエラーとしない
                ' 出勤区分午後=1振替出勤、出勤区分午前=1振替出勤で休日区分午前=5特別休暇または7欠勤のときエラーとしない
                If ((Request.Form("morningwork"     )(i) = "1"   And _
                     Request.Form("afternoonwork"   )(i) = "1"   And _
                    (Request.Form("afternoonholiday")(i) = "5"   Or  _
                     Request.Form("afternoonholiday")(i) = "7")) Or  _
                    (Request.Form("morningwork"     )(i) = "1"   And _
                    (Request.Form("morningholiday"  )(i) = "5"   Or _
                     Request.Form("morningholiday"  )(i) = "7")  And _
                     Request.Form("afternoonwork"   )(i) = "1")) Then
                     ' エラーとしない
                Else
                    err_relation_24                 = 1
                    style_morningwork   (j)         = "errorcolor"
                    style_morningholiday(j)         = "errorcolor"
                    dayErrorFlag(j)                 = "error"
                End If
            End If
            ' 出勤区分午後=1振替出勤のとき、休日区分午後=1公休日, 2振替休日, A法定休日 以外でエラー
            If ((Request.Form("afternoonwork"    )(i) =  "1"   Or  _
                 Request.Form("afternoonwork"    )(i) =  "5")  And _
                (Request.Form("afternoonholiday" )(i) <> "1"   And _
                 Request.Form("afternoonholiday" )(i) <> "2"   And _
                 Request.Form("afternoonholiday" )(i) <> "A")) Then
                ' 出勤区分午前=1振替出勤、出勤区分午後=1振替出勤で休日区分午後=5特別休暇または7欠勤のときエラーとしない
                ' 出勤区分午後=1振替出勤、出勤区分午前=1振替出勤で休日区分午前=5特別休暇または7欠勤のときエラーとしない
                If ((Request.Form("morningwork"     )(i) = "1"   And _
                     Request.Form("afternoonwork"   )(i) = "1"   And _
                    (Request.Form("afternoonholiday")(i) = "5"   Or  _
                     Request.Form("afternoonholiday")(i) = "7")) Or  _
                    (Request.Form("morningwork"     )(i) = "1"   And _
                    (Request.Form("morningholiday"  )(i) = "5"   Or _
                     Request.Form("morningholiday"  )(i) = "7")  And _
                     Request.Form("afternoonwork"   )(i) = "1")) Then
                     ' エラーとしない
                Else
                    err_relation_24                 = 1
                    style_afternoonwork   (j)       = "errorcolor"
                    style_afternoonholiday(j)       = "errorcolor"
                    dayErrorFlag(j)                 = "error"
                End If
            End If
            ' 午前、午後どちらかのに休日区分、出勤区分が入力されている場合、
            ' 午前、午後両方に休日区分、出勤区分の入力がなければエラーとする。
            ' ただし、既に午前、午後の入力チェックでエラーとなっている場合、
            ' 交替勤務で日勤甲と生産会議乙のときはチェックしない。
            If err_relation_01 = 0  And _
               err_relation_02 = 0  And _
               err_relation_03 = 0  And _
               err_relation_04 = 0  And _
               err_relation_24 = 0  And _
               err_relation_50 = 0  And _
               operatorNoCheck = "" Then
                If (Request.Form("morningholiday"  )(i) = "0"  And  _
                    Request.Form("morningwork"     )(i) = "0") Then
                    temp_morningFlag = "0"
                Else
                    temp_morningFlag = "1"
                End If
                If (Request.Form("afternoonholiday")(i) = "0"  And  _
                    Request.Form("afternoonwork"   )(i) = "0") Then
                    temp_afternoonFlag = "0"
                Else
                    temp_afternoonFlag = "1"
                End If
                If temp_morningFlag <> temp_afternoonFlag Then
                    err_relation_27           = 1
                    style_morningwork     (j) = "errorcolor"
                    style_morningholiday  (j) = "errorcolor"
                    style_afternoonwork   (j) = "errorcolor"
                    style_afternoonholiday(j) = "errorcolor"
                    dayErrorFlag(j)           = "error"
                End If
                ' 生産オペレータのときの交替勤務日勤甲と生産会議乙時の出勤区分チェックを行う。
                If is_operator Then
                    ' 生産オペレータで交替勤務が日勤甲('3')のとき、出勤区分は下記のパターンでなければエラー
                    ' パターン１：午前=出勤 And 午後=振替出勤
                    ' パターン２：午前=出勤 And 午後=出勤
                    If (Request.Form("operator")(i) = "3") Then
                        If ((Request.Form("morningwork"  )(i) = "9"   And _
                             Request.Form("afternoonwork")(i) = "1")  Or  _
                            (Request.Form("morningwork"  )(i) = "9"   And _
                             Request.Form("afternoonwork")(i) = "9")) Then
                        Else
                            err_relation_30         = 1
                            style_morningwork   (j) = "errorcolor"
                            style_afternoonwork (j) = "errorcolor"
                            dayErrorFlag        (j) = "error"
                        End if
                    End If
                    ' 生産オペレータで交替勤務が生産会議乙('4')のとき、出勤区分は下記のパターンでなければエラー
                    ' パターン：午前=出勤 And 午後=振替出勤
                    If (Request.Form("operator")(i) = "4") Then
                        If (Request.Form("morningwork"  )(i) = "9"  And _
                            Request.Form("afternoonwork")(i) = "1") Then
                        Else
                            err_relation_31         = 1
                            style_morningwork   (j) = "errorcolor"
                            style_afternoonwork (j) = "errorcolor"
                            dayErrorFlag        (j) = "error"
                        End if
                    End If
                End If
            End If
            ' フレックス勤務時刻チェック
            If workshift = "9" Then
                ' 午前午後とも休暇が入力され、勤務に振替出勤、休出が入力されていないとき、勤務・休憩時間入力はエラー
                If Request.Form("morningholiday"  )(i) <> "0" And _
                   Request.Form("afternoonholiday")(i) <> "0" Then
                    If (Request.Form("morningwork"  )(i) = "1" Or _
                        Request.Form("morningwork"  )(i) = "2" Or _
                        Request.Form("morningwork"  )(i) = "3" Or _
                        Request.Form("morningwork"  )(i) = "4" Or _
                        Request.Form("morningwork"  )(i) = "5" Or _
                        Request.Form("morningwork"  )(i) = "6" Or _
                        Request.Form("morningwork"  )(i) = "9" Or _
                        Request.Form("afternoonwork")(i) = "1" Or _
                        Request.Form("afternoonwork")(i) = "2" Or _
                        Request.Form("afternoonwork")(i) = "3" Or _
                        Request.Form("afternoonwork")(i) = "4" Or _
                        Request.Form("afternoonwork")(i) = "5" Or _
                        Request.Form("afternoonwork")(i) = "6" Or _
                        Request.Form("afternoonwork")(i) = "9") Then
                    Else
                        If Len(Trim(Request.Form("work_begin"  )(i))) > 0 Or _
                           Len(Trim(Request.Form("work_end"    )(i))) > 0 Or _
                           Len(Trim(Request.Form("break_begin1")(i))) > 0 Or _
                           Len(Trim(Request.Form("break_end1"  )(i))) > 0 Or _
                           Len(Trim(Request.Form("break_begin2")(i))) > 0 Or _
                           Len(Trim(Request.Form("break_end2"  )(i))) > 0 Then
                            style_morningholiday   (j) = "errorcolor"
                            style_afternoonholiday (j) = "errorcolor"
                            style_morningwork      (j) = "errorcolor"
                            style_afternoonwork    (j) = "errorcolor"
                            style_work_begin       (j) = "errorcolor"
                            style_work_end         (j) = "errorcolor"
                            style_break_begin1     (j) = "errorcolor"
                            style_break_end1       (j) = "errorcolor"
                            style_break_begin2     (j) = "errorcolor"
                            style_break_end2       (j) = "errorcolor"
                            dayErrorFlag           (j) = "error"
                            err_relation_48            = 1
                        End If
                    End If
                End If
                If Request.Form("morningwork")(i) > "0"  Or  Request.Form("afternoonwork")(i) > "0" Then
                    ' 出勤区分が勤務で勤務時間の自、至どちらかしか入力がないときエラー
                    ' 両方入力無しの時はエラーとしない(標準時間で更新するため)
                    If (Len(Trim(Request.Form("work_begin")(i))) = 0  And Len(Trim(Request.Form("work_end")(i))) = 0) Or _
                       (Len(Trim(Request.Form("work_begin")(i))) > 0  And Len(Trim(Request.Form("work_end")(i))) > 0) Then
                    Else
                        err_relation_39   = 1
                        style_morningwork  (j) = "errorcolor"
                        style_afternoonwork(j) = "errorcolor"
                        style_work_begin   (j) = "errorcolor"
                        style_work_end     (j) = "errorcolor"
                    End If
                Else
                    ' 出勤区分が勤務以外で勤務時間の自、至等に入力があるときエラー
                    If (Len(Trim(Request.Form("work_begin"  )(i))) > 0 Or _
                        Len(Trim(Request.Form("work_end"    )(i))) > 0 Or _
                        Len(Trim(Request.Form("break_begin1")(i))) > 0 Or _
                        Len(Trim(Request.Form("break_end1"  )(i))) > 0 Or _
                        Len(Trim(Request.Form("break_begin2")(i))) > 0 Or _
                        Len(Trim(Request.Form("break_end2"  )(i))) > 0) Then
                        style_work_begin       (j) = "errorcolor"
                        style_work_end         (j) = "errorcolor"
                        style_break_begin1     (j) = "errorcolor"
                        style_break_end1       (j) = "errorcolor"
                        style_break_begin2     (j) = "errorcolor"
                        style_break_end2       (j) = "errorcolor"
                        dayErrorFlag           (j) = "error"
                        err_relation_48            = 1
                    End If
                End If
                ' フレックス勤務の自に入力有りのとき、至に入力無しのときエラー
                ' フレックス勤務の自に入力無しのとき、至に入力有りのときエラー
                If ((Len(Trim(Request.Form("work_begin")(i))) > 0   And _
                     Len(Trim(Request.Form("work_end"  )(i))) = 0)  Or _
                    (Len(Trim(Request.Form("work_begin")(i))) = 0   And _
                     Len(Trim(Request.Form("work_end"  )(i))) > 0)) Then
                    err_relation_42 = 1
                    style_work_begin (j) = "errorcolor"
                    style_work_end   (j) = "errorcolor"
                    dayErrorFlag     (j) = "error"
                End If
                ' フレックス勤務の休憩自に入力有りのとき、休憩至に入力無しのときエラー
                ' フレックス勤務の休憩自に入力無しのとき、休憩至に入力有りのときエラー
                If ((Len(Trim(Request.Form("break_begin1")(i))) > 0   And _
                     Len(Trim(Request.Form("break_end1"  )(i))) = 0)  Or _
                    (Len(Trim(Request.Form("break_begin1")(i))) = 0   And _
                     Len(Trim(Request.Form("break_end1"  )(i))) > 0)) Then
                    err_relation_43 = 1
                    style_break_begin1 (j) = "errorcolor"
                    style_break_end1   (j) = "errorcolor"
                    dayErrorFlag       (j) = "error"
                End If
                ' フレックス勤務の中抜自に入力有りのとき、中抜至に入力無しのときエラー
                ' フレックス勤務の中抜自に入力無しのとき、中抜至に入力有りのときエラー
                If ((Len(Trim(Request.Form("break_begin2")(i))) > 0   And _
                     Len(Trim(Request.Form("break_end2"  )(i))) = 0)  Or _
                    (Len(Trim(Request.Form("break_begin2")(i))) = 0   And _
                     Len(Trim(Request.Form("break_end2"  )(i))) > 0)) Then
                    err_relation_44 = 1
                    style_break_begin2 (j) = "errorcolor"
                    style_break_end2   (j) = "errorcolor"
                    dayErrorFlag       (j) = "error"
                End If
                ' フレックス勤務の休憩に時刻が入っているが、勤務時間に時刻が入っていないとエラー
                If (Len(Trim(Request.Form("break_begin1")(i))) > 0  And _
                    Len(Trim(Request.Form("break_end1"  )(i))) > 0  And _
                    Len(Trim(Request.Form("work_begin"  )(i))) = 0  And _
                    Len(Trim(Request.Form("work_end"    )(i))) = 0) Then
                    err_relation_45 = 1
                    style_work_begin   (j) = "errorcolor"
                    style_work_end     (j) = "errorcolor"
                    style_break_begin1 (j) = "errorcolor"
                    style_break_end1   (j) = "errorcolor"
                    dayErrorFlag       (j) = "error"
                End If
                ' フレックス勤務の中抜に時刻が入っているが、勤務時間に時刻が入っていないとエラー
                If (Len(Trim(Request.Form("break_begin2")(i))) > 0  And _
                    Len(Trim(Request.Form("break_end2"  )(i))) > 0  And _
                    Len(Trim(Request.Form("work_begin"  )(i))) = 0  And _
                    Len(Trim(Request.Form("work_end"    )(i))) = 0) Then
                    err_relation_45 = 1
                    style_work_begin   (j) = "errorcolor"
                    style_work_end     (j) = "errorcolor"
                    style_break_begin2 (j) = "errorcolor"
                    style_break_end2   (j) = "errorcolor"
                    dayErrorFlag       (j) = "error"
                End If
                ' フレックス勤務の勤務時間自と至の整合性チェック
                If err_relation_42 = 0 And err_relation_43 = 0 And _
                   err_relation_44 = 0 And err_relation_45 = 0 Then
                    ' 勤務時間自至と休憩自至に入力有りのとき、
                    ' 勤務時間自 < 休憩自 < 休憩至 < 勤務時間至 になっていなければエラー
                    If Len(Trim(Request.Form("work_begin"  )(i))) > 0 And _
                       Len(Trim(Request.Form("work_end"    )(i))) > 0 And _
                       Len(Trim(Request.Form("break_begin1")(i))) > 0 And _
                       Len(Trim(Request.Form("break_end1"  )(i))) > 0 Then
                        returnCode = checkChronological(Trim(Request.Form("work_begin"  )(i)), _
                                                        Trim(Request.Form("break_begin1")(i)), _
                                                        Trim(Request.Form("break_end1"  )(i)), _
                                                        Trim(Request.Form("work_end"    )(i)))
                        If returnCode <> 0 Then
                            err_relation_47       = 1
                            style_work_begin  (j) = "errorcolor"
                            style_work_end    (j) = "errorcolor"
                            style_break_begin1(j) = "errorcolor"
                            style_break_end1  (j) = "errorcolor"
                            dayErrorFlag      (j) = "error"
                        End If
                    End If
                    ' 勤務時間自至と中抜自至に入力有りのとき、
                    ' 勤務時間自 < 中抜自 < 中抜至 < 勤務時間至 になっていなければエラー
                    If Len(Trim(Request.Form("work_begin"  )(i))) > 0 And _
                       Len(Trim(Request.Form("work_end"    )(i))) > 0 And _
                       Len(Trim(Request.Form("break_begin2")(i))) > 0 And _
                       Len(Trim(Request.Form("break_end2"  )(i))) > 0 Then
                        returnCode = checkChronological(Trim(Request.Form("work_begin"  )(i)), _
                                                        Trim(Request.Form("break_begin2")(i)), _
                                                        Trim(Request.Form("break_end2"  )(i)), _
                                                        Trim(Request.Form("work_end"    )(i)))
                        If returnCode <> 0 Then
                            err_relation_47       = 1
                            style_work_begin  (j) = "errorcolor"
                            style_work_end    (j) = "errorcolor"
                            style_break_begin2(j) = "errorcolor"
                            style_break_end2  (j) = "errorcolor"
                            dayErrorFlag      (j) = "error"
                        End If
                    End If
                    ' 休憩自至と中抜自至に入力有りのとき、
                    ' 休憩自 < 休憩至 <= 中抜自 < 中抜至 または 中抜自 < 中抜至 <= 休憩自 < 休憩至 になっていなければエラー
                    If Len(Trim(Request.Form("break_begin1")(i))) > 0 And _
                       Len(Trim(Request.Form("break_end1"  )(i))) > 0 And _
                       Len(Trim(Request.Form("break_begin2")(i))) > 0 And _
                       Len(Trim(Request.Form("break_end2"  )(i))) > 0 Then
                        returnCode1 = checkChronological(Trim(Request.Form("break_begin1")(i)), _
                                                         Trim(Request.Form("break_end1"  )(i)), _
                                                         Trim(Request.Form("break_begin2")(i)), _
                                                         Trim(Request.Form("break_end2"  )(i)))
                        returnCode2 = checkChronological(Trim(Request.Form("break_begin2")(i)), _
                                                         Trim(Request.Form("break_end2"  )(i)), _
                                                         Trim(Request.Form("break_begin1")(i)), _
                                                         Trim(Request.Form("break_end1"  )(i)))
                        If returnCode1 <> 0 And returnCode2 <> 0 Then
                            err_relation_47       = 1
                            style_break_begin1(j) = "errorcolor"
                            style_break_end1  (j) = "errorcolor"
                            style_break_begin2(j) = "errorcolor"
                            style_break_end2  (j) = "errorcolor"
                            dayErrorFlag      (j) = "error"
                        End If
                    End If
                End If
                ' フレックス勤務の勤務時間自と至の整合性チェック
                ans = 0
                If err_relation_42 = 0 And err_relation_43 = 0 And _
                   err_relation_44 = 0 And err_relation_45 = 0 And _
                   err_relation_47 = 0 Then
                    ' 勤務時間と中抜け1のチェック 勤務開始 < 中抜け1開始 < 中抜け1終了 < 勤務終了
                    ans = ans + checkChronological3(Trim(Request.Form("work_begin"  )(i)), _
                                                    Trim(Request.Form("break_begin1")(i)), _
                                                    Trim(Request.Form("break_end1"  )(i)), _
                                                    Trim(Request.Form("work_end"    )(i)))
'                   ' 勤務時間と中抜け2のチェック 勤務開始 < 中抜け2開始 < 中抜け2終了 < 勤務終了
                    ans = ans + checkChronological3(Trim(Request.Form("work_begin"  )(i)), _
                                                    Trim(Request.Form("break_begin2")(i)), _
                                                    Trim(Request.Form("break_end2"  )(i)), _
                                                    Trim(Request.Form("work_end"    )(i)))
                    ' 勤務時間と時間有給のチェック  勤務開始 <= 時間有給開始 < 時間有給終了 <= 勤務終了
                    ans = ans + checkChronological2_noentry_supported( _
                                                    Trim(Request.Form("work_begin"  )(i)), _
                                                    Trim(Request.Form("vacationtime_begin")(i)), _
                                                    Trim(Request.Form("vacationtime_end"  )(i)), _
                                                    Trim(Request.Form("work_end"    )(i)))
                    ' 中抜け1と中抜け2のチェック
                    ' 中抜け1開始 < 中抜け1終了 <= 中抜け2開始 < 中抜け2終了 Or
                    ' 中抜け2開始 < 中抜け2終了 <= 中抜け1開始 < 中抜け1終了
                    If checkChronological_noentry_supported( _
                            Trim(Request.Form("break_begin1")(i)), _
                            Trim(Request.Form("break_end1"  )(i)), _
                            Trim(Request.Form("break_begin2")(i)), _
                            Trim(Request.Form("break_end2"  )(i))) > 0 And _
                       checkChronological_noentry_supported( _
                            Trim(Request.Form("break_begin2")(i)), _
                            Trim(Request.Form("break_end2"  )(i)), _
                            Trim(Request.Form("break_begin1")(i)), _
                            Trim(Request.Form("break_end1"  )(i))) > 0 Then
                       ans = ans + 1
                    End If
                    ' 中抜け1と時間有給のチェック
                    ' 中抜け1開始  < 中抜け1終了  <= 時間有給開始 < 時間有給終了 Or
                    ' 時間有給開始 < 時間有給終了 <= 中抜け1開始  < 中抜け1終了
                    If checkChronological_noentry_supported( _
                            Trim(Request.Form("break_begin1"      )(i)), _
                            Trim(Request.Form("break_end1"        )(i)), _
                            Trim(Request.Form("vacationtime_begin")(i)), _
                            Trim(Request.Form("vacationtime_end"  )(i))) > 0 And _
                       checkChronological_noentry_supported( _
                            Trim(Request.Form("vacationtime_begin")(i)), _
                            Trim(Request.Form("vacationtime_end"  )(i)), _
                            Trim(Request.Form("break_begin1"      )(i)), _
                            Trim(Request.Form("break_end1"        )(i))) > 0 Then
                       ans = ans + 1
                    End If
                    ' 中抜け2と時間有給のチェック
                    ' 中抜け2開始  < 中抜け2終了  <= 時間有給開始 < 時間有給終了 Or
                    ' 時間有給開始 < 時間有給終了 <= 中抜け2開始  < 中抜け2終了
                    If checkChronological_noentry_supported( _
                            Trim(Request.Form("break_begin2"      )(i)), _
                            Trim(Request.Form("break_end2"        )(i)), _
                            Trim(Request.Form("vacationtime_begin")(i)), _
                            Trim(Request.Form("vacationtime_end"  )(i))) > 0 And _
                       checkChronological_noentry_supported( _
                            Trim(Request.Form("vacationtime_begin")(i)), _
                            Trim(Request.Form("vacationtime_end"  )(i)), _
                            Trim(Request.Form("break_begin2"      )(i)), _
                            Trim(Request.Form("break_end2"        )(i))) > 0 Then
                        ans = ans + 1
                    End If
                    If ans > 0 Then
                        err_relation_52 = 1
                        style_work_begin  (j) = "errorcolor"
                        style_work_end    (j) = "errorcolor"
                        style_break_begin1(j) = "errorcolor"
                        style_break_end1  (j) = "errorcolor"
                        style_break_begin2(j) = "errorcolor"
                        style_break_end2  (j) = "errorcolor"
                        style_vacationtime_begin(j) = "errorcolor"
                        style_vacationtime_end  (j) = "errorcolor"
                        dayErrorFlag      (j) = "error"
                    End If
                End If
                ' フレックス勤務で勤務時間自至と時間外自至がかぶっている場合はエラーとする [2024-09-17追加]
                If Len(Trim(Request.Form("work_begin")(i))) > 0 And _
                   Len(Trim(Request.Form("work_end"  )(i))) > 0 And _
                   Len(Trim(Request.Form("overtime_begin")(i))) > 0 And _
                   Len(Trim(Request.Form("overtime_end"  )(i))) > 0 Then
                   If (checkChronological(Trim(Request.Form("work_begin")(i)), Trim(Request.Form("work_end")(i)), Trim(Request.Form("overtime_begin")(i)), Trim(Request.Form("overtime_end")(i))) <> 0) And _
                      (checkChronological(Trim(Request.Form("overtime_begin")(i)), Trim(Request.Form("overtime_end")(i)), Trim(Request.Form("work_begin")(i)), Trim(Request.Form("work_end")(i))) <> 0) Then
                      err_relation_53         = 1
                      style_work_begin    (j) = "errorcolor"
                      style_work_end      (j) = "errorcolor"
                      style_overtime_begin(j) = "errorcolor"
                      style_overtime_end  (j) = "errorcolor"
                      dayErrorFlag(j)         = "error"
                   End If
                End If
            End If
            ' 時間外(休出)申請分時刻チェック
            ' 時間外(休出)申請分自に入力有りのとき、時間外(休出)申請分至に入力が無し場合はエラー
            ' 時間外(休出)申請分至に入力有りのとき、時間外(休出)申請分自に入力が無し場合はエラー
            If ((Len(Trim(Request.Form("overtime_begin")(i))) > 0   And _
                 Len(Trim(Request.Form("overtime_end"  )(i))) = 0)  Or  _
                (Len(Trim(Request.Form("overtime_begin")(i))) = 0   And _
                 Len(Trim(Request.Form("overtime_end"  )(i))) > 0)) Then
                err_relation_07             = 1
                style_overtime_begin(j)     = "errorcolor"
                style_overtime_end  (j)     = "errorcolor"
                dayErrorFlag(j)             = "error"
            End If
            ' 時間外申請時、出勤区分チェック
            ' 時間外申請時、出勤区分の午前午後どちらかに値がなければエラーとする。
            If (Len(Trim(Request.Form("overtime_begin")(i))) > 0  And _
                Len(Trim(Request.Form("overtime_end"  )(i))) > 0) And _
               (Request.Form("morningwork"   )(i) = "0"    And _
                Request.Form("afternoonwork" )(i) = "0")   Then
                ' フレックス勤務者で宿直、かつ公休日、法定休日時は出勤区分入力チェックは不要
                If (workshift = "9" And Request.Form("nightduty")(i) >= "1") And _
                   (Request.Form("morningholiday")(i) = "1" Or _
                    Request.Form("morningholiday")(i) = "A" ) Then
                Else
                    err_relation_08             = 1
                    style_morningwork   (j)     = "errorcolor"
                    style_afternoonwork (j)     = "errorcolor"
                    style_overtime_begin(j)     = "errorcolor"
                    style_overtime_end  (j)     = "errorcolor"
                    dayErrorFlag(j)             = "error"
                End If
            End If
            ' 時間外(休出)申請分自および、時間外(休出)申請分至に入力無しのとき、
            ' 時間外(休出)申請分休憩自または時間外(休出)申請分休憩至に入力有りのときエラー
            If ((Len(Trim(Request.Form("overtime_begin")(i))) = 0   And _
                 Len(Trim(Request.Form("overtime_end"  )(i))) = 0)  And _
                (Len(Trim(Request.Form("rest_begin"    )(i))) > 0   Or  _
                 Len(Trim(Request.Form("rest_end"      )(i))) > 0)) Then
                err_relation_09             = 1
                style_overtime_begin(j)     = "errorcolor"
                style_overtime_end  (j)     = "errorcolor"
                style_rest_begin    (j)     = "errorcolor"
                style_rest_end      (j)     = "errorcolor"
                dayErrorFlag(j) = "error"
            End If
            ' 時間外(休出)申請分休憩自に入力有りのとき、時間外(休出)申請分休憩至に入力無しのときエラー
            ' 時間外(休出)申請分休憩自に入力無しのとき、時間外(休出)申請分休憩至に入力有りのときエラー
            If ((Len(Trim(Request.Form("rest_begin")(i))) > 0   And _
                 Len(Trim(Request.Form("rest_end"  )(i))) = 0)  Or  _
                (Len(Trim(Request.Form("rest_begin")(i))) = 0   And _
                 Len(Trim(Request.Form("rest_end"  )(i))) > 0)) Then
                err_relation_10             = 1
                style_rest_begin    (j) = "errorcolor"
                style_rest_end      (j) = "errorcolor"
                dayErrorFlag(j) = "error"
            End If
            ' 時間外(休出)申請分自至と時間外(休出)申請分休憩自至に入力有りのとき、
            ' 時間外(休出)申請分自< 時間外(休出)申請分休憩自 < 時間外(休出)申請分休憩至 < 時間外(休出)申請分至
            ' になっていなければエラー
            If Len(Trim(Request.Form("overtime_begin")(i))) > 0 And _
               Len(Trim(Request.Form("overtime_end"  )(i))) > 0 And _
               Len(Trim(Request.Form("rest_begin"    )(i))) > 0 And _
               Len(Trim(Request.Form("rest_end"      )(i))) > 0 Then
                returnCode = checkChronological(Trim(Request.Form("overtime_begin")(i)), _
                                                Trim(Request.Form("rest_begin"    )(i)), _
                                                Trim(Request.Form("rest_end"      )(i)), _
                                                Trim(Request.Form("overtime_end"  )(i)))
                If returnCode <> 0 Then
                    err_relation_11         = 1
                    style_overtime_begin(j) = "errorcolor"
                    style_overtime_end  (j) = "errorcolor"
                    style_rest_begin    (j) = "errorcolor"
                    style_rest_end      (j) = "errorcolor"
                    dayErrorFlag(j)         = "error"
                End If
            End If
            ' 時間外(休出)入力時の休憩時間チェック
            ' 時間外(休出)分算出
            overtimeMin = 0
            If (Len(Trim(Request.Form("overtime_begin")(i))) > 0  And _
                Len(Trim(Request.Form("overtime_end"  )(i))) > 0) Then
                If (legalTime(Request.Form("overtime_begin")(i))  And _
                    legalTime(Request.Form("overtime_end"  )(i))) Then
                    ' 時間外算出
                    overtimeMin = minDif(editTime(Request.Form("overtime_begin")(i)), _
                                         editTime(Request.Form("overtime_end"  )(i)))
                End If
            End If
            ' 時間外(休出)休憩時間分算出
            restMin     = 0
            If (Len(Trim(Request.Form("rest_begin")(i))) > 0  And _
                Len(Trim(Request.Form("rest_end"  )(i))) > 0) Then
                If (legalTime(Request.Form("rest_begin")(i))  And _
                    legalTime(Request.Form("rest_end"  )(i))) Then
                    ' 休憩時間算出
                    restMin = minDif(editTime(Request.Form("rest_begin")(i)), _
                                     editTime(Request.Form("rest_end"  )(i)))
                End If
            End If
            overtimeRealMin = overtimeMin    - restMin          ' 時間外実時間算出
            overtime_count  = overtime_count + overtimeRealMin  ' 時間代休チェック用時間外時間集計
            ' 時間外(休出除く)を求めるための処理
            v_morningholiday   = Trim(Request.Form("morningholiday"  )(i))
            v_afternoonholiday = Trim(Request.Form("afternoonholiday")(i))
            v_morningwork      = Trim(Request.Form("morningwork"     )(i))
            v_afternoonwork    = Trim(Request.Form("afternoonwork"   )(i))
            v_overtime_begin   = editTime(Trim(Request.Form("overtime_begin"  )(i)))
            v_overtime_end     = editTime(Trim(Request.Form("overtime_end"    )(i)))
            v_rest_begin       = editTime(Trim(Request.Form("rest_begin"      )(i)))
            v_rest_end         = editTime(Trim(Request.Form("rest_end"        )(i)))
            v_overtime                  = 0     ' 時間外
            v_overtimelate              = 0     ' 時間外深夜業
            v_holidayshift              = 0     ' 休日出勤
            v_holidayshiftovertime      = 0     ' 休出時間外
            v_holidayshiftlate          = 0     ' 休出深夜業
            v_holidayshiftovertimelate  = 0     ' 休出時間外深夜業
            compOverTimeDetail()
            overtimeonly_count = overtimeonly_count _
                               + time2Min(v_overtime) _
                               + time2Min(v_overtimelate) _
                               + time2Min(v_holidayshiftovertime) _
                               + time2Min(v_holidayshiftovertimelate)
            ' 時間外(休出)申請分至-時間外(休出)申請分自>=(休憩時間除いて)6時間のとき、
            ' 時間外(休出)申請分休憩至-時間外(休出)申請分休憩時間<45分のときエラー
            ' 出勤区分が午前午後に出勤、振替出勤、出張が入力されているときはチェックしない。
            If Not ((Request.Form("morningwork"  )(i) = "1"   Or _
                     Request.Form("morningwork"  )(i) = "4"   Or _
                     Request.Form("morningwork"  )(i) = "5"   Or _
                     Request.Form("morningwork"  )(i) = "9")  Or _
                    (Request.Form("afternoonwork")(i) = "1"   Or _
                     Request.Form("afternoonwork")(i) = "4"   Or _
                     Request.Form("afternoonwork")(i) = "5"   Or _
                     Request.Form("afternoonwork")(i) = "9")) Then
                If (overtimeRealMin > 360 And restMin < 45) Then
                    err_relation_12         = 1
                    style_rest_begin    (j) = "errorcolor"
                    style_rest_end      (j) = "errorcolor"
                    dayErrorFlag(j)         = "error"
                End If
            End If
            ' 時間外(休出)申請分至-時間外(休出)申請分自>=(休憩時間除いて)8時間のとき、
            ' 時間外(休出)申請分休憩至-時間外(休出)申請分休憩自<1時間のときエラー
            ' 出勤区分が午前午後に出勤、振替出勤、出張が入力されているときはチェックしない。
            If Not ((Request.Form("morningwork"  )(i) = "1"   Or _
                     Request.Form("morningwork"  )(i) = "4"   Or _
                     Request.Form("morningwork"  )(i) = "5"   Or _
                     Request.Form("morningwork"  )(i) = "9")  Or _
                    (Request.Form("afternoonwork")(i) = "1"   Or _
                     Request.Form("afternoonwork")(i) = "4"   Or _
                     Request.Form("afternoonwork")(i) = "5"   Or _
                     Request.Form("afternoonwork")(i) = "9")) Then
                If (overtimeRealMin > 480 And restMin < 60) Then
                    err_relation_13         = 1
                    err_relation_12         = 0
                    style_rest_begin    (j) = "errorcolor"
                    style_rest_end      (j) = "errorcolor"
                    dayErrorFlag(j)         = "error"
                End If
            End If
            If Not workshift = "9" Then ' フレックス勤務者以外のとき
                ' 時間単位代休チェック
                ' 時間単位代休申請分自に入力有りのとき、時間単位代休申請分至に入力がなければエラー
                ' 時間単位代休申請分至に入力有りのとき、時間単位代休申請分自に入力がなければエラー
                If ((Len(Trim(Request.Form("requesttime_begin")(i))) > 0   And _
                     Len(Trim(Request.Form("requesttime_end"  )(i))) = 0)  Or  _
                    (Len(Trim(Request.Form("requesttime_begin")(i))) = 0   And _
                     Len(Trim(Request.Form("requesttime_end"  )(i))) > 0)) Then
                    err_relation_14             = 1
                    style_requesttime_begin(j)  = "errorcolor"
                    style_requesttime_end  (j)  = "errorcolor"
                    dayErrorFlag(j)             = "error"
                End If
            End If
            ' 交替勤務入力時、時間単位代休の入力はエラー
            If is_operator Then
                If Request.Form("operator")(i) > "0" Then
                     If (Len(Trim(Request.Form("requesttime_begin")(i))) > 0  Or _
                         Len(Trim(Request.Form("requesttime_end"  )(i))) > 0) Then
                         err_relation_32              = 1
                         style_requesttime_begin (j)  = "errorcolor"
                         style_requesttime_end   (j)  = "errorcolor"
                         dayErrorFlag(j)              = "error"
                     End If
                End If
            End If
            If Not workshift = "9" Then ' フレックス勤務者以外のとき
                ' 時間単位代休申請分自>時間単位代休申請分至のときエラー
                If (Len(Trim(Request.Form("requesttime_begin")(i))) > 0  And _
                    Len(Trim(Request.Form("requesttime_end"  )(i))) > 0) Then
                    If (legalTime(Request.Form("requesttime_begin")(i))  And _
                        legalTime(Request.Form("requesttime_end"  )(i))) Then
                        If (editTime(Request.Form("requesttime_begin")(i))  >= _
                            editTime(Request.Form("requesttime_end"  )(i))) Then
                            err_relation_14             = 1
                            style_requesttime_begin (j) = "errorcolor"
                            style_requesttime_end   (j) = "errorcolor"
                            dayErrorFlag            (j) = "error"
                        Else
                            ' 時間単位代休取得時間チェック
                            ' 時間単位代休は10分単位でなければエラーとする。
                            temp = minDif(editTime(Request.Form("requesttime_begin")(i)),   _
                                          editTime(Request.Form("requesttime_end"  )(i)))
                            If (temp mod 10 > 0) Then
                                err_relation_26             = 1
                                style_requesttime_begin (j) = "errorcolor"
                                style_requesttime_end   (j) = "errorcolor"
                                dayErrorFlag            (j) = "error"
                            End If
                            ' 当月中に時間外労働がある場合のみ時間単位代休は入力可能とし、
                            ' 当月中の時間外労働がマイナスにならないようチェックする。
                            requesttime_count = requesttime_count                                               _
                                              + minDif(editTime(Request.Form("requesttime_begin")(i)),          _
                                                       editTime(Request.Form("requesttime_end"  )(i)))          _
                                              - checkLunchTime(editTime(Request.Form("requesttime_begin")(i)),  _
                                                               editTime(Request.Form("requesttime_end"  )(i)))
                            If requesttime_count > overtime_count Then
                                err_relation_15             = 1
                                style_requesttime_begin (j) = "errorcolor"
                                style_requesttime_end   (j) = "errorcolor"
                                dayErrorFlag            (j) = "error"
                            End If
                        End If
                    End If
                End If
            End If
            ' 時間有給チェック
            ' 時間有給申請分自に入力有りのとき、時間有給申請分至に入力がなければエラー
            ' 時間有給申請分至に入力有りのとき、時間有給申請分自に入力がなければエラー
            If ((Len(Trim(Request.Form("vacationtime_begin")(i))) > 0   And _
                 Len(Trim(Request.Form("vacationtime_end"  )(i))) = 0)  Or  _
                (Len(Trim(Request.Form("vacationtime_begin")(i))) = 0   And _
                 Len(Trim(Request.Form("vacationtime_end"  )(i))) > 0)) Then
                err_relation_22             = 1
                style_vacationtime_begin(j) = "errorcolor"
                style_vacationtime_end  (j) = "errorcolor"
                dayErrorFlag(j)             = "error"
            End If
            ' 時間有給申請分自>時間有給申請分至のときエラー
            If (Len(Trim(Request.Form("vacationtime_begin")(i))) > 0  And _
                Len(Trim(Request.Form("vacationtime_end"  )(i))) > 0) Then
                If (legalTime(Request.Form("vacationtime_begin")(i))  And _
                    legalTime(Request.Form("vacationtime_end"  )(i))) Then
                    If (editTime(Request.Form("vacationtime_begin")(i))   >= _
                        editTime(Request.Form("vacationtime_end"  )(i))) Then
                        err_relation_22                 = 1
                        style_vacationtime_begin    (j) = "errorcolor"
                        style_vacationtime_end      (j) = "errorcolor"
                        dayErrorFlag(j)                 = "error"
                    Else
                        vacationtime_count = vacationtime_count + _
                        minDif(editTime(Request.Form("vacationtime_begin")(i)), _
                               editTime(Request.Form("vacationtime_end"  )(i)))
                    End If
                End If
            End If
            ' 時間有給申請時、出勤区分チェック
            ' 時間有給申請時、出勤区分の午前午後どちらかに値がなければエラーとする。
            If (Len(Trim(Request.Form("vacationtime_begin")(i))) > 0  And _
                Len(Trim(Request.Form("vacationtime_end"  )(i))) > 0) And _
               (Request.Form("morningwork"   )(i) = "0"               And _
                Request.Form("afternoonwork" )(i) = "0")              Then
                err_relation_23             = 1
                style_morningwork   (j)     = "errorcolor"
                style_afternoonwork (j)     = "errorcolor"
                style_vacationtime_begin(j) = "errorcolor"
                style_vacationtime_end  (j) = "errorcolor"
                dayErrorFlag(j)             = "error"
            End If
            ' 時間単位有休取得時間チェック
            ' 時間単位有休は60分単位でなければエラーとする。
            If (Len(Trim(Request.Form("vacationtime_begin")(i))) > 0    And _
                Len(Trim(Request.Form("vacationtime_end"  )(i))) > 0)   Then
                temp = minDif(editTime(Request.Form("vacationtime_begin")(i)),  _
                              editTime(Request.Form("vacationtime_end"  )(i)))
                If (temp mod 60 > 0) Then
                    err_relation_25             = 1
                    style_vacationtime_begin(j) = "errorcolor"
                    style_vacationtime_end  (j) = "errorcolor"
                    dayErrorFlag(j)             = "error"
                End If
            End If
            ' 交替勤務入力時、時間単位有給の入力はエラー
            If is_operator Then
                If Request.Form("operator")(i) > "0" Then
                     If (Len(Trim(Request.Form("vacationtime_begin")(i))) > 0  Or _
                         Len(Trim(Request.Form("vacationtime_end"  )(i))) > 0) Then
                         err_relation_33              = 1
                         style_vacationtime_begin (j) = "errorcolor"
                         style_vacationtime_end   (j) = "errorcolor"
                         dayErrorFlag(j)              = "error"
                     End If
                End If
            End If
            ' 深夜割増チェック
            ' 深夜割増自に入力有りのとき、深夜割増至に入力がなければエラー
            ' 深夜割増至に入力有りのとき、深夜割増自に入力がなければエラー
            If ((Len(Trim(Request.Form("latetime_begin")(i))) > 0   And _
                 Len(Trim(Request.Form("latetime_end"  )(i))) = 0)  Or  _
                (Len(Trim(Request.Form("latetime_begin")(i))) = 0   And _
                 Len(Trim(Request.Form("latetime_end"  )(i))) > 0)) Then
                err_relation_16         = 1
                style_latetime_begin(j) = "errorcolor"
                style_latetime_end  (j) = "errorcolor"
                dayErrorFlag        (j) = "error"
            End If
            ' 深夜割増自及び深夜割増至は22:00～05:00で入力。以外の時はエラー
            If Len(Trim(Request.Form("latetime_begin")(i))) > 0 Then
                If editTime(Request.Form("latetime_begin")(i)) >= "22:00" Or _
                   editTime(Request.Form("latetime_begin")(i)) <= "05:00" Then
                Else
                    err_relation_28         = 1
                    style_latetime_begin(j) = "errorcolor"
                    dayErrorFlag        (j) = "error"
                End If
            End If
            If Len(Trim(Request.Form("latetime_end")(i))) > 0 Then
                If editTime(Request.Form("latetime_end")(i)) >= "22:00" Or _
                   editTime(Request.Form("latetime_end")(i)) <= "05:00" Then
                Else
                    err_relation_28         = 1
                    style_latetime_end  (j) = "errorcolor"
                    dayErrorFlag        (j) = "error"
                End If
            End If
            '時間外入力時間が深夜割増の時間帯にかかっていないとエラーとする。
            If Len(Trim(Request.Form("overtime_begin")(i))) > 0 Then
                If Len(Trim(Request.Form("latetime_begin")(i))) > 0 Then
                    If (checkChronological2(Trim(Request.Form("overtime_begin")(i)), _
                                            Trim(Request.Form("latetime_begin")(i)), _
                                            Trim(Request.Form("latetime_end"  )(i)), _
                                            Trim(Request.Form("overtime_end"  )(i))) <> 0) Then
                        If is_operator Then
                            If Request.Form("operator")(i) <> "0" Then
                                ' 交替勤務のときはこの入力チェックはスルーする
                            Else
                                err_relation_37 = 1
                                style_overtime_begin(j) = "errorcolor"
                                style_overtime_end  (j) = "errorcolor"
                                style_latetime_begin(j) = "errorcolor"
                                style_latetime_end  (j) = "errorcolor"
                                dayErrorFlag        (j) = "error"
                            End If
                        Else
                            err_relation_37 = 1
                            style_overtime_begin(j) = "errorcolor"
                            style_overtime_end  (j) = "errorcolor"
                            style_latetime_begin(j) = "errorcolor"
                            style_latetime_end  (j) = "errorcolor"
                            dayErrorFlag        (j) = "error"
                        End If
                    End If
                End If
            End If
            ' 有給取得日数のカウント
            ' 有休取得日数が有給残日数を超えたときエラーとする。
            If (Request.Form("morningholiday"  )(i) = "3") Then
                vacation_count = vacation_count + 0.5 + operatorAddDays(v_operator)
                If ((Request.Form("sumVacationnumberHidden") - vacation_count) < 0) Then
                    err_relation_17             = 1
                    style_morningholiday(j)     = "errorcolor"
                    dayErrorFlag(j)             = "error"
                End If
            End If
            If (Request.Form("afternoonholiday")(i) = "3") Then
                vacation_count = vacation_count + 0.5
                If ((Request.Form("sumVacationnumberHidden") - vacation_count) < 0) Then
                    err_relation_17             = 1
                    style_afternoonholiday(j)   = "errorcolor"
                    dayErrorFlag(j)             = "error"
                End If
            End If
            ' 振休取得日数のカウント
            ' 振休取得日数が振休残日数を超えたときエラーとするが、
            ' 振替休日を出勤前に取得することもあるので-2日までは入力可能
            If (Request.Form("morningwork"     )(i) = "1"  Or _
                Request.Form("morningwork"     )(i) = "5") Then
                holiday_count = holiday_count + 0.5 + operatorAddDays(v_operator)
            End If
            If (Request.Form("morningholiday"  )(i) = "2") Then
                holiday_count = holiday_count - 0.5 - operatorAddDays(v_operator)
                If ((Request.Form("sumHolidaynumberHidden") + holiday_count) < -2) Then
                    err_relation_18             = 1
                    style_morningholiday(j)     = "errorcolor"
                    dayErrorFlag(j)             = "error"
                End If
            End If
            If (Request.Form("afternoonwork"   )(i) = "1"  Or _
                Request.Form("afternoonwork"   )(i) = "5") Then
                holiday_count = holiday_count + 0.5
            End If
            If (Request.Form("afternoonholiday")(i) = "2") Then
                holiday_count = holiday_count - 0.5
                If ((Request.Form("sumHolidaynumberHidden") + holiday_count) < -2) Then
                    err_relation_18             = 1
                    style_afternoonholiday(j)   = "errorcolor"
                    dayErrorFlag(j)             = "error"
                End If
            End If
            ' 有給取得時、振替残日数>0のときエラーとする。
            If Request.Form("morningholiday"  )(i) = "3"                    And _
               (Request.Form("sumHolidaynumberHidden") + holiday_count) > 1 Then
                err_relation_19             = 1
                style_morningholiday(j)     = "errorcolor"
                dayErrorFlag(j)             = "error"
            End If
            If Request.Form("afternoonholiday")(i) = "3"                    And _
               (Request.Form("sumHolidaynumberHidden") + holiday_count) > 1 Then
                err_relation_19             = 1
                style_afternoonholiday(j)   = "errorcolor"
                dayErrorFlag(j)             = "error"
            End If
            ' 交替勤務入力時、宿直の入力はエラー
            If is_operator Then
                If Request.Form("operator")(i) > "0" Then
                    If (Request.Form("nightduty")(i) > "0") Then
                        err_relation_34     = 1
                        style_nightduty (j) = "errorcolor"
                        dayErrorFlag    (j) = "error"
                    End If
                End If
            End If
            ' 日直入力時、休日区分が公休日、法定休日でないときエラーとする。
            If  Request.Form("dayduty"         )(i)  > "0"  And _
               (Request.Form("morningholiday"  )(i) <> "1"  Or  _
                Request.Form("afternoonholiday")(i) <> "1") And _
               (Request.Form("morningholiday"  )(i) <> "A"  Or  _
                Request.Form("afternoonholiday")(i) <> "A") Then
                err_relation_20             = 1
                style_morningholiday  (j)   = "errorcolor"
                style_afternoonholiday(j)   = "errorcolor"
                style_dayduty         (j)   = "errorcolor"
                dayErrorFlag(j)             = "error"
            End If
            ' 出勤区分=振替出勤 のとき、日直(責任者)のときエラー
            If (Request.Form("dayduty"         )(i) = "1"   And _
               (Request.Form("morningwork"     )(i) = "1"   Or  _
                Request.Form("morningwork"     )(i) = "4"   Or  _
                Request.Form("morningwork"     )(i) = "5"   Or  _
                Request.Form("afternoonwork"   )(i) = "1"   Or  _
                Request.Form("afternoonwork"   )(i) = "4"   Or  _
                Request.Form("afternoonwork"   )(i) = "5")) Then
                err_relation_21         = 1
                style_morningwork  (j)  = "errorcolor"
                style_afternoonwork(j)  = "errorcolor"
                style_dayduty      (j)  = "errorcolor"
                dayErrorFlag(j)         = "error"
            End If
            ' 交替勤務入力時、日直の入力はエラー
            If is_operator Then
                If Request.Form("operator")(i) > "0" Then
                    If (Request.Form("dayduty")(i) > "0") Then
                        err_relation_35   = 1
                        style_dayduty (j) = "errorcolor"
                        dayErrorFlag  (j) = "error"
                    End If
                End If
            End If
            ' 交替勤務入力ありで、休日・出勤区分に入力がないときエラー
            If is_operator Then
                If Request.Form("operator")(i) > "0" Then
                    If (Request.Form("morningholiday"  )(i) = "0"  And _
                        Request.Form("morningwork"     )(i) = "0") Or  _
                       (Request.Form("afternoonholiday")(i) = "0"  And _
                        Request.Form("afternoonwork"   )(i) = "0") Then
                        ' 日勤甲の場合は午後の休日・出勤区分が未入力でもエラートはしない
                        If Request.Form("operator")(i) <> "3" Then
                            err_relation_29           = 1
                            style_morningholiday  (j) = "errorcolor"
                            style_morningwork     (j) = "errorcolor"
                            style_afternoonholiday(j) = "errorcolor"
                            style_afternoonwork   (j) = "errorcolor"
                            style_operator        (j) = "errorcolor"
                            dayErrorFlag(j)           = "error"
                        End If
                    End If
                End If
            End If
            i = i + 1
        Else
            ' 上長チェックされているとき
            ' -----------------------------------------------------------------
            ' 休出回数をカウント(下にも同様の処理有り)
            ' -----------------------------------------------------------------
            If Request.Form("hd_morningwork"  )(x) = "2" Or _
               Request.Form("hd_morningwork"  )(x) = "3" Or _
               Request.Form("hd_morningwork"  )(x) = "6" Or _
               Request.Form("hd_afternoonwork")(x) = "2" Or _
               Request.Form("hd_afternoonwork")(x) = "3" Or _
               Request.Form("hd_afternoonwork")(x) = "6" Then
                holidaywork_count = holidaywork_count + 1
            End If
            ' ------------------------------------------------------------------
            ' 上長チェック済みの時間外と時間代休の集計処理(下にも同様の処理有り)
            ' ------------------------------------------------------------------
            ' 時間外(休出)入力時の休憩時間チェック
            ' 時間外(休出)分算出
            overtimeMin = 0
            If (Len(Trim(Request.Form("hd_overtime_begin")(x))) > 0  And _
                Len(Trim(Request.Form("hd_overtime_end"  )(x))) > 0) Then
                If (legalTime(Request.Form("hd_overtime_begin")(x))  And _
                    legalTime(Request.Form("hd_overtime_end"  )(x))) Then
                    ' 時間外算出
                    overtimeMin = minDif(editTime(Request.Form("hd_overtime_begin")(x)), _
                                         editTime(Request.Form("hd_overtime_end"  )(x)))
                End If
            End If
            ' 時間外(休出)休憩時間分算出
            restMin     = 0
            If (Len(Trim(Request.Form("hd_rest_begin")(x))) > 0  And _
                Len(Trim(Request.Form("hd_rest_end"  )(x))) > 0) Then
                If (legalTime(Request.Form("hd_rest_begin")(x))  And _
                    legalTime(Request.Form("hd_rest_end"  )(x))) Then
                    ' 休憩時間算出
                    restMin = minDif(editTime(Request.Form("hd_rest_begin")(x)), _
                                     editTime(Request.Form("hd_rest_end"  )(x)))
                End If
            End If
            overtimeRealMin = overtimeMin    - restMin          ' 時間外実時間算出
            overtime_count  = overtime_count + overtimeRealMin  ' 時間代休チェック用時間外時間集計
            ' 時間外(休出除く)を求めるための処理
            v_morningholiday   = Trim(Request.Form("hd_morningholiday"  )(x))
            v_afternoonholiday = Trim(Request.Form("hd_afternoonholiday")(x))
            v_morningwork      = Trim(Request.Form("hd_morningwork"     )(x))
            v_afternoonwork    = Trim(Request.Form("hd_afternoonwork"   )(x))
            v_overtime_begin   = editTime(Trim(Request.Form("hd_overtime_begin"  )(x)))
            v_overtime_end     = editTime(Trim(Request.Form("hd_overtime_end"    )(x)))
            v_rest_begin       = editTime(Trim(Request.Form("hd_rest_begin"      )(x)))
            v_rest_end         = editTime(Trim(Request.Form("hd_rest_end"        )(x)))
            v_overtime                  = 0     ' 時間外
            v_overtimelate              = 0     ' 時間外深夜業
            v_holidayshift              = 0     ' 休日出勤
            v_holidayshiftovertime      = 0     ' 休出時間外
            v_holidayshiftlate          = 0     ' 休出深夜業
            v_holidayshiftovertimelate  = 0     ' 休出時間外深夜業
            compOverTimeDetail()
            overtimeonly_count = overtimeonly_count _
                               + time2Min(v_overtime) _
                               + time2Min(v_overtimelate) _
                               + time2Min(v_holidayshiftovertime) _
                               + time2Min(v_holidayshiftovertimelate)
            If Not workshift = "9" Then ' フレックス勤務者以外のとき
                ' 時間代休申請分自>時間代休申請分至のときエラー
                If (Len(Trim(Request.Form("hd_requesttime_begin")(x))) > 0  And _
                    Len(Trim(Request.Form("hd_requesttime_end"  )(x))) > 0) Then
                    If (legalTime(Request.Form("hd_requesttime_begin")(x))  And _
                        legalTime(Request.Form("hd_requesttime_end"  )(x))) Then
                        If (editTime(Request.Form("hd_requesttime_begin")(x))  >= _
                            editTime(Request.Form("hd_requesttime_end"  )(x))) Then
                        Else
                            ' 時間代休取得時間チェック
                            ' 時間代休は10分単位でなければエラーとする。
                            temp = minDif(editTime(Request.Form("hd_requesttime_begin")(x)),   _
                                          editTime(Request.Form("hd_requesttime_end"  )(x)))
                            ' 当月中に時間外労働がある場合のみ時間代休は入力可能とし、
                            ' 当月中の時間外労働がマイナスにならないようチェックする。
                            requesttime_count = requesttime_count                                                  _
                                              + minDif(editTime(Request.Form("hd_requesttime_begin")(x)),          _
                                                       editTime(Request.Form("hd_requesttime_end"  )(x)))          _
                                              - checkLunchTime(editTime(Request.Form("hd_requesttime_begin")(x)),  _
                                                               editTime(Request.Form("hd_requesttime_end"  )(x)))
                        End If
                    End If
                End If
            End If
        End If
    Else
        ' -----------------------------------------------------------------
        ' 休出回数をカウント(上にも同様の処理有り)
        ' -----------------------------------------------------------------
        If Request.Form("hd_morningwork"  )(x) = "2" Or _
           Request.Form("hd_morningwork"  )(x) = "3" Or _
           Request.Form("hd_morningwork"  )(x) = "6" Or _
           Request.Form("hd_afternoonwork")(x) = "2" Or _
           Request.Form("hd_afternoonwork")(x) = "3" Or _
           Request.Form("hd_afternoonwork")(x) = "6" Then
            holidaywork_count = holidaywork_count + 1
        End If
        ' ---------------------------------------------------------------------
        ' 上長チェック済みの時間外と時間代休の集計処理(上にも同様の処理有り)
        ' ---------------------------------------------------------------------
        ' 時間外(休出)入力時の休憩時間チェック
        ' 時間外(休出)分算出
        overtimeMin = 0
        If (Len(Trim(Request.Form("hd_overtime_begin")(x))) > 0  And _
            Len(Trim(Request.Form("hd_overtime_end"  )(x))) > 0) Then
            If (legalTime(Request.Form("hd_overtime_begin")(x))  And _
                legalTime(Request.Form("hd_overtime_end"  )(x))) Then
                ' 時間外算出
                overtimeMin = minDif(editTime(Request.Form("hd_overtime_begin")(x)), _
                                     editTime(Request.Form("hd_overtime_end"  )(x)))
            End If
        End If
        ' 時間外(休出)休憩時間分算出
        restMin     = 0
        If (Len(Trim(Request.Form("hd_rest_begin")(x))) > 0  And _
            Len(Trim(Request.Form("hd_rest_end"  )(x))) > 0) Then
            If (legalTime(Request.Form("hd_rest_begin")(x))  And _
                legalTime(Request.Form("hd_rest_end"  )(x))) Then
                ' 休憩時間算出
                restMin = minDif(editTime(Request.Form("hd_rest_begin")(x)), _
                                 editTime(Request.Form("hd_rest_end"  )(x)))
            End If
        End If
        overtimeRealMin = overtimeMin    - restMin          ' 時間外実時間算出
        overtime_count  = overtime_count + overtimeRealMin  ' 時間代休チェック用時間外時間集計
        If Not workshift = "9" Then ' フレックス勤務者以外のとき
            ' 時間代休申請分自>時間代休申請分至のときエラー
            If (Len(Trim(Request.Form("hd_requesttime_begin")(x))) > 0  And _
                Len(Trim(Request.Form("hd_requesttime_end"  )(x))) > 0) Then
                If (legalTime(Request.Form("hd_requesttime_begin")(x))  And _
                    legalTime(Request.Form("hd_requesttime_end"  )(x))) Then
                    If (editTime(Request.Form("hd_requesttime_begin")(x))  >= _
                        editTime(Request.Form("hd_requesttime_end"  )(x))) Then
                    Else
                        ' 時間代休取得時間チェック
                        ' 時間代休は10分単位でなければエラーとする。
                        temp = minDif(editTime(Request.Form("hd_requesttime_begin")(x)),   _
                                      editTime(Request.Form("hd_requesttime_end"  )(x)))
                        ' 当月中に時間外労働がある場合のみ時間代休は入力可能とし、
                        ' 当月中の時間外労働がマイナスにならないようチェックする。
                        requesttime_count = requesttime_count                                                  _
                                          + minDif(editTime(Request.Form("hd_requesttime_begin")(x)),          _
                                                   editTime(Request.Form("hd_requesttime_end"  )(x)))          _
                                          - checkLunchTime(editTime(Request.Form("hd_requesttime_begin")(x)),  _
                                                           editTime(Request.Form("hd_requesttime_end"  )(x)))
                    End If
                End If
            End If
        End If
    End If
Next
' -------------------------------------------------------------------------
' エラーメッセージの整形
' -------------------------------------------------------------------------
tempIdx = 0
' 時刻チェック
If err_beginTime            = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "出勤時刻"
    tempIdx = tempIdx       + 1
End If
If err_returnTime            = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "戻り時刻"
    tempIdx = tempIdx       + 1
End If
If err_outTime            = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "外出時刻"
    tempIdx = tempIdx       + 1
End If
If err_endTime              = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "退勤時刻"
    tempIdx = tempIdx       + 1
End If
' フレックス勤務
If err_work_begin           = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "勤務時間自"
    tempIdx = tempIdx       + 1
End If
If err_work_end             = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "勤務時間至"
    tempIdx = tempIdx       + 1
End If
If err_break_begin1         = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "休憩時間自"
    tempIdx = tempIdx       + 1
End If
If err_break_end1           = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "休憩時間至"
    tempIdx = tempIdx       + 1
End If
If err_break_begin2         = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "中抜時間自"
    tempIdx = tempIdx       + 1
End If
If err_break_end2           = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "中抜時間至"
    tempIdx = tempIdx       + 1
End If
If err_overtime_begin       = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "時間外自"
    tempIdx = tempIdx       + 1
End If
If err_overtime_end         = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "時間外至"
    tempIdx = tempIdx       + 1
End If
If  err_rest_begin          = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "休憩自"
    tempIdx = tempIdx       + 1
End If
If  err_rest_end            = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "休憩至"
    tempIdx = tempIdx       + 1
End If
If  err_requesttime_begin   = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "時間代休自"
    tempIdx = tempIdx       + 1
End If
If  err_requesttime_end     = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "時間代休至"
    tempIdx = tempIdx       + 1
End If
If  err_vacationtime_begin  = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "時間有給自"
    tempIdx = tempIdx       + 1
End If
If  err_vacationtime_end    = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "時間有給至"
    tempIdx = tempIdx       + 1
End If
If  err_latetime_begin      = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "深夜割増自"
    tempIdx = tempIdx       + 1
End If
If  err_latetime_end        = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "深夜割増至"
    tempIdx = tempIdx       + 1
End If
If  err_weekovertime        = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "週超過時間"
    tempIdx = tempIdx       + 1
End If
If (tempIdx > 0) Then
    tempMessage = "次の項目が時刻として妥当でありません["
    For arrayCounter = 0 To tempIdx - 1 Step 1
        tempMessage = tempMessage + tempArray(arrayCounter) & ","
    Next
    errorMsg = Left(tempMessage, Len(tempMessage)-1) & "]。" & errorMsg
End If
' 関連チェック
If err_relation_01 = 1 Then
    errorMsg = errorMsg & "休日区分の午前と午後が妥当でありません。"
End If
If err_relation_02 = 1 Then
    errorMsg = errorMsg & "休日区分が公休日のとき出勤区分に出勤は選択できません。"
End If
If err_relation_50 = 1 Then
    errorMsg = errorMsg & "休日区分が法定休日のとき出勤区分に出勤は選択できません。"
End If
If err_relation_03 = 1 Then
    errorMsg = errorMsg & "休日区分が有給、特休、保存休、半日欠勤のとき、出勤区分に振替出勤、休出、休出半日未満、出張は選択できません。"
End If
If err_relation_36 = 1 Then
    errorMsg = errorMsg & "「代替休暇」と「時間単位代休」の時刻の組合せが妥当ではありません。"
End If
If err_relation_51 = 1 Then
    errorMsg = errorMsg & "「コアタイム有休」入力時は午前午後とも「コアタイム有休」でなければなりません。"
End If
If err_relation_04 = 1 Then
    errorMsg = errorMsg & "出勤に午前と午後で「振替出勤」と「休出、休出半日未満」の組合せは選択できません。"
End If
If err_relation_24 = 1 Then
    errorMsg = errorMsg & "振替出勤のときは、休日区分には「公休日」または「振替休日」「法定休日」以外選択できません。"
End If
If err_relation_27 = 1 Then
    errorMsg = errorMsg & "午前、午後どちらかのみの入力はできません。"
End If
If err_relation_29 = 1 Then
    errorMsg = errorMsg & "交替勤務を入力するときは、休日・出勤区分にも入力してください。"
End If
If err_relation_30 = 1 Then
    errorMsg = errorMsg & "日勤甲のとき、出勤区分は出勤、振替出勤または出勤、出勤で入力してください。"
End If
If err_relation_31 = 1 Then
    errorMsg = errorMsg & "生産会議乙のとき、出勤区分は出勤、振替出勤で入力してください。"
End If
If err_relation_05 = 1 Then
    errorMsg = errorMsg & "呼出区分が通常のとき、時間外(休出)申請分自が 5:00～22:00でなければなりません。"
End If
If err_relation_42 = "1" Then
    errorMsg = errorMsg & "勤務時間の自と至どちらかしか入力されていません。"
End If
If err_relation_43 = "1" Then
    errorMsg = errorMsg & "休憩時間の自と至どちらかしか入力されていません。"
End If
If err_relation_44 = "1" Then
    errorMsg = errorMsg & "中抜時間の自と至どちらかしか入力されていません。"
End If
If err_relation_45 = "1" Then
    errorMsg = errorMsg & "勤務時間未入力で、休憩時間が入力されています。"
End If
If err_relation_47 = 1 Then
    errorMsg = errorMsg & "勤務、休憩、中抜時刻が妥当でありません。"
End If
If err_relation_48 = 1 Then
    errorMsg = errorMsg & "勤務でないときは、勤務休憩時刻は入力できません。"
End If
If err_relation_06 = 1 Then
    errorMsg = errorMsg & "呼出区分が深夜のとき、時間外(休出)申請分自が" & _
                          "22:00～5:00でなければなりません。"
End If
If err_relation_07 = 1 Then
    errorMsg = errorMsg & "時間外申請の自と至どちらかしか入力されていません。"
End If
If err_relation_08 = 1 Then
    errorMsg = errorMsg & "時間外申請を入力するときは、出勤区分に入力が必要です。"
End If
If err_relation_09 = 1 Then
    errorMsg = errorMsg & "時間外申請に入力が無いとき、時間外休憩の入力ができません。"
End If
If err_relation_10 = 1 Then
    errorMsg = errorMsg & "時間外休憩の自と至どちらかしか入力されていません。"
End If
If err_relation_11 = 1 Then
    errorMsg = errorMsg & "時間外申請と時間外休憩の時刻が妥当でありません。"
End If
If err_relation_12 = 1 Then
    errorMsg = errorMsg & "時間外申請が6時間以上のとき、休憩が45分以上必要です。"
End If
If err_relation_13 = 1 Then
    errorMsg = errorMsg & "時間外申請が8時間以上のとき、休憩が60分以上必要です。"
End If
If err_relation_14 = 1 Then
    errorMsg = errorMsg & "時間代休申請の時刻が妥当でありません。"
End If
If err_relation_26 = 1 Then
    errorMsg = errorMsg & "時間代休は10分単位で取得してください。"
End If
If err_relation_32 = 1 Then
    errorMsg = errorMsg & "交替勤務入力時、時間代休は入力できません。"
End If
If err_relation_15 = 1 Then
    errorMsg = errorMsg & "当月の時間外労働時間内で時間代休は入力可能です。"
End If
If err_relation_22 = 1 Then
    errorMsg = errorMsg & "時間有給の時刻が妥当でありません。"
End If
If err_relation_52 = 1 Then
    errorMsg = errorMsg & "時間有給と勤務、休憩、中抜時刻の関連が妥当でありません。"
End If
If err_relation_23 = 1 Then
    errorMsg = errorMsg & "時間有給を入力するときは、出勤区分に入力が必要です。"
End If
If err_relation_25 = 1 Then
    errorMsg = errorMsg & "時間有休は60分単位で取得してください。"
End If
If err_relation_33 = 1 Then
    errorMsg = errorMsg & "交替勤務入力時、時間有給は入力できません。"
End If
If err_relation_16 = 1 Then
    errorMsg = errorMsg & "深夜割増の時刻が妥当でありません。"
End If
If err_relation_28 = 1 Then
    errorMsg = errorMsg & "深夜割増の時刻は22:00～5:00でなければなりません。"
End If
If err_relation_49 = 1 Then
    errorMsg = errorMsg & "深夜割増入力時、時間外の入力も必要です。"
End If
If err_relation_37 = 1 Then
    errorMsg = errorMsg & "時間外と深夜割増の時刻が重なっていません。"
End If
If err_relation_17 = 1 Then
    errorMsg = errorMsg & "有給残日数を超えて有休を取得しようとしています。"
End If
If err_relation_18 = 1 Then
    errorMsg = errorMsg & "振休残日数を超えて振休を取得しようとしています。"
End If
If err_relation_19 = 1 Then
    errorMsg = errorMsg & "有給休暇より振替休日を優先して取得してください。"
End If
If err_relation_34 = 1 Then
    errorMsg = errorMsg & "交替勤務入力時、宿直は入力できません。"
End If
If err_relation_20 = 1 Then
    errorMsg = errorMsg & "公休日、法定休日でないときに日直が入力されています。"
End If
If err_relation_21 = 1 Then
    errorMsg = errorMsg & "振替出勤 のとき、日直責任者は選択できません。"
End If
If err_relation_35 = 1 Then
    errorMsg = errorMsg & "交替勤務入力時、日直は入力できません。"
End If
If err_relation_38 = 1 Then
    errorMsg = errorMsg & "出勤時、交替勤務区分は必須入力です。"
End If
If err_relation_39 = 1 Then
    errorMsg = errorMsg & "勤務時、勤務時間は必須入力です。"
End If
If err_relation_53 = 1 Then
    errorMsg = errorMsg & "フレックス勤務時間と勤務時間外の入力が重複しています。"
End If
If err_relation_54 = 1 Then
    errorMsg = errorMsg & "公休日に休出は入力できません。"
End If
' 働き方改革対応チェック
If Request.Form("yearlyOvertimeHidden") + mm2Float(overtimeonly_count) > 398 Then
    errorMsg = errorMsg & "当年累積時間外労働が特別条項の上限(398時間)を超えているため入力できません。"
    err_relation_40 = "1"
ElseIf mm2Float(overtimeonly_count) > 60 Then
    errorMsg = errorMsg & "当月の時間外労働が特別条項の上限(60時間)を超えているため入力できません。"
    err_relation_40 = "1"
End If
If Request.Form("yearlyHolidayworkHidden") + holidaywork_count > 42 Then
    errorMsg = errorMsg & "当年度累積休日出勤回数が特別条項の上限(42回)を超えているため入力できません。"
    err_relation_41 = "1"
ElseIf holidaywork_count >= 6 Then
    errorMsg = errorMsg & "当月の休日出勤回数が特別条項の上限(5回)を超えているため入力できません。"
    err_relation_41 = "1"
End If
IF Round((mm2Float(overtime_count) + Request.Form("sumOvertime1")) / 2 , 1) >= 80 Or _
   Round((mm2Float(overtime_count) + Request.Form("sumOvertime1") + Request.Form("sumOvertime2")) / 3 , 1) >= 80 Or _
   Round((mm2Float(overtime_count) + Request.Form("sumOvertime1") + Request.Form("sumOvertime2") + Request.Form("sumOvertime3")) / 4 , 1) >= 80 Or _
   Round((mm2Float(overtime_count) + Request.Form("sumOvertime1") + Request.Form("sumOvertime2") + Request.Form("sumOvertime3") + Request.Form("sumOvertime4")) / 5 , 1) >= 80 Or _
   Round((mm2Float(overtime_count) + Request.Form("sumOvertime1") + Request.Form("sumOvertime2") + Request.Form("sumOvertime3") + Request.Form("sumOvertime4") + Request.Form("sumOvertime5")) / 6 , 1) >= 80 Then
     errorMsg = errorMsg & "複数月平均時間外労働が特別条項の上限(80時間)を超えているため入力できません。"
End If
' エラーメッセージを指定文字数でカットして…で表示する。
If Len(errorMsg)>70 Then
    errorMsg = Left(errorMsg, 70) & "..."
End If
%>
