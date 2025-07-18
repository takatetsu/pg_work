<%
' ------------------------------------------------------------------------------
' worktbl 上長チェック 更新処理
' ------------------------------------------------------------------------------
' ------------------------------------------------------------------------------
' 入力チェック
' ------------------------------------------------------------------------------
' 時刻チェック用フラグ
err_beginTime  = 0
err_endTime    = 0
err_returnTime = 0
err_outTime    = 0
err_iserror    = 0

If Request.QueryString("p")<>"" Then
    For i = 1 To Request.Form("everyday").count Step 1
        j = Right(Request.Form("everyday")(i), 2)
        ' ----------------------------------------------------------------------
        ' 時刻チェック
        ' ----------------------------------------------------------------------
        ' タイムカード出社
        If Not (legalTime(Request.Form("beginTime" & Request.Form("everyday")(i)))) Then
            err_beginTime       = 1
            style_begintime(j)  = "errorcolor"
        End If
        ' タイムカード戻り
        If Not (legalTime(Request.Form("returnTime" & Request.Form("everyday")(i)))) Then
            err_returnTime      = 1
            style_returntime(j) = "errorcolor"
        End If
        ' タイムカード退社
        If Not (legalTime(Request.Form("endTime" & Request.Form("everyday")(i)))) Then
            err_endTime         = 1
            style_endtime(j)    = "errorcolor"
        End If
        ' タイムカード外出
        If Not (legalTime(Request.Form("outTime" & Request.Form("everyday")(i)))) Then
            err_outTime         = 1
            style_outtime(j)    = "errorcolor"
        End If
        ' 上長チェックオンのとき、エラーが0以外だとエラーを設定する。
        If (Request.Form("is_approval" & Request.Form("everyday")(i)) = "on") Then
            If Request.Form("is_error" & Request.Form("everyday")(i)) <> "0" And _
               Request.Form("is_error" & Request.Form("everyday")(i)) <> ""  Then
                err_iserror     = 1
            End If
        End If
    Next
End If

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
If err_returnTime           = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "戻り時刻"
    tempIdx = tempIdx       + 1
End If
If err_outTime              = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "外出時刻"
    tempIdx = tempIdx       + 1
End If
If err_endTime              = 1 Then
    ReDim Preserve tempArray(tempIdx+1)
    tempArray(tempIdx)      = "退勤時刻"
    tempIdx = tempIdx       + 1
End If
If (tempIdx > 0) Then
    tempMessage = "次の項目が時刻として妥当でありません["
    For arrayCounter = 0 To tempIdx - 1 Step 1
        tempMessage = tempMessage + tempArray(arrayCounter) & ","
    Next
    errorMsg = Left(tempMessage, Len(tempMessage)-1) & "]。" & errorMsg
End If
If err_iserror              = 1 Then
    errorMsg = "エラーが立っているため上長チェックを更新できません。" & errorMsg
End If
' エラーメッセージを指定文字数でカットして…で表示する。
If Len(errorMsg)>70 Then
    errorMsg = Left(errorMsg, 70) & "..."
End If

If (errorMsg = "") Then
    ' エラーメッセージが空白の時、更新処理を行う。
    ' --------------------------------------------------------------------------
    ' 上長チェックによる更新
    ' --------------------------------------------------------------------------
    If (Request.QueryString("p")<>"" And Session("MM_is_superior")="1") Then
        For i = 1 To Request.Form("everyday").count Step 1
            ' ------------------------------------------------------------------
            ' timetbl に対しての更新処理
            ' ------------------------------------------------------------------
            If (Request.Form("timetbl_id")(i) = "") Then
                If (Trim(Request.Form("beginTime" & Request.Form("everyday")(i))) <> ""  Or _
                    Trim(Request.Form("endTime"   & Request.Form("everyday")(i))) <> "") Then
                    ' INSERT
                    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
                    MM_editCmd.ActiveConnection = MM_workdbms_STRING
                    MM_editCmd.CommandText = "INSERT INTO dbo.timetbl VALUES" & _
                                             "(DEFAULT, ?, ?, '', ?, '', '', '', '', '', ?)"
                    MM_editCmd.Prepared = true
                    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  5, target_personalcode)
                    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  8, _
                            Request.Form("everyday")(i))
                    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  4, _
                            Left(editTime(Request.Form("beginTime"  & Request.Form("everyday")(i))), 2) & _
                            Right(editTime(Request.Form("beginTime" & Request.Form("everyday")(i))), 2))
                    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  4, _
                            Left(editTime(Request.Form("endTime"   & Request.Form("everyday")(i))), 2) & _
                            Right(editTime(Request.Form("endTime"  & Request.Form("everyday")(i))), 2))
                    MM_editCmd.Execute
                    MM_editCmd.ActiveConnection.Close
                End If
            Else
                ' UPDATE
                Set MM_editCmd = Server.CreateObject ("ADODB.Command")
                MM_editCmd.ActiveConnection = MM_workdbms_STRING
                MM_editCmd.CommandText = "UPDATE dbo.timetbl SET cometime  = ?, " & _
                                                                "leavetime = ?  " & _
                                                                "WHERE id  = ?"
                MM_editCmd.Prepared = true
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  4, _
                        Left(editTime(Request.Form("beginTime"  & Request.Form("everyday")(i))), 2) & _
                        Right(editTime(Request.Form("beginTime" & Request.Form("everyday")(i))), 2))
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  4, _
                        Left(editTime(Request.Form("endTime"  & Request.Form("everyday")(i))), 2) & _
                        Right(editTime(Request.Form("endTime" & Request.Form("everyday")(i))), 2))
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, _
                        Request.Form("timetbl_id")(i))
                MM_editCmd.Execute
                MM_editCmd.ActiveConnection.Close
            End If
        Next

        For i = 1 To Request.Form("approval_ymd").count Step 1
            ' ------------------------------------------------------------------
            ' worktbl に対しての更新処理
            ' ------------------------------------------------------------------
            If (Request.Form("worktbl_approval_id")(i) <> "") Then
                Set MM_editCmd = Server.CreateObject ("ADODB.Command")
                MM_editCmd.ActiveConnection = MM_workdbms_STRING
                ' UPDATE
                MM_editCmd.CommandText = "UPDATE dbo.worktbl SET is_approval = ? " & _
                                         ", is_error = ? , memo = ? " & _
                                         "WHERE id = ? AND CONVERT(int,updatetime) = ?"
                MM_editCmd.Prepared = true
                If (Request.Form("is_approval" & Request.Form("approval_ymd")(i))="on") Then
                    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,, 1, "1")
                Else
                    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,, 1, "0")
                End If
                ' 労働時間適正化エラーフラグ
                If (Request.Form("is_error" & Request.Form("approval_ymd")(i))="0" Or _
                    Request.Form("is_error" & Request.Form("approval_ymd")(i))="") Then
                    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,, 1, "0")
                Else
                    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,, 1, "1")
                End If
                
                ' メモの正しいインデックスを取得
                ' approval_ymdから対象日を取得し、everydayでの位置を特定
                target_date = Request.Form("approval_ymd")(i)
                memo_index = 0
                For j = 1 To Request.Form("everyday").count Step 1
                    If Request.Form("everyday")(j) = target_date Then
                        memo_index = j
                        Exit For
                    End If
                Next
                
                ' メモ（正しいインデックスで取得）
                If memo_index > 0 Then
                    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,100, Request.Form("memo")(memo_index))
                Else
                    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,100, "")
                End If
                
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, _
                                            Request.Form("worktbl_approval_id")(i))
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, _
                                            Request.Form("approval_worktbl_updatetime")(i))
                MM_editCmd.Execute
                MM_editCmd.ActiveConnection.Close
            End If
        Next
        Response.Redirect("checkList.asp?ymb=" & ymb)
    End If
End If
%>