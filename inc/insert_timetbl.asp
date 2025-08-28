<%
' -----------------------------------------------------------------------------
' 出退勤時刻登録処理
' -----------------------------------------------------------------------------
Dim Rs_timetbl
Dim Rs_timetbl_cmd
Dim Rs_timetbl_numRows
If (CStr(Request.Form("button_type")) = "in"   ) Or (CStr(Request.Form("button_type")) = "leave") Then
    nowTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2)
    ' タイムテーブル timetbl 読込
    Set Rs_timetbl_cmd = Server.CreateObject ("ADODB.Command")
    Rs_timetbl_cmd.ActiveConnection = MM_workdbms_STRING
    Rs_timetbl_cmd.CommandText = "SELECT * FROM timetbl WHERE personalcode = ? AND workingdate = ?"
    Rs_timetbl_cmd.Prepared = true
    Rs_timetbl_cmd.Parameters.Append Rs_timetbl_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
    Rs_timetbl_cmd.Parameters.Append Rs_timetbl_cmd.CreateParameter("param2", 200, 1, 8, today)
    Set Rs_timetbl = Rs_timetbl_cmd.Execute
    Rs_timetbl_numRows = 0
    If Rs_timetbl.EOF Then
        ' INSERT
        Set MM_editCmd = Server.CreateObject ("ADODB.Command")
        MM_editCmd.ActiveConnection = MM_workdbms_STRING
        MM_editCmd.CommandText = "INSERT INTO timetbl VALUES(DEFAULT, ?, ?, '', ?, '', '', '', '', '', ?)"
        MM_editCmd.Prepared = true
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  5, Session("MM_Username"))
        MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  8, today)
        If (CStr(Request.Form("button_type")) = "in" ) Then
            MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  4, nowTime)
            MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  4, "")
        Else
            MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  4, "")
            MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  4, nowTime)
        End If
        MM_editCmd.Execute
        MM_editCmd.ActiveConnection.Close
    Else
        If (CStr(Request.Form("button_type")) = "in" ) Then
            If (Len(Trim(Rs_timetbl.Fields.Item("cometime").Value)) = 0) Then
                ' UPDATE 出社時間
                Set MM_editCmd = Server.CreateObject ("ADODB.Command")
                MM_editCmd.ActiveConnection = MM_workdbms_STRING
                MM_editCmd.CommandText = "UPDATE timetbl SET cometime = ? WHERE id = ?"
                MM_editCmd.Prepared = true
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  4, nowTime)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, Rs_timetbl.Fields.Item("id").Value)
                MM_editCmd.Execute
                MM_editCmd.ActiveConnection.Close
            End If
        Else
            If (Len(Trim(Rs_timetbl.Fields.Item("leavetime").Value)) = 0) Then
                ' UPDATE 退社時間
                Set MM_editCmd = Server.CreateObject ("ADODB.Command")
                MM_editCmd.ActiveConnection = MM_workdbms_STRING
                MM_editCmd.CommandText = "UPDATE timetbl SET leavetime = ? WHERE id = ?"
                MM_editCmd.Prepared = true
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  4, nowTime)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, Rs_timetbl.Fields.Item("id").Value)
                MM_editCmd.Execute
                MM_editCmd.ActiveConnection.Close
            End If
        End If
    End If
    Rs_timetbl.Close()
End If
%>
