<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' #######################################
' プログラム仕様書
' #######################################
' 
' ## プログラム概要 ##
' 同一部署職員の勤務状況を確認する画面
' フレックス勤務開始により、出社退社時刻が各自違うため、勤務状況を確認するためのもの。
' 
' ## 注意事項 ##
' 
%>
<!-- #include file="Connections/workdbms.asp" -->
<!-- #include file="inc/RestrictAccess.asp" -->
<%
' 日付の計算
Dim sysDate     'システム日付
Dim dispDate    '表示用日付
Dim dispYear    '表示用年 yyyy
Dim dispMonth   '表示用月 mm
Dim dispLastDay '表示用月の最後の日 dd
Dim nextYmb     '次月移動リンク用
Dim prevYmb     '前月移動リンク用
Dim i           '繰り返し用日付
' stafftbl 読み込み用
Dim Rs_staff
Dim Rs_staff_cmd
' worktbl 読み込み用
Dim Rs_work
Dim Rs_work_cmd

div_size = "3250px"
td_size  = "87px"

' 本日
today    = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2)

If (Request.QueryString("ymb")<>"") Then
    dispDate = DateSerial(Mid(Request.QueryString("ymb"), 1, 4), Mid(Request.QueryString("ymb"), 5, 2), 1)
Else
    'sysDate  = Date
    'dispDate = DateAdd("d", -1, sysDate)    ' 前日日付設定
    dispDate  = Date
End If

dispYear    = Year(dispDate)
dispMonth   = Right("0" & Month(dispDate), 2)
dispLastDay = Day(DateAdd ("d", -1, Year(DateAdd("m", 1, dispDate)) & "/" & Right("0" & Month(DateAdd("m", 1, dispDate)), 2) & "/01"))
nextYmb     = Year(DateAdd("m",  1, dispDate)) & Right("0" & Month(DateAdd("m",  1, dispDate)), 2)
prevYmb     = Year(DateAdd("m", -1, dispDate)) & Right("0" & Month(DateAdd("m", -1, dispDate)), 2)

' stafftblより、表示スタッフ一覧を取得
Set Rs_staff_cmd = Server.CreateObject ("ADODB.Command")
Rs_staff_cmd.ActiveConnection = MM_workdbms_STRING
Rs_staff_cmd.CommandText = "SELECT s.personalcode ,s.staffname " & _
    ",s.orgcode , s.gradecode ,s.workshift ,s.is_operator " & _
    "FROM orgtbl o RIGHT OUTER JOIN stafftbl s " & _
    "ON o.orgcode = s.orgcode " & _
    "WHERE s.is_input = '1' AND s.is_enable = '1' " & _
    "AND o.personalcode = ?  AND o.manageclass = '2' " & _
    "UNION " & _
    "SELECT s.personalcode ,s.staffname ,s.orgcode " & _
    ",s.gradecode ,s.workshift ,s.is_operator " & _
    "FROM stafftbl o RIGHT OUTER JOIN stafftbl s " & _
    "ON o.orgcode = s.orgcode " & _
    "WHERE s.is_input = '1' AND s.is_enable = '1' " & _
    "AND o.personalcode = ? " & _
    "ORDER BY s.orgcode, s.gradecode DESC, s.personalcode"
Rs_staff_cmd.Prepared = true
Rs_staff_cmd.Parameters.Append Rs_staff_cmd.CreateParameter("param1", 200, 1, -1, Session("MM_Username") )
Rs_staff_cmd.Parameters.Append Rs_staff_cmd.CreateParameter("param1", 200, 1, -1, Session("MM_Username") )
Set Rs_staff = Rs_staff_cmd.Execute

' worktblより、表示スタッフ全員分の上長チェックを日付順に取得
Set Rs_work_cmd = Server.CreateObject ("ADODB.Command")
Rs_work_cmd.ActiveConnection = MM_workdbms_STRING
Rs_work_cmd.CommandText = "SELECT w.personalcode " & _
    ",w.workingdate " & _
    ",w.is_approval " & _
    ",w.morningholiday " & _
    ",w.morningwork " & _
    ",w.afternoonholiday " & _
    ",w.afternoonwork " & _
    ",w.work_begin " & _
    ",w.work_end " & _
    ",w.break_begin1 " & _
    ",w.break_end1 " & _
    ",w.break_begin2 " & _
    ",w.break_end2 " & _
    ",vacationtime_begin " & _
    ",vacationtime_end " & _
    ",w.is_error " & _
    "FROM stafftbl s " & _
    "RIGHT OUTER JOIN worktbl w ON s.personalcode = w.personalcode " & _
    "LEFT OUTER JOIN stafftbl o ON o.orgcode = s.orgcode " & _
    "WHERE w.workingdate LIKE ? AND s.is_input = '1' " & _
    "AND s.is_enable = '1' " & _
    "AND o.personalcode = ? " & _
    "ORDER BY s.orgcode, s.gradecode DESC, w.personalcode, w.workingdate"
Rs_work_cmd.Prepared = true
Rs_work_cmd.Parameters.Append Rs_work_cmd.CreateParameter("param1", 200, 1, -1, dispyear & dispmonth & "%")
Rs_work_cmd.Parameters.Append Rs_work_cmd.CreateParameter("param2", 200, 1, -1, Session("MM_Username"))
Set Rs_work = Rs_work_cmd.Execute
%>

<!DOCTYPE HTML>
<html lang="ja">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <title>勤務表管理システム</title>
    <link href="css/default.css" rel="stylesheet" type="text/css">
</head>

<body>
    <div id="container">
        <!-- #include file="inc/header.source" -->
        <div id="contents">
            <p>
                勤務状況確認
                <a href="workstatus.asp?ymb=<%=prevYmb%>">&lt;&lt;</a>&nbsp;
                <a href="workstatus.asp"><%=dispYear%>年<%=dispMonth%>月</a>&nbsp;
                <a href="workstatus.asp?ymb=<%=nextYmb%>">&gt;&gt;</a>
            </p>
            
            <div id="tablediv" class="clear" style="width:<%=div_size%>">
                <table class="data">
                    <tr>
                        <th width="150px;" scope="col">氏名</th>
                        <th width="80px;" scope="col">勤務</th>
                        <%
                        ' 表示月の日付一覧作成
                        For i = 1 To dispLastDay
                            Response.write "<th width='" & td_size & "' scope='col'>"
                            Response.write Right("0" & i, 2)
                            Response.write "</th>"
                            ' 16日目の前に氏名欄を追加
                            If i = 15 Then
                                Response.Write "<th width='150px;' scope='col'>氏名</th>"
                            End If
                        Next
                        %>
                    </tr>
                </table>
                <div id="tbody"  class="tBody" style="width:<%=div_size%>;height:100%;">
                    <table class="data">
                        <%
                        ' 職員とチェック状況の表示
                        If Not Rs_staff.EOF Or Not Rs_staff.BOF Then
                            While (NOT Rs_staff.EOF)
                        %>
                                <tr style="height:45px;">
                                <th width="150px;" nowrap class="permanent" scope="row">
                                    <% If Session("MM_is_superior") = "1" Then %>
                                    <a href="inputwork.asp?p=<%=Rs_staff.Fields.Item("personalcode")%>&ymb=<%=dispYear & Right("0" & dispMonth, 2)%>&c=1">
                                    <% End If %>
                                       <%=RTrim(Rs_staff.Fields.Item("staffname"))%>
                                    <% If Session("MM_is_superior") = "1" Then %>
                                    </a>
                                    <% End If %>
                                </th>
                                <td width="80px;" nowrap style="text-align:center;">
                                    <%
                                    workshift = ""
                                    If Rs_staff.Fields.Item("workshift") = "0" Then
                                        If Rs_staff.Fields.Item("is_operator") = "1" Then
                                            workshift = "オペレータ"
                                        Else
                                            workshift = "通常勤務"
                                        End If
                                    ElseIf Rs_staff.Fields.Item("workshift") = "1" Then
                                        workshift = "コミュ全日"
                                    ElseIf Rs_staff.Fields.Item("workshift") = "2" Then
                                        workshift = "コミュ午前"
                                    ElseIf Rs_staff.Fields.Item("workshift") = "3" Then
                                        workshift = "コミュ午後"
                                    ElseIf Rs_staff.Fields.Item("workshift") = "9" Then
                                        workshift = "フレックス"
                                    End If
                                    response.write(workshift)
                                    %>
                                </td>
                                <%
                                ' 日付部分作成
                                For i = 1 To dispLastDay
                                    loopday    = dispYear & Right("0" & dispMonth, 2) & Right("0" & i, 2)
                                    If today = loopday Then
                                        ' 当日は背景色変更
                                        weekClass = "thisday"
                                    Else
                                        weekNameKanji = WeekdayName(Weekday(DateSerial(dispYear, dispMonth, i)), true)
                                        If weekNameKanji = "日" Then
                                            weekClass = "sunday"
                                        ElseIf weekNameKanji = "土" Then
                                            weekClass = "saturday"
                                        Else
                                            weekClass = ""
                                        End If
                                        
                                        ' フレックス勤務者で 08:00-17:10の勤務時間で昼休みを除く時間で勤務が無い場合、背景色に色付けする
    '                                    If Not Rs_work.EOF And Rs_staff.Fields.Item("workshift") = "9" Then
    '                                        If (Len(Trim(Rs_work.Fields.Item("work_begin"))) > 0 And _
    '                                            Trim(Rs_work.Fields.Item("work_begin"))  > "0830") Then
    '                                            weekClass = "flexcheck"
    '                                        Else If (Len(Trim(Rs_work.Fields.Item("work_end"))) > 0 And _
    '                                                 Trim(Rs_work.Fields.Item("work_end"))  < "1710" And _
    '                                                 Trim(Rs_work.Fields.Item("work_end"))  > "0700") Then
    '                                            weekClass = "flexcheck"
    '                                        End If
    '                                    End If
                                    End If

                                    ' 16日目の前に氏名欄を追加
                                    If i = 16 Then
                                    %>
                                        <th width="150px;" nowrap class="permanent" scope="row">
                                            <% If Session("MM_is_superior") = "1" Then %>
                                            <a href="inputwork.asp?p=<%=Rs_staff.Fields.Item("personalcode")%>&ymb=<%=dispYear & Right("0" & dispMonth, 2)%>&c=1">
                                            <% End If %>
                                               <%=RTrim(Rs_staff.Fields.Item("staffname"))%>
                                            <% If Session("MM_is_superior") = "1" Then %>
                                            </a>
                                            <% End If %>
                                        </th>
                                    <%
                                    End If

                                    Response.Write("<td width='" & td_size & "' align='center' class='" & weekClass & "'>")
                                    If Not Rs_work.EOF  Then
                                        If Rs_work.Fields.Item("workingdate" ) = dispYear & dispMonth & Right("0" & i, 2) And _
                                           Rs_work.Fields.Item("personalcode") = Rs_staff.Fields.Item("personalcode")     Then
                                            txt = ""
                                            If Rs_staff.Fields.Item("workshift") = "9" Then
                                                ' フレックス勤務
                                                If Rs_work.Fields.Item("morningwork") = "" And Rs_work.Fields.Item("afternoonwork") = "" Then
                                                    txt = txt & "&nbsp;"
                                                Else
                                                    ' 出社時間
                                                    If Len(Trim(Rs_work.Fields.Item("work_begin"))) > 0 Then
                                                        txt = txt & "<b>" & editTime(Rs_work.Fields.Item("work_begin")) & "</b>&nbsp;"
                                                    Else
                                                        txt = txt & "&nbsp;"
                                                    End If
                                                    ' 中抜け開始時間1
                                                    If Len(Trim(Rs_work.Fields.Item("break_begin1"))) > 0 Then
                                                        txt = txt & editTime(Rs_work.Fields.Item("break_begin1")) & "&nbsp;"
                                                    End If
                                                    ' 中抜け開始時間2
                                                    If Len(Trim(Rs_work.Fields.Item("break_begin2"))) > 0 Then
                                                        ' 中抜け2入力有
                                                        If Len(Trim(Rs_work.Fields.Item("vacationtime_begin"))) > 0 Then
                                                            ' 時間有給入力有
                                                            If Trim(Rs_work.Fields.Item("break_begin2")) < _
                                                               Trim(Rs_work.Fields.Item("vacationtime_begin")) Then
                                                                ' 中抜け2(開始時間) < 時間有給(開始時間)
                                                                txt = txt & editTime(Rs_work.Fields.Item("break_begin2")) & "&nbsp;"
                                                            Else
                                                                ' 中抜け2(開始時間) > 時間有給(開始時間)
                                                                txt = txt & editTime(Rs_work.Fields.Item("vacationtime_begin")) & "&nbsp;"
                                                            End If
                                                        Else
                                                            ' 時間有給入力無し
                                                            txt = txt & editTime(Rs_work.Fields.Item("break_begin2")) & "&nbsp;"
                                                        End If
                                                    Else
                                                        ' 中抜け2入力無し
                                                        If Len(Trim(Rs_work.Fields.Item("vacationtime_begin"))) > 0 Then
                                                            ' 時間有給入力有
                                                            txt = txt & editTime(Rs_work.Fields.Item("vacationtime_begin")) & "&nbsp;"
                                                        End If
                                                    End If
                                                    txt = txt & "<br>"
                                                    ' 退社時間
                                                    If Len(Trim(Rs_work.Fields.Item("work_end"))) > 0 Then
                                                        txt = txt & "<b>" & editTime(Rs_work.Fields.Item("work_end")) & "</b>&nbsp;"
                                                    Else
                                                        txt = txt & "&nbsp;"
                                                    End If
                                                    ' 中抜け終了時間1
                                                    If Len(Trim(Rs_work.Fields.Item("break_end1"))) > 0 Then
                                                        txt = txt & editTime(Rs_work.Fields.Item("break_end1")) & "&nbsp;"
                                                    End If
                                                    ' 中抜け終了時間2
                                                    If Len(Trim(Rs_work.Fields.Item("break_end2"))) > 0 Then
                                                        ' 中抜け2入力有
                                                        If Len(Trim(Rs_work.Fields.Item("vacationtime_end"))) > 0 Then
                                                            ' 時間有給入力有
                                                            If Trim(Rs_work.Fields.Item("break_end2")) < _
                                                               Trim(Rs_work.Fields.Item("vacationtime_end")) Then
                                                                ' 中抜け2(終了時間) < 時間有給(終了時間)
                                                                txt = txt & editTime(Rs_work.Fields.Item("vacationtime_end")) & "&nbsp;"
                                                            Else
                                                                ' 中抜け2(終了時間) > 時間有給(終了時間)
                                                                txt = txt & editTime(Rs_work.Fields.Item("break_end2")) & "&nbsp;"
                                                            End If
                                                        Else
                                                            ' 時間有給入力無し
                                                            txt = txt & editTime(Rs_work.Fields.Item("break_end2")) & "&nbsp;"
                                                        End If
                                                    Else
                                                        ' 中抜け2入力無し
                                                        If Len(Trim(Rs_work.Fields.Item("vacationtime_end"))) > 0 Then
                                                            ' 時間有給入力有
                                                            txt = txt & editTime(Rs_work.Fields.Item("vacationtime_end")) & "&nbsp;"
                                                        End If
                                                    End If
                                                    txt = "<div style='text-align:left;margin-left:2px;font-size:8px;'>" & txt & "</div>"
                                                End If
                                            ElseIf Rs_staff.Fields.Item("workshift") = "0" And Rs_staff.Fields.Item("is_operator") = "0" Then
                                                ' 一般勤務
                                                If Rs_work.Fields.Item("morningwork") = "0"  Then
                                                    If Rs_work.Fields.Item("afternoonwork") = "0" Then
                                                        If Rs_work.Fields.Item("morningholiday") <> "0" And _
                                                           Rs_work.Fields.Item("afternoonholiday") <> "0" Then
                                                            txt = "－"
                                                        Else
                                                            txt = "..."
                                                        End If
                                                    Else
                                                        txt = "PM"
                                                    End If
                                                Else
                                                    If Rs_work.Fields.Item("afternoonwork") = "0" Then
                                                        txt = "AM"
                                                    Else
                                                        txt = "出"
                                                    End If
                                                End If
                                            Else
                                                ' その他
                                                txt = ""
                                            End If
                                            Response.Write(txt)
                                            Rs_work.MoveNext()
                                        Else
                                            ' worktblのデータと日付セルの情報が一致しない場合
                                            Response.write("&nbsp;")
                                        End If
                                    Else
                                        ' worktblにデータが存在しない
                                        Response.write("&nbsp;")
                                    End If
                                    Response.Write("</td>")
                                Next
                                Response.write "</tr>"
                                Rs_staff.MoveNext()
                            Wend
                        End If
                        %>
                        </tr>
                    </table>
                </div>
            </div>
        </div>
    </div>
</body>
<script type="text/javascript">
    // ウィンドウサイズから div サイズを設定する関数
    function setDivSize(){
        var size_h;
        size_h = document.body.clientHeight;
        size_h = size_h - 130;
        document.getElementById('tablediv').style.height = size_h + "px";
    }
    //読み込み時にサイズを表示
    setDivSize();
</script>
</html>
<%
Rs_staff.Close()
Set Rs_staff = Nothing
Rs_work.Close()
Set Rs_work = Nothing
%>
<!-- #include file="inc/util.asp" -->
