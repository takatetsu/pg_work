<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="Connections/workdbms.asp" -->
<!-- #include file="inc/RestrictAccess.asp" -->
<!-- #include file="inc/properTimeCheck.asp" -->
<!-- #include file="inc/inputCommon1.asp" -->
<%
' -----------------------------------------------------------------------------
' 勤務表入力画面(個人入力用)
' -----------------------------------------------------------------------------
%>
<!-- #include file="inc/select_stafftbl.asp" -->
<!-- #include file="inc/insert_timetbl.asp" -->
<%
' -----------------------------------------------------------------------------
' Form入力項目・ボタン有効無効設定処理
' -----------------------------------------------------------------------------
If screen = 0 Then
    ' 入力画面設定
    ' 未来日付のとき、翌月末までは入力可能 disabled 設定
    If (inputLimitYmb < ymb)Then
        text_disabled = "disabled"
    End If
    ' 給与担当者の処理済年月以前の時は更新ボタン無効化
    If ymb <= proceseed_ymb Or ymb > inputLimitYmb Or screen = 1 Then
        button_submit_disable = "disabled"
    End If
    ' 出退勤ボタン有効無効判定処理
    Set Rs_timetbl_cmd = Server.CreateObject ("ADODB.Command")
    Rs_timetbl_cmd.ActiveConnection = MM_workdbms_STRING
    Rs_timetbl_cmd.CommandText = "SELECT * FROM timetbl WHERE personalcode = ? AND workingdate = ?"
    Rs_timetbl_cmd.Prepared = true
    Rs_timetbl_cmd.Parameters.Append Rs_timetbl_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
    Rs_timetbl_cmd.Parameters.Append Rs_timetbl_cmd.CreateParameter("param2", 200, 1, 8, today)
    Set Rs_timetbl = Rs_timetbl_cmd.Execute
    Rs_timetbl_numRows = 0
    If Not Rs_timetbl.EOF Then
        If (Len(Trim(Rs_timetbl.Fields.Item("cometime").Value)) = 0) Then
            button_come_disabled  = ""
        Else
            button_come_disabled  = "disabled"
        End If
        If (Len(Trim(Rs_timetbl.Fields.Item("leavetime").Value)) = 0) Then
            button_leave_disabled = ""
        Else
            button_leave_disabled = "disabled"
        End If
    End If
    Rs_timetbl.Close()
ElseIf screen = 1 Then
    ' 労務担当者用確認画面設定
    button_come_disabled    = "disabled"
    button_leave_disabled   = "disabled"
    button_submit_disable   = "disabled"
    text_disabled           = "disabled"
ElseIf screen = 2 Then
    ' 上長チェック画面設定
    button_come_disabled    = "disabled"
    button_leave_disabled   = "disabled"
    text_disabled           = "disabled"
    ' 給与担当者の処理済年月以前の時は更新ボタン無効化
    If ymb <= proceseed_ymb Or ymb > inputLimitYmb Or screen = 1 Then
        button_submit_disable = "disabled"
    End If
End If
' -----------------------------------------------------------------------------
' ■ テーブル更新処理
' -----------------------------------------------------------------------------
Dim tmpArray()
If (CStr(Request("MM_update")) = "form1") And screen = "2" Then
%>
    <!-- #include file="inc/update_worktbl_is_approval.asp" -->
<%
ElseIf (CStr(Request("MM_update")) = "form1") And screen = "0" Then
    ' ---------------------------------------------------------------------
    ' ■ 勤務表入力画面 更新処理
    ' ---------------------------------------------------------------------
%>
    <!-- #include file="inc/inputworkCheck.asp" -->
<%
    ' エラーメッセージが空白の時、更新処理を行う。
    If (errorMsg = "") Then
        ' ---------------------------------------------------------------------
        ' 個人勤務表入力による更新
        ' ---------------------------------------------------------------------
        ' stafftbl の締め日付確認
        Set Rs_stafftbl_cmd = Server.CreateObject ("ADODB.Command")
        Rs_stafftbl_cmd.ActiveConnection = MM_workdbms_STRING
        Rs_stafftbl_cmd.CommandText = "SELECT processed_ymb FROM stafftbl " & _
                                      "WHERE personalcode = '" & target_personalcode & "'"
        Set Rs_stafftbl = Rs_stafftbl_cmd.Execute
        processed_ymb = Rs_stafftbl.Fields.Item("processed_ymb").Value
        Rs_stafftbl.Close()
        Set Rs_stafftbl = Nothing
        ' 給与担当者の処理済年月以前の時は更新不可
        If ymb <= proceseed_ymb Then
            errorMsg = "給与担当者処理済のため更新できません"
        Else
            For i = 1 To Request.Form("ymd").count Step 1
%>
                <!-- #include file="inc/upsert_worktbl.asp" -->
<%
            Next
%>
            <!-- #include file="inc/upsert_dutyrostertbl.asp" -->
<%
            ' 正常終了後の画面遷移
            'Response.Redirect("complete.asp")
        End If
    End If
End If

' フレックス勤務の場合、日曜日始まりの1週間内で法定休日が1日無いとエラーとする。
' 1週間内に2日以上法定休日がある場合もエラーとする。
If workshift = "9" And errorMsg = "" Then
    Set Rs_statutory_cmd = Server.CreateObject ("ADODB.Command")
    Rs_statutory_cmd.ActiveConnection = MM_workdbms_STRING
    Rs_statutory_cmd.CommandText = "SELECT begindate, enddate, COUNT(*) AS datacount, SUM(holiday) AS houtei " & _
        "FROM ( SELECT DISTINCT * FROM (SELECT w2.workingdate, w.begindate, w.enddate, " & _
        "CASE w2.morningholiday WHEN 'A' THEN 1 ELSE 0 END AS holiday, old_workshift_last_ymb FROM " & _
        "(SELECT worktbl.personalcode, worktbl.workingdate, worktbl.morningholiday, " & _
        "stafftbl.old_workshift_last_ymb, " & _
        "to_char(to_date(worktbl.workingdate, 'YYYYMMDD') - (extract(DOW FROM to_date(worktbl.workingdate, 'YYYYMMDD'))::text || ' days')::interval + INTERVAL '1 day', 'YYYYMMDD') AS begindate, " & _
        "to_char(to_date(worktbl.workingdate, 'YYYYMMDD') + ((7 - extract(DOW FROM to_date(worktbl.workingdate, 'YYYYMMDD')))::text || ' days')::interval, 'YYYYMMDD') AS enddate " & _
        "FROM worktbl LEFT JOIN stafftbl ON worktbl.personalcode = stafftbl.personalcode " & _
        "WHERE worktbl.personalcode = ? AND (worktbl.workingdate BETWEEN ? AND ? OR " & _
        "to_date(?, 'YYYYMMDD') - (extract(DOW FROM to_date(?, 'YYYYMMDD'))::text || ' days')::interval <= to_date(worktbl.workingdate, 'YYYYMMDD') AND " & _
        "to_date(worktbl.workingdate, 'YYYYMMDD') < to_date(?, 'YYYYMMDD')) AND worktbl.workingdate >= stafftbl.old_workshift_last_ymb) w " & _
        "LEFT JOIN worktbl w2 ON w.personalcode=w2.personalcode AND w.begindate <= w2.workingdate AND " & _
        "w.enddate >= w2.workingdate) x where begindate > old_workshift_last_ymb || '31') y  GROUP BY begindate, enddate HAVING COUNT(*) = 7 AND SUM(holiday) <> 1"
    Rs_statutory_cmd.Prepared = true
    Rs_statutory_cmd.Parameters.Append Rs_statutory_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
    Rs_statutory_cmd.Parameters.Append Rs_statutory_cmd.CreateParameter("param2", 200, 1, 8, ymb & "01")
    Rs_statutory_cmd.Parameters.Append Rs_statutory_cmd.CreateParameter("param3", 200, 1, 8, ymb & "31")
    Rs_statutory_cmd.Parameters.Append Rs_statutory_cmd.CreateParameter("param4", 200, 1, 8, ymb & "01")
    Rs_statutory_cmd.Parameters.Append Rs_statutory_cmd.CreateParameter("param5", 200, 1, 8, ymb & "01")
    Rs_statutory_cmd.Parameters.Append Rs_statutory_cmd.CreateParameter("param6", 200, 1, 8, ymb & "01")
    Set Rs_statutory = Rs_statutory_cmd.Execute
    Rs_statutory_numRows = 0
    If Not Rs_statutory.EOF Then
        errorMsg = errorMsg & "法定休日が入力されていない、または2日以上設定されている週があります。"
    End If
    Rs_statutory.Close()
End If

' 表示エリアで使用する変数のゼロクリア
realworkmin = 0
%>
<!-- #include file="inc/view_init.asp" -->
<!DOCTYPE HTML>
<html lang="ja">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <title>勤務表管理システム</title>
    <link href="css/default.css?<%=Time%>" rel="stylesheet" type="text/css">
</head>

<body>
<div id="container">
<!-- #include file="inc/header.source" -->
<div id="contents">
    <%
    url_param = ""
    If Request.QueryString("c")<>"" Then
        url_param = "&c=" & Request.QueryString("c")
    End If
    If Request.QueryString("s")<>"" Then
        url_param = "&s=" & Request.QueryString("s")
    End If
    If (Request.QueryString("p")<>"") Then
        url_param = url_param & "&p=" & target_personalcode
    End If
    %>
    <a href="inputwork.asp?ymb=<%=lastYmb%><% If (url_param<>"") Then Response.Write(url_param) End If%>">&lt;&lt;</a>&nbsp;
    <%=left(ymb, 4)%>年<%=right(ymb, 2)%>月分&nbsp;
    <a href="inputwork.asp?ymb=<%=nextYmb%><% If (url_param<>"") Then Response.Write(url_param) End If%>">&gt;&gt;</a>&nbsp;

    <div style="width:1300px; position:absolute; top:42px">
        <div style="float:left; width:850px;">
            <table class="data">
                <tr>
                    <th width="60px">個人CD</th>
                    <td width="60px" class="disabled" align="center"><%=target_personalcode%></th>
                    <th width="60px">氏名</th>
                    <td width="185px" class="disabled"><%=name%></th>
                    <th width="60px">所属</th>
                    <td width="339px" class="disabled"><%=orgname%></th>
                </tr>
            </table>
        </div>
        <div style="float:left;width:70px;">
            <%If (Request.QueryString("p")="") Then%>
                <form name="form2" method="post" action="">
                    <input type="submit" name="button" id="button_come" value="出社" <%=button_come_disabled%>/>
                    <input type="hidden" name="button_type" value="in">
                </form>
            <%End If%>
        </div>
        <div style="float:left;width:70px;">
            <%If (Request.QueryString("p")="") Then%>
                <form name="form3" method="post" action="">
                    <input type="submit" name="button" id="button_leave" value="退社" <%=button_leave_disabled%>/>
                    <input type="hidden" name="button_type" value="leave">
                </form>
            <%End If%>
        </div>
        <a href="36協定書.pdf" target="_blank"><b>36協定書.pdf</b></a>&nbsp;/&nbsp;<a href="フレックスタイム制に関する協定.pdf" target="_blank"><b>フレックスタイム制に関する協定.pdf</b></a>
    </div>
    <div style="float:clear;"></div>
    <form name="form1" method="post" action="" class="clear">
    <div id="tablediv" class="clear" style="padding-top:90px; width:1800px; position:absolute; top:60px;">
        <br>
    <div class="tHeader">
        <table class="data">
            <tr>
                <th nowrap rowspan="2" width="25px;">日</th>
                <th nowrap rowspan="2" width="25px;">曜<br>日</th>
                <th nowrap rowspan="2" width="25px;">確<br>認</th>
                <th nowrap colspan="2">
                    <a href="timecard.asp?p=<%=target_personalcode%>&ymb=<%=ymb%>" target="_blank">
                    <font color="#ffffff">タイムカード</font></a>
                </th>
                <% If is_operator Then ' オペレータのとき、交替勤務の入力を出退勤時刻の右に表示%>
                    <th nowrap rowspan="2" width="70px;">交替勤務</th>
                <% End If %>
                <th nowrap colspan="2">休日・出勤区分</th>
                <% If Not is_operator Then %>
                    <th nowrap rowspan="2" width="50px;">勤務時間</th>
                <% End If %>
                <% If workshift = "9" Then %>
                    <th nowrap colspan="3">勤務時間</th>
                <% End If %>
                <th nowrap rowspan="2" width="50px;">呼出</th>
                <% If workshift = "9" Then %>
                <th nowrap colspan="2">ﾌﾚｯｸｽﾀｲﾑ外・休出</th>
                <% Else %>
                <th nowrap colspan="2">時間外(休出)</th>
                <% End If %>
                <% If Not workshift = "9" Then ' フレックス勤務者以外のとき %>
                    <th nowrap colspan="2">時間単位代休</th>
                <% End If %>
                <th nowrap colspan="2">時間単位有休</th>
                <th nowrap colspan="2">深夜割増</th>
                <% If Not workshift = "9" Then %>
                    <th nowrap rowspan="2" width="50px;">週超過<br>時間</th>
                <% End If %>
                <th nowrap rowspan="2" width="60px;">宿直</th>
                <th nowrap rowspan="2" width="60px;">日直</th>
                <th nowrap rowspan="2" width="25px;">承<br>認</th>
                <th nowrap rowspan="2" width="180px;">備考</th>
                <% If workshift = "1" Or workshift = "2" Or workshift = "3" Then ' お客さまセンターオペレータのとき %>
                    <th nowrap rowspan="2" width="50px;">(勤務)<br>時間外</th>
                <% Else %>
                    <th nowrap rowspan="2" width="50px;">時間外</th>
                <% End If %>
                <th nowrap rowspan="2" width="50px;">時間外<br>深夜業</th>
                <th nowrap rowspan="2" width="50px;">休日<br>出勤</th>
                <th nowrap rowspan="2" width="50px;">休出<br>時間外</th>
                <th nowrap rowspan="2" width="50px;">休出<br>深夜業</th>
                <th nowrap rowspan="2" width="50px;">休出時<br>外深夜</th>
            </tr>
            <tr>
                <th nowrap width="50px;">出社</th>
                <th nowrap width="50px;">退社</th>
                <th nowrap width="80px;">午前</th>
                <% If is_operator Then %>
                    <th nowrap width="80px;">午後(0.5)</th>
                <% Else %>
                    <th nowrap width="80px;">午後</th>
                <% End If %>
                <% If workshift = "9" Then %>
                    <th nowrap width="60px;">自至</th>
                    <th nowrap width="60px;">休憩自至</th>
                    <th nowrap width="60px;">中抜自至</th>
                <% End If %>
                <th nowrap width="60px;">自至</th>
                <th nowrap width="60px;">休憩自至</th>
                <% If Not workshift = "9" Then ' フレックス勤務者以外のとき %>
                    <th nowrap width="50px;">時間数</th>
                    <th nowrap width="60px;">自至</th>
                <% End If %>
                <th nowrap width="50px;">時間数</th>
                <th nowrap width="60px;">自至</th>
                <th nowrap width="50px;">時間数</th>
                <th nowrap width="60px;">自至</th>
            </tr>
        </table>
    </div>
    <div id="tbody" class="tBody" style="width:1800px;height:180px;">
        <table id="workdata" class="data">
            <% For i = 1 To lastDay Step 1 ' 1日毎の繰り返し処理 %>
            <!-- #include file="inc/view_proc.asp" -->
            <%
                Dim disableChildCareLeaveOption
                disableChildCareLeaveOption = False

                If workshift = "9" Then
                    Dim currentWeekday
                    currentWeekday = Weekday(DateSerial(left(ymb, 4), right(ymb, 2), i))

                    ' Check for Saturday or Sunday
                    If currentWeekday = 1 Or currentWeekday = 7 Then
                        disableChildCareLeaveOption = True
                    End If

                    ' Check for public holiday (from holidaytbl via view_proc.asp)
                    If disableChildCareLeaveOption = False And v_morningholiday = "1" Then
                        disableChildCareLeaveOption = True
                    End If
                End If
            %>
            <tr>
                <%
                weekNameKanji = WeekdayName(Weekday(DateSerial(left(ymb, 4), right(ymb, 2), i)), true)
                If weekNameKanji = "日" Then
                    weekClass = "sunday"
                ElseIf weekNameKanji = "土" Then
                    weekClass = "saturday"
                Else
                    weekClass = ""
                End If
                If ymb & Right("0" & i, 2) = today Then
                    weekClass = "thisday"
                End If
                %>
                <th class="permanent <%=weekClass%>" width="25px;"><%=right("0" & i, 2)%></th>
                <th class="permanent <%=weekClass%>" width="25px;">
                    <%=weekNameKanji%>
                    <input type="hidden" name="ymd" value="<%=ymb & right("0" & i, 2)%>" <%=text_disabled%>>
                    <input type="hidden" name="everyday" value="<%=ymb & right("0" & i, 2)%>">
                    <input type="hidden" name="approval_ymd" value="<%=ymb & right("0" & i, 2)%>" <%=text_approval_disabled%>>
                    <input type="hidden" name="worktbl_id" value="<%=v_worktbl_id%>" <%=text_disabled%>>
                    <input type="hidden" name="worktbl_approval_id" value="<%=v_worktbl_id%>" <%=text_approval_disabled%>>
                    <input type="hidden" name="worktbl_updatetime" value="<%=v_updatetime%>" <%=text_disabled%>>
                    <input type="hidden" name="approval_worktbl_updatetime" value="<%=v_updatetime%>" <%=text_approval_disabled%>>
                    <input type="hidden" name="timetbl_id" value="<%=v_timetbl_id%>">
                </th>
                <td class="<%=weekClass%>" width="25px;" style="text-align:center;">
                    <input
                        name="is_error<%=ymb & right("0" & i, 2)%>"
                        value="<%=v_is_error%>"
                        maxlength="1"
                        style="text-align:center;width:13px;height:15px;"
                        <%=text_approval_disabled%>
                        class='<% If (v_is_error>"0") Then Response.Write("errorcolor") End If %>'
                        <%=text_is_error_disabled%>>
                </td>
                <td class="<%=weekClass%>" width="50px;" style="text-align:center;">
                    <input
                        name="beginTime<%=ymb & right("0" & i, 2)%>"
                        value="<%=v_cometime%>"
                        maxlength="5"
                        style="text-align:center;width:40px;height:15px;"
                        class="<%=style_beginTime(i)%>"
                        <%=text_timecard_disabled%>>
                    <br><%=pc_ontime%>
                    <input type="hidden" name="cometime"   value="<%=v_cometime  %>">
                    <input type="hidden" name="pc_ontime"  value="<%=pc_ontime   %>">
                    <input type="hidden" name="nightduty2" value="<%=v_nightduty2%>">
                    <input type="hidden" name="operator2"  value="<%=v_operator2 %>">
                </td>
                <td class="<%=weekClass%>" width="50px;" style="text-align:center;">
                    <input
                        name="endTime<%=ymb & right("0" & i, 2)%>"
                        value="<%=v_leavetime%>"
                        maxlength="5"
                        style="text-align:center;width:40px;height:15px;"
                        class="<%=style_endTime(i)%>"
                        <%=text_timecard_disabled%>>
                        <br><%=pc_offtime%>
                    <input type="hidden" name="leavetime"  value="<%=v_leavetime%>">
                    <input type="hidden" name="pc_offtime" value="<%=pc_offtime%>" >
                </td>
                <% If is_operator Then ' オペレータのとき、交替勤務の入力を出退勤時刻の右に表示 %>
                <td class="disabled <%=weekClass%>" width="70px;" style="text-align:center;">
                    <select
                        name="operator"
                        class="<%=style_operator(i)%>"
                        style="width:65px;height:22px;"
                        <%=text_disabled%> 
                        <% If Not is_operator Then Response.Write(" disabled") End If%>>
                        <option value="0">　</option>
                        <option value="1" <% If (v_operator="1") Then Response.Write("selected") End If %>>甲番</option>
                        <option value="2" <% If (v_operator="2") Then Response.Write("selected") End If %>>乙番</option>
                        <option value="3" <% If (v_operator="3") Then Response.Write("selected") End If %>>日勤甲</option>
                        <option value="4" <% If (v_operator="4") Then Response.Write("selected") End If %>>生産会議乙</option>
                        <option value="5" <% If (v_operator="5") Then Response.Write("selected") End If %>>見習(甲)</option>
                        <option value="6" <% If (v_operator="6") Then Response.Write("selected") End If %>>見習(乙)</option>
                        <option value="7" <% If (v_operator="7") Then Response.Write("selected") End If %>>A番</option>
                        <option value="8" <% If (v_operator="8") Then Response.Write("selected") End If %>>B番</option>
                    </select>
                </td>
                <% End If %>
                <td class="<%=weekClass%>" width="80px;" style="text-align:center;">
                    <select
                        name="morningholiday"
                        class="<%=style_morningholiday(i)%>"
                        style="width:75px;height:22px;"
                        <%=text_disabled%>>
                        <option value="0">　</option>
                        <option value="1" <% If (v_morningholiday="1") Then Response.Write("selected") End If %>>公休日</option>
                        <% If workshift = "9" Then %>
                        <option value="A" <% If (v_morningholiday="A") Then Response.Write("selected") End If %>>法定休日</option>
                        <% End If %>
                        <option value="2" <% If (v_morningholiday="2") Then Response.Write("selected") End If %>>振替休日</option>
                        <option value="3" <% If (v_morningholiday="3") Then Response.Write("selected") End If %>>有給休暇</option>
                        <% If workshift = "9" Then %>
                        <option value="9" <% If (v_morningholiday="9") Then Response.Write("selected") End If %>>有給(コアタイム)</option>
                        <% Else %>
                        <option value="4" <% If (v_morningholiday="4") Then Response.Write("selected") End If %>>代替休暇</option>
                        <% End If %>
                        <option value="5" <% If (v_morningholiday="5") Then Response.Write("selected") End If %>>特別休暇</option>
                        <option value="6" <% If (v_morningholiday="6") Then Response.Write("selected") End If %>>保存休暇</option>
                        <option value="7" <% If (v_morningholiday="7") Then Response.Write("selected") End If %>>欠勤</option>
                        <% If Not (workshift = "9" And disableChildCareLeaveOption) Then %>
                        <option value="B" <% If (v_morningholiday="B") Then Response.Write("selected") End If %>>育児休業</option>
                        <% End If %>
                    </select><br>
                    <select
                        name="morningwork"
                        class="<%=style_morningwork(i)%>"
                        style="width:75px;height:22px;"
                        <%=text_disabled%>>
                        <option value="0">　</option>
                        <option value="9" <% If (v_morningwork="9") Then Response.Write("selected") End If %>>出勤</option>
                        <option value="1" <% If (v_morningwork="1") Then Response.Write("selected") End If %>>振替出勤</option>
                        <option value="2" <% If (v_morningwork="2") Then Response.Write("selected") End If %>>休出</option>
                        <option value="3" <% If (v_morningwork="3") Then Response.Write("selected") End If %>>休出半日未満</option>
                        <% If workshift = "0" Or workshift = "9" Then ' 一般社員(お客さまセンターオペレータ以外)のとき) %>
                        <option value="4" <% If (v_morningwork="4") Then Response.Write("selected") End If %>>出張(出勤)</option>
                        <option value="5" <% If (v_morningwork="5") Then Response.Write("selected") End If %>>出張(振替出勤)</option>
                        <option value="6" <% If (v_morningwork="6") Then Response.Write("selected") End If %>>出張(休出)</option>
                        <% End If %>
                    </select>
                    <input type="hidden" name="hd_morningholiday" value="<%=v_morningholiday%>">
                    <input type="hidden" name="hd_morningwork" value="<%=v_morningwork%>">
                </td>
                <td class="<%=weekClass%>" width="80px;" style="text-align:center;">
                    <select
                        name="afternoonholiday"
                        class="<%=style_afternoonholiday(i)%>"
                        style="width:75px;height:22px;"
                        <%=text_disabled%>>
                        <option value="0">　</option>
                        <option value="1" <% If (v_afternoonholiday="1") Then Response.Write("selected") End If %>>公休日</option>
                        <% If workshift = "9" Then %>
                        <option value="A" <% If (v_afternoonholiday="A") Then Response.Write("selected") End If %>>法定休日</option>
                        <% End If %>
                        <option value="2" <% If (v_afternoonholiday="2") Then Response.Write("selected") End If %>>振替休日</option>
                        <option value="3" <% If (v_afternoonholiday="3") Then Response.Write("selected") End If %>>有給休暇</option>
                        <% If workshift = "9" Then %>
                        <option value="9" <% If (v_afternoonholiday="9") Then Response.Write("selected") End If %>>有給(コアタイム)</option>
                        <% Else %>
                        <option value="4" <% If (v_afternoonholiday="4") Then Response.Write("selected") End If %>>代替休暇</option>
                        <% End If %>
                        <option value="5" <% If (v_afternoonholiday="5") Then Response.Write("selected") End If %>>特別休暇</option>
                        <option value="6" <% If (v_afternoonholiday="6") Then Response.Write("selected") End If %>>保存休暇</option>
                        <option value="7" <% If (v_afternoonholiday="7") Then Response.Write("selected") End If %>>欠勤</option>
                        <% If Not (workshift = "9" And disableChildCareLeaveOption) Then %>
                        <option value="B" <% If (v_afternoonholiday="B") Then Response.Write("selected") End If %>>育児休業</option>
                        <% End If %>
                    </select><br>
                    <select
                        name="afternoonwork"
                        class="<%=style_afternoonwork(i)%>"
                        style="width:75px;height:22px;"
                        <%=text_disabled%>>
                        <option value="0">　</option>
                        <option value="9" <% If (v_afternoonwork="9") Then Response.Write("selected") End If %>>出勤</option>
                        <option value="1" <% If (v_afternoonwork="1") Then Response.Write("selected") End If %>>振替出勤</option>
                        <option value="2" <% If (v_afternoonwork="2") Then Response.Write("selected") End If %>>休出</option>
                        <option value="3" <% If (v_afternoonwork="3") Then Response.Write("selected") End If %>>休出半日未満</option>
                        <% If workshift = "0" Or workshift = "9" Then ' 一般社員(お客さまセンターオペレータ以外)のとき) %>
                        <option value="4" <% If (v_afternoonwork="4") Then Response.Write("selected") End If %>>出張(出勤)</option>
                        <option value="5" <% If (v_afternoonwork="5") Then Response.Write("selected") End If %>>出張(振替出勤)</option>
                        <option value="6" <% If (v_afternoonwork="6") Then Response.Write("selected") End If %>>出張(休出)</option>
                        <% End If %>
                    </select>
                    <input type="hidden" name="hd_afternoonholiday" value="<%=v_afternoonholiday%>">
                    <input type="hidden" name="hd_afternoonwork" value="<%=v_afternoonwork%>">
                </td>
                <% If Not is_operator Then %>
                    <%
                    If v_workmin = Empty Then
                        Response.Write("&nbsp;")
                    Else
                        
                        'Response.Write("<span class='" & v_overwork & "'>" & min2Time(v_workmin) & "</span>")
                        realworkmin = realworkmin + v_workmin
                    End If
                    If workshift = "9" Then v_overwork = "" End If
                    %>
                    <td class="<%=weekClass & " " & v_overwork%>" width="50px;" style="text-align:center;">
                        <%=min2Time(v_workmin)%>
                    </td>
                <% End If %>
                <% If workshift = "9" Then ' フレックス勤務者のとき %>
                <td class="<%=weekClass%>" width="60px;" style="text-align:center;">
                    <input
                        name="work_begin"
                        class="<%=style_work_begin(i)%>"
                        maxlength="5"
                        style="text-align:center;width:45px;height:15px;"
                        <%=text_disabled%>
                        value="<%=v_work_begin%>"><br>
                    <input
                        name="work_end"
                        class="<%=style_work_end(i)%>"
                        maxlength="5"
                        style="text-align:center;width:45px;height:15px;"
                        <%=text_disabled%>
                        value="<%=v_work_end%>">
                    <input type="hidden" name="hd_work_begin" value="<%=v_work_begin%>">
                    <input type="hidden" name="hd_work_end" value="<%=v_work_end%>">
                </td>
                <td class="<%=weekClass%>" width="60px;" style="text-align:center;">
                    <input
                        name="break_begin1"
                        class="<%=style_break_begin1(i)%>"
                        maxlength="5"
                        style="text-align:center;width:45px;height:15px;"
                        <%=text_disabled%>
                        value="<%=v_break_begin1%>"><br>
                    <input
                        name="break_end1"
                        class="<%=style_break_end1(i)%>"
                        maxlength="5"
                        style="text-align:center;width:45px;height:15px;"
                        <%=text_disabled%>
                        value="<%=v_break_end1%>">
                    <input type="hidden" name="hd_break_begin1" value="<%=v_break_begin1%>">
                    <input type="hidden" name="hd_break_end1" value="<%=v_break_end1%>">
                </td>
                <td class="<%=weekClass%>" width="60px;" style="text-align:center;">
                    <input
                        name="break_begin2"
                        class="<%=style_break_begin2(i)%>"
                        maxlength="5"
                        style="text-align:center;width:45px;height:15px;"
                        <%=text_disabled%>
                        value="<%=v_break_begin2%>"><br>
                    <input
                        name="break_end2"
                        class="<%=style_break_end2(i)%>"
                        maxlength="5"
                        style="text-align:center;width:45px;height:15px;"
                        <%=text_disabled%>
                        value="<%=v_break_end2%>">
                    <input type="hidden" name="hd_break_begin2" value="<%=v_break_begin2%>">
                    <input type="hidden" name="hd_break_end2" value="<%=v_break_end2%>">
                </td>
                <% End If %>
                <td class="<%=weekClass%>" width="50px;" style="text-align:center;">
                    <select
                        name="summons"
                        class="<%=style_summons(i)%>"
                        style="width:45px;height:22px;"
                        <%=text_disabled%>>
                        <option value="0">　</option>
                        <option value="1" <% If (v_summons="1") Then Response.Write("selected") End If %>>通常</option>
                        <option value="2" <% If (v_summons="2") Then Response.Write("selected") End If %>>深夜</option>
                    </select>
                </td>
                <td class="<%=weekClass%>" width="60px;" style="text-align:center;">
                    <input
                        name="overtime_begin"
                        class="<%=style_overtime_begin(i)%>"
                        maxlength="5"
                        style="text-align:center;width:45px;height:15px;"
                        <%=text_disabled%>
                        value="<%=v_overtime_begin%>"><br>
                    <input
                        name="overtime_end"
                        class="<%=style_overtime_end(i)%>"
                        maxlength="5"
                        style="text-align:center;width:45px;height:15px;"
                        <%=text_disabled%>
                        value="<%=v_overtime_end%>">
                    <input
                        type="hidden"
                        name="hd_overtime_begin"
                        value="<%=v_overtime_begin%>">
                    <input
                        type="hidden"
                        name="hd_overtime_end"
                        value="<%=v_overtime_end%>">
                </td>
                <td class="<%=weekClass%>" width="60px;" style="text-align:center;">
                    <input
                        name="rest_begin"
                        class="<%=style_rest_begin(i)%>"
                        maxlength="5"
                        style="text-align:center;width:45px;height:15px;"
                        <%=text_disabled%>
                        value="<%=v_rest_begin%>"><br>
                    <input
                        name="rest_end"
                        class="<%=style_rest_end(i)%>"
                        maxlength="5"
                        style="text-align:center;width:45px;height:15px;"
                        <%=text_disabled%>
                        value="<%=v_rest_end%>">
                    <input
                        type="hidden"
                        name="hd_rest_begin"
                        value="<%=v_rest_begin%>">
                    <input
                        type="hidden"
                        name="hd_rest_end"
                        value="<%=v_rest_end%>">
                </td>
                <% If Not workshift = "9" Then ' フレックス勤務者以外のとき %>
                <td class="<%=weekClass%>" width="50px;" style="text-align:center;">
                    <%
                    If ((style_requesttime_begin(i) = "" And style_requesttime_end(i) = "")  And _
                        (Len(v_requesttime_begin)   > 0  And Len(v_requesttime_end) > 0   )) Then
                        v_requesttime = min2Time(minDif(v_requesttime_begin, v_requesttime_end) - _
                                            checkLunchTime(v_requesttime_begin, v_requesttime_end))
                    Else
                        v_requesttime = ""
                    End If
                    %>
                    <%=v_requesttime%>
                </td>
                <td class="<%=weekClass%>" width="60px;" style="text-align:center;">
                    <input
                        name="requesttime_begin"
                        class="<%=style_requesttime_begin(i)%>"
                        maxlength="5"
                        style="text-align:center;width:45px;height:15px;"
                        <%=text_disabled%>
                        value="<%=v_requesttime_begin%>"><br>
                    <input
                        name="requesttime_end"
                        class="<%=style_requesttime_end(i)%>"
                        maxlength="5"
                        style="text-align:center;width:45px;height:15px;"
                        <%=text_disabled%>
                        value="<%=v_requesttime_end%>">
                    <input
                        type="hidden"
                        name="hd_requesttime_begin"
                        value="<%=v_requesttime_begin%>">
                    <input
                        type="hidden"
                        name="hd_requesttime_end"
                        value="<%=v_requesttime_end%>">
                </td>
                <% End If %>
                <td class="<%=weekClass%>" width="50px;" style="text-align:center;">
                    <%
                    If ((style_vacationtime_begin(i) = "" And style_vacationtime_end(i) = "")  And _
                        (Len(v_vacationtime_begin)   > 0  And Len(v_vacationtime_end) > 0   )) Then
                        v_vacationtime = min2Time(minDif(v_vacationtime_begin, v_vacationtime_end) - _
                                            checkLunchTime(v_vacationtime_begin, v_vacationtime_end))
                    Else
                        v_vacationtime = ""
                    End If
                    %>
                    <%=v_vacationtime%>
                </td>
                <td class="<%=weekClass%>" width="60px;" style="text-align:center;">
                    <input
                        name="vacationtime_begin"
                        class="<%=style_vacationtime_begin(i)%>"
                        maxlength="5"
                        style="text-align:center;width:45px;height:15px;"
                        <%=text_disabled%>
                        value="<%=v_vacationtime_begin%>"><br>
                    <input
                        name="vacationtime_end"
                        class="<%=style_vacationtime_end(i)%>"
                        maxlength="5"
                        style="text-align:center;width:45px;height:15px;"
                        <%=text_disabled%>
                        value="<%=v_vacationtime_end%>">
                    <input
                        type="hidden"
                        name="hd_vacationtime_begin"
                        value="<%=v_vacationtime_begin%>">
                    <input
                        type="hidden"
                        name="hd_vacationtime_end"
                        value="<%=v_vacationtime_end%>">
                </td>
                <td class="<%=weekClass%>" width="50px;" style="text-align:center;">
                    <%
                    If ((style_latetime_begin(i) = "" And style_latetime_end(i) = "")  And _
                        (Len(v_latetime_begin)   > 0  And Len(v_latetime_end) > 0   )) Then
                        v_latetime = min2Time(minDif(v_latetime_begin, v_latetime_end))
                    Else
                        v_latetime = ""
                    End If
                    %>
                    <%=v_latetime%>
                </td>
                <td class="<%=weekClass%>" width="60px;" style="text-align:center;">
                    <input
                        name="latetime_begin"
                        class="<%=style_latetime_begin(i)%>"
                        maxlength="5"
                        style="text-align:center;width:45px;height:15px;"
                        <%=text_disabled%>
                        value="<%=v_latetime_begin%>"><br>
                    <input
                        name="latetime_end"
                        class="<%=style_latetime_end(i)%>"
                        maxlength="5"
                        style="text-align:center;width:45px;height:15px;"
                        <%=text_disabled%>
                        value="<%=v_latetime_end%>">
                    <input
                        type="hidden"
                        name="hd_latetime_begin"
                        value="<%=v_latetime_begin%>">
                    <input
                        type="hidden"
                        name="hd_latetime_end"
                        value="<%=v_latetime_end%>">
                </td>
                <% If Not workshift = "9" Then %>
                <td class="<%=weekClass%>" width="50px;" style="text-align:center;">
                    <input
                        name="weekovertime"
                        class="<%=style_weekovertime(i)%>"
                        maxlength="5"
                        style="text-align:center;width:45px;height:15px;"
                        <%=text_disabled%>
                        value="<%=v_weekovertime%>">
                </td>
                <% End If %>
                <td class="disabled <%=weekClass%>" width="60px;" style="text-align:center;">
                    <select
                        name="nightduty"
                        class="<%=style_nightduty(i)%>"
                        style="width:55px;height:22px;"
                        <%=text_disabled%>>
                        <option value="0">　</option>
                        <option value="1" <% If (v_nightduty="1") Then Response.Write("selected") End If %>>責任者</option>
                        <option value="2" <% If (v_nightduty="2") Then Response.Write("selected") End If %>>処理者</option>
                    </select>
                </td>
                <td class="disabled <%=weekClass%>" align="center" width="60px;">
                    <select
                        name="dayduty"
                        class="<%=style_dayduty(i)%>"
                        style="width:55px;height:22px;"
                        <%=text_disabled%>>
                        <option value="0">　</option>
                        <option value="2" <% If (v_dayduty="2") Then Response.Write("selected") End If %>>処理者</option>
                        <option value="3" <% If (v_dayduty="3") Then Response.Write("selected") End If %>>休日出番</option>
                    </select>
                </td>
                <td class="disabled <%=weekClass%>" align="center" width="25px;">
                    <input
                        name="is_approval<%=ymb & right("0" & i, 2)%>"
                        <%=text_approval_disabled%>
                        type="checkbox"
                        <% If (v_is_approval="1") Then Response.Write("checked") End If %>>
                </td>
                <td class="<%=weekClass%>" width="180px;" style="text-align:center;">
                    <input
                        name="memo"
                        class="<%=style_memo(i)%>"
                        maxlength="50"
                        style=" text-align:left; width:160px; height:15px;"
                        <% If screen = 0 Then Response.Write(text_disabled) End If %>
                        value="<%=v_memo%>"><br >
                        <select name="memo2" <%=text_disabled%> style=" text-align:left; width:168px;">
                            <option value='0'>-</option>
                            <%
                            array_idx = 1
                            Do While (memoArray(array_idx) <> "")
                                memo2sel = ""
                                If (v_memo2 = (array_idx & "")) Then memo2sel = "selected" End If
                                Response.Write("<option value='" & array_idx & "' " & memo2sel & " >" & memoArray(array_idx) & "</option>")
                                array_idx = array_idx + 1
                            Loop
                            %>
                        </select>
                </td>
                <% compOverTime() %>
                <td class="<%=weekClass%>" width="50px;" style="text-align:center;">
                    <% 
                    If workshift = "1" Or workshift = "2" Or workshift = "3" Then ' お客さまセンターオペレータのとき、勤務時間を表示
                        tempMin = 0
                        If (v_morningwork > "0" Or v_afternoonwork > "0") Then
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
                        
                        Response.Write("(" & mm2Float(tempMin) & ")<br>" & v_overtime)
                    Else
                        Response.write(v_overtime)
                    End If
                    %>
                </td>
                <td class="<%=weekClass%>" width="50px;">
                    <p style="text-align:center;"><%=v_overtimelate%></p>
                </td>
                <td class="<%=weekClass%>" width="50px;">
                    <p style="text-align:center;"><%=v_holidayshift%></p>
                </td>
                <td class="<%=weekClass%>" width="50px;">
                    <p style="text-align:center;"><%=v_holidayshiftovertime%></p>
                </td>
                <td class="<%=weekClass%>" width="50px;">
                    <p style="text-align:center;"><%=v_holidayshiftlate%></p>
                </td>
                <td class="<%=weekClass%>" width="50px;">
                    <p style="text-align:center;"><%=v_holidayshiftovertimelate%></p>
                </td>
            </tr>
            <% Next %>
        </table>
    </div>
</div>
<div style="position:absolute;top:70px;width:1800px;">
    <div class="left" style="width:1550px;">
        <div>
            <table class="data">
                <tr>
                    <th rowspan="2" width="25px">当<br>月</th>
                    <th width="60px" style="font-size:8pt;">可出勤日数</th>
                    <th width="60px">欠勤日数</th>
                    <th width="60px">有休日数</th>
                    <th width="60px" style="font-size:8pt;">保存休日数</th>
                    <th width="60px">特休日数</th>
                    <th width="60px">休出日数</th>
                    <th width="60px" style="font-size:8pt;">実出勤日数</th>
                    <th width="60px">呼出回数</th>
                    <th width="60px">呼出深夜</th>
                    <th width="60px">宿直</th>
                    <th width="60px">日直</th>
                    <% If workshift <> "9" Then %>
                    <th width="60px">時間代休</th>
                    <% End If %>
                    <th width="60px">深夜割増</th>
                    <% If is_operator Then %>
                    <th width="60px">甲番勤務</th>
                    <th width="60px">乙番勤務</th>
                    <% End If %>
                    <th width="60px">休出回数</th>
                    <th width="60px">時間外計</th>
                    <th width="60px">休出計</th>
                    <% If Not workshift = "9" And Not is_operator Then %>
                    <th width="60px">週超過計</th>
                    <th width="60px" style="font-size:8pt;">AM基本時間</th>
                    <th width="60px" style="font-size:8pt;">PM基本時間</th>
                    <% End If %>
                    <% If workshift = "1" Or workshift = "2" Or workshift = "3" Then ' コミュニケータ勤務 %>
                    <th width="70px">土曜+100円</th>
                    <th width="70px">平日(分)</th>
                    <% End If %>
                    <% If workshift = "9" Then ' フレックス勤務 %>
                    <th width="70px" style="font-size:8pt;">基準労働時間</th>
                    <th width="70px" style="font-size:8pt;">当月労働時間</th>
                    <th width="60px">勤務実績</th>
                    <th width="60px">勤務差分</th>
                    <% End If %>
                </tr>
                <tr>
                    <td width="60px" class="disabled" style="text-align:right;">
                        <%=sumWorkDays&strDay%>
                        <input
                            type="hidden"
                            name="dutyrostertbl_id"
                            value="<%=dutyrostertbl_id%>">
                    </td>
                    <td width="60px" class="disabled" style="text-align:right;">
                        <%=sumAbsenceDays&strDay%>
                    </td>
                    <td width="60px" class="disabled" style="text-align:right;">
                        <%=sumPaidvacations&strDay%>
                    </td>
                    <td width="60px" class="disabled" style="text-align:right;">
                        <%=sumPreservevacations&strDay%>
                    </td>
                    <td width="60px" class="disabled" style="text-align:right;">
                        <%=sumSpecialvacations&strDay%>
                    </td>
                    <%
                    ' 休出日数の小数点計算
                    If workshift = "9" Then
                        temp_holidayshifts = mm2FloatDay(sumFlex_holidayshift)
                    Else
                        temp_holidayshifts = mm2FloatDay(sumHolidayshifttime + sumHolidayshiftlate)
                    End If
                    If temp_holidayshifts > 2 Then
                        warningClass = "abnormality"
                    ElseIf temp_holidayshifts >= 1.5 Then
                        warningClass = "warning"
                    Else
                        warningClass = ""
                    End If
                    %>
                    <td width="60px" class="disabled <%=warningClass%>" style="text-align:right;">
                        <%=temp_holidayshifts&strDay%>
                    </td>
                    <td width="60px" class="disabled" style="text-align:right;">
                        <%
                        If workshift = "9" Then
                            ' フレックス勤務者の場合は既にsumRealworkdaysに休出分が含まれているため加算しない
                            Response.Write(sumRealworkdays & strDay)
                        Else
                            ' 通常勤務者の場合も休出日数は実出勤日数に含めない
                            Response.Write(sumRealworkdays & strDay)
                        End If
                        %>
                    </td>
                    <td width="60px" class="disabled" style="text-align:right;">
                        <%=sumSummons&strCount%>
                    </td>
                    <td width="60px" class="disabled" style="text-align:right;">
                        <%=sumSummonslate&strCount%>
                    </td>
                    <td width="60px" class="disabled" style="text-align:right;">
                        <%=sumNightdutyCount&strCount%>
                    </td>
                    <td width="60px" class="disabled" style="text-align:right;">
                        <%=sumDaydutyCount&strCount%>
                    </td>
                    <% If workshift <> "9" Then %>
                    <td width="60px" class="disabled" style="text-align:right;">
                        <%=mm2Float(sumWorkholidaytime)%>&nbsp;時間
                    </td>
                    <% End If %>
                    <td width="60px" class="disabled" style="text-align:right;">
                        <%=mm2Float(sumLatepremium)%>&nbsp;時間
                    </td>
                    <% If is_operator Then %>
                    <td width="60px" class="disabled" style="text-align:right;">
                        <%=sumOperatorKou&strCount%>
                    </td>
                    <td width="60px" class="disabled" style="text-align:right;">
                        <%=sumOperatorOtsu&strCount%>
                    </td>
                    <%
                    End If
                    If sumHolidayWork >= 5 Then
                        warningClass = "abnormality"
                    ElseIf sumHolidayWork >= 4 Then
                        warningClass = "warning"
                    Else
                        warningClass = ""
                    End If
                    %>
                    <td width="60px" class="disabled <%=warningClass%>" style="text-align:right;">
                        <%=sumHolidayWork%>&nbsp;回
                    </td>
                    <%
                    ' 時間外労働計 = 時間外 + 時間外深夜業 + 休出時間外 + 休出時間外深夜
                    sumTotalOvertime = mm2Float(sumOvertime                ) _
                                     + mm2Float(sumOvertimelate            ) _
                                     + mm2Float(sumHolidayshiftovertime    ) _
                                     + mm2Float(sumHolidayshiftovertimelate)
                                     '- mm2Float(sumWorkholidaytime         )   時間代休の減算を取りやめる
                    If workshift = "9" Then
                        sumTotalOvertime = realworkmin - currentworkmin
                        If sumTotalOvertime > 0 Then
                            sumTotalOvertime = mm2Float(sumTotalOvertime)
                        Else
                            sumTotalOvertime = 0
                        End If
                    End If
                    
                    If sumTotalOvertime + init_weekovertime >= 29 Then
                        overtime_warningClass = "abnormality"
                    ElseIf sumTotalOvertime + init_weekovertime >= 25 Then
                        overtime_warningClass = "warning"
                    Else
                        overtime_warningClass = ""
                    End If
                    %>
                    <td width="60px" class="disabled <%=overtime_warningClass%>" style="text-align:right;">
                        <%
                        If sumTotalOvertime < 0 Then
                            Response.Write("0&nbsp;時間")
                        Else
                            Response.Write(sumTotalOvertime & "&nbsp;時間")
                        End If
                        %>
                    </td>
                    <%
                    If workshift = 9 Then
                        sumHolidayshifttime = sumFlex_holidayshift
                        sumHolidayshiftlate = 0
                    End If
                    If mm2Float(sumHolidayshifttime) + mm2Float(sumHolidayshiftlate) >= 15.4 Then
                        warningClass = "abnormality"
                    ElseIf mm2Float(sumHolidayshifttime) + mm2Float(sumHolidayshiftlate) >= 10 Then
                        warningClass = "warning"
                    Else
                        warningClass = ""
                    End If
                    %>
                    <td width="60px" class="disabled <%=warningClass%>" style="text-align:right;">
                        <%=mm2Float(sumHolidayshifttime) + mm2Float(sumHolidayshiftlate)%>&nbsp;時間
                    </td>
                    <% If Not workshift = "9" And Not is_operator Then %>
                        <td style="text-align:right;" class="disabled <%=overtime_warningClass%>"><%=init_weekovertime & strTime%></td>
                        <td style="text-align:center;"><%=min2time(base_am_workmin)%></td>
                        <td style="text-align:center;"><%=min2time(base_pm_workmin)%></td>
                    <% End If %>
                    <% If workshift = "1" Or workshift = "2" Or workshift = "3" Then ' コミュニケータ勤務 %>
                    <td style="text-align:right;"><%=mm2Float(sumSaturdayWorkMin) & strTime%></td>
                    <td style="text-align:right;"><%=mm2Float(sumWeekdaysWorkMin) & strTime%></td>
                    <% End If %>
                    <% If workshift = "9" Then ' フレックス勤務 %>
                    <td style="text-align:right;">
                        <%=min2Time(thismonth_basemin)%>
                        <input type="hidden" name="thismonth_baseminHidden" value="<%=thismonth_basemin%>">
                        <input type="hidden" name="lastmonth_currentworkminHidden" value="<%=lastmonth_currentworkmin%>">
                        <input type="hidden" name="lastmonth_workingminsHidden" value="<%=lastmonth_workingmins%>">
                    </td>
                    <td style="text-align:right;">
                        <%=min2Time(currentworkmin)%>
                    </td>
                    <td style="text-align:right;"><%=min2Time(realworkmin)%></td>
                    <%
                    ' フレックス勤務 勤務差分
                    temp = realworkmin - currentworkmin
                    flexclass = ""
                    If temp >= 0 Then
                        temp = min2Time(temp)
                    Else
                        temp = "-" + min2Time(temp * -1)
                        flexclass = "flexcheck"
                    End If
                    %>
                    <td style="text-align:right;" class="<%=flexclass%>">
                        <%=temp%>
                    </td>
                    <% End If %>
                </tr>
            </table>
        </div>
        <div>
            <div class="left">
                <table class="data">
                    <tr>
                        <th rowspan="2" width="25px">累<br>積</th>
                        <th width="60px" style="font-size:8pt;">有休付与日</th>
                        <th width="60px">有休取得</th>
                        <th width="60px">有休残数</th>
                        <th width="60px">振替残数</th>
                        <th width="60px">時間有休</th>
                        <th width="60px" style="font-size:8pt;">保存休残数</th>
                        <th width="60px" style="font-size:8pt;">時間外累積</th>
                        <th width="60px">休出時間</th>
                        <th width="60px">休出回数</th>
                        <th width="60px" style="font-size:8pt;">2月平均時外</th>
                        <th width="60px" style="font-size:8pt;">3月平均時外</th>
                        <th width="60px" style="font-size:8pt;">4月平均時外</th>
                        <th width="60px" style="font-size:8pt;">5月平均時外</th>
                        <th width="60px" style="font-size:8pt;">6月平均時外</th>
                    </tr>
                    <tr>
                        <td width="60px" class="disabled" style="text-align:center;">
                            <%=Left(grantdate,2) * 1 & "月" & Right(grantdate,2) * 1 & "日"%>
                        </td>
                        <%
                        warningClass = ""
                        If totalPaidvacations < 5 Then
                            If Left(grantdate,2) = checkHolidayMM1 Then
                                warningClass = "abnormality"
                            ElseIf UBound(Filter(checkHolidayMM3, Left(grantdate, 2))) <> -1 Then
                                warningClass = "warning"
                            End If
                        End If
                        %>
                        <td width="60px" class="disabled <%=warningClass%>" style="text-align:right;">
                           <%=totalPaidvacations & strDay%>
                        </td>
                        <td width="60px" class="disabled" style="text-align:right;">
                            <%=(vacationnumber - sumPaidvacations) & strDay%>
                            <input
                                type="hidden"
                                name="sumVacationnumberHidden"
                                value="<%=vacationnumber + sumVacationnumberHidden%>">
                            <input
                                type="hidden"
                                name="vacationnumberHidden"
                                value="<%=vacationnumber%>">
                        </td>
                        <td width="60px" class="disabled" style="text-align:right;">
                            <%=(holidaynumber + sumHolidaynumber) & strDay%>
                            <input
                                type="hidden"
                                name="sumHolidaynumberHidden"
                                value="<%=holidaynumber + sumHolidaynumberHidden%>">
                            <input
                                type="hidden"
                                name="holidaynumberHidden"
                                value="<%=holidaynumber%>">
                        </td>
                        <td width="60px" class="disabled" style="text-align:right;">
                            <%=mm2Float(vacationtime + sumVacationtime) & strTime%>
                            <input
                                type="hidden"
                                name="vacationtimeHidden"
                                value="<%=vacationtime%>">
                        </td>
                        <td width="60px" class="disabled" style="text-align:right;">
                            <%
                            ' =================================================================
                            ' 保存休テーブル処理
                            ' =================================================================
                            If Not Rs_remainvacationtbl.EOF Then
                                remainvacation = Rs_remainvacationtbl.Fields.Item("remainvacation").Value - _
                                                 Rs_remainvacationtbl.Fields.Item("preservevacations").Value
                            Else
                                remainvacation = 0
                            End If
                            %>
                            <%=remainvacation & strDay%>
                        </td>
                        <%
                        If sumTotalOvertime + init_weekovertime + yearlyOvertime >= 176 Then
                            warningClass = "abnormality"
                        ElseIf sumTotalOvertime + init_weekovertime + yearlyOvertime >= 150 Then
                            warningClass = "warning"
                        Else
                            warningClass = ""
                        End If
                        %>
                        <td width="60px" class="disabled <%=warningClass%>" style="text-align:right;">
                            <%
                            If sumTotalOvertime + init_weekovertime + yearlyOvertime < 0 Then
                                Response.Write(yearlyOvertime & strTime)
                            Else
                                Response.Write(Round(sumTotalOvertime + init_weekovertime + yearlyOvertime, 1) & strTime)
                            End If
                            %>
                            <input type="hidden" name="yearlyOvertimeHidden" value="<%=yearlyOvertime%>">
                        </td>
                        <%
                        If mm2Float(sumHolidayshifttime) + mm2Float(sumHolidayshiftlate) + yearlyHolidaytime >= 184 Then
                            warningClass = "abnormality"
                        ElseIf mm2Float(sumHolidayshifttime) + mm2Float(sumHolidayshiftlate) + yearlyHolidaytime >= 150 Then
                            warningClass = "warning"
                        Else
                            warningClass = ""
                        End If
                        %>
                        <td width="60px" class="disabled <%=warningClass%>" style="text-align:right;">
                            <%=mm2Float(sumHolidayshifttime) + mm2Float(sumHolidayshiftlate) + yearlyHolidaytime & strTime%>
                        </td>
                        <%
                        If yearlyHolidaywork + monthlyHolidaywork >= 42 Then
                            warningClass = "abnormality"
                        ElseIf yearlyHolidaywork + monthlyHolidaywork >= 35 Then
                            warningClass = "warning"
                        Else
                            warningClass = ""
                        End If
                        %>
                        <td width="60px" class="disabled <%=warningClass%>" style="text-align:right;">
                            <%=yearlyHolidaywork+monthlyHolidaywork%>&nbsp;回
                            <input type="hidden" name="yearlyHolidayworkHidden" value="<%=yearlyHolidaywork%>">
                        </td>
                        <%
                        If Round((sumOvertime0 + sumOvertime1) / 2, 1) >= 80 Then
                            warningClass = "abnormality"
                        ElseIf Round((sumOvertime0 + sumOvertime1) / 2, 1) >= 70 Then
                            warningClass = "warning"
                        Else
                            warningClass = ""
                        End If
                        %>
                        <td width="60px" class="disabled <%=warningClass%>" style="text-align:right;">
                            <%=Round((sumOvertime0 + sumOvertime1) / 2, 1) & strTime%>
                            <input type="hidden" name="sumOvertime1" value="<%=sumOvertime1%>">
                            <input type="hidden" name="sumOvertime2" value="<%=sumOvertime2%>">
                            <input type="hidden" name="sumOvertime3" value="<%=sumOvertime3%>">
                            <input type="hidden" name="sumOvertime4" value="<%=sumOvertime4%>">
                            <input type="hidden" name="sumOvertime5" value="<%=sumOvertime5%>">
                        </td>
                        <%
                        If Round((sumOvertime0 + sumOvertime1 + sumOvertime2) / 3, 1) >= 80 Then
                            warningClass = "abnormality"
                        ElseIf Round((sumOvertime0 + sumOvertime1 + sumOvertime2) / 3, 1) >= 70 Then
                            warningClass = "warning"
                        Else
                            warningClass = ""
                        End If
                        %>
                        <td width="60px" class="disabled <%=warningClass%>" style="text-align:right;">
                            <%=Round((sumOvertime0 + sumOvertime1 + sumOvertime2) / 3, 1) & strTime%>
                        </td>
                        <%
                        If Round((sumOvertime0 + sumOvertime1 + sumOvertime2 + sumOvertime3) /4, 1) >= 80 Then
                            warningClass = "abnormality"
                        ElseIf Round((sumOvertime0 + sumOvertime1 + sumOvertime2 + sumOvertime3) /4, 1) >= 70 Then
                            warningClass = "warning"
                        Else
                            warningClass = ""
                        End If
                        %>
                        <td width="60px" class="disabled <%=warningClass%>" style="text-align:right;">
                            <%=Round((sumOvertime0 + sumOvertime1 + sumOvertime2 + sumOvertime3) /4, 1) & strTime%>
                        </td>
                        <%
                        If Round((sumOvertime0 + sumOvertime1 + sumOvertime2 + sumOvertime3 + sumOvertime4) / 5, 1) >= 80 Then
                            warningClass = "abnormality"
                        ElseIf Round((sumOvertime0 + sumOvertime1 + sumOvertime2 + sumOvertime3 + sumOvertime4) / 5, 1) >= 70 Then
                            warningClass = "warning"
                        Else
                            warningClass = ""
                        End If
                        %>
                        <td width="60px" class="disabled <%=warningClass%>" style="text-align:right;">
                            <%=Round((sumOvertime0 + sumOvertime1 + sumOvertime2 + sumOvertime3 + sumOvertime4) / 5, 1) & strTime%>
                        </td>
                        <%
                        If Round((sumOvertime0 + sumOvertime1 + sumOvertime2 + sumOvertime3 + sumOvertime4 + sumOvertime5) / 6, 1) >= 80 Then
                            warningClass = "abnormality"
                        ElseIf Round((sumOvertime0 + sumOvertime1 + sumOvertime2 + sumOvertime3 + sumOvertime4 + sumOvertime5) / 6, 1) >= 70 Then
                            warningClass = "warning"
                        Else
                            warningClass = ""
                        End If
                        %>
                        <td width="60px" class="disabled <%=warningClass%>" style="text-align:right;">
                            <%=Round((sumOvertime0 + sumOvertime1 + sumOvertime2 + sumOvertime3 + sumOvertime4 + sumOvertime5) / 6, 1) & strTime%>
                        </td>
                    </tr>
                </table>
            </div>
        <div class="right" style="padding-left: 5px;">
            <div class="left" style="padding-top: 10px;">
                <input type="submit" name="button" id="button" value="登録" <%=button_submit_disable%> style="width:50px;height:30px;">
                <input type="hidden" name="MM_update" value="form1">
            </div>
            <div class="right" style="padding-top: 5px; padding-left: 5px; width:570px;">
                <% If workshift = "0" Or workshift = "8" Or workshift = "9" Then ' 一般社員(お客さまセンターオペレータ以外)のとき) %>
                    <% If is_operator Then %>
                        &nbsp;<br>&nbsp;<br>
                    <% Else %>
                        <div style="font-size:8pt;">有休残日数は、時間有休を集計していません。
                        <% If workshift <> "9" Then ' 一般社員(お客さまセンターオペレータ以外)のとき) %>
                        <!-- #include file="message.asp" -->
                        <% End If %>
                        <br>深夜時間帯とは22：00～05：00の時間になります。</div>
                    <% End If %>
                <% End If %>
            </div>
        </div>
        </div
    </div>
    </form>
</div>
</div>
<div style="position:absolute;top:22px;left:180px;width:2000px;">
    <%
    ' 当日時間外14時間超チェック
    If warn_time14over = 1 Then
        errorMsg = errorMsg & "当日時間外労働が36協定の限度時間(14時間)を超えています。"
    End If
    ' 時間外労働計(休出含まず)と週超過労働時間の合計が29超のとき警告
    If sumTotalOvertime + init_weekovertime > 29 Then
        errorMsg = errorMsg + "当月時間外労働が36協定の限度時間(29時間)を超えています。"
    End If
    ' 当年度累積時間外労働時間(休出除く)が176超のとき警告
    If sumTotalOvertime + init_weekovertime + yearlyOvertime > 176 Then
        errorMsg = errorMsg + "当年度累積時間外労働が36協定の限度時間(176時間)を超えています。"
    End If
    ' 当月休日出勤時間が15時間20分を超えているとき警告
    If mm2Float(sumHolidayshifttime) + mm2Float(sumHolidayshiftlate) > 15.4 Then
        errorMsg = errorMsg + "当月休日出勤が36協定の限度時間(15時間20分)を超えています。"
    End If
    ' 当年度累積休日出勤時間が184超のとき警告
    If mm2Float(sumHolidayshifttime) + mm2Float(sumHolidayshiftlate) + yearlyHolidaytime > 184 Then
        errorMsg = errorMsg + "当年度累積休日出勤が36協定の限度時間(184時間)を超えています。"
    End If
    If warn_time14over = 1 Or sumTotalOvertime + init_weekovertime > 29 Or _
       sumTotalOvertime + init_weekovertime + yearlyOvertime > 176 Or _
       mm2Float(sumHolidayshifttime) + mm2Float(sumHolidayshiftlate) > 15.4 Or _
       mm2Float(sumHolidayshifttime) + mm2Float(sumHolidayshiftlate) + yearlyHolidaytime > 184 Then
        errorMsg = errorMsg + "特別条項に該当しているか確認してください。"
    End If
    %>
    <font color="red"><b><%=errorMsg%></b></font>
</div>

</body>
<script type="text/javascript">
function setDivSize(){
    // ------------------------------------------------------------------------
    // ウィンドウサイズから div サイズを設定する関数
    // ------------------------------------------------------------------------
    var size_h;
    size_h = document.body.clientHeight;
    if (size_h < 600) {
        size_h = 320;
    } else {
        size_h = size_h - 250;
    }
    document.getElementById('tablediv').style.height = size_h + "px";
    document.getElementById('tbody').style.height = size_h + "px";
}
setDivSize();
</script>
</html>
<%
Rs_worktbl.Close()
Set Rs_worktbl = Nothing
Rs_timetbl.Close()
Set Rs_timetbl = Nothing
Rs_holidaytbl.Close()
Set Rs_holidaytbl = Nothing
Rs_remainvacationtbl.Close()
Set Rs_remainvacationtbl = Nothing
%>
<!-- #include file="inc/util.asp" -->
