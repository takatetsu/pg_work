<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' #######################################
' プログラム仕様書
' #######################################
'
' ## プログラム概要 ##
' 閲覧者（上長）と同じ組織に所属する職員の、上長チェック状況を月別に閲覧するページです。
' 上長チェックがされている日付セルにはチェックマークが表示されます。
'
' （機能）
' ・勤怠入力フラグが「1:入力不要」、有効フラグが「1:無効データ」の職員は表示されません
' ・毎月１５日までは初期表示月は前月となります
' ・月の左右のリンクで、表示月を遷移します
' ・職員名のリンクで、各職員の勤務表上長チェック画面へ遷移します
'
'
' ## 出力項目 ##
'   表示月      ：サーバシステム日付より出力。毎月１５日までは前月が表示月となります。
'   氏名        ：セッション情報の組織コードを元にスタッフテーブルより出力。
'                 オプション付のリンクとなっています。
'   日付        ：関数により表示月の月末を計算して出力。
'   日付セル    ：氏名のパーソナルコード、表示月を元にワークテーブルより上長チェック状況を出力。
'
' ## 入力チェック ##
'
' ## 注意事項 ##
'
'
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

If (Request.QueryString("ymb")<>"") Then
    dispDate = DateSerial(Mid(Request.QueryString("ymb"), 1, 4), Mid(Request.QueryString("ymb"), 5, 2), 1)
Else
    sysDate  = Date
    dispDate = DateAdd("d", -1, sysDate)    ' 前日日付設定
End If

dispYear    = Year(dispDate)
dispMonth   = Right("0" & Month(dispDate), 2)
dispLastDay = Day(DateAdd ("d", -1, Year(DateAdd("m", 1, dispDate)) & "/" & Right("0" & Month(DateAdd("m", 1, dispDate)), 2) & "/01"))
nextYmb     = Year(DateAdd("m",  1, dispDate)) & Right("0" & Month(DateAdd("m",  1, dispDate)), 2)
prevYmb     = Year(DateAdd("m", -1, dispDate)) & Right("0" & Month(DateAdd("m", -1, dispDate)), 2)

' stafftblより、表示スタッフ一覧を取得
Dim Rs_staff
Dim Rs_staff_cmd
Set Rs_staff_cmd = Server.CreateObject ("ADODB.Command")
Rs_staff_cmd.ActiveConnection = MM_workdbms_STRING
Rs_staff_cmd.CommandText = "SELECT stafftbl.personalcode ,stafftbl.staffname " & _
    ",stafftbl.orgcode , stafftbl.gradecode " & _
    "FROM orgtbl RIGHT OUTER JOIN stafftbl stafftbl " & _
    "ON orgtbl.orgcode = stafftbl.orgcode " & _
    "WHERE stafftbl.is_input = '1' AND stafftbl.is_enable = '1' " & _
    "AND orgtbl.personalcode = ?  AND orgtbl.manageclass = '2' " & _
    "ORDER BY stafftbl.orgcode, stafftbl.gradecode DESC, stafftbl.personalcode"
Rs_staff_cmd.Prepared = true
Rs_staff_cmd.Parameters.Append Rs_staff_cmd.CreateParameter("param1", 200, 1, -1, Session("MM_Username") )
Set Rs_staff = Rs_staff_cmd.Execute

' worktblより、表示スタッフ全員分の上長チェックを日付順に取得
Dim Rs_work
Dim Rs_work_cmd
Set Rs_work_cmd = Server.CreateObject ("ADODB.Command")
Rs_work_cmd.ActiveConnection = MM_workdbms_STRING
Rs_work_cmd.CommandText = "SELECT worktbl.personalcode " & _
    ",worktbl.workingdate " & _
    ",worktbl.is_approval " & _
    ",worktbl.morningholiday " & _
    ",worktbl.morningwork " & _
    ",worktbl.is_error " & _
    "FROM stafftbl " & _
    "RIGHT OUTER JOIN worktbl ON stafftbl.personalcode = worktbl.personalcode " & _
    "LEFT OUTER JOIN orgtbl ON orgtbl.orgcode = stafftbl.orgcode " & _
    "WHERE worktbl.workingdate LIKE ? AND stafftbl.is_input = '1' " & _
    "AND stafftbl.is_enable = '1' AND orgtbl.manageclass = '2' " & _
    "AND orgtbl.personalcode = ? " & _
    "ORDER BY stafftbl.orgcode, stafftbl.gradecode DESC, worktbl.personalcode, worktbl.workingdate"
Rs_work_cmd.Prepared = true
Rs_work_cmd.Parameters.Append Rs_work_cmd.CreateParameter("param1", 200, 1, -1, dispyear & dispmonth & "%")
Rs_work_cmd.Parameters.Append Rs_work_cmd.CreateParameter("param2", 200, 1, -1, Session("MM_Username"))
Set Rs_work = Rs_work_cmd.Execute
%>

<!DOCTYPE HTML>
<html>
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
        勤務表確認対象者　勤務表確認　
        <a href="checklist.asp?ymb=<%=prevYmb%>">&lt;&lt;</a>&nbsp;
        <%=dispYear%>年<%=dispMonth%>月&nbsp;
        <a href="checklist.asp?ymb=<%=nextYmb%>">&gt;&gt;</a>
        　　<a href="checklist.asp?ymb=<%=dispYear & dispMonth%>">上長チェック</a>
        　　<a href="check_holiday.asp?ymb=<%=dispYear & dispMonth%>">休暇確認</a>
        　　<a href="check_overtime.asp?ymb=<%=dispYear & dispMonth%>">時間外確認</a>
        　　<a href="check_holidaywork.asp?ymb=<%=dispYear & dispMonth%>">休出確認</a>
    </p>
    <div id="tablediv" class="clear" style="width:1470px;">
        <table class="data">
            <tr>
                <th width="150px;" scope="col">氏名</th>
                <th width="40px;" scope="col">有給残</th>
                <th width="40px;" scope="col">振休残</th>
                <th width="40px;" scope="col">時間外</th>
                <th width="40px;" scope="col">休出</th>
                <th width="40px;" scope="col">甲番<br />乙番</th>
                <%
                ' 表示月の日付一覧作成
                For i = 1 To dispLastDay
                    Response.write "<th width='31px;' scope='col'>"
                    Response.write Right("0" & i, 2)
                    Response.write "</th>"
                Next
                %>
            </tr>
        </table>
        <div id="tbody"  class="tBody" style="width:1470px;height:100%;">
            <table class="data">
                <%
                all_holidaynumber   = 0
                all_overtime        = 0
                all_holidayshift    = 0
                all_shiftwork_kou   = 0
                all_shiftwork_otsu  = 0
                
                ' 職員とチェック状況の表示
                If Not Rs_staff.EOF Or Not Rs_staff.BOF Then
                    While (NOT Rs_staff.EOF)
                %>
                        <tr style="height:45px;">
                        <th width="150px;" nowrap class="permanent" scope="row">
                           <a href="inputwork.asp?p=<%=Rs_staff.Fields.Item("personalcode")%>&ymb=<%=dispYear & Right("0" & dispMonth, 2)%>&s=1">
                               <%=RTrim(Rs_staff.Fields.Item("staffname"))%>
                           </a>
                        </th>
                        <%
                        ' 最新の勤務表テーブルを読み、当月末有給休暇残、当月末振替休日残、時間外を取得
                        Dim Rs_dutyrostertbl
                        Dim Rs_dutyrostertbl_cmd
                        Dim Rs_dutyrostertbl_numRows
                        Set Rs_dutyrostertbl_cmd = Server.CreateObject ("ADODB.Command")
                        Rs_dutyrostertbl_cmd.ActiveConnection = MM_workdbms_STRING
                        Rs_dutyrostertbl_cmd.CommandText = "SELECT * FROM dutyrostertbl " & _
                            "WHERE personalcode = ? AND ymb <= ? ORDER BY ymb DESC"
                        Rs_dutyrostertbl_cmd.Prepared = true
                        Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param1", 200, 1, 5, Trim(Rs_staff.Fields.Item("personalcode").Value))
                        Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param2", 200, 1, 6, dispYear & Right("0" & dispMonth, 2))
                        Set Rs_dutyrostertbl = Rs_dutyrostertbl_cmd.Execute
                        Rs_dutyrostertbl_numRows = 0
                        If Rs_dutyrostertbl.BOF And Rs_dutyrostertbl.EOF Then
                            Rs_dutyrostertbl_vacationnumber  = 0
                            Rs_dutyrostertbl_holidaynumber   = 0
                            Rs_dutyrostertbl_sumOvertime     = 0
                            Rs_dutyrostertbl_sumHolidayshift = 0
                            Rs_dutyrostertbl_shiftwork_kou   = 0
                            Rs_dutyrostertbl_shiftwork_otsu  = 0
                        Else
                            Rs_dutyrostertbl_vacationnumber = Rs_dutyrostertbl.Fields.Item("vacationnumber").Value
                            Rs_dutyrostertbl_holidaynumber  = Rs_dutyrostertbl.Fields.Item("holidaynumber" ).Value
                            If Rs_dutyrostertbl.Fields.Item("ymb" ).Value = (dispYear & Right("0" & dispMonth, 2)) Then
                                ' 時間外労働計 = 時間外 + 時間外深夜業 + 休出時間外 + 休出時間外深夜
                                Rs_dutyrostertbl_sumOvertime    = floatTime2min(Rs_dutyrostertbl.Fields.Item("overtime").Value) _
                                                                + floatTime2min(Rs_dutyrostertbl.Fields.Item("overtimelate").Value) _
                                                                + floatTime2min(Rs_dutyrostertbl.Fields.Item("holidayshiftovertime").Value) _
                                                                + floatTime2min(Rs_dutyrostertbl.Fields.Item("holidayshiftovertimelate").Value)
                                Rs_dutyrostertbl_sumHolidayshift= Rs_dutyrostertbl.Fields.Item("holidayshifttime").Value _
                                                                + Rs_dutyrostertbl.Fields.Item("holidayshiftlate").Value
                                Rs_dutyrostertbl_shiftwork_kou  = Rs_dutyrostertbl.Fields.Item("shiftwork_kou").Value
                                Rs_dutyrostertbl_shiftwork_otsu = Rs_dutyrostertbl.Fields.Item("shiftwork_otsu").Value
                            Else
                                Rs_dutyrostertbl_sumOvertime    = 0
                                Rs_dutyrostertbl_sumHolidayshift= 0
                                Rs_dutyrostertbl_shiftwork_kou  = 0
                                Rs_dutyrostertbl_shiftwork_otsu = 0
                            End If
                        End If
                        Rs_dutyrostertbl.Close()
                        Set Rs_dutyrostertbl = Nothing
                        
                        all_holidaynumber   = all_holidaynumber  + Rs_dutyrostertbl_holidaynumber
                        all_overtime        = all_overtime       + Rs_dutyrostertbl_sumOvertime
                        all_holidayshift    = all_holidayshift   + Rs_dutyrostertbl_sumHolidayshift
                        all_shiftwork_kou   = all_shiftwork_kou  + Rs_dutyrostertbl_shiftwork_kou
                        all_shiftwork_otsu  = all_shiftwork_otsu + Rs_dutyrostertbl_shiftwork_otsu
                        
                        %>
                        <td width="40px;" nowrap style="text-align:right;">
                            <%=Rs_dutyrostertbl_vacationnumber%>
                        </td>
                        <td width="40px;" nowrap style="text-align:right;">
                            <%
                            If (Rs_dutyrostertbl_holidaynumber = 0) Then
                                Response.Write("&nbsp;")
                            Else
                                Response.Write(Rs_dutyrostertbl_holidaynumber)
                            End If
                            %>
                        </td>
                        <td width="40px;" nowrap style="text-align:right;">
                            <%
                            If (Rs_dutyrostertbl_sumOvertime = 0) Then
                                Response.Write("&nbsp;")
                            Else
                                Response.Write(min2Time(Rs_dutyrostertbl_sumOvertime))
                            End If
                            %>
                        </td>
                        <td width="40px;" nowrap style="text-align:right;">
                            <%
                            If (Rs_dutyrostertbl_sumHolidayshift = 0) Then
                                Response.Write("&nbsp;")
                            Else
                                Response.Write(min2Time(floatTime2min(Rs_dutyrostertbl_sumHolidayshift)))
                            End If
                            %>
                            <br />
                            <%
                            temp = floatTime2min(Rs_dutyrostertbl_sumHolidayshift)
                            If temp = 0 Then
                                Response.Write("&nbsp;")
                            Else
                                Response.Write("(" & mm2FloatDay(floatTime2min(Rs_dutyrostertbl_sumHolidayshift)) & ")")
                            End If
                            %>
                        </td>
                        <td width="40px;" nowrap style="text-align:right;">
                            <% If ((Rs_dutyrostertbl_shiftwork_kou + Rs_dutyrostertbl_shiftwork_otsu) > 0) Then %>
                                <%=Rs_dutyrostertbl_shiftwork_kou%><br /><%=Rs_dutyrostertbl_shiftwork_otsu%>
                            <% Else %>
                                &nbsp;<br />&nbsp;
                            <% End If %>
                        </td>
                        <%
                        ' 日付部分作成
                        For i = 1 To dispLastDay
                            weekNameKanji = WeekdayName(Weekday(DateSerial(dispYear, dispMonth, i)), true)
                            If weekNameKanji = "日" Then
                                weekClass = "sunday"
                            ElseIf weekNameKanji = "土" Then
                                weekClass = "saturday"
                            Else
                                weekClass = ""
                            End If
                            strTdNodate      = "<td width='31px;' align='center' class='" & _
                                                weekClass & "'>&nbsp;</td>"
                            strTdChecked     = "<td width='31px;' align='center' class='" & _
                                                weekClass & "' valign='middle'><font size='4'>◯</font></td>"
                            strTdNoneChecked = "<td width='31px;' align='center' class='" & _
                                                weekClass & "' valign='middle'><font size='4'>✕</font></td>"
                            strTdError       = "<td width='31px;' align='center' class='" & _
                                                weekClass & "' valign='middle'><font size='4' color='#FF0000'>✕</font></td>"
                            If Not Rs_work.EOF  Then
                                If Rs_work.Fields.Item("workingdate" ) = dispYear & dispMonth & Right("0" & i, 2) And _
                                   Rs_work.Fields.Item("personalcode") = Rs_staff.Fields.Item("personalcode")     Then
                                    If Rs_work.Fields.Item("is_approval").Value = "1" Then
                                        '上長チェックが済の場合
                                        Response.write(strTdChecked)
                                    Else
                                        If Rs_staff.Fields.Item("gradecode").Value >= "033" Or _
                                           Rs_staff.Fields.Item("gradecode").Value  = "000" Then
                                            ' 課長以上または等級コード01のとき
                                            Response.write(strTdNodate)
                                        Else
                                            '上長チェックがされてない場合
                                            If (Rs_work.Fields.Item("morningholiday").Value = "0"  And _
                                                Rs_work.Fields.Item("morningwork"   ).Value = "0") Then
                                                ' 空データのとき
                                                Response.write(strTdNodate)
                                            Else
                                                ' 空データでないとき
                                                If Rs_work.Fields.Item("is_error"   ).Value = "1" Then
                                                    ' 労働時間適正化エラーのとき
                                                    Response.write(strTdError      )
                                                Else
                                                    ' 労働時間適正化エラーでないとき
                                                    Response.write(strTdNoneChecked)
                                                End If
                                            End If
                                        End If
                                    End If
                                    Rs_work.MoveNext()
                                Else
                                    'worktblのデータと日付セルの情報が一致しない場合
                                    Response.write(strTdNodate)
                                End If
                            Else
                                'worktblにデータが存在しない場合
                                Response.write(strTdNodate)
                            End If
                        Next
                        Response.write "</tr>"
                        Rs_staff.MoveNext()
                    Wend
                End If
                Rs_staff.Close()
                Set Rs_staff = Nothing
                Rs_work.Close()
                Set Rs_work = Nothing
                %>
                <tr>
                    <th width="150px;" nowrap class="permanent" scope="row">計</th>
                    <td width="40px;" nowrap style="text-align:right;">-</td>
                    <td width="40px;" nowrap style="text-align:right;"><%=all_holidaynumber%></td>
                    <td width="40px;" nowrap style="text-align:right;"><%=min2Time(all_overtime)%></td>
                    <td width="40px;" nowrap style="text-align:right;">
                        <%=min2Time(floatTime2min(all_holidayshift))%><br />
                        <%
                        temp = floatTime2min(all_holidayshift)
                        If temp = 0 Then
                            Response.Write("(0)")
                        Else
                            Response.Write("(" & mm2FloatDay(floatTime2min(all_holidayshift)) & ")")
                        End If
                        %>
                    </td>
                    <td width="40px;" nowrap style="text-align:right;"><%=all_shiftwork_kou%><br /><%=all_shiftwork_otsu%></td>
                    <%
                    For i = 1 To dispLastDay
                        weekNameKanji = WeekdayName(Weekday(DateSerial(dispYear, dispMonth, i)), true)
                        If weekNameKanji = "日" Then
                            weekClass = "sunday"
                        ElseIf weekNameKanji = "土" Then
                            weekClass = "saturday"
                        Else
                            weekClass = ""
                        End If
                        Response.write("<td width='31px;' align='center' class='" & weekClass & "'>-</td>")
                    Next
                    %>
                </tr>
            </table>
        </div>
    </div>
</div>
<!-- #include file="inc/footer.source" -->
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
<!-- #include file="inc/util.asp" -->
