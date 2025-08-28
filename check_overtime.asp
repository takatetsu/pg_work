<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' #######################################
' プログラム仕様書
' #######################################
'
' ## プログラム概要 ##
' 閲覧者（上長）と同じ組織に所属する職員の、時間外状況を月別に閲覧するページです。
'
' （機能）
' ・勤怠入力フラグが「1:入力不要」、有効フラグが「1:無効データ」の職員は表示されません
' ・初期表示月はシステム日付の月となります
' ・月の左右のリンクで、表示月を遷移します
' ・職員名のリンクで、各職員の勤務表上長チェック画面へ遷移します
'
'
' ## 出力項目 ##
'   表示月      ：サーバシステム日付より出力。
'   氏名        ：セッション情報の組織コードを元にスタッフテーブルより出力。
'                 オプション付のリンクとなっています。
'   日付        ：関数により表示月の月末を計算して出力。
'   日付セル    ：氏名のパーソナルコード、表示月を元にワークテーブルより時間外状況を出力。
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

sumTotalOvertime = 0' 時間外労働計
sumOvertime0 = 0    ' 当月時間外計(休出含む)
sumOvertime1 = 0    ' 前月時間外計(休出含む)
sumOvertime2 = 0    ' 2カ月前時間外計(休出含む)
sumOvertime3 = 0    ' 3カ月前時間外計(休出含む)
sumOvertime4 = 0    ' 4カ月前時間外計(休出含む)
sumOvertime5 = 0    ' 5カ月前時間外計(休出含む)

If (Request.QueryString("ymb")<>"") Then
    dispDate = DateSerial(Mid(Request.QueryString("ymb"), 1, 4), Mid(Request.QueryString("ymb"), 5, 2), 1)
    ymb      = Request.QueryString("ymb")
Else
    dispDate = Date
    ymb      = Year(dispDate) & Right("0" & Month(dispDate), 2)
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

' worktblより、表示スタッフ全員分の worktbl 表示月分を日付順に取得
Dim Rs_work
Dim Rs_work_cmd
Set Rs_work_cmd = Server.CreateObject ("ADODB.Command")
Rs_work_cmd.ActiveConnection = MM_workdbms_STRING
Rs_work_cmd.CommandText = "SELECT worktbl.personalcode "        & _
    ",worktbl.workingdate ,worktbl.morningholiday "             & _
    ",worktbl.afternoonholiday ,worktbl.overtime "              & _
    ",worktbl.overtimelate ,worktbl.holidayshiftovertime "      & _
    ",worktbl.holidayshiftovertimelate ,worktbl.requesttime "   & _
    "FROM stafftbl RIGHT OUTER JOIN worktbl "                   & _
    "ON stafftbl.personalcode = worktbl.personalcode "          & _
    "LEFT OUTER JOIN orgtbl ON "                                & _
    "orgtbl.orgcode = stafftbl.orgcode "                        & _
    "WHERE worktbl.workingdate LIKE ? AND "                     & _
    "stafftbl.is_input = '1' "                              & _
    "AND stafftbl.is_enable = '1' AND "                     & _
    "orgtbl.manageclass = '2' "                                 & _
    "AND orgtbl.personalcode = ? "                              & _
    "ORDER BY stafftbl.orgcode, stafftbl.gradecode "            & _
    "DESC, worktbl.personalcode, worktbl.workingdate"
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
        勤務表確認対象者　時間外確認　
        <a href="check_overtime.asp?ymb=<%=prevYmb%>">&lt;&lt;</a>&nbsp;
        <%=dispYear%>年<%=dispMonth%>月&nbsp;
        <a href="check_overtime.asp?ymb=<%=nextYmb%>">&gt;&gt;</a>
        　　<a href="checklist.asp?ymb=<%=dispYear & dispMonth%>">上長チェック</a>
        　　<a href="check_holiday.asp?ymb=<%=dispYear & dispMonth%>">休暇確認</a>
        　　<a href="check_overtime.asp?ymb=<%=dispYear & dispMonth%>">時間外確認</a>
        　　<a href="check_holidaywork.asp?ymb=<%=dispYear & dispMonth%>">休出確認</a>
    </p>
    <div id="tablediv" class="clear" style="width:1550px;">
        <table class="data">
            <tr>
                <th width="150px;" scope="col">氏名</th>
                <th width="40px;" scope="col">時間外</th>
                <th width="40px;" scope="col">時間外<br />累積</th>
                <th width="40px;" scope="col">2ヶ月<br />平均</th>
                <th width="40px;" scope="col">3ヶ月<br />平均</th>
                <th width="40px;" scope="col">4ヶ月<br />平均</th>
                <th width="40px;" scope="col">5ヶ月<br />平均</th>
                <th width="40px;" scope="col">6ヶ月<br />平均</th>
                <%
                ' 表示月の日付一覧作成
                For i = 1 To dispLastDay
                    Response.write "<th width=""31px;""scope=""col"">"
                    Response.write Right("0" & i, 2)
                    Response.write "</th>"
                Next
                %>
            </tr>
        </table>
        <div id="tbody"  class="tBody" style="width:1550px;height:100%;">
            <table class="data">
                <%
                all_holidaynumber   = 0
                all_overtime        = 0
                all_totalOvertime   = 0
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
                                                                '- floatTime2min(Rs_dutyrostertbl.Fields.Item("workholidaytime").Value) ' - 時間代休
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

                        ' 当月時間外警告表示設定
                        If Rs_dutyrostertbl_sumOvertime >= 1740 Then
                            classTemp = "abnormality"
                        ElseIf Rs_dutyrostertbl_sumOvertime >= 1500 Then
                            classTemp = "warning"
                        Else
                            classTemp = ""
                        End If
                        %>
                        <td width="40px;" nowrap style="text-align:right;" class="<%=classTemp%>">
                            <%
                            If (Rs_dutyrostertbl_sumOvertime = 0) Then
                                Response.Write("&nbsp;")
                            Else
                                Response.Write(min2Time(Rs_dutyrostertbl_sumOvertime))
                            End If
                            %>
                        </td>
                        <% ' 時間外累積
                        ' -----------------------------------------------------------------------------
                        ' 勤務表テーブル dutyrostertbl 読込
                        ' -----------------------------------------------------------------------------
                        ' 対象年度4月算出
                        If Right(ymb, 2) > "03" Then
                            businessYear = Left(ymb, 4) & "04"
                        Else
                            businessYear = (Left(ymb, 4) - 1) & "04"
                        End If
                        Set Rs_dutyrostertbl_cmd = Server.CreateObject ("ADODB.Command")
                        Rs_dutyrostertbl_cmd.ActiveConnection = MM_workdbms_STRING
                        Rs_dutyrostertbl_cmd.CommandText = "SELECT SUM(overtime + holidayshiftovertime + overtimelate + holidayshiftovertimelate) AS sumovertime " & _
                               ",SUM(holidayshifttime + holidayshiftlate) AS sumholidaytime " & _
                               "FROM dutyrostertbl WHERE personalcode = ? AND ymb >= ? AND ymb <= ?"
                        Rs_dutyrostertbl_cmd.Prepared = true
                        Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param1", 200, 1, 5, Trim(Rs_staff.Fields.Item("personalcode").Value))
                        Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param2", 200, 1, 6, businessYear)
                        Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param3", 200, 1, 6, ymb)
                        Set Rs_dutyrostertbl = Rs_dutyrostertbl_cmd.Execute
                        Rs_dutyrostertbl_numRows = 0
                        If Rs_dutyrostertbl.EOF And Rs_dutyrostertbl.BOF Then
                        Else
                            sumTotalOvertime = Rs_dutyrostertbl.Fields.Item("sumovertime").Value '+ Rs_dutyrostertbl.Fields.Item("sumholidaytime").Value 
                        End If
                        Rs_dutyrostertbl.Close()
                        all_totalOvertime = all_totalOvertime + sumTotalOvertime
                        
                        ' 累積時間外警告表示設定
                        If sumTotalOvertime >= 176 Then
                            classTemp = "abnormality"
                        ElseIf sumTotalOvertime >= 150 Then
                            classTemp = "warning"
                        Else
                            classTemp = ""
                        End If
                        %>
                        <td width="40px;" nowrap style="text-align:right;" class="<%=classTemp%>">
                            <%=sumTotalOvertime%>
                        </td>
                        <%
                        ' 時間外計(休出含む)求める(2, 3, 4, 5カ月前)
                        sumOvertime0 = 0
                        baseYMD = CDate(Left(ymb,4) & "/" & Right(ymb,2) & "/01")
                        ' -----------------------------------------------------------------------------
                        ' 勤務表テーブル dutyrostertbl 読込
                        ' -----------------------------------------------------------------------------
                        ' 当月分
                        Set Rs_dutyrostertbl_cmd = Server.CreateObject ("ADODB.Command")
                        Rs_dutyrostertbl_cmd.ActiveConnection = MM_workdbms_STRING
                        Rs_dutyrostertbl_cmd.CommandText = "SELECT personalcode, ymb, overtime + holidayshiftovertime + " & _
                                                           "overtimelate + holidayshiftovertimelate + holidayshifttime + " & _
                                                           "holidayshiftlate AS sumovertime FROM dutyrostertbl " & _
                                                           "WHERE personalcode=? AND ymb=?"
                        Rs_dutyrostertbl_cmd.Prepared = true
                        Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param1", 200, 1, 5, Trim(Rs_staff.Fields.Item("personalcode").Value))
                        Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param2", 200, 1, 6, Year(baseYMD) & Right("0" & Month(baseYMD), 2))
                        Set Rs_dutyrostertbl = Rs_dutyrostertbl_cmd.Execute
                        Rs_dutyrostertbl_numRows = 0
                        If Rs_dutyrostertbl.EOF And Rs_dutyrostertbl.BOF Then
                        Else
                            sumOvertime0 = Rs_dutyrostertbl.Fields.Item("sumovertime").Value 
                        End If
                        Rs_dutyrostertbl.Close()

                        ' 前月分
                        sumOvertime1 = 0
                        baseYMD = DateAdd("m", -1, baseYMD)
                        Set Rs_dutyrostertbl_cmd = Server.CreateObject ("ADODB.Command")
                        Rs_dutyrostertbl_cmd.ActiveConnection = MM_workdbms_STRING
                        Rs_dutyrostertbl_cmd.CommandText = "SELECT personalcode, ymb, overtime + holidayshiftovertime + " & _
                                                           "overtimelate + holidayshiftovertimelate + holidayshifttime + " & _
                                                           "holidayshiftlate AS sumovertime FROM dutyrostertbl " & _
                                                           "WHERE personalcode=? AND ymb=?"
                        Rs_dutyrostertbl_cmd.Prepared = true
                        Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param1", 200, 1, 5, Trim(Rs_staff.Fields.Item("personalcode").Value))
                        Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param2", 200, 1, 6, Year(baseYMD) & Right("0" & Month(baseYMD), 2))
                        Set Rs_dutyrostertbl = Rs_dutyrostertbl_cmd.Execute
                        Rs_dutyrostertbl_numRows = 0
                        If Rs_dutyrostertbl.EOF And Rs_dutyrostertbl.BOF Then
                        Else
                            sumOvertime1 = Rs_dutyrostertbl.Fields.Item("sumovertime").Value 
                        End If
                        Rs_dutyrostertbl.Close()
                        
                        ' 2カ月前分
                        sumOvertime2 = 0
                        baseYMD = DateAdd("m", -1, baseYMD)
                        Set Rs_dutyrostertbl_cmd = Server.CreateObject ("ADODB.Command")
                        Rs_dutyrostertbl_cmd.ActiveConnection = MM_workdbms_STRING
                        Rs_dutyrostertbl_cmd.CommandText = "SELECT personalcode, ymb, overtime + holidayshiftovertime + " & _
                                                           "overtimelate + holidayshiftovertimelate + holidayshifttime + " & _
                                                           "holidayshiftlate AS sumovertime FROM dutyrostertbl " & _
                                                           "WHERE personalcode=? AND ymb=?"
                        Rs_dutyrostertbl_cmd.Prepared = true
                        Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param1", 200, 1, 5, Trim(Rs_staff.Fields.Item("personalcode").Value))
                        Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param2", 200, 1, 6, Year(baseYMD) & Right("0" & Month(baseYMD), 2))
                        Set Rs_dutyrostertbl = Rs_dutyrostertbl_cmd.Execute
                        Rs_dutyrostertbl_numRows = 0
                        If Rs_dutyrostertbl.EOF And Rs_dutyrostertbl.BOF Then
                        Else
                            sumOvertime2 = Rs_dutyrostertbl.Fields.Item("sumovertime").Value 
                        End If
                        Rs_dutyrostertbl.Close()
                        ' 3カ月前分
                        sumOvertime3 = 0
                        baseYMD = DateAdd("m", -1, baseYMD)
                        Set Rs_dutyrostertbl_cmd = Server.CreateObject ("ADODB.Command")
                        Rs_dutyrostertbl_cmd.ActiveConnection = MM_workdbms_STRING
                        Rs_dutyrostertbl_cmd.CommandText = "SELECT personalcode, ymb, overtime + holidayshiftovertime + " & _
                                                           "overtimelate + holidayshiftovertimelate + holidayshifttime + " & _
                                                           "holidayshiftlate AS sumovertime FROM dutyrostertbl " & _
                                                           "WHERE personalcode=? AND ymb=?"
                        Rs_dutyrostertbl_cmd.Prepared = true
                        Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param1", 200, 1, 5, Trim(Rs_staff.Fields.Item("personalcode").Value))
                        Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param2", 200, 1, 6, Year(baseYMD) & Right("0" & Month(baseYMD), 2))
                        Set Rs_dutyrostertbl = Rs_dutyrostertbl_cmd.Execute
                        Rs_dutyrostertbl_numRows = 0
                        If Rs_dutyrostertbl.EOF And Rs_dutyrostertbl.BOF Then
                        Else
                            sumOvertime3 = Rs_dutyrostertbl.Fields.Item("sumovertime").Value 
                        End If
                        Rs_dutyrostertbl.Close()
                        ' 4カ月前分
                        sumOvertime4 = 0
                        baseYMD = DateAdd("m", -1, baseYMD)
                        Set Rs_dutyrostertbl_cmd = Server.CreateObject ("ADODB.Command")
                        Rs_dutyrostertbl_cmd.ActiveConnection = MM_workdbms_STRING
                        Rs_dutyrostertbl_cmd.CommandText = "SELECT personalcode, ymb, overtime + holidayshiftovertime + " & _
                                                           "overtimelate + holidayshiftovertimelate + holidayshifttime + " & _
                                                           "holidayshiftlate AS sumovertime FROM dutyrostertbl " & _
                                                           "WHERE personalcode=? AND ymb=?"
                        Rs_dutyrostertbl_cmd.Prepared = true
                        Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param1", 200, 1, 5, Trim(Rs_staff.Fields.Item("personalcode").Value))
                        Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param2", 200, 1, 6, Year(baseYMD) & Right("0" & Month(baseYMD), 2))
                        Set Rs_dutyrostertbl = Rs_dutyrostertbl_cmd.Execute
                        Rs_dutyrostertbl_numRows = 0
                        If Rs_dutyrostertbl.EOF And Rs_dutyrostertbl.BOF Then
                        Else
                            sumOvertime4 = Rs_dutyrostertbl.Fields.Item("sumovertime").Value 
                        End If
                        Rs_dutyrostertbl.Close()
                        ' 5カ月前分
                        sumOvertime5 = 0
                        baseYMD = DateAdd("m", -1, baseYMD)
                        Set Rs_dutyrostertbl_cmd = Server.CreateObject ("ADODB.Command")
                        Rs_dutyrostertbl_cmd.ActiveConnection = MM_workdbms_STRING
                        Rs_dutyrostertbl_cmd.CommandText = "SELECT personalcode, ymb, overtime + holidayshiftovertime + " & _
                                                           "overtimelate + holidayshiftovertimelate + holidayshifttime + " & _
                                                           "holidayshiftlate AS sumovertime FROM dutyrostertbl " & _
                                                           "WHERE personalcode=? AND ymb=?"
                        Rs_dutyrostertbl_cmd.Prepared = true
                        Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param1", 200, 1, 5, Trim(Rs_staff.Fields.Item("personalcode").Value))
                        Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param2", 200, 1, 6, Year(baseYMD) & Right("0" & Month(baseYMD), 2))
                        Set Rs_dutyrostertbl = Rs_dutyrostertbl_cmd.Execute
                        Rs_dutyrostertbl_numRows = 0
                        If Rs_dutyrostertbl.EOF And Rs_dutyrostertbl.BOF Then
                        Else
                            sumOvertime5 = Rs_dutyrostertbl.Fields.Item("sumovertime").Value 
                        End If
                        Rs_dutyrostertbl.Close()
                        Set Rs_dutyrostertbl = Nothing
                        %>
                        <%
                        ' 2カ月平均時間外警告表示設定
                        If Round((sumOvertime0+sumOvertime1)/2,1) >= 80 Then
                            classTemp = "abnormality"
                        ElseIf Round((sumOvertime0+sumOvertime1)/2,1) >= 70 Then
                            classTemp = "warning"
                        Else
                            classTemp = ""
                        End If
                        %>
                        <td width="40px;" nowrap style="text-align:right;" class="<%=classTemp%>">
                            <% ' 2カ月平均
                            Response.Write(Round((sumOvertime0+sumOvertime1)/2,1))
                            %>
                        </td>
                        <%
                        ' 3カ月平均時間外警告表示設定
                        If Round((sumOvertime0+sumOvertime1+sumOvertime2)/3,1) >= 80 Then
                            classTemp = "abnormality"
                        ElseIf Round((sumOvertime0+sumOvertime1+sumOvertime2)/3,1) >= 70 Then
                            classTemp = "warning"
                        Else
                            classTemp = ""
                        End If
                        %>
                        <td width="40px;" nowrap style="text-align:right;" class="<%=classTemp%>">
                            <% ' 3カ月平均
                            Response.Write(Round((sumOvertime0+sumOvertime1+sumOvertime2)/3,1))
                            %>
                        </td>
                        <%
                        ' 4カ月平均時間外警告表示設定
                        If Round((sumOvertime0+sumOvertime1+sumOvertime2+sumOvertime3)/4,1) >= 80 Then
                            classTemp = "abnormality"
                        ElseIf Round((sumOvertime0+sumOvertime1+sumOvertime2+sumOvertime3)/4,1) >= 70 Then
                            classTemp = "warning"
                        Else
                            classTemp = ""
                        End If
                        %>
                        <td width="40px;" nowrap style="text-align:right;" class="<%=classTemp%>">
                            <% ' 4カ月平均
                            Response.Write(Round((sumOvertime0+sumOvertime1+sumOvertime2+sumOvertime3)/4,1))
                            %>
                        </td>
                        <%
                        ' 5カ月平均時間外警告表示設定
                        If Round((sumOvertime0+sumOvertime1+sumOvertime2+sumOvertime3+sumOvertime4)/5,1) >= 80 Then
                            classTemp = "abnormality"
                        ElseIf Round((sumOvertime0+sumOvertime1+sumOvertime2+sumOvertime3+sumOvertime4)/5,1) >= 70 Then
                            classTemp = "warning"
                        Else
                            classTemp = ""
                        End If
                        %>
                        <td width="40px;" nowrap style="text-align:right;" class="<%=classTemp%>">
                            <% ' 5カ月平均
                            Response.Write(Round((sumOvertime0+sumOvertime1+sumOvertime2+sumOvertime3+sumOvertime4)/5,1))
                            %>
                        </td>
                        <%
                        ' 6カ月平均時間外警告表示設定
                        If Round((sumOvertime0+sumOvertime1+sumOvertime2+sumOvertime3+sumOvertime4+sumOvertime5)/6,1) >= 80 Then
                            classTemp = "abnormality"
                        ElseIf Round((sumOvertime0+sumOvertime1+sumOvertime2+sumOvertime3+sumOvertime4+sumOvertime5)/6,1) >= 70 Then
                            classTemp = "warning"
                        Else
                            classTemp = ""
                        End If
                        %>
                        <td width="40px;" nowrap style="text-align:right;" class="<%=classTemp%>">
                            <% ' 6カ月平均
                            Response.Write(Round((sumOvertime0+sumOvertime1+sumOvertime2+sumOvertime3+sumOvertime4+sumOvertime5)/6,1))
                            %>
                        </td>
                        <%
                        ' 日付部分作成
                        For i = 1 To dispLastDay
                            sumOvertime = 0
                            If Not Rs_work.EOF  Then
                                If Rs_work.Fields.Item("workingdate" ) = dispYear & dispMonth & Right("0" & i, 2) And _
                                   Rs_work.Fields.Item("personalcode") = Rs_staff.Fields.Item("personalcode")     Then
                                    If Len(Trim(Rs_work.Fields.Item("overtime").Value)) = 0 Then
                                        workOvertime = 0
                                    Else
                                        workOvertime = time2Min(editTime(Rs_work.Fields.Item("overtime").Value))
                                    End If
                                    If Len(Trim(Rs_work.Fields.Item("overtimelate").Value)) = 0 Then
                                        workOvertimelate = 0
                                    Else
                                        workOvertimelate = _
                                            time2Min(editTime(Rs_work.Fields.Item("overtimelate").Value))
                                    End If
                                    If Len(Trim(Rs_work.Fields.Item("holidayshiftovertime").Value)) = 0 Then
                                        workHolidayshiftovertime = 0
                                    Else
                                        workHolidayshiftovertime = _
                                            time2Min(editTime(Rs_work.Fields.Item("holidayshiftovertime").Value))
                                    End If
                                    If Len(Trim(Rs_work.Fields.Item("holidayshiftovertimelate").Value)) = 0 Then
                                        workHolidayshiftovertimelate = 0
                                    Else
                                        workHolidayshiftovertimelate = _
                                            time2Min(editTime(Rs_work.Fields.Item("holidayshiftovertimelate").Value))
                                    End If
                                    If (Len(Trim(Rs_work.Fields.Item("requesttime").Value)) = 0 Or _
                                       Rs_work.Fields.Item("requesttime").Value = "0000")  Then
                                        workRequesttime = 0
                                    Else
                                        workRequesttime = _
                                            time2Min(editTime(Rs_work.Fields.Item("requesttime").Value))
                                    End If
                                    sumOvertime = workOvertime                  _
                                                + workOvertimelate              _
                                                + workHolidayshiftovertime      _
                                                + workHolidayshiftovertimelate '_
                                                '- workRequesttime

                                    Rs_work.MoveNext()
                                End If
                            End If
                            If Len(Trim(sumOvertime)) = 0 Then
                                sumOvertime = "&nbsp;"
                            Else
                                If sumOvertime = "0" Then
                                    sumOvertime = "&nbsp;"
                                End If
                            End If
                            weekNameKanji = WeekdayName(Weekday(DateSerial(dispYear, dispMonth, i)), true)
                            If weekNameKanji = "日" Then
                                weekClass = "sunday"
                            ElseIf weekNameKanji = "土" Then
                                weekClass = "saturday"
                            Else
                                weekClass = ""
                            End If
                            Response.write("<td width='31px;' align='right' class='" & _
                                                weekClass & "'>" & min2Time(sumOvertime) & "</td>")
                        Next
                        Response.write "</tr>"
                        Rs_staff.MoveNext()
                    Wend
                End If
                Rs_staff.Close()
                Rs_work.Close()
                Set Rs_staff = Nothing
                Set Rs_work = Nothing
                %>
                <tr style="height:45px;">
                    <th width="150px;" nowrap class="permanent" scope="row">計</th>
                    <td width="40px;" nowrap style="text-align:right;"><%=min2Time(all_overtime)%></td>
                    <td width="40px;" nowrap style="text-align:right;"><%=all_totalOvertime%></td>
                    <td width="40px;" nowrap style="text-align:right;">-</td>
                    <td width="40px;" nowrap style="text-align:right;">-</td>
                    <td width="40px;" nowrap style="text-align:right;">-</td>
                    <td width="40px;" nowrap style="text-align:right;">-</td>
                    <td width="40px;" nowrap style="text-align:right;">-</td>
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
