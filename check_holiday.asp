<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' #######################################
' プログラム仕様書
' #######################################
'
' ## プログラム概要 ##
' 閲覧者（上長）と同じ組織に所属する職員の、休暇状況を月別に閲覧するページです。
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
'   日付セル    ：氏名のパーソナルコード、表示月を元にワークテーブルより休暇状況を出力。
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
Dim ymb         '表示年月
Dim i           '繰り返し用日付

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

checkHolidayMM1 = ""        ' 表示月の1カ月後チェック用月
checkHolidayMM3 = Array("") ' 表示月の3カ月後チェック用月配列
' 有給休暇取得チェック用年月の設定
If        dispMonth = "01" Then
    checkHolidayMM1 = "02"
    checkHolidayMM3 = Array("02","03","04")
ElseIf    dispMonth = "02" Then
    checkHolidayMM1 = "03"
    checkHolidayMM3 = Array("03","04","05")
ElseIf    dispMonth = "03" Then
    checkHolidayMM1 = "04"
    checkHolidayMM3 = Array("04","05","06")
ElseIf    dispMonth = "04" Then
    checkHolidayMM1 = "05"
    checkHolidayMM3 = Array("05","06","07")
ElseIf    dispMonth = "05" Then
    checkHolidayMM1 = "06"
    checkHolidayMM3 = Array("06","07","08")
ElseIf    dispMonth = "06" Then
    checkHolidayMM1 = "07"
    checkHolidayMM3 = Array("07","08","09")
ElseIf    dispMonth = "07" Then
    checkHolidayMM1 = "08"
    checkHolidayMM3 = Array("08","09","10")
ElseIf    dispMonth = "08" Then
    checkHolidayMM1 = "09"
    checkHolidayMM3 = Array("09","10","11")
ElseIf    dispMonth = "09" Then
    checkHolidayMM1 = "10"
    checkHolidayMM3 = Array("10","11","12")
ElseIf    dispMonth = "10" Then
    checkHolidayMM1 = "11"
    checkHolidayMM3 = Array("11","12","01")
ElseIf    dispMonth = "11" Then
    checkHolidayMM1 = "12"
    checkHolidayMM3 = Array("12","01","02")
ElseIf    dispMonth = "12" Then
    checkHolidayMM1 = "01"
    checkHolidayMM3 = Array("01","02","03")
End If

' stafftblより、表示スタッフ一覧を取得
Dim Rs_staff
Dim Rs_staff_cmd
Set Rs_staff_cmd = Server.CreateObject ("ADODB.Command")
Rs_staff_cmd.ActiveConnection = MM_workdbms_STRING
Rs_staff_cmd.CommandText = "SELECT stafftbl.personalcode ,stafftbl.staffname " & _
    ",stafftbl.orgcode ,stafftbl.gradecode ,stafftbl.grantdate " & _
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
Rs_work_cmd.CommandText = "SELECT worktbl.personalcode " & _
    ",worktbl.workingdate ,worktbl.morningholiday "      & _
    ",worktbl.afternoonholiday "                         & _
    ",worktbl.morningwork, worktbl.afternoonwork "       & _
    "FROM stafftbl "                                     & _
    "RIGHT OUTER JOIN worktbl "                          & _
    "ON stafftbl.personalcode = worktbl.personalcode "   & _
    "LEFT OUTER JOIN orgtbl ON "                         & _
    "orgtbl.orgcode = stafftbl.orgcode "                 & _
    "WHERE worktbl.workingdate LIKE ? AND "              & _
    "stafftbl.is_input = '1' "                       & _
    "AND stafftbl.is_enable = '1' AND "              & _
    "orgtbl.manageclass = '2' "                          & _
    "AND orgtbl.personalcode = ? "                       & _
    "ORDER BY stafftbl.orgcode, stafftbl.gradecode "     & _
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
        勤務表確認対象者　休暇確認　　
        <a href="check_holiday.asp?ymb=<%=prevYmb%>">&lt;&lt;</a>&nbsp;
        <%=dispYear%>年<%=dispMonth%>月&nbsp;
        <a href="check_holiday.asp?ymb=<%=nextYmb%>">&gt;&gt;</a>
        　　<a href="checklist.asp?ymb=<%=dispYear & dispMonth%>">上長チェック</a>
        　　<a href="check_holiday.asp?ymb=<%=dispYear & dispMonth%>">休暇確認</a>
        　　<a href="check_overtime.asp?ymb=<%=dispYear & dispMonth%>">時間外確認</a>
        　　<a href="check_holidaywork.asp?ymb=<%=dispYear & dispMonth%>">休出確認</a>
    </p>
    <div id="tablediv" class="clear" style="width:1510px;">
        <table class="data">
            <tr>
                <th width="150px;" scope="col">氏名</th>
                <th width="40px;" scope="col">有休<br />付与日</th>
                <th width="40px;" scope="col">取得<br />有休</th>
                <th width="40px;" scope="col">有給残</th>
                <th width="40px;" scope="col">振休残</th>
                <th width="40px;" scope="col">保休残</th>
                <th width="40px;" scope="col">時間<br />有休</th>
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
        <div id="tbody"  class="tBody" style="width:1510px;height:100%;">
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
                                                                '- floatTime2min(Rs_dutyrostertbl.Fields.Item("workholidaytime").Value) '  - 時間代休
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
                        <td width="40px;" nowrap style="text-align:center;">
                            <% ' 有休付与日
                            Response.Write(Left(Rs_staff.Fields.Item("grantdate").Value,2)*1 & "/" & Right(Rs_staff.Fields.Item("grantdate").Value,2)*1)
                            %>
                        </td>
                        <% ' 有休取得
                        searchGrantDate1 = ""   ' 検索開始日付
                        searchGrantDate2 = ""   ' 検索終了日付
                        If (Right(ymb, 2) & "01" < Rs_staff.Fields.Item("grantdate").Value) Then
                            ' 有休付与日は前年
                            searchGrantDate1 = Left(ymb,4) - 1 & Trim(Rs_staff.Fields.Item("grantdate").Value)
                        Else
                            ' 有休付与日は当年
                            searchGrantDate1 = Left(ymb,4) & Trim(Rs_staff.Fields.Item("grantdate").Value)
                        End If
                        ' 検索対象終了日付
                        If (Right(ymb, 2) = Left(Rs_staff.Fields.Item("grantdate").Value, 2) And _
                            Right(Rs_staff.Fields.Item("grantdate").Value, 2) > "01") Then
                            searchGrantDate2 = ymb & Right("0" & Right(Rs_staff.Fields.Item("grantdate").Value, 2) - 1, 2)
                        Else
                            searchGrantDate2 = ymb & "99"
                        End If
                        holidayCount = 0
                        ' 午前有休集計
                        Set Rs_worktbl_cmd = Server.CreateObject ("ADODB.Command")
                        Rs_worktbl_cmd.ActiveConnection = MM_workdbms_STRING
                        Rs_worktbl_cmd.CommandText = "SELECT COUNT(morningholiday) AS holiday FROM worktbl " & _
                            "WHERE morningholiday IN ('3', '9') AND personalcode=? AND workingdate>=? AND workingdate<=?"
                        Rs_worktbl_cmd.Prepared = true
                        Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param1", 200, 1, 5, Trim(Rs_staff.Fields.Item("personalcode").Value))
                        Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param2", 200, 1, 8, searchGrantDate1)
                        Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param3", 200, 1, 8, searchGrantDate2)
                        Set Rs_worktbl = Rs_worktbl_cmd.Execute
                        Rs_worktbl_numRows = 0
                        If Rs_worktbl.BOF And Rs_worktbl.EOF Then
                        Else
                            holidayCount = Rs_worktbl.Fields.Item("holiday").Value
                        End If
                        Rs_worktbl.Close()
                        Set Rs_worktbl = Nothing
                        ' 午後有休集計
                        Set Rs_worktbl_cmd = Server.CreateObject ("ADODB.Command")
                        Rs_worktbl_cmd.ActiveConnection = MM_workdbms_STRING
                        Rs_worktbl_cmd.CommandText = "SELECT COUNT(afternoonholiday) AS holiday FROM worktbl " & _
                            "WHERE afternoonholiday='3' AND personalcode=? AND workingdate>=? AND workingdate<=?"
                        Rs_worktbl_cmd.Prepared = true
                        Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param1", 200, 1, 5, Trim(Rs_staff.Fields.Item("personalcode").Value))
                        Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param2", 200, 1, 8, searchGrantDate1)
                        Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param3", 200, 1, 8, searchGrantDate2)
                        Set Rs_worktbl = Rs_worktbl_cmd.Execute
                        Rs_worktbl_numRows = 0
                        If Rs_worktbl.BOF And Rs_worktbl.EOF Then
                        Else
                            holidayCount = holidayCount + Rs_worktbl.Fields.Item("holiday").Value
                        End If
                        Rs_worktbl.Close()
                        Set Rs_worktbl = Nothing
                        
                        
                        ' 有休取得の警告表示設定
                        classholiday = ""
                        If (holidayCount / 2 < 5) Then
                            If Left(Rs_staff.Fields.Item("grantdate").Value,2) = checkHolidayMM1 Then
                                classholiday = "abnormality"
                            ElseIf UBound(Filter(checkHolidayMM3, Left(Rs_staff.Fields.Item("grantdate").Value, 2))) <> -1 Then
                                classholiday = "warning"
                            End If
                        End If
                        %>
                        <td width="40px;" nowrap style="text-align:right;" class="<%=classholiday%>">
                            <% Response.Write(holidayCount / 2) ' 集計結果を2で割ることで実際の有休日数となる %>
                        </td>
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
                            ' 保存休暇残
                            Set Rs_remainvacationtbl_cmd = Server.CreateObject ("ADODB.Command")
                            Rs_remainvacationtbl_cmd.ActiveConnection = MM_workdbms_STRING
                            Rs_remainvacationtbl_cmd.CommandText = "SELECT r.personalcode, COALESCE(r.remainvacation, 0) AS remainvacation, " & _
                                    "COALESCE(SUM(p.preservevacations), 0) AS preservevacations FROM " & _
                                    "(SELECT * FROM (" & _
                                    " SELECT personalcode, ymb, remainvacation, ROW_NUMBER() OVER (ORDER BY ymb DESC) AS rownum FROM remainvacationtbl " & _
                                    "WHERE personalcode= ? AND ymb <= ? " & _
                                    " ) rv WHERE rownum=1) r " & _
                                    "LEFT JOIN " & _
                                    "(SELECT personalcode, ymb, preservevacations FROM dutyrostertbl WHERE personalcode= ? ) p " & _
                                    "ON r.personalcode=p.personalcode AND r.ymb<=p.ymb AND ? >=p.ymb " & _
                                    "GROUP BY r.personalcode, r.ymb, r.remainvacation"
                            Rs_remainvacationtbl_cmd.Prepared = true
                            Rs_remainvacationtbl_cmd.Parameters.Append Rs_remainvacationtbl_cmd.CreateParameter("param1", 200, 1, 5, Trim(Rs_staff.Fields.Item("personalcode").Value))
                            Rs_remainvacationtbl_cmd.Parameters.Append Rs_remainvacationtbl_cmd.CreateParameter("param2", 200, 1, 6, ymb)
                            Rs_remainvacationtbl_cmd.Parameters.Append Rs_remainvacationtbl_cmd.CreateParameter("param3", 200, 1, 5, Trim(Rs_staff.Fields.Item("personalcode").Value))
                            Rs_remainvacationtbl_cmd.Parameters.Append Rs_remainvacationtbl_cmd.CreateParameter("param4", 200, 1, 6, ymb)
                            Set Rs_remainvacationtbl = Rs_remainvacationtbl_cmd.Execute
                            Rs_remainvacationtbl_numRows = 0
                            If Not Rs_remainvacationtbl.EOF Then
                                remCount = Rs_remainvacationtbl.Fields.Item("remainvacation"   ).Value - _
                                           Rs_remainvacationtbl.Fields.Item("preservevacations").Value
                            Else
                                remCount = 0
                            End If
                            Rs_remainvacationtbl.Close()
                            Set Rs_remainvacationtbl = Nothing
                            Response.Write(remCount)
                            %>
                        </td>
                        <td width="40px;" nowrap style="text-align:right;">
                            <%
                            ' 時間有休
                            Set Rs_worktbl_cmd = Server.CreateObject ("ADODB.Command")
                            Rs_worktbl_cmd.ActiveConnection = MM_workdbms_STRING
                            Rs_worktbl_cmd.CommandText = "SELECT personalcode, workingdate, vacationtime FROM worktbl " & _
                                "WHERE vacationtime > 0 AND personalcode=? AND workingdate LIKE ?"
                            Rs_worktbl_cmd.Prepared = true
                            Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param1", 200, 1, 5, Trim(Rs_staff.Fields.Item("personalcode").Value))
                            Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param2", 200, 1, 8, ymb & "%")
                            Set Rs_worktbl = Rs_worktbl_cmd.Execute
                            Rs_worktbl_numRows = 0
                            vacationtime       = 0
                            While (NOT Rs_worktbl.EOF)
                                vacationtime = vacationtime + hhmm2Float(Rs_worktbl.Fields.Item("vacationtime").Value)
                                Rs_worktbl.MoveNext()
                            Wend
                            Rs_worktbl.Close()
                            Set Rs_worktbl = Nothing
                            Response.Write(vacationtime)
                            %>
                        </td>
                        <%
                        ' 日付部分作成
                        For i = 1 To dispLastDay
                            strTd = "&nbsp;"
                            If Not Rs_work.EOF  Then
                                If Rs_work.Fields.Item("workingdate" ) = dispYear & dispMonth & Right("0" & i, 2) And _
                                   Rs_work.Fields.Item("personalcode") = Rs_staff.Fields.Item("personalcode")     Then
                                    If Trim(Rs_work.Fields.Item("morningwork"     ).Value) <> "1" And _
                                       Trim(Rs_work.Fields.Item("morningwork"     ).Value) <> "2" And _
                                       Trim(Rs_work.Fields.Item("morningwork"     ).Value) <> "3" And _
                                       Trim(Rs_work.Fields.Item("morningwork"     ).Value) <> "4" And _
                                       Trim(Rs_work.Fields.Item("morningwork"     ).Value) <> "5" And _
                                       Trim(Rs_work.Fields.Item("morningwork"     ).Value) <> "6" And _
                                       Trim(Rs_work.Fields.Item("morningwork"     ).Value) <> "9" Then
                                        work_am = "0"
                                    Else
                                        work_am = "1"
                                    End If
                                    If Trim(Rs_work.Fields.Item("afternoonwork"   ).Value) <> "1" And _
                                       Trim(Rs_work.Fields.Item("afternoonwork"   ).Value) <> "2" And _
                                       Trim(Rs_work.Fields.Item("afternoonwork"   ).Value) <> "3" And _
                                       Trim(Rs_work.Fields.Item("afternoonwork"   ).Value) <> "4" And _
                                       Trim(Rs_work.Fields.Item("afternoonwork"   ).Value) <> "5" And _
                                       Trim(Rs_work.Fields.Item("afternoonwork"   ).Value) <> "6" And _
                                       Trim(Rs_work.Fields.Item("afternoonwork"   ).Value) <> "9" Then
                                        work_pm = "0"
                                    Else
                                        work_pm = "1"
                                    End If
                                    
                                    If work_am = "0" And work_pm = "0" Then
                                        If Trim(Rs_work.Fields.Item("morningholiday"  ).Value) = "1" And _
                                           Trim(Rs_work.Fields.Item("afternoonholiday").Value) = "1" Then
                                            strTd = "公"
                                        ElseIf Trim(Rs_work.Fields.Item("morningholiday"  ).Value) <> "0" And _
                                               Trim(Rs_work.Fields.Item("morningholiday"  ).Value) <> "7" And _
                                               Trim(Rs_work.Fields.Item("afternoonholiday").Value) <> "0" And _
                                               Trim(Rs_work.Fields.Item("afternoonholiday").Value) <> "7" Then
                                               strTd = "休"
                                        End If
                                    Else
                                        If work_am = "0" Then
                                            If Trim(Rs_work.Fields.Item("morningholiday"  ).Value) <> "0" And _
                                               Trim(Rs_work.Fields.Item("morningholiday"  ).Value) <> "7" Then
                                                '午前休暇有り
                                                strTd = "AM<br />休"
                                            End If
                                        ElseIf work_pm = "0" Then
                                            If Trim(Rs_work.Fields.Item("afternoonholiday").Value) <> "0" And _
                                               Trim(Rs_work.Fields.Item("afternoonholiday").Value) <> "7" Then
                                                '午後休暇有り
                                                strTd = "PM<br />休"
                                            End If
                                        End if
                                    End If
                                    Rs_work.MoveNext()
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
                            Response.write("<td width='31px;' align='center'  class='" & _
                                                weekClass & "'>" & strTd & "</td>")
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
                    <td width="40px;" nowrap style="text-align:center;">-</td>
                    <td width="40px;" nowrap style="text-align:right;">-</td>
                    <td width="40px;" nowrap style="text-align:right;">-</td>
                    <td width="40px;" nowrap style="text-align:right;"><%=all_holidaynumber%></td>
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
