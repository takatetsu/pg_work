<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="Connections/workdbms.asp" -->
<!-- #include file="inc/RestrictAccess.asp" -->
<%
' -----------------------------------------------------------------------------
' 初期処理
' -----------------------------------------------------------------------------
' 固定値
Dim strErrorMsg
strErrorMsg = "入力内容に誤りがあります。確認してください。"

Dim errorMsg    ' エラーメッセージ
' 日付計算
Dim sysDate     ' システム日付
Dim dispDate    ' 表示用日付
Dim dispYear    ' 表示用年 yyyy
Dim dispMonth   ' 表示用月 mm
Dim lastDay     ' 対象年月末日
Dim lastYmb     ' 対象年月の前月'

Dim v_workdays

errorMsg = ""
If (Request.QueryString("ymb")<>"") Then
    dispDate = DateSerial(                             _
                Mid(Request.QueryString("ymb"), 1, 4), _
                Mid(Request.QueryString("ymb"), 5, 2), _
                1)
Else
    dispDate = Date
End If
dispYear  = Year(dispDate)
dispMonth = Right("0" & Month(dispDate), 2)
lastDay   = right(DateSerial(dispYear, dispMonth + 1, 0), 2)
' 対象月前月設定
temp      = DateSerial(dispYear, dispMonth, 0)
lastYmb   = left(temp, 4) & mid(temp, 6, 2)
' 対象月翌月設定
temp      = DateSerial(dispYear, dispMonth , 32)
nextYmb   = left(temp, 4) & mid(temp, 6, 2)

' 入力可能月（システム日付の前月まで）
inputMaxYmb = Year(Date) & Right("0" & Month(Date), 2)
temp        = DateSerial(Year(Date), Month(Date), 0)
inputMinYmb = Left(temp, 4) & Mid(temp, 6, 2)
inputDisable = ""
If (dispYear & dispMonth) < inputMinYmb Then
'    inputDisable = "Disabled"
End If
If (dispYear & dispMonth) >= inputMaxYmb Then
'    inputDisable = "Disabled"
End If

' -----------------------------------------------------------------------------
' 給与担当者が管理する対象者一覧に上長チェックされていないデータがないか確認
' -----------------------------------------------------------------------------
Dim Rs_counttbl
Dim Rs_counttbl_cmd
Dim Rs_counttbl_numRows
Set Rs_counttbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_counttbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_counttbl_cmd.CommandText = "SELECT COALESCE(SUM(count),0) AS count FROM " & _
    "(SELECT orgcode FROM orgtbl WHERE personalcode='" & Session("MM_Username") & _
    "' AND manageclass='1') ORG " & _
    "LEFT JOIN " & _
    "(SELECT personalcode, gradecode, orgcode FROM stafftbl " & _
    "WHERE is_enable='1' AND gradecode<'033' And gradecode != '000') STAFF " & _
    "ON ORG.orgcode=STAFF.orgcode " & _
    "LEFT JOIN " & _
    "(SELECT personalcode AS pcode, COUNT(*) AS count FROM worktbl " & _
    "WHERE workingdate LIKE '" & dispYear & dispMonth & "%' AND is_approval = '0' " & _
    "group by personalcode) APPROVAL " & _
    "ON STAFF.personalcode=APPROVAL.pcode"
Rs_counttbl_cmd.Prepared = true

Set Rs_counttbl = Rs_counttbl_cmd.Execute
Rs_counttbl_numRows = 0
If Not Rs_counttbl.EOF Or Not Rs_counttbl.BOF Then
    If CLng(Rs_counttbl.Fields.Item("count").Value) > 0 Then
        inputDisable = "Disabled"
    End If
End If
Rs_counttbl.Close()
Set Rs_counttbl = Nothing


'Option Explicit
Dim style_workdays                ()
Dim style_workholidays            ()
Dim style_absencedays             ()
Dim style_paidvacations           ()
Dim style_preservevacations       ()
Dim style_specialvacations        ()
Dim style_holidayshifts           ()
Dim style_realworkdays            ()
Dim style_shortdays               ()
Dim style_nightduty_a             ()
Dim style_nightduty_b             ()
Dim style_nightduty_c             ()
Dim style_nightduty_d             ()
Dim style_holidaypremium          ()
Dim style_dayduty                 ()
Dim style_shiftwork_kou           ()
Dim style_shiftwork_otsu          ()
Dim style_shiftwork_hei           ()
Dim style_shiftwork_a             ()
Dim style_shiftwork_b             ()
Dim style_summons                 ()
Dim style_summonslate             ()
Dim style_yearend1230             ()
Dim style_yearend1231             ()
Dim style_workholidaytime         ()
Dim style_latepremium             ()
Dim style_overtime                ()
Dim style_holidayshifttime        ()
Dim style_holidayshiftovertime    ()
Dim style_holidayshiftlate        ()
Dim style_overtimelate            ()
Dim style_holidayshiftovertimelate()
Dim style_saturdayworkmin         ()
Dim style_weekdaysworkmin         ()
Dim style_workingmins             ()
Dim style_currentworkmin          ()
Dim style_legalholiday_extra_min  ()
Dim style_weekovertime            ()
ReDim Preserve style_workdays                (0)
ReDim Preserve style_workholidays            (0)
ReDim Preserve style_absencedays             (0)
ReDim Preserve style_paidvacations           (0)
ReDim Preserve style_preservevacations       (0)
ReDim Preserve style_specialvacations        (0)
ReDim Preserve style_holidayshifts           (0)
ReDim Preserve style_realworkdays            (0)
ReDim Preserve style_shortdays               (0)
ReDim Preserve style_nightduty_a             (0)
ReDim Preserve style_nightduty_b             (0)
ReDim Preserve style_nightduty_c             (0)
ReDim Preserve style_nightduty_d             (0)
ReDim Preserve style_holidaypremium          (0)
ReDim Preserve style_dayduty                 (0)
ReDim Preserve style_shiftwork_kou           (0)
ReDim Preserve style_shiftwork_otsu          (0)
ReDim Preserve style_shiftwork_hei           (0)
ReDim Preserve style_shiftwork_a             (0)
ReDim Preserve style_shiftwork_b             (0)
ReDim Preserve style_summons                 (0)
ReDim Preserve style_summonslate             (0)
ReDim Preserve style_yearend1230             (0)
ReDim Preserve style_yearend1231             (0)
ReDim Preserve style_workholidaytime         (0)
ReDim Preserve style_latepremium             (0)
ReDim Preserve style_overtime                (0)
ReDim Preserve style_holidayshifttime        (0)
ReDim Preserve style_holidayshiftovertime    (0)
ReDim Preserve style_holidayshiftlate        (0)
ReDim Preserve style_overtimelate            (0)
ReDim Preserve style_holidayshiftovertimelate(0)
ReDim Preserve style_saturdayworkmin         (0)
ReDim Preserve style_weekdaysworkmin         (0)
ReDim Preserve style_workingmins             (0)
ReDim Preserve style_currentworkmin          (0)
ReDim Preserve style_legalholiday_extra_min  (0)
ReDim Preserve style_weekovertime            (0)

Dim id
Dim personalcode
Dim workdays
Dim workholidays
Dim absencedays
Dim paidvacations
Dim preservevacations
Dim specialvacations
Dim holidayshifts
Dim realworkdays
Dim shortdays
Dim nightduty_a
Dim nightduty_b
Dim nightduty_c
Dim nightduty_d
Dim holidaypremium
Dim dayduty
Dim shiftwork_kou
Dim shiftwork_otsu
Dim shiftwork_hei
Dim shiftwork_a
Dim shiftwork_b
Dim summons
Dim summonslate
Dim yearend1230
Dim yearend1231
Dim workholidaytime
Dim latepremium
Dim overtime
Dim holidayshifttime
Dim holidayshiftovertime
Dim holidayshiftlate
Dim overtimelate
Dim holidayshiftovertimelate
Dim vacationnumber
Dim holidaynumber
Dim vacationtime
Dim saturdayworkmin
Dim weekdaysworkmin
Dim workingmins
Dim currentworkmin
Dim legalholiday_extra_min
Dim weekovertime

If (CStr(Request("MM_update")) = "form1") Then
    ' -------------------------------------------------------------------------
    ' 更新処理
    ' -------------------------------------------------------------------------
    ' -------------------------------------------------------------------------
    ' 入力チェック
    ' -------------------------------------------------------------------------
    For i = 1 To Request.Form("personalcode").count Step 1
        setData()
    Next

    ' 入力チェックでエラーが無いとき、dutyrostertbl の更新処理を行う。
    If (errorMsg = "") Then
        For i = 1 To Request.Form("personalcode").count Step 1
            setData()
            ' stafftbl の締め日付を更新
            Set MM_editCmd = Server.CreateObject ("ADODB.Command")
            MM_editCmd.ActiveConnection = MM_workdbms_STRING
            MM_editCmd.CommandText = "UPDATE stafftbl SET processed_ymb = '" & _
                                        dispYear & dispMonth & _
                                        "' WHERE personalcode = '" & personalcode & "'"
            MM_editCmd.Prepared = true
            MM_editCmd.Execute
            MM_editCmd.ActiveConnection.Close

            If id = "" Then
                ' -------------------------------------------------------------
                ' データ登録処理 INSERT
                ' -------------------------------------------------------------
                Set MM_editCmd = Server.CreateObject ("ADODB.Command")
                MM_editCmd.ActiveConnection = MM_workdbms_STRING

                MM_editCmd.CommandText = "INSERT INTO dutyrostertbl VALUES(DEFAULT, " & _
                                        "?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, " & _
                                        "?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, " & _
                                        "?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                MM_editCmd.Prepared = true
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  5, personalcode)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  6, dispYear & dispMonth)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, workdays)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, workholidays)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, absencedays)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, paidvacations)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, preservevacations)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, specialvacations)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, holidayshifts)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, realworkdays)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, shortdays)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, nightduty_a)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, nightduty_b)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, nightduty_c)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, nightduty_d)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, holidaypremium)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, dayduty)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, shiftwork_kou)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, shiftwork_otsu)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, shiftwork_hei)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, summons)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, summonslate)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, yearend1230)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, yearend1231)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, workholidaytime)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, latepremium)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, overtime)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, holidayshifttime)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, holidayshiftovertime)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, holidayshiftlate)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, overtimelate)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, holidayshiftovertimelate)
                If IsNumeric(vacationnumber) Then
                    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, vacationnumber)
                Else
                    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, 0)
                End if
                If IsNumeric(holidaynumber) Then
                    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, holidaynumber)
                Else
                    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, 0)
                End if
                If IsNumeric(vacationtime) Then
                    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, (vacationtime * 1))
                Else
                    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, 0)
                End If
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, shiftwork_a)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, shiftwork_b)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, saturdayworkmin)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, weekdaysworkmin)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, workingmins)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, currentworkmin)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, legalholiday_extra_min)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, weekovertime)
                MM_editCmd.Execute
                MM_editCmd.ActiveConnection.Close
            Else
                ' -------------------------------------------------------------
                ' データ更新処理 UPDATE
                ' -------------------------------------------------------------
                Set MM_editCmd = Server.CreateObject ("ADODB.Command")
                MM_editCmd.ActiveConnection = MM_workdbms_STRING

                MM_editCmd.CommandText = "UPDATE dutyrostertbl SET "   & _
                                         "personalcode              = ?, " & _
                                         "ymb                       = ?, " & _
                                         "workdays                  = ?, " & _
                                         "workholidays              = ?, " & _
                                         "absencedays               = ?, " & _
                                         "paidvacations             = ?, " & _
                                         "preservevacations         = ?, " & _
                                         "specialvacations          = ?, " & _
                                         "holidayshifts             = ?, " & _
                                         "realworkdays              = ?, " & _
                                         "shortdays                 = ?, " & _
                                         "nightduty_a               = ?, " & _
                                         "nightduty_b               = ?, " & _
                                         "nightduty_c               = ?, " & _
                                         "nightduty_d               = ?, " & _
                                         "holidaypremium            = ?, " & _
                                         "dayduty                   = ?, " & _
                                         "shiftwork_kou             = ?, " & _
                                         "shiftwork_otsu            = ?, " & _
                                         "shiftwork_hei             = ?, " & _
                                         "summons                   = ?, " & _
                                         "summonslate               = ?, " & _
                                         "yearend1230               = ?, " & _
                                         "yearend1231               = ?, " & _
                                         "workholidaytime           = ?, " & _
                                         "latepremium               = ?, " & _
                                         "overtime                  = ?, " & _
                                         "holidayshifttime          = ?, " & _
                                         "holidayshiftovertime      = ?, " & _
                                         "holidayshiftlate          = ?, " & _
                                         "overtimelate              = ?, " & _
                                         "holidayshiftovertimelate  = ?, " & _
                                         "vacationnumber            = ?, " & _
                                         "holidaynumber             = ?, " & _
                                         "vacationtime              = ?, " & _
                                         "shiftwork_a               = ?, " & _
                                         "shiftwork_b               = ?, " & _
                                         "saturday_workmin          = ?, " & _
                                         "weekdays_workmin          = ?, " & _
                                         "workingmins               = ?, " & _
                                         "currentworkmin            = ?, " & _
                                         "legalholiday_extra_min    = ?, " & _
                                         "weekovertime              = ?  " & _
                                         "WHERE id                  = ?"
                MM_editCmd.Prepared = true
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  5, personalcode)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,201,,  6, dispYear & dispMonth)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, workdays)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, workholidays)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, absencedays)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, paidvacations)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, preservevacations)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, specialvacations)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, holidayshifts)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, realworkdays)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, shortdays)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, nightduty_a)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, nightduty_b)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, nightduty_c)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, nightduty_d)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, holidaypremium)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, dayduty)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, shiftwork_kou)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, shiftwork_otsu)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, shiftwork_hei)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, summons)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, summonslate)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, yearend1230)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, yearend1231)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, workholidaytime)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, latepremium)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, overtime)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, holidayshifttime)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, holidayshiftovertime)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, holidayshiftlate)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, overtimelate)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, holidayshiftovertimelate)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, vacationnumber)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, holidaynumber)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, vacationtime)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, shiftwork_a)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, shiftwork_b)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, saturdayworkmin)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, weekdaysworkmin)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, workingmins)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, currentworkmin)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, legalholiday_extra_min)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, weekovertime)
                MM_editCmd.Parameters.Append MM_editCmd.CreateParameter(,  5,, -1, id)
                MM_editCmd.Execute
                MM_editCmd.ActiveConnection.Close
            End If
        Next
        Response.Redirect("complete.asp")
    End If
End If

' -----------------------------------------------------------------------------
' 入力画面表示処理
' -----------------------------------------------------------------------------

' -----------------------------------------------------------------------------
' 給与担当者が管理する対象者一覧取得
' -----------------------------------------------------------------------------
Dim Rs_worktbl
Dim Rs_worktbl_cmd
Dim Rs_worktbl_numRows

Set Rs_worktbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_worktbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_worktbl_cmd.CommandText = "SELECT * FROM "                                           & _
    "(SELECT orgcode FROM orgtbl "                                                  & _
    "WHERE personalcode='" & Session("MM_Username") & "' AND manageclass='1') ORG "     & _
    "LEFT JOIN "                                                                        & _
    "(SELECT personalcode AS pcode, staffname, orgcode AS org, gradecode AS grade "     & _
    "FROM stafftbl "                                                                & _
    "WHERE is_enable='1') STAFF "                                                       & _
    "ON ORG.orgcode=STAFF.org "                                                         & _
    "LEFT JOIN "                                                                        & _
    "(SELECT * FROM dutyrostertbl WHERE ymb='" & dispYear & dispMonth & "') DUTY "  & _
    "ON STAFF.pcode=DUTY.personalcode "                                                 & _
    "LEFT JOIN "                                                                        & _
    "(SELECT personalcode AS countpcode, COUNT(*) AS count FROM worktbl "           & _
    "WHERE workingdate LIKE '" & dispYear & dispMonth & "%' AND "                       & _
    "is_approval = '1' group by personalcode) APPROVAL "                                & _
    "ON STAFF.pcode=APPROVAL.countpcode "                                               & _
    "WHERE pcode IS NOT NULL "                                                          & _
    "ORDER BY STAFF.org ASC, STAFF.grade DESC, STAFF.pcode ASC"
Rs_worktbl_cmd.Prepared = true

Set Rs_worktbl = Rs_worktbl_cmd.Execute
Rs_worktbl_numRows = 0
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
    <form name="inputAll" method="post" action="">
        <div style="width:1400px;">
            <br />
            勤務表全体入力&nbsp;
            <a href="inputall.asp?ymb=<%=lastYmb%>">&lt;&lt;</a>&nbsp;
            <%=dispYear%>年<%=dispMonth%>月分&nbsp;
            <a href="inputall.asp?ymb=<%=nextYmb%>">&gt;&gt;</a>&nbsp;
            　[<a target="_blank" href="check_holiday_charge.asp?ymb=<%=dispYear & dispMonth%>">休暇確認</a>
            　<a target="_blank" href="check_overtime_charge.asp?ymb=<%=dispYear & dispMonth%>">時間外確認</a>
            　<a target="_blank" href="check_holidaywork_charge.asp?ymb=<%=dispYear & dispMonth%>">休出確認</a>]　
            <input type="button" name="button" id="button" value="登録" onClick="inputSubmit()" <%=inputDisable%>>&nbsp;
            <input type="button" name="buttonData" id="buttonData" value="エクセル" onClick="clickDownloadCSV()">&nbsp;
            <input type="hidden" name="MM_update" value="form1">&nbsp;
            <font color="red"><b><%=errorMsg%></b></font>
        </div>
        <div id="tablediv" class="clear">
            <div class="tHeader" style="width:1990px;">
                <table class="data">
                    <tr>
                        <th nowrap width="38px">個人<br />CD</th>
                        <th nowrap width="110px">氏名</th>
                        <th nowrap width="28px">上長<br />確認</th>
                        <th nowrap width="40px">可出勤<br />日数</th>
                        <th nowrap width="40px">代休<br />日数</th>
                        <th nowrap width="40px">欠勤<br />日数</th>
                        <th nowrap width="40px">有給<br />日数</th>
                        <th nowrap width="40px">保存<br />休暇<br />日数</th>
                        <th nowrap width="40px">特休<br />日数</th>
                        <th nowrap width="40px">休出<br />日数</th>
                        <th nowrap width="40px">実<br />出勤<br />日数</th>
                        <th nowrap width="40px">遅早<br />回数</th>
                        <th nowrap width="40px">宿直A<br />回数</th>
                        <th nowrap width="40px">宿直B<br />回数</th>
                        <th nowrap width="40px">宿直C<br />回数</th>
                        <th nowrap width="40px">宿直D<br />回数</th>
                        <th nowrap width="40px">休日<br />割増</th>
                        <th nowrap width="40px">日直</th>
                        <th nowrap width="40px">交替<br />甲番</th>
                        <th nowrap width="40px">交替<br />乙番</th>
                        <th nowrap width="40px">交替<br />丙番</th>
                        <th nowrap width="40px">交替<br />A番</th>
                        <th nowrap width="40px">交替<br />B番</th>
                        <th nowrap width="110px">氏名</th>
                        <th nowrap width="40px">呼出<br />通常<br />回数</th>
                        <th nowrap width="40px">呼出<br />深夜<br />回数</th>
                        <th nowrap width="40px">年末<br />年始<br />1230</th>
                        <th nowrap width="40px">年末<br />年始<br />1231</th>
                        <th nowrap width="40px">時間<br />代休</th>
                        <th nowrap width="40px">深夜<br />割増</th>
                        <th nowrap width="40px">時間外</th>
                        <th nowrap width="40px">休日<br />出勤</th>
                        <th nowrap width="40px">休出<br />時間外</th>
                        <th nowrap width="40px">休出<br />深夜</th>
                        <th nowrap width="40px">時間外<br />深夜</th>
                        <th nowrap width="40px">休出<br />時間外<br />深夜</th>
                        <th nowrap width="40px">時間外<br />合計</th>
                        <th nowrap width="40px">土曜<br />時間<br />+100円</th>
                        <th nowrap width="40px">平日<br />労働<br />時間</th>
                        <th nowrap width="40px">労働<br />時間<br />(分)</th>
                        <th nowrap width="40px">基準<br />労働<br />時間<br />(分)</th>
                        <th nowrap width="40px">法定<br />休日<br />割増<br />(分)</th>
                        <th nowrap width="40px">週超過<br />時間</th>
                    </tr>
                </table>
            </div>
            <div id="tbody" class="tBody" style="width:1990px;height:500px;">
                <table id="workdata" class="data">
                    <%
                    i = 0
                    While (NOT Rs_worktbl.EOF)
                        i = i + 1
                        ' -------------------------------------------------
                        ' 上長チェック状況の確認
                        ' -------------------------------------------------
                        If Not Rs_worktbl.EOF Or Not Rs_worktbl.BOF Then
                            If (Rs_worktbl.Fields.Item("grade").Value >= "033" Or _
                                Rs_worktbl.Fields.Item("grade").Value  = "000") Then
                                approval         = "－"
                            Else
                                If (IsNull(Rs_worktbl.Fields.Item("count").Value)) Then
                                    approval     = "×"
                                Else
                                    If (lastDay - CLng(Rs_worktbl.Fields.Item("count").Value) = 0) Then
                                        approval = "○"
                                    Else
                                        approval = "×"
                                    End If
                                End If
                            End If
                        Else
                            approval = "－"
                        End If

                        ' -------------------------------------------------
                        ' 表示項目設定
                        ' 入力チェックエラー時は、前回入力情報を表示
                        ' -------------------------------------------------
                        id           = Rs_worktbl.Fields.Item("id"            ).Value
                        personalcode = Trim(Rs_worktbl.Fields.Item("pcode"    ).Value)
                        staffname    = Trim(Rs_worktbl.Fields.Item("staffname").Value)
                        If (errorMsg <> "") Then
                            ' エラー有り
                            workdays                = Request.Form("workdays"                )(i)
                            workholidays            = Request.Form("workholidays"            )(i)
                            absencedays             = Request.Form("absencedays"             )(i)
                            paidvacations           = Request.Form("paidvacations"           )(i)
                            preservevacations       = Request.Form("preservevacations"       )(i)
                            specialvacations        = Request.Form("specialvacations"        )(i)
                            holidayshifts           = Request.Form("holidayshifts"           )(i)
                            realworkdays            = Request.Form("realworkdays"            )(i)
                            shortdays               = Request.Form("shortdays"               )(i)
                            nightduty_a             = Request.Form("nightduty_a"             )(i)
                            nightduty_b             = Request.Form("nightduty_b"             )(i)
                            nightduty_c             = Request.Form("nightduty_c"             )(i)
                            nightduty_d             = Request.Form("nightduty_d"             )(i)
                            holidaypremium          = Request.Form("holidaypremium"          )(i)
                            dayduty                 = Request.Form("dayduty"                 )(i)
                            shiftwork_kou           = Request.Form("shiftwork_kou"           )(i)
                            shiftwork_otsu          = Request.Form("shiftwork_otsu"          )(i)
                            shiftwork_hei           = Request.Form("shiftwork_hei"           )(i)
                            shiftwork_a             = Request.Form("shiftwork_a"             )(i)
                            shiftwork_b             = Request.Form("shiftwork_b"             )(i)
                            summons                 = Request.Form("summons"                 )(i)
                            summonslate             = Request.Form("summonslate"             )(i)
                            yearend1230             = Request.Form("yearend1230"             )(i)
                            yearend1231             = Request.Form("yearend1231"             )(i)
                            workholidaytime         = Request.Form("workholidaytime"         )(i)
                            latepremium             = Request.Form("latepremium"             )(i)
                            overtime                = Request.Form("overtime"                )(i)
                            holidayshifttime        = Request.Form("holidayshifttime"        )(i)
                            holidayshiftovertime    = Request.Form("holidayshiftovertime"    )(i)
                            holidayshiftlate        = Request.Form("holidayshiftlate"        )(i)
                            overtimelate            = Request.Form("overtimelate"            )(i)
                            holidayshiftovertimelate= Request.Form("holidayshiftovertimelate")(i)
                            saturdayworkmin         = Request.Form("saturdayworkmin"         )(i)
                            weekdaysworkmin         = Request.Form("weekdaysworkmin"         )(i)
                            workingmins             = Request.Form("workingmins"             )(i)
                            currentworkmin          = Request.Form("currentworkmin"          )(i)
                            legalholiday_extra_min  = Request.Form("legalholiday_extra_min"  )(i)
                            weekovertime            = Request.Form("weekovertime"            )(i)
                        Else
                            ' エラー無し(初期表示時)
                            workdays                = Rs_worktbl.Fields.Item("workdays"                ).Value
                            workholidays            = Rs_worktbl.Fields.Item("workholidays"            ).Value
                            absencedays             = Rs_worktbl.Fields.Item("absencedays"             ).Value
                            paidvacations           = Rs_worktbl.Fields.Item("paidvacations"           ).Value
                            preservevacations       = Rs_worktbl.Fields.Item("preservevacations"       ).Value
                            specialvacations        = Rs_worktbl.Fields.Item("specialvacations"        ).Value
                            holidayshifts           = Rs_worktbl.Fields.Item("holidayshifts"           ).Value
                            realworkdays            = Rs_worktbl.Fields.Item("realworkdays"            ).Value
                            shortdays               = Rs_worktbl.Fields.Item("shortdays"               ).Value
                            nightduty_a             = Rs_worktbl.Fields.Item("nightduty_a"             ).Value
                            nightduty_b             = Rs_worktbl.Fields.Item("nightduty_b"             ).Value
                            nightduty_c             = Rs_worktbl.Fields.Item("nightduty_c"             ).Value
                            nightduty_d             = Rs_worktbl.Fields.Item("nightduty_d"             ).Value
                            holidaypremium          = Rs_worktbl.Fields.Item("holidaypremium"          ).Value
                            dayduty                 = Rs_worktbl.Fields.Item("dayduty"                 ).Value
                            shiftwork_kou           = Rs_worktbl.Fields.Item("shiftwork_kou"           ).Value
                            shiftwork_otsu          = Rs_worktbl.Fields.Item("shiftwork_otsu"          ).Value
                            shiftwork_hei           = Rs_worktbl.Fields.Item("shiftwork_hei"           ).Value
                            shiftwork_a             = Rs_worktbl.Fields.Item("shiftwork_a"             ).Value
                            shiftwork_b             = Rs_worktbl.Fields.Item("shiftwork_b"             ).Value
                            summons                 = Rs_worktbl.Fields.Item("summons"                 ).Value
                            summonslate             = Rs_worktbl.Fields.Item("summonslate"             ).Value
                            yearend1230             = Rs_worktbl.Fields.Item("yearend1230"             ).Value
                            yearend1231             = Rs_worktbl.Fields.Item("yearend1231"             ).Value
                            workholidaytime         = Rs_worktbl.Fields.Item("workholidaytime"         ).Value
                            latepremium             = Rs_worktbl.Fields.Item("latepremium"             ).Value
                            ' 管理職は時間外ゼロクリア
                            If Rs_worktbl.Fields.Item("grade").Value >= "033" Then
                                overtime            = 0
                            Else
                                overtime            = Rs_worktbl.Fields.Item("overtime"                ).Value
                            End If
                            ' 休日出勤はゼロクリア対象外
                            holidayshifttime        = Rs_worktbl.Fields.Item("holidayshifttime"        ).Value
                            ' 管理職は休出時間外ゼロクリア
                            If Rs_worktbl.Fields.Item("grade").Value >= "033" Then
                                holidayshiftovertime = 0
                            Else
                                holidayshiftovertime = Rs_worktbl.Fields.Item("holidayshiftovertime"    ).Value
                            End If
                            ' 管理職は休出深夜ゼロクリア
                            If Rs_worktbl.Fields.Item("grade").Value >= "033" Then
                                holidayshiftlate    = 0
                            Else
                                holidayshiftlate    = Rs_worktbl.Fields.Item("holidayshiftlate"        ).Value
                            End If
                            ' 管理職は時間外深夜ゼロクリア
                            If Rs_worktbl.Fields.Item("grade").Value >= "033" Then
                                overtimelate        = 0
                            Else
                                overtimelate        = Rs_worktbl.Fields.Item("overtimelate"            ).Value
                            End If
                            ' 管理職は休出時間外深夜ゼロクリア
                            If Rs_worktbl.Fields.Item("grade").Value >= "033" Then
                                holidayshiftovertimelate = 0
                            Else
                                holidayshiftovertimelate = Rs_worktbl.Fields.Item("holidayshiftovertimelate").Value
                            End If
                            
                            saturdayworkmin         = Rs_worktbl.Fields.Item("saturday_workmin"        ).Value
                            weekdaysworkmin         = Rs_worktbl.Fields.Item("weekdays_workmin"        ).Value
                            workingmins             = Rs_worktbl.Fields.Item("workingmins"             ).Value
                            currentworkmin          = Rs_worktbl.Fields.Item("currentworkmin"          ).Value
                            legalholiday_extra_min  = Rs_worktbl.Fields.Item("legalholiday_extra_min"  ).Value
                            weekovertime            = Rs_worktbl.Fields.Item("weekovertime"            ).Value

                            If workdays                 = 0 Then workdays                 = ""
                            If workholidays             = 0 Then workholidays             = ""
                            If absencedays              = 0 Then absencedays              = ""
                            If paidvacations            = 0 Then paidvacations            = ""
                            If preservevacations        = 0 Then preservevacations        = ""
                            If specialvacations         = 0 Then specialvacations         = ""
                            If holidayshifts            = 0 Then holidayshifts            = ""
                            If realworkdays             = 0 Then realworkdays             = ""
                            If shortdays                = 0 Then shortdays                = ""
                            If nightduty_a              = 0 Then nightduty_a              = ""
                            If nightduty_b              = 0 Then nightduty_b              = ""
                            If nightduty_c              = 0 Then nightduty_c              = ""
                            If nightduty_d              = 0 Then nightduty_d              = ""
                            If holidaypremium           = 0 Then holidaypremium           = ""
                            If dayduty                  = 0 Then dayduty                  = ""
                            If shiftwork_kou            = 0 Then shiftwork_kou            = ""
                            If shiftwork_otsu           = 0 Then shiftwork_otsu           = ""
                            If shiftwork_hei            = 0 Then shiftwork_hei            = ""
                            If shiftwork_a              = 0 Then shiftwork_a              = ""
                            If shiftwork_b              = 0 Then shiftwork_b              = ""
                            If summons                  = 0 Then summons                  = ""
                            If summonslate              = 0 Then summonslate              = ""
                            If yearend1230              = 0 Then yearend1230              = ""
                            If yearend1231              = 0 Then yearend1231              = ""
                            If workholidaytime          = 0 Then workholidaytime          = ""
                            If latepremium              = 0 Then latepremium              = ""
                            If overtime                 = 0 Then overtime                 = ""
                            If holidayshifttime         = 0 Then holidayshifttime         = ""
                            If holidayshiftovertime     = 0 Then holidayshiftovertime     = ""
                            If holidayshiftlate         = 0 Then holidayshiftlate         = ""
                            If overtimelate             = 0 Then overtimelate             = ""
                            If holidayshiftovertimelate = 0 Then holidayshiftovertimelate = ""
                            If saturdayworkmin          = 0 Then saturdayworkmin          = ""
                            If weekdaysworkmin          = 0 Then weekdaysworkmin          = ""
                            If workingmins              = 0 Then workingmins              = ""
                            If currentworkmin           = 0 Then currentworkmin           = ""
                            If legalholiday_extra_min   = 0 Then legalholiday_extra_min   = ""
                            If weekovertime             = 0 Then weekovertime             = ""
                        End If
                        %>
                        <tr>
                            <th nowrap width="38px" class="permanent">
                                <a href="inputwork.asp?p=<%=personalcode%>&ymb=<%=dispYear & dispMonth%>&c=1" target="_blank">
                                    <%=personalcode%>
                                </a>
                                <input type="hidden" name="personalcode" value="<%=personalcode%>">
                                <input type="hidden" name="id" value='<%=id%>'>
                                <input type="hidden" name="vacationnumber" value='<%=Rs_worktbl.Fields.Item("vacationnumber").Value%>'>
                                <input type="hidden" name="holidaynumber" value='<%=Rs_worktbl.Fields.Item("holidaynumber").Value%>'>
                                <input type="hidden" name="vacationtime" value='<%=Rs_worktbl.Fields.Item("vacationtime").Value%>'>
                            </th>
                            <th nowrap width="110px" class="permanent"><%=staffname%></th>
                            <td align="center" nowrap width="28px" ><%=approval%></td>
                            <%
                            If UBound(style_workdays) < i Then
                                style = ""
                            Else
                                style = style_workdays(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="4"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="workdays<%=personalcode%>"
                                    name="workdays"
                                    value="<%=workdays%>"
                                    onBlur="sum('<%=personalcode%>', 'workdays')"
                                    onFocus="this.select()"
                                    >
                                </td>
                            <%
                            If UBound(style_workholidays) < i Then
                                style = ""
                            Else
                                style = style_workholidays(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="4"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="workholidays<%=personalcode%>"
                                    name="workholidays"
                                    value="<%=workholidays%>"
                                    onBlur="sum('<%=personalcode%>', 'workholidays')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_absencedays) < i Then
                                style = ""
                            Else
                                style = style_absencedays(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="4"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="absencedays<%=personalcode%>"
                                    name="absencedays"
                                    value="<%=absencedays%>"
                                    onBlur="sum('<%=personalcode%>', 'absencedays')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_paidvacations) < i Then
                                style = ""
                            Else
                                style = style_paidvacations(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="4"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="paidvacations<%=personalcode%>"
                                    name="paidvacations"
                                    value="<%=paidvacations%>"
                                    onBlur="sum('<%=personalcode%>', 'paidvacations')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_preservevacations) < i Then
                                style = ""
                            Else
                                style = style_preservevacations(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="4"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="preservevacations<%=personalcode%>"
                                    name="preservevacations"
                                    value="<%=preservevacations%>"
                                    onBlur="sum('<%=personalcode%>', 'preservevacations')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_specialvacations) < i Then
                                style = ""
                            Else
                                style = style_specialvacations(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="4"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="specialvacations<%=personalcode%>"
                                    name="specialvacations"
                                    value="<%=specialvacations%>"
                                    onBlur="sum('<%=personalcode%>', 'specialvacations')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_holidayshifts) < i Then
                                style = ""
                            Else
                                style = style_holidayshifts(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="4"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="holidayshifts<%=personalcode%>"
                                    name="holidayshifts"
                                    value="<%=holidayshifts%>"
                                    onBlur="sum('<%=personalcode%>', 'holidayshifts')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_realworkdays) < i Then
                                style = ""
                            Else
                                style = style_realworkdays(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="4"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="realworkdays<%=personalcode%>"
                                    name="realworkdays"
                                    value="<%=realworkdays%>"
                                    onBlur="sum('<%=personalcode%>', 'realworkdays')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_shortdays) < i Then
                                style = ""
                            Else
                                style = style_shortdays(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="4"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="shortdays<%=personalcode%>"
                                    name="shortdays"
                                    value="<%=shortdays%>"
                                    onBlur="sum('<%=personalcode%>', 'shortdays')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_nightduty_a) < i Then
                                style = ""
                            Else
                                style = style_nightduty_a(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="2"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="nightduty_a<%=personalcode%>"
                                    name="nightduty_a"
                                    value="<%=nightduty_a%>"
                                    onBlur="sum('<%=personalcode%>', 'nightduty_a')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_nightduty_b) < i Then
                                style = ""
                            Else
                                style = style_nightduty_b(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="2"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="nightduty_b<%=personalcode%>"
                                    name="nightduty_b"
                                    value="<%=nightduty_b%>"
                                    onBlur="sum('<%=personalcode%>', 'nightduty_b')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_nightduty_c) < i Then
                                style = ""
                            Else
                                style = style_nightduty_c(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="2"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="nightduty_c<%=personalcode%>"
                                    name="nightduty_c"
                                    value="<%=nightduty_c%>"
                                    onBlur="sum('<%=personalcode%>', 'nightduty_c')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_nightduty_d) < i Then
                                style = ""
                            Else
                                style = style_nightduty_d(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="2"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="nightduty_d<%=personalcode%>"
                                    name="nightduty_d"
                                    value="<%=nightduty_d%>"
                                    onBlur="sum('<%=personalcode%>', 'nightduty_d')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_holidaypremium) < i Then
                                style = ""
                            Else
                                style = style_holidaypremium(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="5"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="holidaypremium<%=personalcode%>"
                                    name="holidaypremium"
                                    value="<%=holidaypremium%>"
                                    onBlur="sum('<%=personalcode%>', 'holidaypremium')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_dayduty) < i Then
                                style = ""
                            Else
                                style = style_dayduty(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="2"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="dayduty<%=personalcode%>"
                                    name="dayduty"
                                    value="<%=dayduty%>"
                                    onBlur="sum('<%=personalcode%>', 'dayduty')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_shiftwork_kou) < i Then
                                style = ""
                            Else
                                style = style_shiftwork_kou(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="4"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="shiftwork_kou<%=personalcode%>"
                                    name="shiftwork_kou"
                                    value="<%=shiftwork_kou%>"
                                    onBlur="sum('<%=personalcode%>', 'shiftwork_kou')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_shiftwork_otsu) < i Then
                                style = ""
                            Else
                                style = style_shiftwork_otsu(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="4"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="shiftwork_otsu<%=personalcode%>"
                                    name="shiftwork_otsu"
                                    value="<%=shiftwork_otsu%>"
                                    onBlur="sum('<%=personalcode%>', 'shiftwork_otsu')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_shiftwork_hei) < i Then
                                style = ""
                            Else
                                style = style_shiftwork_hei(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="4"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="shiftwork_hei<%=personalcode%>"
                                    name="shiftwork_hei"
                                    value="<%=shiftwork_hei%>"
                                    onBlur="sum('<%=personalcode%>', 'shiftwork_hei')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_shiftwork_a) < i Then
                                style = ""
                            Else
                                style = style_shiftwork_a(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="4"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="shiftwork_a<%=personalcode%>"
                                    name="shiftwork_a"
                                    value="<%=shiftwork_a%>"
                                    onBlur="sum('<%=personalcode%>', 'shiftwork_a')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_shiftwork_b) < i Then
                                style = ""
                            Else
                                style = style_shiftwork_b(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="4"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="shiftwork_b<%=personalcode%>"
                                    name="shiftwork_b"
                                    value="<%=shiftwork_b%>"
                                    onBlur="sum('<%=personalcode%>', 'shiftwork_b')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <th nowrap width="110px" class="permanent"><%=staffname%></th>
                            <%
                            If UBound(style_summons) < i Then
                                style = ""
                            Else
                                style = style_summons(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="2"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="summons<%=personalcode%>"
                                    name="summons"
                                    value="<%=summons%>"
                                    onBlur="sum('<%=personalcode%>', 'summons')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_summonslate) < i Then
                                style = ""
                            Else
                                style = style_summonslate(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="2"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="summonslate<%=personalcode%>"
                                    name="summonslate"
                                    onBlur="sum('<%=personalcode%>', 'summonslate')"
                                    value="<%=summonslate%>">
                            </td>
                            <%
                            If UBound(style_yearend1230) < i Then
                                style = ""
                            Else
                                style = style_yearend1230(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="3"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="yearend1230<%=personalcode%>"
                                    name="yearend1230"
                                    value="<%=yearend1230%>"
                                    onBlur="sum('<%=personalcode%>', 'yearend1230')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_yearend1231) < i Then
                                style = ""
                            Else
                                style = style_yearend1231(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="3"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="yearend1231<%=personalcode%>"
                                    name="yearend1231"
                                    value="<%=yearend1231%>"
                                    onBlur="sum('<%=personalcode%>', 'yearend1231')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_workholidaytime) < i Then
                                style = ""
                            Else
                                style = style_workholidaytime(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="5"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="workholidaytime<%=personalcode%>"
                                    name="workholidaytime"
                                    value="<%=workholidaytime%>"
                                    onBlur="sum('<%=personalcode%>', 'workholidaytime')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_latepremium) < i Then
                                style = ""
                            Else
                                style = style_latepremium(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="5"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="latepremium<%=personalcode%>"
                                    name="latepremium"
                                    onBlur="sum('<%=personalcode%>', 'latepremium')"
                                    value="<%=latepremium%>"
                                    >
                                </td>
                            <%
                            If UBound(style_overtime) < i Then
                                style = ""
                            Else
                                style = style_overtime(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="5"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="overtime<%=personalcode%>"
                                    name="overtime"
                                    value="<%=overtime%>"
                                    onBlur="sum('<%=personalcode%>', 'overtime')"
                                    onFocus="this.select()"
                                    >
                                </td>
                            <%
                            If UBound(style_holidayshifttime) < i Then
                                style = ""
                            Else
                                style = style_holidayshifttime(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="5"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="holidayshifttime<%=personalcode%>"
                                    name="holidayshifttime"
                                    value="<%=holidayshifttime%>"
                                    onBlur="sum('<%=personalcode%>', 'holidayshifttime')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_holidayshiftovertime) < i Then
                                style = ""
                            Else
                                style = style_holidayshiftovertime(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="5"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="holidayshiftovertime<%=personalcode%>"
                                    name="holidayshiftovertime"
                                    value="<%=holidayshiftovertime%>"
                                    onBlur="sum('<%=personalcode%>', 'holidayshiftovertime')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_holidayshiftlate) < i Then
                                style = ""
                            Else
                                style = style_holidayshiftlate(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="5"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="holidayshiftlate<%=personalcode%>"
                                    name="holidayshiftlate"
                                    value="<%=holidayshiftlate%>"
                                    onBlur="sum('<%=personalcode%>', 'holidayshiftlate')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_overtimelate) < i Then
                                style = ""
                            Else
                                style = style_overtimelate(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="5"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="overtimelate<%=personalcode%>"
                                    name="overtimelate"
                                    value="<%=overtimelate%>"
                                    onBlur="sum('<%=personalcode%>', 'overtimelate')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <%
                            If UBound(style_holidayshiftovertimelate) < i Then
                                style = ""
                            Else
                                style = style_holidayshiftovertimelate(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="5"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="holidayshiftovertimelate<%=personalcode%>"
                                    name="holidayshiftovertimelate"
                                    value="<%=holidayshiftovertimelate%>"
                                    onBlur="sum('<%=personalcode%>', 'holidayshiftovertimelate')"
                                    onFocus="this.select()"
                                    >
                            </td>
                            <td align="right" nowrap width="40px">
                                <div id="sumOvertime<%=personalcode%>"></div>
                            </td>
                            <%
                            If UBound(style_saturdayworkmin) < i Then
                                style = ""
                            Else
                                style = style_saturdayworkmin(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="5"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="saturdayworkmin<%=personalcode%>"
                                    name="saturdayworkmin"
                                    value="<%=saturdayworkmin%>"
                                    >
                            </td>
                            <%
                            If UBound(style_weekdaysworkmin) < i Then
                                style = ""
                            Else
                                style = style_weekdaysworkmin(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="5"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="weekdaysworkmin<%=personalcode%>"
                                    name="weekdaysworkmin"
                                    value="<%=weekdaysworkmin%>"
                                    >
                            </td>
                            <%
                            If UBound(style_workingmins) < i Then
                                style = ""
                            Else
                                style = style_workingmins(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="5"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="workingmins<%=personalcode%>"
                                    name="workingmins"
                                    value="<%=workingmins%>"
                                    >
                            </td>
                            <%
                            If UBound(style_currentworkmin) < i Then
                                style = ""
                            Else
                                style = style_currentworkmin(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="5"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="currentworkmin<%=personalcode%>"
                                    name="currentworkmin"
                                    value="<%=currentworkmin%>"
                                    >
                            </td>
                            <%
                            If UBound(style_legalholiday_extra_min) < i Then
                                style = ""
                            Else
                                style = style_legalholiday_extra_min(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="5"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="legalholiday_extra_min<%=personalcode%>"
                                    name="legalholiday_extra_min"
                                    value="<%=legalholiday_extra_min%>"
                                    >
                            </td>
                            <%
                            If UBound(style_weekovertime) < i Then
                                style = ""
                            Else
                                style = style_weekovertime(i)
                            End If
                            %>
                            <td align="center" nowrap width="40px">
                                <input class="<%=style%>"
                                    maxlength="5"
                                    style="ime-mode:disabled;text-align:right;width:32px;"
                                    id="weekovertime<%=personalcode%>"
                                    name="weekovertime"
                                    value="<%=weekovertime%>"
                                    >
                            </td>
                        </tr>
                    <%
                        Rs_worktbl.MoveNext()
                    Wend
                    %>
                    <tr>
                        <td align="center" width="38px" class="permanent">-</td>
                        <td align="center" width="110px" class="permanent">合計</td>
                        <td align="center" width="28px">-</td>
                        <td align="right" width="40px"><div id="_workdays"></div></td>
                        <td align="right" width="40px"><div id="_workholidays"></div></td>
                        <td align="right" width="40px"><div id="_absencedays"></div></td>
                        <td align="right" width="40px"><div id="_paidvacations"></div></td>
                        <td align="right" width="40px"><div id="_preservevacations"></div></td>
                        <td align="right" width="40px"><div id="_specialvacations"></div></td>
                        <td align="right" width="40px"><div id="_holidayshifts"></div></td>
                        <td align="right" width="40px"><div id="_realworkdays"></div></td>
                        <td align="right" width="40px"><div id="_shortdays"></div></td>
                        <td align="right" width="40px"><div id="_nightduty_a"></div></td>
                        <td align="right" width="40px"><div id="_nightduty_b"></div></td>
                        <td align="right" width="40px"><div id="_nightduty_c"></div></td>
                        <td align="right" width="40px"><div id="_nightduty_d"></div></td>
                        <td align="right" width="40px"><div id="_holidaypremium"></div></td>
                        <td align="right" width="40px"><div id="_dayduty"></div></td>
                        <td align="right" width="40px"><div id="_shiftwork_kou"></div></td>
                        <td align="right" width="40px"><div id="_shiftwork_otsu"></div></td>
                        <td align="right" width="40px"><div id="_shiftwork_hei"></div></td>
                        <td align="right" width="40px"><div id="_shiftwork_a"></div></td>
                        <td align="right" width="40px"><div id="_shiftwork_b"></div></td>
                        <td align="center" width="110px" class="permanent">合計</td>
                        <td align="right" width="40px"><div id="_summons"></div></td>
                        <td align="right" width="40px"><div id="_summonslate"></div></td>
                        <td align="right" width="40px"><div id="_yearend1230"></div></td>
                        <td align="right" width="40px"><div id="_yearend1231"></div></td>
                        <td align="right" width="40px"><div id="_workholidaytime"></div></td>
                        <td align="right" width="40px"><div id="_latepremium"></div></td>
                        <td align="right" width="40px"><div id="_overtime"></div></td>
                        <td align="right" width="40px"><div id="_holidayshifttime"></div></td>
                        <td align="right" width="40px"><div id="_holidayshiftovertime"></div></td>
                        <td align="right" width="40px"><div id="_holidayshiftlate"></div></td>
                        <td align="right" width="40px"><div id="_overtimelate"></div></td>
                        <td align="right" width="40px"><div id="_holidayshiftovertimelate"></div></td>
                        <td align="right" width="40px"><div id="_sumOvertime">-</div></td>
                        <td align="right" width="40px"><div id="_sumSaturdyworkmin">-</div></td>
                        <td align="right" width="40px"><div id="_sumWeekdaysworkmin">-</div></td>
                        <td align="right" width="40px"><div id="_sumWorkingmins">-</div></td>
                        <td align="right" width="40px"><div id="_sumCurrentworkmin">-</div></td>
                        <td align="right" width="40px"><div id="_sumLegalholiday_extra_min">-</div></td>
                        <td align="right" width="40px"><div id="_weekovertime">-</div></td>
                    </tr>
                </table>
            </div>
        </div>
    </form>
</div>
<!-- #include file="inc/footer.source" -->
</div>
</body>
<script type="text/javascript">
    /* ************************************************************************
     * 変数定義
     * ************************************************************************/
    // 表示職員配列
    var staffList = new Array(
    <%
    If Not Rs_worktbl.EOF Or Not Rs_worktbl.BOF Then
        Rs_worktbl.MoveFirst()
        While (NOT Rs_worktbl.EOF)
            Response.write """" & Rs_worktbl.Fields.Item("pcode") & ""","
            Rs_worktbl.MoveNext()
        Wend
    End If
    %>
    "-");
    staffList.pop();    // 配列最後が""のため削除を行う。
    var overtimeList    = new Array ("overtime", "holidayshifttime",
                                    "holidayshiftovertime", "holidayshiftlate",
                                    "overtimelate", "holidayshiftovertimelate");
    var columnList      = new Array("workdays", "workholidays", "absencedays", 
                                    "paidvacations", "preservevacations", 
                                    "specialvacations", "holidayshifts", 
                                    "realworkdays", "shortdays", 
                                    "nightduty_a", "nightduty_b", 
                                    "nightduty_c", "nightduty_d", 
                                    "holidaypremium", "dayduty", 
                                    "shiftwork_kou", "shiftwork_otsu", 
                                    "shiftwork_hei", "shiftwork_a", "shiftwork_b",
                                    "summons", "summonslate", 
                                    "yearend1230", "yearend1231", 
                                    "workholidaytime", "latepremium",
                                    "overtime", "holidayshifttime",
                                    "holidayshiftovertime", "holidayshiftlate",
                                    "overtimelate", "holidayshiftovertimelate"
                                    //,"sumOvertime"
                                    );
    var overtime                 = 0;
    var holidayshifttime         = 0;
    var holidayshiftovertime     = 0;
    var holidayshiftlate         = 0;
    var overtimelate             = 0;
    var holidayshiftovertimelate = 0;
    var sumOvertime              = 0;
    
    var holidayshifts            = 0;
    var nightduty_a              = 0;
    var nightduty_b              = 0;
    var nightduty_c              = 0;
    var nightduty_d              = 0;
    var sumNightduty             = 0;
    
    var sumColumn                = 0;
    var columnName               = "";
    
    function setDivSize(){
        /* ********************************************************************
         * ウィンドウサイズから div サイズを設定
         * ********************************************************************/
        var size_h;
        size_h = document.body.clientHeight;
        if (size_h < 500) {
            size_h = 320;
        } else {
            size_h = size_h - 120;
        }
        document.getElementById('tablediv').style.height = size_h + "px";
        document.getElementById('tbody').style.height = size_h - 60 + "px";
    }
    function inputSubmit(){
        /* ********************************************************************
         * 登録ボタン押下時の処理
         * ********************************************************************/
        ans=confirm("情報を登録します。\nよろしいですか？");
        if(ans){
            document.inputAll.submit();
        }
    }
    function clickDownloadCSV(){
        /* ********************************************************************
         * CSVデータダウンロードボタン押下時の処理
         * ********************************************************************/
        ans=confirm("データをダウンロードします。\n入力途中の内容は登録されません。\nよろしいですか？");
        if(ans) {
            location.href="csv_inputall.asp?ymb=<%=dispYear & dispMonth%>";
        }
    }
    function clickDownloadXls(){
        /* ********************************************************************
         * XLSデータダウンロードボタン押下時の処理
         * 現在(2012-12-20)は使用されていません。サンプルとして残しておきます。
         * ********************************************************************/
        ans=confirm("エクセルファイルをダウンロードします。\n入力途中の内容は登録されません。\nよろしいですか？");
        if(ans) {
            location.href="xls_inputall.asp";
        }
    }
    function sum(personalcode, column) {
        /* ********************************************************************
         * 変更された職員の項目の集計処理
         * 引数：変更された職員コード, 変更された項目名
         * ********************************************************************/
        sumLine(personalcode);
        sumRow(column);
    }
    function sumRow(column) {
        /* ********************************************************************
         * 列(項目)を合計行に集計
         * 引数：項目名
         * ********************************************************************/
        var i;
        sumColumn = 0;
        for (i=0; i<staffList.length; i++) {
            columnName = column + staffList[i];
            if (isNaN(document.getElementById(columnName).value)) {
            } else {
                sumColumn = sumColumn
                          + (1 * document.getElementById(columnName).value);
            }
            if (isNaN(document.getElementById(columnName).value)) {
            } else {
                document.getElementById(columnName).value =
                    Number(document.getElementById(columnName).value);
            }
            if (Number(document.getElementById(columnName).value) == "0") {
                document.getElementById(columnName).value = "";
            }
        }
        if (column == "workdays"                    ||
            column == "workholidays"                ||
            column == "absencedays"                 ||
            column == "paidvacations"               ||
            column == "preservevacations"           ||
            column == "specialvacations"            ||
            column == "holidayshifts"               ||
            column == "realworkdays"                ||
            column == "yearend1230"                 ||
            column == "yearend1231"                 ||
            column == "workholidaytime"             ||
            column == "latepremium"                 ||
            column == "overtime"                    ||
            column == "holidayshifttime"            ||
            column == "holidayshiftovertime"        ||
            column == "holidayshiftlate"            ||
            column == "overtimelate"                ||
            column == "holidayshiftovertimelate"    ||
            column == "sumOvertime")                {
            // 小数点1桁までの表示
            sumColumn = sumColumn.toFixed(1);
        } else {
            // 整数として表示
            sumColumn = sumColumn.toFixed(0);
        }
        document.getElementById("_"+column).innerHTML = sumColumn;
        // 時間外のとき
        /*
        if (column == "overtime"                    ||
            column == "holidayshiftovertime"        ||
            column == "overtimelate"                ||
            column == "holidayshiftovertimelate" )  {
            sumRow("sumOvertime");
        }
        */
    }
    function sumLine(personalcode) {
        /* ********************************************************************
         * 行の時間外を集計
         * 引数：集計対象行の職員コード
         * ********************************************************************/
        var i;
        overtime                 = document.getElementById("overtime"                 + personalcode).value;
        holidayshifttime         = document.getElementById("holidayshifttime"         + personalcode).value;
        holidayshiftovertime     = document.getElementById("holidayshiftovertime"     + personalcode).value;
        holidayshiftlate         = document.getElementById("holidayshiftlate"         + personalcode).value;
        overtimelate             = document.getElementById("overtimelate"             + personalcode).value;
        holidayshiftovertimelate = document.getElementById("holidayshiftovertimelate" + personalcode).value;
        overtime                 = (isNaN(overtime                ) || !overtime                ) ? 0 : 1 * overtime;
        holidayshifttime         = (isNaN(holidayshifttime        ) || !holidayshifttime        ) ? 0 : 1 * holidayshifttime;
        holidayshiftovertime     = (isNaN(holidayshiftovertime    ) || !holidayshiftovertime    ) ? 0 : 1 * holidayshiftovertime;
        holidayshiftlate         = (isNaN(holidayshiftlate        ) || !holidayshiftlate        ) ? 0 : 1 * holidayshiftlate;
        overtimelate             = (isNaN(overtimelate            ) || !overtimelate            ) ? 0 : 1 * overtimelate;
        holidayshiftovertimelate = (isNaN(holidayshiftovertimelate) || !holidayshiftovertimelate) ? 0 : 1 * holidayshiftovertimelate;
        // 時間外集計とカンマ編集
        sumOvertime = (overtime             +
                       holidayshiftovertime +
                       overtimelate         +
                       holidayshiftovertimelate).toFixed(1);
        // 時間外計設定
        document.getElementById("sumOvertime" + personalcode).innerHTML = sumOvertime;
        // 時間外計30以上の時、赤反転表示
        if (sumOvertime >= 30) {
            document.getElementById("sumOvertime" + personalcode).className = 'warning';
        } else {
            document.getElementById("sumOvertime" + personalcode).className = '';
        }
        // 各項目の前0と、0値のとき0表示消去
        for (i=0; i<overtimeList.length; i++) {
            if (isNaN(document.getElementById(overtimeList[i] + personalcode).value)) {
            } else {
                document.getElementById(overtimeList[i] + personalcode).value =
                    Number(document.getElementById(overtimeList[i] + personalcode).value);
            }
            if (Number(document.getElementById(overtimeList[i] + personalcode).value) == "0") {
                document.getElementById(overtimeList[i] + personalcode).value = "";
            }
        }
        // 休出2以上の時、赤反転表示
        holidayshifts = document.getElementById("holidayshifts" + personalcode).value;
        holidayshifts = (isNaN(holidayshifts) || !holidayshifts) ? 0 : 1 * holidayshifts;
        if (holidayshifts >= 2) {
            document.getElementById("holidayshifts" + personalcode).className = 'warning';
        } else {
            document.getElementById("holidayshifts" + personalcode).className = '';
        }
        // 宿直回数4回以上の時、赤反転表示
        nightduty_a = document.getElementById("nightduty_a" + personalcode).value;
        nightduty_b = document.getElementById("nightduty_b" + personalcode).value;
        nightduty_c = document.getElementById("nightduty_c" + personalcode).value;
        nightduty_d = document.getElementById("nightduty_d" + personalcode).value;
        nightduty_a = (isNaN(nightduty_a) || !nightduty_a) ? 0 : 1 * nightduty_a;
        nightduty_b = (isNaN(nightduty_b) || !nightduty_b) ? 0 : 1 * nightduty_b;
        nightduty_c = (isNaN(nightduty_c) || !nightduty_c) ? 0 : 1 * nightduty_c;
        nightduty_d = (isNaN(nightduty_d) || !nightduty_d) ? 0 : 1 * nightduty_d;
        sumNightduty = (nightduty_a + nightduty_b + nightduty_c + nightduty_d).toFixed(1);
        if (sumNightduty >= 4) {
            document.getElementById("nightduty_a" + personalcode).className = 'warning';
            document.getElementById("nightduty_b" + personalcode).className = 'warning';
            document.getElementById("nightduty_c" + personalcode).className = 'warning';
            document.getElementById("nightduty_d" + personalcode).className = 'warning';
        } else {
            document.getElementById("nightduty_a" + personalcode).className = '';
            document.getElementById("nightduty_b" + personalcode).className = '';
            document.getElementById("nightduty_c" + personalcode).className = '';
            document.getElementById("nightduty_d" + personalcode).className = '';
        }
    }
    function sumInit() {
        /* ********************************************************************
         * 初期表示時の集計処理
         * 引数：なし
         * ********************************************************************/
        var i;
        // 各行の時間外を集計
        for (i=0; i<staffList.length; i++) {
            sumLine(staffList[i]);
        }
        // 合計行の集計
        for (i=0; i<columnList.length; i++) {
            sumRow(columnList[i]);
        }
    }
    /* ************************************************************************
     * 画面初期表示時の処理
     * ************************************************************************/
    setDivSize();
    sumInit();
</script>
</html>
<%
Rs_worktbl.Close()
Set Rs_worktbl = Nothing

' -----------------------------------------------------------------------------
' 入力チェックと値の設定を行う。
' -----------------------------------------------------------------------------
Sub setData()
    id                       = Trim(Request.Form("id"                      )(i))
    personalcode             = Trim(Request.Form("personalcode"            )(i))
    workdays                 = Trim(Request.Form("workdays"                )(i))
    workholidays             = Trim(Request.Form("workholidays"            )(i))
    absencedays              = Trim(Request.Form("absencedays"             )(i))
    paidvacations            = Trim(Request.Form("paidvacations"           )(i))
    preservevacations        = Trim(Request.Form("preservevacations"       )(i))
    specialvacations         = Trim(Request.Form("specialvacations"        )(i))
    holidayshifts            = Trim(Request.Form("holidayshifts"           )(i))
    realworkdays             = Trim(Request.Form("realworkdays"            )(i))
    shortdays                = Trim(Request.Form("shortdays"               )(i))
    nightduty_a              = Trim(Request.Form("nightduty_a"             )(i))
    nightduty_b              = Trim(Request.Form("nightduty_b"             )(i))
    nightduty_c              = Trim(Request.Form("nightduty_c"             )(i))
    nightduty_d              = Trim(Request.Form("nightduty_d"             )(i))
    holidaypremium           = Trim(Request.Form("holidaypremium"          )(i))
    dayduty                  = Trim(Request.Form("dayduty"                 )(i))
    shiftwork_kou            = Trim(Request.Form("shiftwork_kou"           )(i))
    shiftwork_otsu           = Trim(Request.Form("shiftwork_otsu"          )(i))
    shiftwork_hei            = Trim(Request.Form("shiftwork_hei"           )(i))
    shiftwork_a              = Trim(Request.Form("shiftwork_a"             )(i))
    shiftwork_b              = Trim(Request.Form("shiftwork_b"             )(i))
    summons                  = Trim(Request.Form("summons"                 )(i))
    summonslate              = Trim(Request.Form("summonslate"             )(i))
    yearend1230              = Trim(Request.Form("yearend1230"             )(i))
    yearend1231              = Trim(Request.Form("yearend1231"             )(i))
    workholidaytime          = Trim(Request.Form("workholidaytime"         )(i))
    latepremium              = Trim(Request.Form("latepremium"             )(i))
    overtime                 = Trim(Request.Form("overtime"                )(i))
    holidayshifttime         = Trim(Request.Form("holidayshifttime"        )(i))
    holidayshiftovertime     = Trim(Request.Form("holidayshiftovertime"    )(i))
    holidayshiftlate         = Trim(Request.Form("holidayshiftlate"        )(i))
    overtimelate             = Trim(Request.Form("overtimelate"            )(i))
    holidayshiftovertimelate = Trim(Request.Form("holidayshiftovertimelate")(i))
    vacationnumber           = Trim(Request.Form("vacationnumber"          )(i))
    holidaynumber            = Trim(Request.Form("holidaynumber"           )(i))
    vacationtime             = Trim(Request.Form("vacationtime"            )(i))
    saturdayworkmin          = Trim(Request.Form("saturdayworkmin"         )(i))
    weekdaysworkmin          = Trim(Request.Form("weekdaysworkmin"         )(i))
    workingmins              = Trim(Request.Form("workingmins"             )(i))
    currentworkmin           = Trim(Request.Form("currentworkmin"          )(i))
    legalholiday_extra_min   = Trim(Request.Form("legalholiday_extra_min"  )(i))
    weekovertime             = Trim(Request.Form("weekovertime"            )(i))

    ' 前月勤務表テーブルを読み、当月末有給休暇残と当月末振替休日残と
    ' 入力内容の妥当性チェックに使用する。
    Dim Rs_lastmonth_dutyrostertbl
    Dim Rs_lastmonth_dutyrostertbl_cmd
    Dim Rs_lastmonth_dutyrostertbl_numRows
    Set Rs_lastmonth_dutyrostertbl_cmd = Server.CreateObject ("ADODB.Command")
    Rs_lastmonth_dutyrostertbl_cmd.ActiveConnection = MM_workdbms_STRING
    Rs_lastmonth_dutyrostertbl_cmd.CommandText = "SELECT * FROM dutyrostertbl WHERE personalcode = ? AND ymb = ?"
    Rs_lastmonth_dutyrostertbl_cmd.Prepared = true
    Rs_lastmonth_dutyrostertbl_cmd.Parameters.Append Rs_lastmonth_dutyrostertbl_cmd.CreateParameter("param1", 200, 1, 5, personalcode)
    Rs_lastmonth_dutyrostertbl_cmd.Parameters.Append Rs_lastmonth_dutyrostertbl_cmd.CreateParameter("param2", 200, 1, 7, lastYmb)
    Set Rs_lastmonth_dutyrostertbl = Rs_lastmonth_dutyrostertbl_cmd.Execute
    Rs_lastmonth_dutyrostertbl_numRows = 0

    ' 可出勤日数
    ReDim Preserve style_workdays(i)
    If (Len(workdays) = 0) Then
        workdays = 0
    Else
        If (Not IsNumeric(workdays)) Then
            errorMsg          = strErrorMsg
            style_workdays(i) = "errorcolor"
        ElseIf (workdays < 0) Then
            errorMsg          = strErrorMsg
            style_workdays(i) = "errorcolor"
        End If
    End If
    ' 欠勤日数
    ReDim Preserve style_absencedays(i)
    If (Len(absencedays) = 0) Then
        absencedays = 0
    Else
        If (Not IsNumeric(absencedays)) Then
            errorMsg             = strErrorMsg
            style_absencedays(i) = "errorcolor"
        ElseIf (absencedays < 0) Then
            errorMsg             = strErrorMsg
            style_absencedays(i) = "errorcolor"
        End If
    End If
    ' 保存休暇日数
    ReDim Preserve style_preservevacations(i)
    If (Len(preservevacations) = 0) Then
        preservevacations = 0
    Else
        If (Not IsNumeric(preservevacations)) Then
            errorMsg                   = strErrorMsg
            style_preservevacations(i) = "errorcolor"
        ElseIf (preservevacations < 0) Then
            errorMsg                   = strErrorMsg
            style_preservevacations(i) = "errorcolor"
        End If
    End If
    ' 特休日数
    ReDim Preserve style_specialvacations(i)
    If (Len(specialvacations) = 0) Then
        specialvacations = 0
    Else
        If (Not IsNumeric(specialvacations)) Then
            errorMsg                  = strErrorMsg
            style_specialvacations(i) = "errorcolor"
        ElseIf (specialvacations < 0) Then
            errorMsg                  = strErrorMsg
            style_specialvacations(i) = "errorcolor"
        End If
    End If
    ' 休出日数
    ReDim Preserve style_holidayshifts(i)
    If (Len(holidayshifts) = 0) Then
        holidayshifts = 0
    Else
        If (Not IsNumeric(holidayshifts)) Then
            errorMsg               = strErrorMsg
            style_holidayshifts(i) = "errorcolor"
        ElseIf (holidayshifts < 0) Then
            errorMsg               = strErrorMsg
            style_holidayshifts(i) = "errorcolor"
        End If
    End If
    ' 実出勤日数
    ReDim Preserve style_realworkdays(i)
    If (Len(realworkdays) = 0) Then
        realworkdays = 0
    Else
        If (Not IsNumeric(realworkdays)) Then
            errorMsg              = strErrorMsg
            style_realworkdays(i) = "errorcolor"
        ElseIf (realworkdays < 0) Then
            errorMsg              = strErrorMsg
            style_realworkdays(i) = "errorcolor"
        End If
    End If
    ' 遅早回数
    ReDim Preserve style_shortdays(i)
    If (Len(shortdays) = 0) Then
        shortdays = 0
    Else
        If (Not IsNumeric(shortdays)) Then
            errorMsg           = strErrorMsg
            style_shortdays(i) = "errorcolor"
        ElseIf (shortdays < 0) Then
            errorMsg           = strErrorMsg
            style_shortdays(i) = "errorcolor"
        End If
    End If
    ' 宿直A回数
    ReDim Preserve style_nightduty_a(i)
    If (Len(nightduty_a) = 0) Then
        nightduty_a = 0
    Else
        If (Not IsNumeric(nightduty_a)) Then
            errorMsg             = strErrorMsg
            style_nightduty_a(i) = "errorcolor"
        ElseIf (nightduty_a < 0) Then
            errorMsg             = strErrorMsg
            style_nightduty_a(i) = "errorcolor"
        End If
        ' 小数点はエラー
        If (InStr(nightduty_a, ".") <> 0) Then
            errorMsg             = strErrorMsg
            style_nightduty_a(i) = "errorcolor"
        End If
    End If
    ' 宿直B回数
    ReDim Preserve style_nightduty_b(i)
    If (Len(nightduty_b) = 0) Then
        nightduty_b = 0
    Else
        If (Not IsNumeric(nightduty_b)) Then
            errorMsg             = strErrorMsg
            style_nightduty_b(i) = "errorcolor"
        ElseIf (nightduty_b < 0) Then
            errorMsg             = strErrorMsg
            style_nightduty_b(i) = "errorcolor"
        End If
        ' 小数点はエラー
        If (InStr(nightduty_b, ".") <> 0) Then
            errorMsg             = strErrorMsg
            style_nightduty_b(i) = "errorcolor"
        End If
    End If
    ' 宿直C回数
    ReDim Preserve style_nightduty_c(i)
    If (Len(nightduty_c) = 0) Then
        nightduty_c = 0
    Else
        If (Not IsNumeric(nightduty_c)) Then
            errorMsg             = strErrorMsg
            style_nightduty_c(i) = "errorcolor"
        ElseIf (nightduty_c < 0) Then
            errorMsg             = strErrorMsg
            style_nightduty_c(i) = "errorcolor"
        End If
        ' 小数点はエラー
        If (InStr(nightduty_c, ".") <> 0) Then
            errorMsg             = strErrorMsg
            style_nightduty_c(i) = "errorcolor"
        End If
    End If
    ' 宿直D回数
    ReDim Preserve style_nightduty_d(i)
    If (Len(nightduty_d) = 0) Then
        nightduty_d = 0
    Else
        If (Not IsNumeric(nightduty_d)) Then
            errorMsg             = strErrorMsg
            style_nightduty_d(i) = "errorcolor"
        ElseIf (nightduty_d < 0) Then
            errorMsg             = strErrorMsg
            style_nightduty_d(i) = "errorcolor"
        End If
        ' 小数点はエラー
        If (InStr(nightduty_d, ".") <> 0) Then
            errorMsg             = strErrorMsg
            style_nightduty_d(i) = "errorcolor"
        End If
    End If
    ' 休日割増
    ReDim Preserve style_holidaypremium(i)
    If (Len(holidaypremium) = 0) Then
        holidaypremium = 0
    Else
        If (Not IsNumeric(holidaypremium)) Then
            errorMsg                = strErrorMsg
            style_holidaypremium(i) = "errorcolor"
        ElseIf (holidaypremium < 0) Then
            errorMsg                = strErrorMsg
            style_holidaypremium(i) = "errorcolor"
        End If
    End If
    ' 日直
    ReDim Preserve style_dayduty(i)
    If (Len(dayduty) = 0) Then
        dayduty = 0
    Else
        If (Not IsNumeric(dayduty)) Then
            errorMsg         = strErrorMsg
            style_dayduty(i) = "errorcolor"
        ElseIf (dayduty < 0) Then
            errorMsg         = strErrorMsg
            style_dayduty(i) = "errorcolor"
        End If
        ' 小数点はエラー
        If (InStr(dayduty, ".") <> 0) Then
            errorMsg         = strErrorMsg
            style_dayduty(i) = "errorcolor"
        End If
    End If
    ' 交代甲番
    ReDim Preserve style_shiftwork_kou(i)
    If (Len(shiftwork_kou) = 0) Then
        shiftwork_kou = 0
    Else
        If (Not IsNumeric(shiftwork_kou)) Then
            errorMsg               = strErrorMsg
            style_shiftwork_kou(i) = "errorcolor"
        ElseIf (shiftwork_kou < 0) Then
            errorMsg               = strErrorMsg
            style_shiftwork_kou(i) = "errorcolor"
        End If
    End If
    ' 交代乙番
    ReDim Preserve style_shiftwork_otsu(i)
    If (Len(shiftwork_otsu) = 0) Then
        shiftwork_otsu = 0
    Else
        If (Not IsNumeric(shiftwork_otsu)) Then
            errorMsg                = strErrorMsg
            style_shiftwork_otsu(i) = "errorcolor"
        ElseIf (shiftwork_otsu < 0) Then
            errorMsg                = strErrorMsg
            style_shiftwork_otsu(i) = "errorcolor"
        End If
    End If
    ' 交代丙番
    ReDim Preserve style_shiftwork_hei(i)
    If (Len(shiftwork_hei) = 0) Then
        shiftwork_hei = 0
    Else
        If (Not IsNumeric(shiftwork_hei)) Then
            errorMsg               = strErrorMsg
            style_shiftwork_hei(i) = "errorcolor"
        ElseIf (shiftwork_hei < 0) Then
            errorMsg               = strErrorMsg
            style_shiftwork_hei(i) = "errorcolor"
        End If
    End If
    ' 交代A番
    ReDim Preserve style_shiftwork_a(i)
    If (Len(shiftwork_a) = 0) Then
        shiftwork_a = 0
    Else
        If (Not IsNumeric(shiftwork_a)) Then
            errorMsg               = strErrorMsg
            style_shiftwork_a(i) = "errorcolor"
        ElseIf (shiftwork_a < 0) Then
            errorMsg               = strErrorMsg
            style_shiftwork_a(i) = "errorcolor"
        End If
    End If
    ' 交代B番
    ReDim Preserve style_shiftwork_b(i)
    If (Len(shiftwork_b) = 0) Then
        shiftwork_b = 0
    Else
        If (Not IsNumeric(shiftwork_b)) Then
            errorMsg               = strErrorMsg
            style_shiftwork_b(i) = "errorcolor"
        ElseIf (shiftwork_b < 0) Then
            errorMsg               = strErrorMsg
            style_shiftwork_b(i) = "errorcolor"
        End If
    End If

    ' 呼出通常回数
    ReDim Preserve style_summons(i)
    If (Len(summons) = 0) Then
        summons = 0
    Else
        If (Not IsNumeric(summons)) Then
            errorMsg         = strErrorMsg
            style_summons(i) = "errorcolor"
        ElseIf (summons < 0) Then
            errorMsg         = strErrorMsg
            style_summons(i) = "errorcolor"
        End If
        ' 小数点はエラー
        If (InStr(summons, ".") <> 0) Then
            errorMsg         = strErrorMsg
            style_summons(i) = "errorcolor"
        End If
    End If
    ' 呼出深夜回数
    ReDim Preserve style_summonslate(i)
    If (Len(summonslate) = 0) Then
        summonslate = 0
    Else
        If (Not IsNumeric(summonslate)) Then
            errorMsg             = strErrorMsg
            style_summonslate(i) = "errorcolor"
        ElseIf (summonslate < 0) Then
            errorMsg             = strErrorMsg
            style_summonslate(i) = "errorcolor"
        End If
        ' 小数点はエラー
        If (InStr(summonslate, ".") <> 0) Then
            errorMsg             = strErrorMsg
            style_summonslate(i) = "errorcolor"
        End If
    End If
    ' 年末年始1230
    ReDim Preserve style_yearend1230(i)
    If (Len(yearend1230) = 0) Then
        yearend1230 = 0
    Else
        If (Not IsNumeric(yearend1230)) Then
            errorMsg             = strErrorMsg
            style_yearend1230(i) = "errorcolor"
        ElseIf (yearend1230 < 0) Then
            errorMsg             = strErrorMsg
            style_yearend1230(i) = "errorcolor"
        ElseIf ((yearend1230 * 10 mod 5) > 0) Then
            errorMsg             = strErrorMsg
            style_yearend1230(i) = "errorcolor"
        End If
    End If
    ' 年末年始1231
    ReDim Preserve style_yearend1231(i)
    If (Len(yearend1231) = 0) Then
        yearend1231 = 0
    Else
        If (Not IsNumeric(yearend1231)) Then
            errorMsg             = strErrorMsg
            style_yearend1231(i) = "errorcolor"
        ElseIf (yearend1231 < 0) Then
            errorMsg             = strErrorMsg
            style_yearend1231(i) = "errorcolor"
        ElseIf ((yearend1231 * 10 mod 5) > 0) Then
            errorMsg             = strErrorMsg
            style_yearend1231(i) = "errorcolor"
        End If
    End If
    ' 時間代休
    ReDim Preserve style_workholidaytime(i)
    If (Len(workholidaytime) = 0) Then
        workholidaytime = 0
    Else
        If (Not IsNumeric(workholidaytime)) Then
            errorMsg                 = strErrorMsg
            style_workholidaytime(i) = "errorcolor"
        ElseIf (workholidaytime < 0) Then
            errorMsg                 = strErrorMsg
            style_workholidaytime(i) = "errorcolor"
        End If
    End If
    ' 深夜割増
    ReDim Preserve style_latepremium(i)
    If (Len(latepremium) = 0) Then
        latepremium = 0
    Else
        If (Not IsNumeric(latepremium)) Then
            errorMsg             = strErrorMsg
            style_latepremium(i) = "errorcolor"
        ElseIf (latepremium < 0) Then
            errorMsg             = strErrorMsg
            style_latepremium(i) = "errorcolor"
        End If
    End If
    ' 時間外
    ReDim Preserve style_overtime(i)
    If (Len(overtime) = 0) Then
        overtime = 0
    Else
        If (Not IsNumeric(overtime)) Then
            errorMsg          = strErrorMsg
            style_overtime(i) = "errorcolor"
        ElseIf (overtime < 0) Then
            errorMsg          = strErrorMsg
            style_overtime(i) = "errorcolor"
        End If
    End If
    ' 休日出勤
    ReDim Preserve style_holidayshifttime(i)
    If (Len(holidayshifttime) = 0) Then
        holidayshifttime = 0
    Else
        If (Not IsNumeric(holidayshifttime)) Then
            errorMsg                  = strErrorMsg
            style_holidayshifttime(i) = "errorcolor"
        ElseIf (holidayshifttime < 0) Then
            errorMsg                  = strErrorMsg
            style_holidayshifttime(i) = "errorcolor"
        End If
    End If
    ' 休出時外
    ReDim Preserve style_holidayshiftovertime(i)
    If (Len(holidayshiftovertime) = 0) Then
        holidayshiftovertime = 0
    Else
        If (Not IsNumeric(holidayshiftovertime)) Then
            errorMsg                      = strErrorMsg
            style_holidayshiftovertime(i) = "errorcolor"
        ElseIf (holidayshiftovertime < 0) Then
            errorMsg                      = strErrorMsg
            style_holidayshiftovertime(i) = "errorcolor"
        End If
    End If
    ' 休出深夜
    ReDim Preserve style_holidayshiftlate(i)
    If (Len(holidayshiftlate) = 0) Then
        holidayshiftlate = 0
    Else
        If (Not IsNumeric(holidayshiftlate)) Then
            errorMsg                  = strErrorMsg
            style_holidayshiftlate(i) = "errorcolor"
        ElseIf (holidayshiftlate < 0) Then
            errorMsg                  = strErrorMsg
            style_holidayshiftlate(i) = "errorcolor"
        End If
    End If
    ' 時外深夜
    ReDim Preserve style_overtimelate(i)
    If (Len(overtimelate) = 0) Then
        overtimelate = 0
    Else
        If (Not IsNumeric(overtimelate)) Then
            errorMsg              = strErrorMsg
            style_overtimelate(i) = "errorcolor"
        ElseIf (overtimelate < 0) Then
            errorMsg              = strErrorMsg
            style_overtimelate(i) = "errorcolor"
        End If
    End If
    ' 休出時外深夜
    ReDim Preserve style_holidayshiftovertimelate(i)
    If (Len(holidayshiftovertimelate) = 0) Then
        holidayshiftovertimelate = 0
    Else
        If (Not IsNumeric(holidayshiftovertimelate)) Then
            errorMsg                          = strErrorMsg
            style_holidayshiftovertimelate(i) = "errorcolor"
        ElseIf (holidayshiftovertimelate < 0) Then
            errorMsg                          = strErrorMsg
            style_holidayshiftovertimelate(i) = "errorcolor"
        End If
    End If
    ' 代休日数
    ReDim Preserve style_workholidays(i)
    If (Len(workholidays) = 0) Then
        workholidays = 0
    Else
        If (Not IsNumeric(workholidays)) Then
            errorMsg              = strErrorMsg
            style_workholidays(i) = "errorcolor"
        ElseIf (workholidays < 0) Then
            errorMsg              = strErrorMsg
            style_workholidays(i) = "errorcolor"
        End If
    End If
    ' 有給日数
    ReDim Preserve style_paidvacations(i)
    If (Len(paidvacations) = 0) Then
        paidvacations = 0
    Else
        If (Not IsNumeric(paidvacations)) Then
            errorMsg               = strErrorMsg
            style_paidvacations(i) = "errorcolor"
        ElseIf (paidvacations < 0) Then
            errorMsg               = strErrorMsg
            style_paidvacations(i) = "errorcolor"
        Else
            ' 前月末有給休暇残日数 - 当月取得有給休暇日数 < 0 のときエラー
            If IsNumeric(vacationnumber) Then
                If vacationnumber < 0 Then
                    errorMsg                = strErrorMsg
                    style_paidvacations(i)  = "errorcolor"
                End If
            Else
                errorMsg                = strErrorMsg
                style_paidvacations(i)  = "errorcolor"
            End If
        End If
    End If
    ' 時間有給数
    If Trim(Request.Form("vacationtime")(i)) = "" Then
        vacationtime = Rs_lastmonth_dutyrostertbl_vacationtime
    Else
        vacationtime = Trim(Request.Form("vacationtime")(i))
    End If
    ' 土曜日勤務時間(分)
    ReDim Preserve style_saturdayworkmin(i)
    If (Len(saturdayworkmin) = 0) Then
        saturdayworkmin = 0
    Else
        If (Not IsNumeric(saturdayworkmin)) Then
            errorMsg                          = strErrorMsg
            style_saturdayworkmin(i) = "errorcolor"
        ElseIf (saturdayworkmin < 0) Then
            errorMsg                          = strErrorMsg
            style_saturdayworkmin(i) = "errorcolor"
        End If
    End If
    ' 平日勤務時間(分)
    ReDim Preserve style_weekdaysworkmin(i)
    If (Len(weekdaysworkmin) = 0) Then
        weekdaysworkmin = 0
    Else
        If (Not IsNumeric(weekdaysworkmin)) Then
            errorMsg                          = strErrorMsg
            style_weekdaysworkminn(i) = "errorcolor"
        ElseIf (weekdaysworkmin < 0) Then
            errorMsg                          = strErrorMsg
            style_weekdaysworkmin(i) = "errorcolor"
        End If
    End If

    ' 労働時間(分)
    ReDim Preserve style_workingmins(i)
    If (Len(workingmins) = 0) Then
        workingmins = 0
    Else
        If (Not IsNumeric(workingmins)) Then
            errorMsg = strErrorMsg
            style_workingmins(i) = "errorcolor"
        ElseIf (workingmins < 0) Then
            errorMsg = strErrorMsg
            style_workingmins(i) = "errorcolor"
        End If
    End If
    ' 当月労働時間(分)
    ReDim Preserve style_currentworkmin(i)
    If (Len(currentworkmin) = 0) Then
        currentworkmin = 0
    Else
        If (Not IsNumeric(currentworkmin)) Then
            errorMsg = strErrorMsg
            style_currentworkmin(i) = "errorcolor"
        ElseIf (currentworkmin < 0) Then
            errorMsg = strErrorMsg
            style_currentworkmin(i) = "errorcolor"
        End If
    End If
    ' 法定休日割増時間
    ReDim Preserve style_legalholiday_extra_min(i)
    If (Len(legalholiday_extra_min) = 0) Then
        legalholiday_extra_min = 0
    Else
        If (Not IsNumeric(legalholiday_extra_min)) Then
            errorMsg = strErrorMsg
            style_legalholiday_extra_min(i) = "errorcolor"
        ElseIf (legalholiday_extra_min < 0) Then
            errorMsg = strErrorMsg
            style_legalholiday_extra_min(i) = "errorcolor"
        End If
    End If
    ' 週超過時間
    ReDim Preserve style_weekovertime(i)
    If (Len(weekovertime) = 0) Then
        weekovertime = 0
    Else
        If (Not IsNumeric(weekovertime)) Then
            errorMsg = strErrorMsg
            style_weekovertime(i) = "errorcolor"
        ElseIf (weekovertime < 0) Then
            errorMsg = strErrorMsg
            style_weekovertime(i) = "errorcolor"
        End If
    End If

    Rs_lastmonth_dutyrostertbl.Close()
    Set Rs_lastmonth_dutyrostertbl = Nothing

End Sub
%>