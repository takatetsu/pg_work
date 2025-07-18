<%
' -----------------------------------------------------------------------------
' 集計項目初期化
' -----------------------------------------------------------------------------
sumWorkDays                 = lastDay   ' 可出勤日数
sumWorkholidays             = 0         ' 代休日数
sumAbsenceDays              = 0         ' 欠勤日数
sumPaidvacations            = 0         ' 有休日数
sumPreservevacations        = 0         ' 保存休日数
sumSpecialvacations         = 0         ' 特休日数
sumHolidayshifts            = 0         ' 休出日数
sumRealworkdays             = 0         ' 実出勤日数
sumSummons                  = 0         ' 呼出回数
sumSummonslate              = 0         ' 呼出深夜
sumOvertime                 = 0         ' 時間外
sumHolidayshifttime         = 0         ' 休日出勤
sumHolidayshiftovertime     = 0         ' 休出時間外
sumHolidayshiftlate         = 0         ' 休出深夜
sumOvertimelate             = 0         ' 時間外深夜
sumHolidayshiftovertimelate = 0         ' 休出時間外深夜
sumWorkholidaytime          = 0         ' 時間代休
sumVacationtime             = 0         ' 時間有給
sumLatepremium              = 0         ' 深夜割増
sumTotalOvertime            = 0         ' 時間外労働計
sumHolidaynumber            = 0         ' 振替残日数
sumNightdutyCount           = 0         ' 宿直回数
sumDaydutyCount             = 0         ' 日直回数
sumOperatorKou              = 0         ' 交替勤務甲
sumOperatorOtsu             = 0         ' 交替勤務乙
sumHolidayWork              = 0         ' 休出回数
sumSaturdayWorkMin          = 0         ' 土曜日勤務時間(コミュニケータ用)
sumWeekdaysWorkMin          = 0         ' 平日勤務時間(コミュニケータ用)

sumFlex_holidayshift        = 0         ' フレックス勤務休出日数集計用(分)

sumOvertime0                = 0         ' 当月時間外計(休出含む)
sumOvertime1                = 0         ' 前月時間外計(休出含む)
sumOvertime2                = 0         ' 2カ月前時間外計(休出含む)
sumOvertime3                = 0         ' 3カ月前時間外計(休出含む)
sumOvertime4                = 0         ' 4カ月前時間外計(休出含む)
sumOvertime5                = 0         ' 5カ月前時間外計(休出含む)
totalPaidvacations          = 0         ' 有休付与日からの有休取得日数

checkHolidayMM1             = ""        ' 表示月の1カ月後チェック用月
checkHolidayMM3             = Array("") ' 表示月の3カ月後チェック用月配列

' -----------------------------------------------------------------------------
' 勤務表テーブル dutyrostertbl 読込
' -----------------------------------------------------------------------------
' 対象月データ読込み
Dim Rs_dutyrostertbl
Dim Rs_dutyrostertbl_cmd
Dim Rs_dutyrostertbl_numRows
Set Rs_dutyrostertbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_dutyrostertbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_dutyrostertbl_cmd.CommandText = "SELECT * FROM dbo.dutyrostertbl WHERE personalcode = ? AND ymb = ? ORDER BY ymb DESC"
Rs_dutyrostertbl_cmd.Prepared = true
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param2", 200, 1, 6, ymb)
Set Rs_dutyrostertbl = Rs_dutyrostertbl_cmd.Execute
Rs_dutyrostertbl_numRows = 0
If Rs_dutyrostertbl.EOF And Rs_dutyrostertbl.BOF Then
    dutyrostertbl_id = ""                                       ' テーブルキー
    init_weekovertime = 0
Else
    dutyrostertbl_id = Rs_dutyrostertbl.Fields.Item("id").Value ' テーブルキー
    sumOvertime0 = Rs_dutyrostertbl.Fields.Item("overtime"                ).Value + _
                   Rs_dutyrostertbl.Fields.Item("holidayshiftovertime"    ).Value + _
                   Rs_dutyrostertbl.Fields.Item("overtimelate"            ).Value + _
                   Rs_dutyrostertbl.Fields.Item("holidayshiftovertimelate").Value + _
                   Rs_dutyrostertbl.Fields.Item("holidayshifttime"        ).Value + _
                   Rs_dutyrostertbl.Fields.Item("holidayshiftlate"        ).Value + _
                   Rs_dutyrostertbl.Fields.Item("weekovertime"            ).Value
    init_weekovertime = Rs_dutyrostertbl.Fields.Item("weekovertime").Value
End If
Rs_dutyrostertbl.Close()

' 対象月前月データ読込み
Set Rs_dutyrostertbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_dutyrostertbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_dutyrostertbl_cmd.CommandText = "SELECT * FROM dbo.dutyrostertbl WHERE personalcode = ? AND ymb < ? ORDER BY ymb DESC"
Rs_dutyrostertbl_cmd.Prepared = true
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param2", 200, 1, 6, ymb)
Set Rs_dutyrostertbl = Rs_dutyrostertbl_cmd.Execute
Rs_dutyrostertbl_numRows = 0
If Rs_dutyrostertbl.EOF And Rs_dutyrostertbl.BOF Then
    vacationnumber  = 0 ' 有休残日数
    holidaynumber   = 0 ' 振替残日数
    vacationtime    = 0 ' 時間有給取得日数
    lastmonth_workingmins = 0 ' 労働時間(分)
    lastmonth_currentworkmin = 0 ' 月内労働時間(分)
Else
    vacationnumber = Rs_dutyrostertbl.Fields.Item("vacationnumber").Value   ' 有休残日数
    holidaynumber  = Rs_dutyrostertbl.Fields.Item("holidaynumber" ).Value   ' 振替残日数
'    If Right(ymb,2) = "04" Then    4月時の時間有給0クリアは人事で3月分データを更新して行う。
'        vacationtime   = 0
'    Else
        vacationtime   = Rs_dutyrostertbl.Fields.Item("vacationtime" ).Value ' 時間有給(分)
'    End If
    sumOvertime1 = Rs_dutyrostertbl.Fields.Item("overtime"                ).Value + _
                   Rs_dutyrostertbl.Fields.Item("holidayshiftovertime"    ).Value + _
                   Rs_dutyrostertbl.Fields.Item("overtimelate"            ).Value + _
                   Rs_dutyrostertbl.Fields.Item("holidayshiftovertimelate").Value + _
                   Rs_dutyrostertbl.Fields.Item("holidayshifttime"        ).Value + _
                   Rs_dutyrostertbl.Fields.Item("holidayshiftlate"        ).Value
    lastmonth_workingmins = Rs_dutyrostertbl.Fields.Item("workingmins").Value   ' 労働時間(分)
    lastmonth_currentworkmin = Rs_dutyrostertbl.Fields.Item("currentworkmin").Value ' 月内労働時間(分)
End If
sumHolidaynumber        = 0     ' 振替残日数集計項目
sumHolidaynumberHidden  = 0     ' 振替残日数上長承認分集計項目
sumVacationnumberHidden = 0     ' 有休残日数上長承認分集計項目
sumVacationtimeHidden   = 0     ' 時間有休取得日数上長承認分集計項目
Rs_dutyrostertbl.Close()

' 当年度データの集計読込み(時間外累積時間, 休日出勤累積時間)
Set Rs_dutyrostertbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_dutyrostertbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_dutyrostertbl_cmd.CommandText = "SELECT SUM(overtime + holidayshiftovertime + overtimelate + holidayshiftovertimelate + weekovertime) AS sumovertime " & _
                                   ",SUM(holidayshifttime + holidayshiftlate) AS sumholidaytime " & _
                                   "FROM dbo.dutyrostertbl WHERE personalcode = ? AND ymb >= ? AND ymb < ?"
Rs_dutyrostertbl_cmd.Prepared = true
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param2", 200, 1, 6, businessYear)
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param3", 200, 1, 6, ymb)
Set Rs_dutyrostertbl = Rs_dutyrostertbl_cmd.Execute
Rs_dutyrostertbl_numRows = 0
If Rs_dutyrostertbl.EOF And Rs_dutyrostertbl.BOF Then
    yearlyOvertime    = 0   ' 時間外累積時間
    yearlyHolidaytime = 0   ' 休日出勤累積時間
Else
    yearlyOvertime    = Rs_dutyrostertbl.Fields.Item("sumovertime").Value       ' 時間外累積時間
    yearlyHolidaytime = Rs_dutyrostertbl.Fields.Item("sumholidaytime").Value    ' 休日出勤累積時間
End If
' 集計結果が数値でないとき、ゼロを設定
If Not(IsNumeric(yearlyOvertime)) Then
    yearlyOvertime    = 0   ' 時間外累積時間
End If
If Not(IsNumeric(yearlyHolidaytime)) Then
    yearlyHolidaytime = 0   ' 休日出勤累積時間
End If
Rs_dutyrostertbl.Close()

' 時間外計(休出含む)求める(2, 3, 4, 5カ月前)
baseYMD = CDate(Left(ymb,4) & "/" & Right(ymb,2) & "/01")
' 2カ月前分
baseYMD = DateAdd("m", -2, baseYMD)
Set Rs_dutyrostertbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_dutyrostertbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_dutyrostertbl_cmd.CommandText = "SELECT personalcode, ymb, overtime + holidayshiftovertime + overtimelate + " & _
                                   "holidayshiftovertimelate + holidayshifttime + holidayshiftlate + weekovertime AS sumovertime FROM dutyrostertbl " & _
                                   "WHERE personalcode=? AND ymb=?"
Rs_dutyrostertbl_cmd.Prepared = true
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param2", 200, 1, 6, Year(baseYMD) & Right("0" & Month(baseYMD), 2))
Set Rs_dutyrostertbl = Rs_dutyrostertbl_cmd.Execute
Rs_dutyrostertbl_numRows = 0
If Rs_dutyrostertbl.EOF And Rs_dutyrostertbl.BOF Then
Else
    sumOvertime2 = Rs_dutyrostertbl.Fields.Item("sumovertime").Value 
End If
Rs_dutyrostertbl.Close()
' 3カ月前分
baseYMD = DateAdd("m", -1, baseYMD)
Set Rs_dutyrostertbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_dutyrostertbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_dutyrostertbl_cmd.CommandText = "SELECT personalcode, ymb, overtime + holidayshiftovertime + overtimelate + " & _
                                   "holidayshiftovertimelate + holidayshifttime + holidayshiftlate + weekovertime AS sumovertime FROM dutyrostertbl " & _
                                   "WHERE personalcode=? AND ymb=?"
Rs_dutyrostertbl_cmd.Prepared = true
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param2", 200, 1, 6, Year(baseYMD) & Right("0" & Month(baseYMD), 2))
Set Rs_dutyrostertbl = Rs_dutyrostertbl_cmd.Execute
Rs_dutyrostertbl_numRows = 0
If Rs_dutyrostertbl.EOF And Rs_dutyrostertbl.BOF Then
Else
    sumOvertime3 = Rs_dutyrostertbl.Fields.Item("sumovertime").Value 
End If
Rs_dutyrostertbl.Close()
' 4カ月前分
baseYMD = DateAdd("m", -1, baseYMD)
Set Rs_dutyrostertbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_dutyrostertbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_dutyrostertbl_cmd.CommandText = "SELECT personalcode, ymb, overtime + holidayshiftovertime + overtimelate + " & _
                                   "holidayshiftovertimelate + holidayshifttime + holidayshiftlate + weekovertime AS sumovertime FROM dutyrostertbl " & _
                                   "WHERE personalcode=? AND ymb=?"
Rs_dutyrostertbl_cmd.Prepared = true
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param2", 200, 1, 6, Year(baseYMD) & Right("0" & Month(baseYMD), 2))
Set Rs_dutyrostertbl = Rs_dutyrostertbl_cmd.Execute
Rs_dutyrostertbl_numRows = 0
If Rs_dutyrostertbl.EOF And Rs_dutyrostertbl.BOF Then
Else
    sumOvertime4 = Rs_dutyrostertbl.Fields.Item("sumovertime").Value 
End If
Rs_dutyrostertbl.Close()
' 5カ月前分
baseYMD = DateAdd("m", -1, baseYMD)
Set Rs_dutyrostertbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_dutyrostertbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_dutyrostertbl_cmd.CommandText = "SELECT personalcode, ymb, overtime + holidayshiftovertime + overtimelate + " & _
                                   "holidayshiftovertimelate + holidayshifttime + holidayshiftlate + weekovertime AS sumovertime FROM dutyrostertbl " & _
                                   "WHERE personalcode=? AND ymb=?"
Rs_dutyrostertbl_cmd.Prepared = true
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param2", 200, 1, 6, Year(baseYMD) & Right("0" & Month(baseYMD), 2))
Set Rs_dutyrostertbl = Rs_dutyrostertbl_cmd.Execute
Rs_dutyrostertbl_numRows = 0
If Rs_dutyrostertbl.EOF And Rs_dutyrostertbl.BOF Then
Else
    sumOvertime5 = Rs_dutyrostertbl.Fields.Item("sumovertime").Value 
End If
Rs_dutyrostertbl.Close()

' 累計有休取得日数
searchGrantymb = ""
If (Right(ymb, 2) < Left(grantdate, 2)) Then
    ' 付与日は前年
    searchGrantymb = Left(ymb,4) - 1 & Left(grantdate,2)
Else
    ' 付与日は当年
    searchGrantymb = Left(ymb,4) & Left(grantdate,2)
End If
Set Rs_dutyrostertbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_dutyrostertbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_dutyrostertbl_cmd.CommandText = "SELECT SUM(paidvacations) AS totalPaidvacations FROM dutyrostertbl " & _
                                   "WHERE personalcode=? AND ymb>=? AND ymb<=?"
Rs_dutyrostertbl_cmd.Prepared = true
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param2", 200, 1, 6, searchGrantymb)
Rs_dutyrostertbl_cmd.Parameters.Append Rs_dutyrostertbl_cmd.CreateParameter("param3", 200, 1, 6, ymb)
Set Rs_dutyrostertbl = Rs_dutyrostertbl_cmd.Execute
Rs_dutyrostertbl_numRows = 0
If Rs_dutyrostertbl.EOF And Rs_dutyrostertbl.BOF Then
Else
    totalPaidvacations = Rs_dutyrostertbl.Fields.Item("totalPaidvacations").Value 
End If
Rs_dutyrostertbl.Close()

Set Rs_dutyrostertbl = Nothing

' 当年度データの集計読込み(休出回数)当月含まず
Set Rs_worktbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_worktbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_worktbl_cmd.CommandText = "SELECT COUNT(*) AS holidaywork FROM dbo.worktbl " & _
                             "WHERE personalcode = ? AND workingdate >= ? AND workingdate < ? AND " & _
                             "(morningwork IN ('2', '3', '6') OR afternoonwork IN ('2', '3', '6'))"
Rs_worktbl_cmd.Prepared = true
Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param2", 200, 1, 8, businessYear & "00")
Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param3", 200, 1, 8, ymb & "00")
Set Rs_worktbl = Rs_worktbl_cmd.Execute
Rs_worktbl_numRows = 0
If Rs_worktbl.EOF And Rs_worktbl.BOF Then
    yearlyHolidaywork = 0   ' 休出累積回数
Else
    yearlyHolidaywork = Rs_worktbl.Fields.Item("holidaywork").Value ' 休出累積回数
End If
' 集計結果が数値でないとき、ゼロを設定
If Not(IsNumeric(yearlyHolidaywork)) Then
    yearlyHolidaywork = 0   ' 休出累積回数
End If
Rs_worktbl.Close()
' 当月休出回数
Set Rs_worktbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_worktbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_worktbl_cmd.CommandText = "SELECT COUNT(*) AS holidaywork FROM dbo.worktbl " & _
                             "WHERE personalcode = ? AND workingdate LIKE ? AND " & _
                             "(morningwork IN ('2', '3', '6') OR afternoonwork IN ('2', '3', '6'))"
Rs_worktbl_cmd.Prepared = true
Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param2", 200, 1, 8, ymb & "%")
Set Rs_worktbl = Rs_worktbl_cmd.Execute
Rs_worktbl_numRows = 0
If Rs_worktbl.EOF And Rs_worktbl.BOF Then
    monthlyHolidaywork = 0   ' 休出累積回数
Else
    monthlyHolidaywork = Rs_worktbl.Fields.Item("holidaywork").Value ' 休出累積回数
End If
' 集計結果が数値でないとき、ゼロを設定
If Not(IsNumeric(monthlyHolidaywork)) Then
    monthlyHolidaywork = 0   ' 休出累積回数
End If
Rs_worktbl.Close()

Set Rs_worktbl = Nothing

' -----------------------------------------------------------------------------
' 勤怠テーブル worktbl 読込
' -----------------------------------------------------------------------------
Dim Rs_worktbl
Dim Rs_worktbl_cmd
Dim Rs_worktbl_numRows
Set Rs_worktbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_worktbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_worktbl_cmd.CommandText = "SELECT CONVERT(int,w1.updatetime) AS inttime, w1.id, w1.updatetime, w1.personalcode, w1.workingdate, w1.morningwork, " & _
    "w1.afternoonwork, w1.morningholiday, w1.afternoonholiday, w1.summons, w1.overtime_begin, w1.overtime_end, " & _
    "w1.rest_begin, w1.rest_end, w1.overtime, w1.overtimelate, w1.holidayshift, w1.holidayshiftovertime, " & _
    "w1.holidayshiftlate, w1.holidayshiftovertimelate, w1.requesttime, w1.requesttime_begin, w1.requesttime_end, " & _
    "w1.latetime, w1.latetime_begin, w1.latetime_end, w1.is_approval, w1.nightduty, w1.dayduty, w1.operator, " & _
    "w1.vacationtime, w1.vacationtime_begin, w1.vacationtime_end, w1.memo, w1.memo2, w1.is_error, " & _
    "w1.work_begin, w1.work_end, w1.break_begin1, w1.break_end1, w1.break_begin2, w1.break_end2, w1.workmin, " & _
    "w2.nightduty AS nightduty2, w2.operator AS operator2, w3.cumulative_workmin, w3.overwork, " & _
    "w1.weekovertime " & _
    "FROM dbo.worktbl w1 " & _
    "LEFT JOIN dbo.worktbl w2 ON w1.personalcode = w2.personalcode AND CONVERT(NVARCHAR, DATEADD(day, -1, CONVERT(nvarchar, w1.workingdate, 112)), 112) = w2.workingdate " & _
    "LEFT JOIN (SELECT personalcode, workingdate, cumulative_workmin, " & _
    "CASE WHEN CEILING(cumulative_workmin / 60.0) > 40 THEN 'overwork' ELSE '' END AS overwork " & _
    "FROM (SELECT personalcode, workingdate, workmin, SUM(workmin) OVER (PARTITION BY personalcode, " & _
    "FORMAT(DATEADD(day, 1-DATEPART(WEEKDAY, workingdate), workingdate), 'yyyyMMdd') ORDER BY workingdate ) AS cumulative_workmin " & _
    "FROM worktbl " & _
    "WHERE personalcode = ? AND workingdate BETWEEN FORMAT(DATEADD(day, -1*DATEPART(WEEKDAY, ?)+1, ?), 'yyyyMMdd') AND ?) r " & _
    "WHERE workingdate >= ? " & _
    ") w3 ON w1.personalcode = w3.personalcode AND w1.workingdate=w3.workingdate " & _
    "WHERE w1.personalcode = ? AND w1.workingdate LIKE ? ORDER BY w1.workingdate ASC"
Rs_worktbl_cmd.Prepared = true
Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param2", 200, 1, 8, ymb & "01")
Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param3", 200, 1, 8, ymb & "01")
Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param4", 200, 1, 8, ymb & "31")
Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param5", 200, 1, 8, ymb & "01")
Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param6", 200, 1, 5, target_personalcode)
Rs_worktbl_cmd.Parameters.Append Rs_worktbl_cmd.CreateParameter("param7", 200, 1, 7, ymb & "%")
Set Rs_worktbl = Rs_worktbl_cmd.Execute
Rs_worktbl_numRows = 0

' -----------------------------------------------------------------------------
' タイムテーブル timetbl 読込
' -----------------------------------------------------------------------------
Set Rs_timetbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_timetbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_timetbl_cmd.CommandText = "SELECT * FROM dbo.timetbl WHERE personalcode = ? AND workingdate LIKE ? ORDER BY workingdate ASC"
Rs_timetbl_cmd.Prepared = true
Rs_timetbl_cmd.Parameters.Append Rs_timetbl_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
Rs_timetbl_cmd.Parameters.Append Rs_timetbl_cmd.CreateParameter("param2", 200, 1, 7, ymb & "%")
Set Rs_timetbl = Rs_timetbl_cmd.Execute
Rs_timetbl_numRows = 0

' -----------------------------------------------------------------------------
' 公休日テーブル holidaytbl 読込
' -----------------------------------------------------------------------------
Set Rs_holidaytbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_holidaytbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_holidaytbl_cmd.CommandText = "SELECT * FROM dbo.holidaytbl " & "WHERE holidaydate LIKE ? AND holidaytype = ? ORDER BY holidaydate ASC"
Rs_holidaytbl_cmd.Prepared = true
Rs_holidaytbl_cmd.Parameters.Append Rs_holidaytbl_cmd.CreateParameter("param1", 200, 1, 7, ymb & "%")
Rs_holidaytbl_cmd.Parameters.Append Rs_holidaytbl_cmd.CreateParameter("param2", 200, 1, 1, holidaytype)
Set Rs_holidaytbl = Rs_holidaytbl_cmd.Execute
Rs_holidaytbl_numRows = 0

' -----------------------------------------------------------------------------
' 保存休暇残日数
' -----------------------------------------------------------------------------
Set Rs_remainvacationtbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_remainvacationtbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_remainvacationtbl_cmd.CommandText = "SELECT r.personalcode, IsNULL(r.remainvacation, 0) AS remainvacation, " & _
        "IsNULL(SUM(p.preservevacations), 0) AS preservevacations FROM " & _
        "(SELECT * FROM (" & _
        " SELECT personalcode, ymb, remainvacation, ROW_NUMBER() OVER (ORDER BY ymb DESC) AS rownum FROM remainvacationtbl " & _
        "WHERE personalcode= ? AND ymb <= ? " & _
        " ) rv WHERE rownum=1) r " & _
        "LEFT JOIN " & _
        "(SELECT personalcode, ymb, preservevacations FROM dutyrostertbl WHERE personalcode= ? ) p " & _
        "ON r.personalcode=p.personalcode AND r.ymb<=p.ymb AND ? >=p.ymb " & _
        "GROUP BY r.personalcode, r.ymb, r.remainvacation"
Rs_remainvacationtbl_cmd.Prepared = true
Rs_remainvacationtbl_cmd.Parameters.Append Rs_remainvacationtbl_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
Rs_remainvacationtbl_cmd.Parameters.Append Rs_remainvacationtbl_cmd.CreateParameter("param2", 200, 1, 6, ymb)
Rs_remainvacationtbl_cmd.Parameters.Append Rs_remainvacationtbl_cmd.CreateParameter("param3", 200, 1, 5, target_personalcode)
Rs_remainvacationtbl_cmd.Parameters.Append Rs_remainvacationtbl_cmd.CreateParameter("param4", 200, 1, 6, ymb)
Set Rs_remainvacationtbl = Rs_remainvacationtbl_cmd.Execute
Rs_remainvacationtbl_numRows = 0

' -----------------------------------------------------------------------------
' typetbl 備考リストボックス項目を読み込み、配列(memoArray)作成
' -----------------------------------------------------------------------------
Set Rs_typetbl_memo_cmd = Server.CreateObject ("ADODB.Command")
Rs_typetbl_memo_cmd.ActiveConnection = MM_workdbms_STRING
Rs_typetbl_memo_cmd.CommandText = "SELECT * FROM typetbl WHERE CODETYPE='memo' ORDER BY dispseq"
Rs_typetbl_memo_cmd.Prepared = true
Set Rs_typetbl_memo = Rs_typetbl_memo_cmd.Execute
Rs_typetbl_memo_numRows = 0

Dim memoArray(10)
If Not Rs_typetbl_memo.EOF Or Not Rs_typetbl_memo.BOF Then
    While (NOT Rs_typetbl_memo.EOF)
        memoArray(Rs_typetbl_memo.Fields.Item("code").Value) = Trim(Rs_typetbl_memo.Fields.Item("codetext").Value)
        Rs_typetbl_memo.MoveNext()
    Wend
End If
Rs_typetbl_memo.Close()
Set Rs_typetbl_memo = Nothing

' 電源オン時間データ読み込み
Dim Rs_pctimetbl_on
Dim Rs_pctimetbl_on_cmd
Dim Rs_pctimetbl_on_numRows
Set Rs_pctimetbl_on_cmd = Server.CreateObject ("ADODB.Command")
Rs_pctimetbl_on_cmd.ActiveConnection = MM_workdbms_STRING
pc_ontime = " "
Rs_pctimetbl_on_cmd.CommandText = "SELECT * FROM (" & _
"SELECT i.personalcode, t.pcdate, t.pctime, " & _
"row_number() over(partition by t.pcdate order by t.pcdate, t.pctime) row_num " & _
"FROM iptbl i LEFT JOIN pctimetbl t " & _
"ON i.ipnumber = t.ipnumber AND '" & ymb & "' = SUBSTRING(t.pcdate,1,6) AND '電源ON' = t.pcstatus " & _
"WHERE i.personalcode='" & target_personalcode & "' " & _
"AND SUBSTRING(i.begindate,1,6) <= '" & ymb & "'" & _
"AND SUBSTRING(i.enddate,1,6) >= '" & ymb & "' ) j WHERE row_num=1 ORDER BY pcdate"
Rs_pctimetbl_on_cmd.Prepared = true
Set Rs_pctimetbl_on = Rs_pctimetbl_on_cmd.Execute
' 電源オフ時間データ読み込み
Dim Rs_pctimetbl_off
Dim Rs_pctimetbl_off_cmd
Dim Rs_pctimetbl_off_numRows
Set Rs_pctimetbl_off_cmd = Server.CreateObject ("ADODB.Command")
Rs_pctimetbl_off_cmd.ActiveConnection = MM_workdbms_STRING
pc_ontime = " "
Rs_pctimetbl_off_cmd.CommandText = "SELECT * FROM (" & _
"SELECT i.personalcode, t.pcdate, t.pctime, " & _
"row_number() over(partition by t.pcdate order by t.pcdate, t.pctime DESC) row_num " & _
"FROM iptbl i LEFT JOIN pctimetbl t " & _
"ON i.ipnumber = t.ipnumber AND '" & ymb & "' = SUBSTRING(t.pcdate,1,6) AND '電源OFF' = t.pcstatus " & _
"WHERE i.personalcode='" & target_personalcode & "' " & _
"AND SUBSTRING(i.begindate,1,6) <= '" & ymb & "'" & _
"AND SUBSTRING(i.enddate,1,6) >= '" & ymb & "' ) j WHERE row_num=1 ORDER BY pcdate"
Rs_pctimetbl_off_cmd.Prepared = true
Set Rs_pctimetbl_off = Rs_pctimetbl_off_cmd.Execute

checkHolidayMM1 = ""        ' 表示月の1カ月後チェック用月
checkHolidayMM3 = Array("") ' 表示月の3カ月後チェック用月配列
' 有給休暇取得チェック用年月の設定
If     Right(ymb,2) = "01" Then
    checkHolidayMM1 = "02"
    checkHolidayMM3 = Array("02","03","04")
ElseIf Right(ymb,2) = "02" Then
    checkHolidayMM1 = "03"
    checkHolidayMM3 = Array("03","04","05")
ElseIf Right(ymb,2) = "03" Then
    checkHolidayMM1 = "04"
    checkHolidayMM3 = Array("04","05","06")
ElseIf Right(ymb,2) = "04" Then
    checkHolidayMM1 = "05"
    checkHolidayMM3 = Array("05","06","07")
ElseIf Right(ymb,2) = "05" Then
    checkHolidayMM1 = "06"
    checkHolidayMM3 = Array("06","07","08")
ElseIf Right(ymb,2) = "06" Then
    checkHolidayMM1 = "07"
    checkHolidayMM3 = Array("07","08","09")
ElseIf Right(ymb,2) = "07" Then
    checkHolidayMM1 = "08"
    checkHolidayMM3 = Array("08","09","10")
ElseIf Right(ymb,2) = "08" Then
    checkHolidayMM1 = "09"
    checkHolidayMM3 = Array("09","10","11")
ElseIf Right(ymb,2) = "09" Then
    checkHolidayMM1 = "10"
    checkHolidayMM3 = Array("10","11","12")
ElseIf Right(ymb,2) = "10" Then
    checkHolidayMM1 = "11"
    checkHolidayMM3 = Array("11","12","01")
ElseIf Right(ymb,2) = "11" Then
    checkHolidayMM1 = "12"
    checkHolidayMM3 = Array("12","01","02")
ElseIf Right(ymb,2) = "12" Then
    checkHolidayMM1 = "01"
    checkHolidayMM3 = Array("01","02","03")
End If

' -----------------------------------------------------------------------------
' 基準労働時間テーブル baseworktimetbl 読込
' -----------------------------------------------------------------------------
' 対象月データ読込み
Dim Rs_baseworktimetbl
Dim Rs_baseworktimetbl_cmd
Dim Rs_baseworktimetbl_numRows
Set Rs_baseworktimetbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_baseworktimetbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_baseworktimetbl_cmd.CommandText = "SELECT * FROM dbo.baseworktimetbl WHERE personalcode IN ('00000', ?) AND ymb = ? ORDER BY personalcode DESC"
Rs_baseworktimetbl_cmd.Prepared = true
Rs_baseworktimetbl_cmd.Parameters.Append Rs_baseworktimetbl_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
Rs_baseworktimetbl_cmd.Parameters.Append Rs_baseworktimetbl_cmd.CreateParameter("param2", 200, 1, 6, ymb)
Set Rs_baseworktimetbl = Rs_baseworktimetbl_cmd.Execute
Rs_baseworktimetbl_numRows = 0
If Rs_baseworktimetbl.EOF And Rs_baseworktimetbl.BOF Then
    baseworkmin       = 0            ' 当月基準労働時間
    thismonth_basemin = baseworkmin  ' 当月基準労働時間(更新時に使用するためのhidden項目に設定)
Else
    baseworkmin       = Rs_baseworktimetbl.Fields.Item("basemin").Value  ' 当月基準労働時間
    thismonth_basemin = baseworkmin  ' 当月基準労働時間(更新時に使用するためのhidden項目に設定)
    ' 当月基準労働時間算出(前月労働時間不足分加算)
    If (lastmonth_currentworkmin > lastmonth_workingmins) Then
        baseworkmin = baseworkmin + lastmonth_currentworkmin - lastmonth_workingmins
    End If
End If
Rs_baseworktimetbl.Close()
' 当月基準労働時間
currentworkmin = baseworkmin
%>
