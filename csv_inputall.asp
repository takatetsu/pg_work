<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<!-- #include file="Connections/workdbms.asp" -->
<%
' #######################################
' プログラム仕様書
' #######################################
'
' ## プログラム概要 ##
' 全体入力画面で登録した内容を CSV ファイルとしてダウンロードする
'
' ## 出力項目 ##
' CSV ファイルは全体入力画面に準じています。
'
' ## 入力チェック ##
'
' ## 注意事項 ##
' EXCEL ファイルとしてダウンロードも可能だが、ファイルを開くたびに
' 破損しているか確認…のメッセージが表示されるので、 CSV ファイル
' としてダウンロードする。
'
' -----------------------------------------------------------------------------
' 初期処理
' -----------------------------------------------------------------------------
' 日付計算
Dim sysDate     ' システム日付
Dim dispDate    ' 表示用日付
Dim dispYear    ' 表示用年 yyyy
Dim dispMonth   ' 表示用月 mm
Dim lastDay     ' 対象年月末日

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

' -----------------------------------------------------------------------------
' エクセル出力指示
' -----------------------------------------------------------------------------
Response.BUFFER=TRUE
Response.ContentType = "application/vnd.ms-excel"
Response.Charset = "utf-8"
Response.AddHeader "Content-Disposition","attachment; filename=勤務表_" & _
                    dispYear & "年" & dispMonth & "月分.csv"

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
    "(SELECT orgcode FROM dbo.orgtbl "                                                  & _
    "WHERE personalcode='" & Session("MM_Username") & "' AND manageclass='1') ORG "     & _
    "LEFT JOIN "                                                                        & _
    "(SELECT personalcode AS pcode, staffname, orgcode AS org, gradecode AS grade FROM dbo.stafftbl "       & _
    "WHERE is_enable='1') STAFF "                                                       & _
    "ON ORG.orgcode=STAFF.org "                                                         & _
    "LEFT JOIN "                                                                        & _
    "(SELECT * FROM dbo.dutyrostertbl WHERE ymb='" & dispYear & dispMonth & "') DUTY "  & _
    "ON STAFF.pcode=DUTY.personalcode "                                                 & _
    "WHERE pcode IS NOT NULL "                                                          & _
    "ORDER BY org, grade DESC, STAFF.pcode ASC"
Rs_worktbl_cmd.Prepared = true

Set Rs_worktbl = Rs_worktbl_cmd.Execute
Rs_worktbl_numRows = 0
Response.Write("個人CD,氏名,可出勤日数,代休日数,欠勤日数,有給日数,保存休暇日数,"        & _
    "特休日数,休出日数,実出勤日数,遅早回数,宿直A回数,宿直B回数,宿直C回数,"              & _
    "宿直D回数,休日割増,日直,交替甲番,交替乙番,交替丙番,交替A番,交替B番,呼出通常回数,"  & _
    "呼出深夜回数,年末年始1230,年末年始1231,時間代休,深夜割増,時間外,休日出勤,"         & _
    "休出時外,休出深夜,時外深夜,休出時外深夜,土曜時間+100円,平日労働時間,法定休日割増" & vbNewLine)
While (NOT Rs_worktbl.EOF)
    personalcode = Trim(Rs_worktbl.Fields.Item("pcode"    ).Value)
    staffname    = Trim(Rs_worktbl.Fields.Item("staffname").Value)
    workdays                = Rs_worktbl.Fields.Item("workdays"                  ).Value
    workholidays            = Rs_worktbl.Fields.Item("workholidays"              ).Value
    absencedays             = Rs_worktbl.Fields.Item("absencedays"               ).Value
    paidvacations           = Rs_worktbl.Fields.Item("paidvacations"             ).Value
    preservevacations       = Rs_worktbl.Fields.Item("preservevacations"         ).Value
    specialvacations        = Rs_worktbl.Fields.Item("specialvacations"          ).Value
    holidayshifts           = Rs_worktbl.Fields.Item("holidayshifts"             ).Value
    realworkdays            = Rs_worktbl.Fields.Item("realworkdays"              ).Value
    shortdays               = Rs_worktbl.Fields.Item("shortdays"                 ).Value
    nightduty_a             = Rs_worktbl.Fields.Item("nightduty_a"               ).Value
    nightduty_b             = Rs_worktbl.Fields.Item("nightduty_b"               ).Value
    nightduty_c             = Rs_worktbl.Fields.Item("nightduty_c"               ).Value
    nightduty_d             = Rs_worktbl.Fields.Item("nightduty_d"               ).Value
    holidaypremium          = Rs_worktbl.Fields.Item("holidaypremium"            ).Value
    dayduty                 = Rs_worktbl.Fields.Item("dayduty"                   ).Value
    shiftwork_kou           = Rs_worktbl.Fields.Item("shiftwork_kou"             ).Value
    shiftwork_otsu          = Rs_worktbl.Fields.Item("shiftwork_otsu"            ).Value
    shiftwork_hei           = Rs_worktbl.Fields.Item("shiftwork_hei"             ).Value
    shiftwork_a             = Rs_worktbl.Fields.Item("shiftwork_a"               ).Value
    shiftwork_b             = Rs_worktbl.Fields.Item("shiftwork_b"               ).Value
    summons                 = Rs_worktbl.Fields.Item("summons"                   ).Value
    summonslate             = Rs_worktbl.Fields.Item("summonslate"               ).Value
    yearend1230             = Rs_worktbl.Fields.Item("yearend1230"               ).Value
    yearend1231             = Rs_worktbl.Fields.Item("yearend1231"               ).Value
    ' 時間代休に週超過労働時間を足し込む
    workholidaytime         = Rs_worktbl.Fields.Item("workholidaytime"           ).Value _
                            + Rs_worktbl.Fields.Item("weekovertime"              ).Value
    latepremium             = Rs_worktbl.Fields.Item("latepremium"               ).Value
    ' 時間外に週超過労働時間を足し込む
    overtime                = Rs_worktbl.Fields.Item("overtime"                  ).Value _
                            + Rs_worktbl.Fields.Item("weekovertime"              ).Value
    holidayshifttime        = Rs_worktbl.Fields.Item("holidayshifttime"          ).Value
    holidayshiftovertime    = Rs_worktbl.Fields.Item("holidayshiftovertime"      ).Value
    holidayshiftlate        = Rs_worktbl.Fields.Item("holidayshiftlate"          ).Value
    overtimelate            = Rs_worktbl.Fields.Item("overtimelate"              ).Value
    holidayshiftovertimelate= Rs_worktbl.Fields.Item("holidayshiftovertimelate"  ).Value
    saturdayworkmin         = Rs_worktbl.Fields.Item("saturday_workmin"          ).Value
    weekdaysworkmin         = Rs_worktbl.Fields.Item("weekdays_workmin"          ).Value
    legalholiday_extra_min  = Rs_worktbl.Fields.Item("legalholiday_extra_min"    ).Value
    Response.Write("=""" & personalcode & """," & staffname & "," & workdays & _
        "," & workholidays & "," & absencedays & "," & paidvacations & _
        "," & preservevacations & "," & specialvacations & "," & holidayshifts & _
        "," & realworkdays & "," & shortdays & "," & nightduty_a & _
        "," & nightduty_b & "," & nightduty_c & "," & nightduty_d & _
        "," & holidaypremium & "," & dayduty & "," & shiftwork_kou & _
        "," & shiftwork_otsu & "," & shiftwork_hei & "," & shiftwork_a & _
        "," & shiftwork_b & "," & summons & _
        "," & summonslate & "," & yearend1230 & "," & yearend1231 & _
        "," & workholidaytime & "," & latepremium & "," & overtime & _
        "," & holidayshifttime & "," & holidayshiftovertime & _
        "," & holidayshiftlate & "," & overtimelate & _
        "," & holidayshiftovertimelate & "," & saturdayworkmin & _
        "," & weekdaysworkmin & "," & legalholiday_extra_min & vbNewLine)
    Rs_worktbl.MoveNext()
Wend
Rs_worktbl.Close()
Set Rs_worktbl = Nothing
Response.Flush
Response.End
Response.Redirect("inputall.asp")
%>
