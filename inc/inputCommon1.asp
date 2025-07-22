<%
' 固定文言
strDay      = "&nbsp;日"
strTime     = "&nbsp;時間"
strCount    = "&nbsp;回"
strErrorMsg = "入力内容に誤りがあります。確認してください。"

' 本日
today    = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2)
' 表示対象の年月を設定
If Request.QueryString("ymb")<>"" Then
  ymb    = Request.QueryString("ymb")
Else
  ymb    = Year(Now) & Right("0" & Month(Now), 2)
End If
' 対象月前月設定
temp     = DateSerial(left(ymb, 4), right(ymb, 2) , 0)
lastYmb  = left(temp, 4) & mid(temp, 6, 2)
' 対象月翌月設定
temp     = DateSerial(left(ymb, 4), right(ymb, 2) , 32)
nextYmb  = left(temp, 4) & mid(temp, 6, 2)
' 対象年月末日
lastDay  = right(DateSerial(left(ymb, 4), right(ymb, 2) + 1, 0), 2)
' 対象年度4月算出
If Right(ymb, 2) > "03" Then
    businessYear = Left(ymb, 4) & "04"
Else
    businessYear = (Left(ymb, 4) - 1) & "04"
End If

' 勤務表入力可能未来年月（翌月まで)
temp    = DateSerial(Year(Now), right("0" & Month(Now), 2) , 32)
inputLimitYmb = left(temp, 4) & mid(temp, 6, 2)
' 勤務表チェック可能未来年月（当月まで)
checkLimitYmb = Year(Now) & Right("0" & Month(Now), 2)
' 上長チェック可能年月日（前日まで)
temp    = DateAdd("d", -1, Now())
inputLimitApprovalYmd = left(temp, 4) & mid(temp, 6, 2) & mid(temp, 9, 2)

' エラーメッセージ
errorMsg = ""

Dim style_beginTime                 (31)    ' タイムカード出社 スタイル
Dim style_endTime                   (31)    ' タイムカード退社 スタイル
Dim style_returnTime                (31)    ' タイムカード戻り スタイル
Dim style_outTime                   (31)    ' タイムカード外出 スタイル

Dim style_morningwork               (31)    ' 出勤区分(午前)
Dim style_afternoonwork             (31)    ' 出勤区分(午後)
Dim style_morningholiday            (31)    ' 休日区分(午前)
Dim style_afternoonholiday          (31)    ' 休日区分(午後)
Dim style_work_begin                (31)    ' フレックス勤務時間 自
Dim style_work_end                  (31)    ' フレックス勤務時間 至
Dim style_break_begin1              (31)    ' フレックス休憩 自
Dim style_break_end1                (31)    ' フレックス休憩 至
Dim style_break_begin2              (31)    ' フレックス中抜 自
Dim style_break_end2                (31)    ' フレックス中抜 至
Dim style_summons                   (31)    ' 呼出
Dim style_overtime_begin            (31)    ' 時間外(休出)申請分 自
Dim style_overtime_end              (31)    ' 時間外(休出)申請分 至
Dim style_rest_begin                (31)    ' 時間外(休出)申請分 休憩自
Dim style_rest_end                  (31)    ' 時間外(休出)申請分 休憩至
Dim style_requesttime_begin         (31)    ' 時間代休申請分 自
Dim style_requesttime_end           (31)    ' 時間代休申請分 至
Dim style_vacationtime_begin        (31)    ' 時間有給 自
Dim style_vacationtime_end          (31)    ' 時間有給 至
Dim style_latetime_begin            (31)    ' 深夜割増 自
Dim style_latetime_end              (31)    ' 深夜割増 至
Dim style_weekovertime              (31)    ' 週超過時間
Dim style_nightduty                 (31)    ' 宿直
Dim style_dayduty                   (31)    ' 日直
Dim style_operator                  (31)    ' オペレータ
Dim style_memo                      (31)    ' 備考欄
Dim dayErrorFlag                    (31)    ' エラーフラグ
                                            ' (時間外、時間外深夜業などを算出するとき、
                                            ' エラーが無い行のみ算出するためのフラグ)
Dim v_overtime
Dim v_overtimelate
Dim v_holidayshift
Dim v_holidayshiftovertime
Dim v_holidayshiftlate
Dim v_holidayshiftovertimelate
Dim v_flexovermin
Dim sepTimeAry
Dim sepTime1
Dim sepTime2

Dim op  ' 労働時間適正化：当日勤務のオペレータ判定用フラグ
Dim op2 ' 労働時間適正化：前日勤務のオペレータ判定用フラグ

' -----------------------------------------------------------------------------
' 画面表示の切り分け (入力画面、上長チェック画面、労務担当者確認画面)
' -----------------------------------------------------------------------------
screen = 0 ' 画面フラグ 0:入力画面 1:参照画面 2:上長チェック画面
If Request.QueryString("p")<>"" And Session("MM_Username") <> Request.QueryString("p") Then
    ' 参照系
    screen = 1
End If
If (Request.QueryString("c")<>"" And Session("MM_is_charge"   )="1") Then
    screen = 1  ' 労務担当者
End If
If (Request.QueryString("s")<>"" And Session("MM_is_superior" )="1") Then
    screen = 2  ' 上長
End If

' 個人勤務表入力画面の設定
If screen = 0 Then
    target_personalcode = Session("MM_Username")
Else
    ' 上長チェック、労務担当者確認画面時はURLパラメータの個人コードを設定
    target_personalcode = Request.QueryString("p")
End If
personalcode = Session("MM_Username")
checkUser()

' 交替勤務画面変更切替年月
Dim beforeOpLimit '当画面利用可能下限年月
beforeOpLimit = "201509"

' 当日時間外14時間超チェックフラグ
Dim warn_time14over
warn_time14over = "0"

' フレックス勤務用変数
Dim baseworkmin     ' 基準労働時間(分)
Dim currentworkmin  ' 当月基準労働時間(分)
Dim realworkmin     ' 勤務実績(分)
baseworkmin    = 0
currentworkmin = 0
realworkmin    = 0

' 入力項目無効化設定項目
Dim button_come_disabled    ' 出社ボタン無効化
Dim button_leave_disabled   ' 退社ボタン無効化
Dim button_submit_disable   ' 登録ボタン無効化
Dim text_approval_disabled  ' 上長チェック無効化
Dim text_timecard_disabled  ' タイムカード項目無効化
Dim text_disabled           ' 入力画面項目
button_come_disabled   = ""
button_leave_disabled  = ""
button_submit_disable  = ""
text_approval_disabled = ""
text_timecard_disabled = ""
text_disabled          = ""


' -------------------------------------------------------------------------
' 入力チェック用変数
' -------------------------------------------------------------------------
Dim vacation_count
Dim holiday_count
Dim overtime_count
Dim requesttime_count
Dim vacationtime_count
Dim overtimeonly_count  ' 当月時間外労働時間(休出除く)
vacation_count     = 0
holiday_count      = 0
overtime_count     = 0
requesttime_count  = 0
vacationtime_count = 0
overtimeonly_count = 0
' 時刻チェック用フラグ
err_beginTime           = 0
err_returnTime          = 0
err_outTime             = 0
err_endTime             = 0

' フレックス時間
err_work_begin          = 0
err_work_end            = 0
err_break_begin1        = 0
err_break_end1          = 0
err_break_begin2        = 0
err_break_end2          = 0

err_overtime_begin      = 0
err_overtime_end        = 0
err_rest_begin          = 0
err_rest_end            = 0
err_requesttime_begin   = 0
err_requesttime_end     = 0
err_vacationtime_begin  = 0
err_vacationtime_end    = 0
err_latetime_begin      = 0
err_latetime_end        = 0
err_weekovertime        = 0
' 休出カウント 半日未満だけでも午前午後とも休出でも1日1回とカウントする
holidaywork_count       = 0
' 関連チェック用フラグ
err_relation_01 = 0
err_relation_02 = 0
err_relation_03 = 0
err_relation_04 = 0
err_relation_05 = 0
err_relation_06 = 0
err_relation_07 = 0
err_relation_08 = 0
err_relation_09 = 0
err_relation_10 = 0
err_relation_20 = 0
err_relation_21 = 0
err_relation_22 = 0
err_relation_23 = 0
err_relation_24 = 0
err_relation_25 = 0
err_relation_26 = 0
err_relation_27 = 0
err_relation_28 = 0
err_relation_29 = 0
err_relation_30 = 0
err_relation_31 = 0
err_relation_32 = 0
err_relation_33 = 0
err_relation_34 = 0
err_relation_35 = 0
err_relation_36 = 0
err_relation_37 = 0
err_relation_38 = 0
err_relation_39 = 0
err_relation_40 = 0
err_relation_41 = 0
err_relation_42 = 0 ' フレックス勤務 勤務自至時間関連チェック
err_relation_43 = 0 ' フレックス勤務 休憩自至時間関連チェック
err_relation_44 = 0 ' フレックス勤務 中抜自至時間関連チェック
err_relation_45 = 0 ' フレックス勤務 勤務時間と休憩時間の関連チェック
'err_relation_46 = 0 ' フレックス勤務 休憩時間と中抜時間の関連チェック
err_relation_47 = 0 ' フレックス勤務 勤務時間と休憩時間の整合性チェック
err_relation_48 = 0 ' フレックス勤務で休暇時に勤務休憩時間入力はエラー
err_relation_49 = 0 ' 時間外未入力で深夜割増入力時はエラー
err_relation_50 = 0 ' 法定休日のとき、勤務入力はエラー
err_relation_51 = 0 ' コアタイム有休入力時は、午前午後ともコアタイム有休でないとエラー
err_relation_52 = 0 ' フレックス勤務者の時間有給と勤務時間、休憩時間、中抜け時間との整合性チェック
err_relation_53 = 0 ' 勤務時間と時間外の時間重複をチェック
err_relation_54 = 0 ' フレックス勤務者の公休日に休出は入力できない
err_relation_55 = 0 ' 育児休業には出勤区分は入力できない
%>
