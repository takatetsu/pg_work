<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<!-- #include file="Connections/workdbms.asp" -->
<%
' #######################################
' プログラム仕様書
' #######################################
'
' ## プログラム概要 ##
' 支店控除入力画面で登録した内容を CSV ファイルとしてダウンロードする
'
' ## 出力項目 ##
' CSV ファイルは支店控除入力画面に準じています。
'
' ## 入力チェック ##
'
' ## 注意事項 ##
' EXCEL ファイルとしてダウンロードも可能だが、ファイルを開くたびに
' 破損しているか確認…のメッセージが表示されるので、 CSV ファイル
' としてダウンロードする。

' 日付の計算
Dim sysDate     'システム日付
Dim dispDate    '表示用日付
Dim dispYear    '表示用年 yyyy
Dim dispMonth   '表示用月 mm
Dim i           '繰り返し用日付

If (Request.QueryString("ymb")<>"") Then
    dispDate = DateSerial(Mid(Request.QueryString("ymb"), 1, 4), Mid(Request.QueryString("ymb"), 5, 2), 1)
Else
    dispDate = Date
End If
dispYear    = Year(dispDate)
dispMonth   = Right("0" & Month(dispDate), 2)

' -----------------------------------------------------------------------------
' エクセル出力指示
' -----------------------------------------------------------------------------
Response.BUFFER=TRUE
Response.ContentType = "application/vnd.ms-excel"
Response.Charset = "utf-8"
Response.AddHeader "Content-Disposition","attachment; filename=支店控除_" & _
                    dispYear & "年" & dispMonth & "月分.csv"

' deductiontblより、担当者のymd最新レコードを取得(未登録、登録済　判定用)
Dim Rs_dedu_ck
Dim Rs_dedu_ck_cmd
Set Rs_dedu_ck_cmd = Server.CreateObject ("ADODB.Command")
Rs_dedu_ck_cmd.ActiveConnection = MM_workdbms_STRING
Rs_dedu_ck_cmd.CommandText = "SELECT * FROM deductiontbl WHERE ymb = (SELECT MAX(ymb) FROM deductiontbl WHERE personalcode = ? ) AND personalcode = ?"
Rs_dedu_ck_cmd.Prepared = true
Rs_dedu_ck_cmd.Parameters.Append Rs_dedu_ck_cmd.CreateParameter("param1", 200, 1, -1, Session("MM_Username") )
Rs_dedu_ck_cmd.Parameters.Append Rs_dedu_ck_cmd.CreateParameter("param1", 200, 1, -1, Session("MM_Username") )
Set Rs_dedu_ck = Rs_dedu_ck_cmd.Execute
If Not Rs_dedu_ck.EOF Or Not Rs_dedu_ck.BOF Then
    other1Name = Trim(Rs_dedu_ck.Fields.Item("amount08ncr"))
    other2Name = Trim(Rs_dedu_ck.Fields.Item("amount09ncr"))
    other3Name = Trim(Rs_dedu_ck.Fields.Item("amount10ncr"))
End If
'担当者のymd最新レコードをクローズ
Rs_dedu_ck.Close()
Set Rs_dedu_ck = Nothing

Response.Write("個人CD,氏名,控除額合計,火災共済,交通災害,駐車場代," & _
    "住宅共益費,水道代,合格祝金,支部費(組合),その他１:" & other1Name & _
    ",その他２:" & other2Name & ",その他３:" & other3Name & vbNewLine)

' stafftblより、表示職員一覧を取得
Dim Rs_staff
Dim Rs_staff_cmd
Set Rs_staff_cmd = Server.CreateObject ("ADODB.Command")
Rs_staff_cmd.ActiveConnection = MM_workdbms_STRING
Rs_staff_cmd.CommandText = "SELECT stafftbl.personalcode ,stafftbl.staffname FROM orgtbl " & _
        "RIGHT OUTER JOIN stafftbl stafftbl ON orgtbl.orgcode = stafftbl.orgcode " & _
        "WHERE stafftbl.is_enable = '1' AND orgtbl.personalcode = ?  AND " & _
        "orgtbl.manageclass = '0' ORDER BY stafftbl.orgcode, stafftbl.gradecode DESC, stafftbl.personalcode"
Rs_staff_cmd.Prepared = true
Rs_staff_cmd.Parameters.Append Rs_staff_cmd.CreateParameter("param1", 200, 1, -1, Session("MM_Username") )
'表示職員の一覧取得SQLの発行
Set Rs_staff = Rs_staff_cmd.Execute

'表示職員の数だけ、データの更新を行う
Dim perCod
While (NOT Rs_staff.EOF)
    perCod = Rs_staff.Fields.Item("personalcode")

    'deductiontblより、最新月の支店控除読み込み
    Dim Rs_dedu
    Dim Rs_dedu_cmd
    Set Rs_dedu_cmd = Server.CreateObject ("ADODB.Command")
    Rs_dedu_cmd.ActiveConnection = MM_workdbms_STRING
    Rs_dedu_cmd.CommandText = "SELECT * FROM deductiontbl WHERE ymb = ? AND personalcode = ?"
    Rs_dedu_cmd.Prepared = true
    Rs_dedu_cmd.Parameters.Append Rs_dedu_cmd.CreateParameter("param1", 200, 1, -1, Request.QueryString("ymb") )
    Rs_dedu_cmd.Parameters.Append Rs_dedu_cmd.CreateParameter("param1", 200, 1, -1, perCod )
    Set Rs_dedu = Rs_dedu_cmd.Execute

    sumAmount = Rs_dedu.Fields.Item("amount01") + _
                Rs_dedu.Fields.Item("amount02") + _
                Rs_dedu.Fields.Item("amount03") + _
                Rs_dedu.Fields.Item("amount04") + _
                Rs_dedu.Fields.Item("amount05") + _
                Rs_dedu.Fields.Item("amount06") + _
                Rs_dedu.Fields.Item("amount07") + _
                Rs_dedu.Fields.Item("amount08") + _
                Rs_dedu.Fields.Item("amount09") + _
                Rs_dedu.Fields.Item("amount10")

    Response.Write("=""" & perCod & """," & _
        Trim(Rs_staff.Fields.Item("staffname")) & _
        "," & sumAmount                         & _
        "," & Rs_dedu.Fields.Item("amount01")   & _
        "," & Rs_dedu.Fields.Item("amount02")   & _
        "," & Rs_dedu.Fields.Item("amount03")   & _
        "," & Rs_dedu.Fields.Item("amount04")   & _
        "," & Rs_dedu.Fields.Item("amount05")   & _
        "," & Rs_dedu.Fields.Item("amount06")   & _
        "," & Rs_dedu.Fields.Item("amount07")   & _
        "," & Rs_dedu.Fields.Item("amount08")   & _
        "," & Rs_dedu.Fields.Item("amount09")   & _
        "," & Rs_dedu.Fields.Item("amount10") & vbNewLine)
    Rs_staff.MoveNext()
Wend
'表示職員一覧を取得するSQLをクローズ
Rs_staff.Close()
Set Rs_staff = Nothing
Response.Flush
Response.End
Response.Redirect("inputdeduction.asp")
%>
