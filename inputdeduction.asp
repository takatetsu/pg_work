<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="Connections/workdbms.asp" -->
<!-- #include file="inc/RestrictAccess.asp" -->
<%
UpdateSuccess = "complete.asp"

' 日付の計算
Dim sysDate     'システム日付
Dim dispDate    '表示用日付
Dim dispYear    '表示用年 yyyy
Dim dispMonth   '表示用月 mm
Dim i           '繰り返し用日付
sysDate =Date
If (Request.QueryString("ymb")<>"") Then
    dispDate = DateSerial(Mid(Request.QueryString("ymb"), 1, 4), Mid(Request.QueryString("ymb"), 5, 2), 1)
Else
    dispDate = Date
End If
dispYear    = Year(dispDate)
dispMonth   = Right("0" & Month(dispDate), 2)
' 対象月前月設定
temp      = DateSerial(dispYear, dispMonth, 0)
lastYmb   = left(temp, 4) & mid(temp, 6, 2)
' 対象月翌月設定
temp      = DateSerial(dispYear, dispMonth , 32)
nextYmb   = left(temp, 4) & mid(temp, 6, 2)

' 入力可能月（システム日付の当月）
inputDisable = ""
If ((dispYear & dispMonth) = (Year(sysDate) & Right("0" & Month(sysDate), 2))) Then
    inputDisable = ""
Else
    inputDisable = "Disabled"
End If

Dim strErrorMsg
strErrorMsg = "入力内容に誤りがあります。確認してください。"

Dim ymb
Dim personalcode
Dim kasai
Dim koutu
Dim park
Dim kyoeki
Dim water
Dim gokaku
Dim kumiai
Dim ta1_
Dim ta2_
Dim ta3_

Dim style_kasai ()
Dim style_koutu ()
Dim style_park  ()
Dim style_kyoeki()
Dim style_water ()
Dim style_gokaku()
Dim style_kumiai()
Dim style_ta1_  ()
Dim style_ta2_  ()
Dim style_ta3_  ()

ReDim Preserve style_kasai (0)
ReDim Preserve style_koutu (0)
ReDim Preserve style_park  (0)
ReDim Preserve style_kyoeki(0)
ReDim Preserve style_water (0)
ReDim Preserve style_gokaku(0)
ReDim Preserve style_kumiai(0)
ReDim Preserve style_ta1_  (0)
ReDim Preserve style_ta2_  (0)
ReDim Preserve style_ta3_  (0)
%>
<%
' stafftblより、表示職員一覧を取得するSQL
Dim Rs_staff
Dim Rs_staff_cmd
Set Rs_staff_cmd = Server.CreateObject ("ADODB.Command")
Rs_staff_cmd.ActiveConnection = MM_workdbms_STRING
Rs_staff_cmd.CommandText = "SELECT stafftbl.personalcode ,stafftbl.staffname " & _
    "FROM orgtbl RIGHT OUTER JOIN stafftbl stafftbl ON orgtbl.orgcode = stafftbl.orgcode " & _
    "WHERE stafftbl.is_enable = '1' AND orgtbl.personalcode = ?  AND orgtbl.manageclass = '0' " & _
    "ORDER BY stafftbl.orgcode, stafftbl.gradecode DESC, stafftbl.personalcode"
Rs_staff_cmd.Prepared = true
Rs_staff_cmd.Parameters.Append Rs_staff_cmd.CreateParameter("param1", 200, 1, -1, Session("MM_Username") )

%>
<%
'フォームが送信されてきた場合、フォーム内容のチェック・更新を行う
If Request.Form("i_deduction") = "i_deduction" Then
    errorMsg = ""
    For i = 1 To Request.Form("personalcode").count Step 1
        setData()
    Next
    
    ' 入力チェックでエラーが無いとき、dutyrostertbl の更新処理を行う。
    If (errorMsg = "") Then
        For i = 1 To Request.Form("personalcode").count Step 1
            setData()
            If (Request.Form("ymb" & personalcode) <> (dispYear & dispMonth)) Then
                ' -------------------------------------------------------------
                ' データ登録処理 INSERT
                ' -------------------------------------------------------------
                MM_Dedu_SQL = "INSERT INTO dbo.deductiontbl(personalcode, ymb, amount01, amount02, amount03, amount04, amount05, amount06, amount07, amount08, amount09, amount10, amount08ncr, amount09ncr, amount10ncr) VALUES( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                Set MM_Dedu_cmd = Server.CreateObject ("ADODB.Command")
                MM_Dedu_cmd.ActiveConnection = MM_workdbms_STRING
                MM_Dedu_cmd.CommandText = MM_Dedu_SQL
                MM_Dedu_cmd.Prepared = true
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, personalcode)
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, dispYear & dispMonth)
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, kasai       )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, koutu       )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, park        )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, kyoeki      )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, water       )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, gokaku      )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, kumiai      )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, ta1_        )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, ta2_        )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, ta3_        )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, Request.Form("ta1_ncr"))
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, Request.Form("ta2_ncr"))
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, Request.Form("ta3_ncr"))
                Set MM_Dedu = MM_Dedu_cmd.Execute
            Else
                ' -------------------------------------------------------------
                ' データ更新処理 UPDATE
                ' -------------------------------------------------------------
                MM_Dedu_SQL = "UPDATE dbo.deductiontbl SET amount01 = ?, amount02 = ?, amount03 = ?, amount04 = ?, amount05 = ?, amount06 = ?, amount07 = ?, amount08 = ?, amount09 = ?, amount10 = ?, amount08ncr = ?, amount09ncr = ?, amount10ncr = ? WHERE personalcode = ? AND ymb = ?"
                Set MM_Dedu_cmd = Server.CreateObject ("ADODB.Command")
                MM_Dedu_cmd.ActiveConnection = MM_workdbms_STRING
                MM_Dedu_cmd.CommandText = MM_Dedu_SQL
                MM_Dedu_cmd.Prepared = true
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, kasai   )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, koutu   )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, park    )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, kyoeki  )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, water   )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, gokaku  )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, kumiai  )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, ta1_    )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, ta2_    )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, ta3_    )
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, Request.Form("ta1_ncr"))
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, Request.Form("ta2_ncr"))
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, Request.Form("ta3_ncr"))
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, personalcode)
                MM_Dedu_cmd.Parameters.Append MM_Dedu_cmd.CreateParameter("param1", 200, 1, -1, dispyear & dispmonth)
                Set MM_Dedu = MM_Dedu_cmd.Execute
            End If
        Next
        '更新処理が終われば、ページを移動する
        Response.Redirect(UpdateSuccess)
    End If
End If
%>

<%
' deductiontblより、担当者のymd最新レコードを取得(未登録、登録済　判定用)
Dim Rs_dedu_ck
Dim Rs_dedu_ck_cmd
Set Rs_dedu_ck_cmd = Server.CreateObject ("ADODB.Command")
Rs_dedu_ck_cmd.ActiveConnection = MM_workdbms_STRING
Rs_dedu_ck_cmd.CommandText = "SELECT * FROM deductiontbl WHERE ymb = ? AND personalcode = ?"
Rs_dedu_ck_cmd.Prepared = true
Rs_dedu_ck_cmd.Parameters.Append Rs_dedu_ck_cmd.CreateParameter("param1", 200, 1, -1, dispyear & dispmonth   )
Rs_dedu_ck_cmd.Parameters.Append Rs_dedu_ck_cmd.CreateParameter("param1", 200, 1, -1, Session("MM_Username") )
Set Rs_dedu_ck = Rs_dedu_ck_cmd.Execute
%>

<!DOCTYPE HTML>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--[if lt IE 9]><script src="dist/html5shiv.js"></script><![endif]-->
<title>勤務表管理システム</title>
<link href="css/default.css" rel="stylesheet" type="text/css">
<link href="css/superTables_compressed.css" rel="stylesheet" type="text/css">
</head>

<body>
<div id="container">
    <!-- #include file="inc/header.source" -->

    <div id="contents">
        <form name="deduction" method="post" action="">
            <div style="width:990px;">
                <br />
                <div style="float:left;">
                    支店控除入力&nbsp;
                    <a href="inputdeduction.asp?ymb=<%=lastYmb%>">&lt;&lt;</a>&nbsp;
                    <%=dispYear%>年<%=dispMonth%>月分&nbsp;
                    <a href="inputdeduction.asp?ymb=<%=nextYmb%>">&gt;&gt;</a>&nbsp;
                    <%
                    '表示職員の一覧取得SQLの発行
                    Set Rs_staff = Rs_staff_cmd.Execute
                    '対象職員が存在しない場合はボタンを表示しない
                    If Not Rs_staff.EOF Or Not Rs_staff.BOF Then
                        '担当者の最新ymbが表示月と一致していれば「登録済」
                        '一致しなかったり、レコードがなければ「未登録」を表示
                        button_disable = ""
                        If Not Rs_dedu_ck.EOF Or Not Rs_dedu_ck.BOF Then
                            If Rs_dedu_ck.Fields.Item("ymb") = (dispYear & dispMonth ) Then
                                Response.write "[登録済]&nbsp;"
                            Else
                                Response.write "<font class=""fontred"">[未登録]</font>&nbsp;"
                                button_disable = "disabled"
                            End If
                        Else
                            Response.write "<font class=""fontred"">[未登録]</font>&nbsp;"
                            button_disable = "disabled"
                        End If
                        %>
                        <input type="button" name="dedu_submit" id="dedu_submit" value="登録" onClick="deduSubmit()" <%=inputDisable%>>&nbsp;
                        <input type="button" name="button3" id="button3" value="エクセル" onClick="clickDownloadCSV()" <%=button_disable%>>&nbsp;
                        <font color="red"><b><%=errorMsg%></b></font>
                    <% End If%>
                </div>
                <div style="float:right;text-align:right;width:230px;">
                    <input type="button" name="clear1" id="clear1" value="他1消去" onclick="clearOther(1)">
                    <input type="button" name="clear2" id="clear2" value="他2消去" onclick="clearOther(2)">
                    <input type="button" name="clear2" id="clear3" value="他3消去" onclick="clearOther(3)">
                </div>
            </div>
            <div id="tablediv" class="clear">
                <div class="tHeader" style="width:1100px;">
                    <table class="data" >
                        <tr>
                            <th width="50px" rowspan="2" nowrap scope="col">個人CD</th>
                            <th width="140px" rowspan="2" nowrap scope="col">氏名</th>
                            <th width="75px" rowspan="2" nowrap scope="col">控除額合計</th>
                            <th colspan="7" nowrap scope="col">項目</th>
                            <th nowrap scope="col">その他１</th>
                            <th nowrap scope="col">その他２</th>
                            <th nowrap scope="col">その他３</th>
                        </tr>
                        <tr>
                            <th width="69px" nowrap>火災共済</th>
                            <th width="69px" nowrap>交通災害</th>
                            <th width="69px" nowrap>駐車場代</th>
                            <th width="69px" nowrap>住宅共益費</th>
                            <th width="69px" nowrap>水道代</th>
                            <th width="69px" nowrap>合格祝金</th>
                            <th width="69px" nowrap>支部費<br />(組合)</th>
                            <th width="69px" nowrap>
                                <input name="ta1_ncr"
                                         id="ta1_ncr"
                                      style="text-align:center;font-size: 9pt;width:58px;" type="text" maxlength="5"
                                      value=<%If Request.Form("i_deduction") = "i_deduction" Then%>
                                                "<%=Request.Form("ta1_ncr")%>"
                                            <%Else%>
                                                "<%If Not Rs_dedu_ck.EOF Or Not Rs_dedu_ck.BOF Then%><%=Trim(Rs_dedu_ck.Fields.Item("amount08ncr"))%><% End If%>"
                                            <%End If%>
                                >
                            </th>
                            <th width="69px" nowrap>
                                <input name="ta2_ncr"
                                         id="ta2_ncr"
                                      style="text-align:center;font-size: 9pt;width:58px;" type="text" maxlength="5"
                                      value=<%If Request.Form("i_deduction") = "i_deduction" Then%>
                                                "<%=Request.Form("ta2_ncr")%>"
                                            <%Else%>
                                                "<%If Not Rs_dedu_ck.EOF Or Not Rs_dedu_ck.BOF Then%><%=Trim(Rs_dedu_ck.Fields.Item("amount09ncr"))%><% End If%>"
                                            <%End If%>
                                >
                            </th>
                            <th width="69px" nowrap>
                                <input name="ta3_ncr"
                                         id="ta3_ncr"
                                      style="text-align:center;font-size: 9pt;width:58px;" type="text" maxlength="5"
                                      value=<%If Request.Form("i_deduction") = "i_deduction" Then%>
                                                "<%=Request.Form("ta3_ncr")%>"
                                            <%Else%>
                                                "<%If Not Rs_dedu_ck.EOF Or Not Rs_dedu_ck.BOF Then%><%=Trim(Rs_dedu_ck.Fields.Item("amount10ncr"))%><% End If%>"
                                            <%End If%>
                                >
                            </th>
                        </tr>
                    </table>
                </div>
                <% '担当者のymd最新レコードをクローズ
                Rs_dedu_ck.Close()
                Set Rs_dedu_ck = Nothing %>

                <% '表示職員の一覧取得SQLの発行
                Set Rs_staff = Rs_staff_cmd.Execute
                If Not Rs_staff.EOF Or Not Rs_staff.BOF Then %>
                    <div id="tbody"  class="tBody" style="width:1030px;height:100%;">
                        <table id="workdata" class="data" style="table-layout: fixed; ">
                            <% '表示職員の数だけ、データの更新を行う
                            Dim perCod
                            While (NOT Rs_staff.EOF)
                                perCod = Rs_staff.Fields.Item("personalcode")
                                ymb    = ""
                                'フォーム情報がない時(エラーではないとき)deductiontblの読み込みを行う
                                'If Request.Form("i_deduction") <> "i_deduction" Then
                                If (errorMsg = "") Then
                                    'deductiontblより、最新月の支店控除読み込み
                                    Dim Rs_dedu
                                    Dim Rs_dedu_cmd
                                    Set Rs_dedu_cmd = Server.CreateObject ("ADODB.Command")
                                    Rs_dedu_cmd.ActiveConnection = MM_workdbms_STRING
                                    Rs_dedu_cmd.CommandText = "SELECT * FROM deductiontbl WHERE ymb = (SELECT MAX(ymb) FROM deductiontbl WHERE personalcode = ? AND ymb <= ?) AND personalcode = ?"
                                    Rs_dedu_cmd.Prepared = true
                                    Rs_dedu_cmd.Parameters.Append Rs_dedu_cmd.CreateParameter("param1", 200, 1, -1, perCod )
                                    Rs_dedu_cmd.Parameters.Append Rs_dedu_cmd.CreateParameter("param1", 200, 1, -1, dispyear & dispmonth )
                                    Rs_dedu_cmd.Parameters.Append Rs_dedu_cmd.CreateParameter("param1", 200, 1, -1, perCod )
                                    Set Rs_dedu = Rs_dedu_cmd.Execute
                                    If Not Rs_dedu.EOF Or Not Rs_dedu.BOF Then
                                        ymb      = Rs_dedu.Fields.Item("ymb")
                                    End If
                                End If
                                
                                ' -------------------------------------------------
                                ' 表示項目設定
                                ' 入力チェックエラー時は、前回入力情報を表示
                                ' -------------------------------------------------
                                personalcode = Trim(Rs_staff.Fields.Item("personalcode").Value)
                                staffname    = Trim(Rs_staff.Fields.Item("staffname"   ).Value)
                                kasai        = 0
                                koutu        = 0
                                park         = 0
                                kyoeki       = 0
                                water        = 0
                                gokaku       = 0
                                kumiai       = 0
                                ta1_         = 0
                                ta2_         = 0
                                ta3_         = 0
                                If (errorMsg <> "") Then
                                    ' エラー有り
                                    kasai   = Request.Form("kasai"  )(i)
                                    koutu   = Request.Form("koutu"  )(i)
                                    park    = Request.Form("park"   )(i)
                                    kyoeki  = Request.Form("kyoeki" )(i)
                                    water   = Request.Form("water"  )(i)
                                    gokaku  = Request.Form("gokaku" )(i)
                                    kumiai  = Request.Form("kumiai" )(i)
                                    ta1_    = Request.Form("ta1_"   )(i)
                                    ta2_    = Request.Form("ta2_"   )(i)
                                    ta3_    = Request.Form("ta3_"   )(i)
                                Else
                                    ' エラー無し(初期表示時)
                                    If Not Rs_dedu.EOF Or Not Rs_dedu.BOF Then
                                        kasai   = Rs_dedu.Fields.Item("amount01").Value
                                        koutu   = Rs_dedu.Fields.Item("amount02").Value
                                        park    = Rs_dedu.Fields.Item("amount03").Value
                                        kyoeki  = Rs_dedu.Fields.Item("amount04").Value
                                        water   = Rs_dedu.Fields.Item("amount05").Value
                                        gokaku  = Rs_dedu.Fields.Item("amount06").Value
                                        kumiai  = Rs_dedu.Fields.Item("amount07").Value
                                        ta1_    = Rs_dedu.Fields.Item("amount08").Value
                                        ta2_    = Rs_dedu.Fields.Item("amount09").Value
                                        ta3_    = Rs_dedu.Fields.Item("amount10").Value
                                    End If
                                End If
                                If kasai    = 0 Then kasai  = "" End If
                                If koutu    = 0 Then koutu  = "" End If
                                If park     = 0 Then park   = "" End If
                                If kyoeki   = 0 Then kyoeki = "" End If
                                If water    = 0 Then water  = "" End If
                                If gokaku   = 0 Then gokaku = "" End If
                                If kumiai   = 0 Then kumiai = "" End If
                                If ta1_     = 0 Then ta1_   = "" End If
                                If ta2_     = 0 Then ta2_   = "" End If
                                If ta3_     = 0 Then ta3_   = "" End If
                            %>
                            <tr>
                                <th width="50px" nowrap class="permanent" scope="row" style="table-layout: fixed;">
                                    <%=perCod%>
                                    <input  name="personalcode"
                                              id="personalcode<%=perCod%>"
                                            type="hidden"
                                           value="<%=perCod%>"
                                    >
                                </th>
                                <th width="140px" class="permanent" style="table-layout: fixed;"><%=RTrim(Rs_staff.Fields.Item("staffname"))%></th>
                                <td width="75px" align="right" style="table-layout: fixed;"><div id="sum<%=perCod%>"></div></td>
                                <td width="69px" align="center">
                                    <%
                                    If UBound(style_kasai) < i Then
                                        style = ""
                                    Else
                                        style = style_kasai(i)
                                    End If
                                    %>
                                    <input  name="kasai"
                                              id="kasai<%=perCod%>"
                                           style="text-align:right;width:60px;" type="text" maxlength="7"
                                          onBlur="sum('<%=perCod%>', 'kasai')"
                                           class="<%=style%>"
                                           value="<%=kasai%>"
                                    >
                                </td>
                                <td width="69px" align="center">
                                    <%
                                    If UBound(style_koutu) < i Then
                                        style = ""
                                    Else
                                        style = style_koutu(i)
                                    End If
                                    %>
                                    <input  name="koutu"
                                              id="koutu<%=perCod%>"
                                           style="text-align:right;width:60px;" type="text" maxlength="7"
                                          onBlur="sum('<%=perCod%>', 'koutu')"
                                           class="<%=style%>"
                                           value="<%=koutu%>"
                                    >
                                </td>
                                <td width="69px" align="center">
                                    <%
                                    If UBound(style_park) < i Then
                                        style = ""
                                    Else
                                        style = style_park(i)
                                    End If
                                    %>
                                    <input  name="park"
                                              id="park<%=perCod%>"
                                           style="text-align:right;width:60px;" type="text" maxlength="7"
                                          onBlur="sum('<%=perCod%>', 'park')"
                                           class="<%=style%>"
                                           value="<%=park%>"
                                    >
                                </td>
                                <td width="69px" align="center">
                                    <%
                                    If UBound(style_kyoeki) < i Then
                                        style = ""
                                    Else
                                        style = style_kyoeki(i)
                                    End If
                                    %>
                                    <input  name="kyoeki"
                                              id="kyoeki<%=perCod%>"
                                           style="text-align:right;width:60px;" type="text" maxlength="7"
                                          onBlur="sum('<%=perCod%>', 'kyoeki')"
                                           class="<%=style%>"
                                           value="<%=kyoeki%>"
                                    >
                                </td>
                                <td width="69px" align="center">
                                    <%
                                    If UBound(style_water) < i Then
                                        style = ""
                                    Else
                                        style = style_water(i)
                                    End If
                                    %>
                                    <input  name="water"
                                              id="water<%=perCod%>"
                                           style="text-align:right;width:60px;" type="text" maxlength="7"
                                          onBlur="sum('<%=perCod%>', 'water')"
                                           class="<%=style%>"
                                           value="<%=water%>"
                                    >
                                </td>
                                <td width="69px" align="center">
                                    <%
                                    If UBound(style_gokaku) < i Then
                                        style = ""
                                    Else
                                        style = style_gokaku(i)
                                    End If
                                    %>
                                    <input  name="gokaku"
                                              id="gokaku<%=perCod%>"
                                           style="text-align:right;width:60px;" type="text" maxlength="7"
                                          onBlur="sum('<%=perCod%>', 'gokaku')"
                                           class="<%=style%>"
                                           value="<%=gokaku%>"
                                    >
                                </td>
                                <td width="69px" align="center">
                                    <%
                                    If UBound(style_kumiai) < i Then
                                        style = ""
                                    Else
                                        style = style_kumiai(i)
                                    End If
                                    %>
                                    <input  name="kumiai"
                                              id="kumiai<%=perCod%>"
                                           style="text-align:right;width:60px;" type="text" maxlength="7"
                                          onBlur="sum('<%=perCod%>', 'kumiai')"
                                           class="<%=style%>"
                                           value="<%=kumiai%>"
                                    >
                                </td>
                                <td width="69px" align="center">
                                    <%
                                    If UBound(style_ta1_) < i Then
                                        style = ""
                                    Else
                                        style = style_ta1_(i)
                                    End If
                                    %>
                                    <input  name="ta1_"
                                              id="ta1_<%=perCod%>"
                                           style="text-align:right;width:60px;" type="text" maxlength="7"
                                          onBlur="sum('<%=perCod%>', 'ta1_')"
                                           class="<%=style%>"
                                           value="<%=ta1_%>"
                                    >
                                </td>
                                <td width="69px" align="center">
                                    <%
                                    If UBound(style_ta2_) < i Then
                                        style = ""
                                    Else
                                        style = style_ta2_(i)
                                    End If
                                    %>
                                    <input  name="ta2_"
                                              id="ta2_<%=perCod%>"
                                           style="text-align:right;width:60px;" type="text" maxlength="7"
                                          onBlur="sum('<%=perCod%>', 'ta2_')"
                                           class="<%=style%>"
                                           value="<%=ta2_%>"
                                    >
                                </td>
                                    <%
                                    If UBound(style_ta3_) < i Then
                                        style = ""
                                    Else
                                        style = style_ta3_(i)
                                    End If
                                    %>
                                <td width="69px" align="center">
                                    <input  name="ta3_"
                                              id="ta3_<%=perCod%>"
                                           style="text-align:right;width:60px;" type="text" maxlength="7"
                                          onBlur="sum('<%=perCod%>', 'ta3_')"
                                           class="<%=style%>"
                                           value="<%=ta3_%>"
                                    >
                                    <input  name="ymb<%=perCod%>"
                                              id="ymb<%=perCod%>"
                                            type="hidden"
                                           value="<%=ymb%>"
                                    >
                                </td>
                            </tr>
                            <%
                            Rs_staff.MoveNext()
                            Wend
                            %>
                            <tr>
                                <th nowrap class="permanent" scope="row" style="table-layout: fixed;">-</th>
                                <th class="permanent" style="table-layout: fixed;">合計</th>
                                <td align="right"><div id="_sum"></div></td>
                                <td align="right"><div id="_kasai"></div></td>
                                <td align="right"><div id="_koutu"></div></td>
                                <td align="right"><div id="_park"></div></td>
                                <td align="right"><div id="_kyoeki"></div></td>
                                <td align="right"><div id="_water"></div></td>
                                <td align="right"><div id="_gokaku"></div></td>
                                <td align="right"><div id="_kumiai"></div></td>
                                <td align="right"><div id="_ta1_"></div></td>
                                <td align="right"><div id="_ta2_"></div></td>
                                <td align="right"><div id="_ta3_"></div></td>
                            </tr>
                        </table>
                    </div>
                <%End If%>
            </div>
            <input type="hidden" name="i_deduction" id="i_deduction" value="i_deduction">
        </form>
    </div>

    <!-- #include file="inc/footer.source" -->
</div>
</body>

<script type="text/javascript" >
    // 表示職員配列
    var staffList = new Array(
    <%
    Set Rs_staff = Rs_staff_cmd.Execute
    While (NOT Rs_staff.EOF)
        Response.write """" & Rs_staff.Fields.Item("personalcode") & ""","
        Rs_staff.MoveNext()
    Wend
    %>
    "-");
    staffList.pop();    // 配列最後が""のため削除を行う。
    
    // ウィンドウサイズから div サイズを設定する関数
    function setDivSize(){
        var size_h;
        size_h = document.body.clientHeight;
        size_h = size_h - 155;
        document.getElementById('tablediv').style.height = size_h + "px";
    }

    //読み込み時にサイズを表示
    setDivSize();

    function sum(code, column){
        sumLine(code);
        sumRow(column);
    }

    // 控除額列合計
    function sumRow(column){
        /* ********************************************************************
         * 列(項目)を合計行に集計
         * 引数：項目名
         * ********************************************************************/
        var i;
        sumColumn = 0;
        for (i=0; i<staffList.length-1; i++) {
            columnName = column + staffList[i];
            if (isNaN(document.getElementById(columnName).value)) {
            } else {
                sumColumn = sumColumn
                          + (1 * document.getElementById(columnName).value);
                document.getElementById(columnName).value =
                    Number(document.getElementById(columnName).value);
            }
            if (Number(document.getElementById(columnName).value) == "0") {
                document.getElementById(columnName).value = "";
            }
        }
        // 整数として表示
        document.getElementById("_"+column).innerHTML = addFigure(sumColumn.toFixed(0));
        
        // 合計行の合計欄
        sumAll = 0
        for (i=0; i<staffList.length-1; i++) {
            if (isNaN(document.getElementById("kasai"  + staffList[i]).value)) {
            } else {
                sumAll = sumAll + (1 * document.getElementById("kasai" + staffList[i]).value);
            }
            if (isNaN(document.getElementById("koutu"  + staffList[i]).value)) {
            } else {
                sumAll = sumAll + (1 * document.getElementById("koutu" + staffList[i]).value);
            }
            if (isNaN(document.getElementById("park"   + staffList[i]).value)) {
            } else {
                sumAll = sumAll + (1 * document.getElementById("park"  + staffList[i]).value);
            }
            if (isNaN(document.getElementById("kyoeki" + staffList[i]).value)) {
            } else {
                sumAll = sumAll + (1 * document.getElementById("kyoeki"+ staffList[i]).value);
            }
            if (isNaN(document.getElementById("water"  + staffList[i]).value)) {
            } else {
                sumAll = sumAll + (1 * document.getElementById("water" + staffList[i]).value);
            }
            if (isNaN(document.getElementById("gokaku" + staffList[i]).value)) {
            } else {
                sumAll = sumAll + (1 * document.getElementById("gokaku"+ staffList[i]).value);
            }
            if (isNaN(document.getElementById("kumiai" + staffList[i]).value)) {
            } else {
                sumAll = sumAll + (1 * document.getElementById("kumiai"+ staffList[i]).value);
            }
            if (isNaN(document.getElementById("ta1_"   + staffList[i]).value)) {
            } else {
                sumAll = sumAll + (1 * document.getElementById("ta1_"  + staffList[i]).value);
            }
            if (isNaN(document.getElementById("ta2_"   + staffList[i]).value)) {
            } else {
                sumAll = sumAll + (1 * document.getElementById("ta2_"  + staffList[i]).value);
            }
            if (isNaN(document.getElementById("ta3_"   + staffList[i]).value)) {
            } else {
                sumAll = sumAll + (1 * document.getElementById("ta3_"  + staffList[i]).value);
            }
        }
        document.getElementById("_sum").innerHTML = addFigure(sumAll.toFixed(0));
    }
    
    // 控除額各自合計
    function sumLine(code){
        //計算項目の配列を作成
        var deduList = new Array ("kasai", "koutu", "park", "kyoeki", "water", "gokaku", "kumiai", "ta1_", "ta2_", "ta3_");
        //項目内容の整数化
        var kasai   = document.getElementById("kasai"   + code).value;
        var koutu   = document.getElementById("koutu"   + code).value;
        var park    = document.getElementById("park"    + code).value;
        var kyoeki  = document.getElementById("kyoeki"  + code).value;
        var water   = document.getElementById("water"   + code).value;
        var gokaku  = document.getElementById("gokaku"  + code).value;
        var kumiai  = document.getElementById("kumiai"  + code).value;
        var ta1     = document.getElementById("ta1_"    + code).value;
        var ta2     = document.getElementById("ta2_"    + code).value;
        var ta3     = document.getElementById("ta3_"    + code).value;
        var sum;
        kasai   = (isNaN(kasai)     || !kasai   ) ? 0 : ~~(kasai    );
        koutu   = (isNaN(koutu)     || !koutu   ) ? 0 : ~~(koutu    );
        park    = (isNaN(park)      || !park    ) ? 0 : ~~(park     );
        kyoeki  = (isNaN(kyoeki)    || !kyoeki  ) ? 0 : ~~(kyoeki   );
        water   = (isNaN(water)     || !water   ) ? 0 : ~~(water    );
        gokaku  = (isNaN(gokaku)    || !gokaku  ) ? 0 : ~~(gokaku   );
        kumiai  = (isNaN(kumiai)    || !kumiai  ) ? 0 : ~~(kumiai   );
        ta1     = (isNaN(ta1)       || !ta1     ) ? 0 : ~~(ta1      );
        ta2     = (isNaN(ta2)       || !ta2     ) ? 0 : ~~(ta2      );
        ta3     = (isNaN(ta3)       || !ta3     ) ? 0 : ~~(ta3      );
        //合計の計算とカンマの配置
        sum = addFigure(kasai + koutu + park + kyoeki + water + gokaku + kumiai + ta1 + ta2 + ta3);
        //控除額合計へ動的書き込み
        document.getElementById("sum" + code).innerHTML = sum;

        //各項目の頭に0がある場合と、0のみが入力されていた場合は0の表示を消す
        for(n=0; n<deduList.length; n++) {
            if (isNaN(document.getElementById(deduList[n] + code).value)){
            } else {
                if ( Number(document.getElementById(deduList[n] + code).value) == "0" ){
                    document.getElementById(deduList[n] + code).value = "";
                } else {
                    document.getElementById(deduList[n] + code).value = Number(document.getElementById(deduList[n] + code).value);
                }
            }
        }
    }

    //数値にカンマを配置する
    function addFigure(str) {
        var num = new String(str).replace(/,/g, "");
        while(num != (num = num.replace(/^(-?\d+)(\d{3})/, "$1,$2")));
        return num;
    }

    //表示職員の配列を作成
    var staffList = new Array(<%
    Set Rs_staff = Rs_staff_cmd.Execute
    If Not Rs_staff.EOF Or Not Rs_staff.BOF Then
        While (NOT Rs_staff.EOF)
            Response.write """" & Rs_staff.Fields.Item("personalcode") & ""","
            Rs_staff.MoveNext()
        Wend
    End If
    %>"");

    //表示職員全ての合計計算
    function sumlist(){
        for(i=0; i<staffList.length-1; i++) {
            sumLine(staffList[i]);
        }
        var deduList = new Array ("kasai", "koutu", "park", "kyoeki", "water", "gokaku", "kumiai", "ta1_", "ta2_", "ta3_");
        for(n=0; n<deduList.length; n++) {
            sumRow(deduList[n]);
        }
    }

    //ページ読み込み時にsumlistを実行する
    sumlist();

    //登録確認メッセージ
    function deduSubmit(){
        ans=confirm("支店控除情報を登録いたします。\nよろしいですか？");
        if(ans==true){
            document.deduction.submit();
        }
    }

    //CSVデータダウンロードボタン押下時の処理
    function clickDownloadCSV(){
        ans=confirm("データをダウンロードします。\n入力途中の内容は登録されません。\nよろしいですか？");
        if(ans) {
            location.href="csv_inputdeduction.asp?ymb=<%=dispyear & dispmonth%>";
        }
    }

    // その他消去ボタン押下時の処理
    function clearOther(otherNumber) {
        var strName = ""
        switch (otherNumber) {
            case 1:
                strName = "ta1_"
                break;
            case 2:
                strName = "ta2_"
                break;
            case 3:
                strName = "ta3_"
                break;
            default:
        }
        document.getElementsByName(strName + "ncr")[0].value = "";
        //var es = document.getElementsByName("ta1");
        var elements = document.getElementsByTagName("input");
        for (i=0; i<elements.length; i++) {
            if (elements[i].name.substring(0, 4) == strName) {
                var code = elements[i].name.substring(4, 9);
                if (code != "ncr") {
                    elements[i].value = "0";
                }
            }
        }
        for (i=0; i<staffList.length-1; i++) {
            sumLine(staffList[i]);
        }
        sumRow(strName);
    }

</script>

<%
' -----------------------------------------------------------------------------
' 入力チェックと値の設定を行う。
' -----------------------------------------------------------------------------
Sub setData()
    personalcode    = Trim(Request.Form("personalcode" )(i))
    kasai           = Trim(Request.Form("kasai"        )(i))
    koutu           = Trim(Request.Form("koutu"        )(i))
    park            = Trim(Request.Form("park"         )(i))
    kyoeki          = Trim(Request.Form("kyoeki"       )(i))
    water           = Trim(Request.Form("water"        )(i))
    gokaku          = Trim(Request.Form("gokaku"       )(i))
    kumiai          = Trim(Request.Form("kumiai"       )(i))
    ta1_            = Trim(Request.Form("ta1_"         )(i))
    ta2_            = Trim(Request.Form("ta2_"         )(i))
    ta3_            = Trim(Request.Form("ta3_"         )(i))
    
    ReDim Preserve style_kasai(i)
    If (Len(kasai) = 0) Then
        kasai = 0
    Else
        If (Not IsNumeric(kasai)) Then
            errorMsg        = strErrorMsg
            style_kasai(i)  = "errorcolor"
        ElseIf (kasai < 0) Then
            errorMsg        = strErrorMsg
            style_kasai(i)  = "errorcolor"
        End If
        ' 小数点はエラー
        If (InStr(kasai, ".") <> 0) Then
            errorMsg        = strErrorMsg
            style_kasai(i)  = "errorcolor"
        End If
    End If

    ReDim Preserve style_koutu(i)
    If (Len(koutu) = 0) Then
        koutu = 0
    Else
        If (Not IsNumeric(koutu)) Then
            errorMsg        = strErrorMsg
            style_koutu(i)  = "errorcolor"
        ElseIf (koutu < 0) Then
            errorMsg        = strErrorMsg
            style_koutu(i)  = "errorcolor"
        End If
        ' 小数点はエラー
        If (InStr(koutu, ".") <> 0) Then
            errorMsg        = strErrorMsg
            style_koutu(i)  = "errorcolor"
        End If
    End If

    ReDim Preserve style_park(i)
    If (Len(park) = 0) Then
        park = 0
    Else
        If (Not IsNumeric(park)) Then
            errorMsg        = strErrorMsg
            style_park(i)   = "errorcolor"
        ElseIf (park < 0) Then
            errorMsg        = strErrorMsg
            style_park(i)   = "errorcolor"
        End If
        ' 小数点はエラー
        If (InStr(park, ".") <> 0) Then
            errorMsg        = strErrorMsg
            style_park(i)   = "errorcolor"
        End If
    End If

    ReDim Preserve style_kyoeki(i)
    If (Len(kyoeki) = 0) Then
        kyoeki = 0
    Else
        If (Not IsNumeric(kyoeki)) Then
            errorMsg        = strErrorMsg
            style_kyoeki(i) = "errorcolor"
        ElseIf (kyoeki < 0) Then
            errorMsg        = strErrorMsg
            style_kyoeki(i) = "errorcolor"
        End If
        ' 小数点はエラー
        If (InStr(kyoeki, ".") <> 0) Then
            errorMsg        = strErrorMsg
            style_kyoeki(i) = "errorcolor"
        End If
    End If

    ReDim Preserve style_water(i)
    If (Len(water) = 0) Then
        water = 0
    Else
        If (Not IsNumeric(water)) Then
            errorMsg        = strErrorMsg
            style_water(i)  = "errorcolor"
        ElseIf (water < 0) Then
            errorMsg        = strErrorMsg
            style_water(i)  = "errorcolor"
        End If
        ' 小数点はエラー
        If (InStr(water, ".") <> 0) Then
            errorMsg        = strErrorMsg
            style_water(i)  = "errorcolor"
        End If
    End If

    ReDim Preserve style_gokaku(i)
    If (Len(gokaku) = 0) Then
        gokaku = 0
    Else
        If (Not IsNumeric(gokaku)) Then
            errorMsg        = strErrorMsg
            style_gokaku(i) = "errorcolor"
        ElseIf (gokaku < 0) Then
            errorMsg        = strErrorMsg
            style_gokaku(i) = "errorcolor"
        End If
        ' 小数点はエラー
        If (InStr(gokaku, ".") <> 0) Then
            errorMsg        = strErrorMsg
            style_gokaku(i) = "errorcolor"
        End If
    End If

    ReDim Preserve style_kumiai(i)
    If (Len(kumiai) = 0) Then
        kumiai = 0
    Else
        If (Not IsNumeric(kumiai)) Then
            errorMsg        = strErrorMsg
            style_kumiai(i) = "errorcolor"
        ElseIf (kumiai < 0) Then
            errorMsg        = strErrorMsg
            style_kumiai(i) = "errorcolor"
        End If
        ' 小数点はエラー
        If (InStr(kumiai, ".") <> 0) Then
            errorMsg        = strErrorMsg
            style_kumiai(i) = "errorcolor"
        End If
    End If

    ReDim Preserve style_ta1_(i)
    If (Len(ta1_) = 0) Then
        ta1_ = 0
    Else
        If (Not IsNumeric(ta1_)) Then
            errorMsg        = strErrorMsg
            style_ta1_(i)   = "errorcolor"
        ElseIf (ta1_ < 0) Then
            errorMsg        = strErrorMsg
            style_ta1_(i)   = "errorcolor"
        End If
        ' 小数点はエラー
        If (InStr(ta1_, ".") <> 0) Then
            errorMsg        = strErrorMsg
            style_ta1_(i)   = "errorcolor"
        End If
    End If

    ReDim Preserve style_ta2_(i)
    If (Len(ta2_) = 0) Then
        ta2_ = 0
    Else
        If (Not IsNumeric(ta2_)) Then
            errorMsg        = strErrorMsg
            style_ta2_(i)   = "errorcolor"
        ElseIf (ta2_ < 0) Then
            errorMsg        = strErrorMsg
            style_ta2_(i)   = "errorcolor"
        End If
        ' 小数点はエラー
        If (InStr(ta2_, ".") <> 0) Then
            errorMsg        = strErrorMsg
            style_ta2_(i)   = "errorcolor"
        End If
    End If
    
    ReDim Preserve style_ta3_(i)
    If (Len(ta3_) = 0) Then
        ta3_ = 0
    Else
        If (Not IsNumeric(ta3_)) Then
            errorMsg        = strErrorMsg
            style_ta3_(i)   = "errorcolor"
        ElseIf (ta3_ < 0) Then
            errorMsg        = strErrorMsg
            style_ta3_(i)   = "errorcolor"
        End If
        ' 小数点はエラー
        If (InStr(ta3_, ".") <> 0) Then
            errorMsg        = strErrorMsg
            style_ta3_(i)   = "errorcolor"
        End If
    End If
End Sub

'表示職員一覧を取得するSQLをクローズ
Rs_staff.Close()
Set Rs_staff = Nothing
%>
</html>
