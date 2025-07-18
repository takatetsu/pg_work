<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' #######################################
' プログラム仕様書
' #######################################
'
' ## プログラム概要 ##
' 個人労働時間の変更を行う
'
' ## 出力項目 ##
'
' ## 入力チェック ##
' 
' 
'
' ## 注意事項 ##
'
'
'
%>
<!-- #include file="Connections/workdbms.asp" -->
<!-- #include file="inc/RestrictAccess.asp" -->
<!-- #include file="inc/util.asp" -->
<%
Dim messege
Dim e_cod1
Dim e_cod2
Dim e_color1
Dim e_color2
e_cod1 = 0
e_cod2 = 0
e_color1 = ""
e_color2 = ""
msg = ""
' フォームの入力内容のチェック
If Request.Form("MM_update") = "form1" Then
    If Len(Request.Form("base_am_workmin")) = 0 Or _
       Not legalTime(Request.Form("base_am_workmin")) Then
        e_color1 = "errorcolor"
    End If
    If Len(Request.Form("base_pm_workmin")) = 0 Or _
       Not legalTime(Request.Form("base_pm_workmin")) Then
        e_color2 = "errorcolor"
    End If
    If e_color1 = "" And e_color2 = "" Then
        ' 入力正常、労働時間を更新
        Dim MM_updateSuccess
        Dim MM_updateSQL
        Dim MM_cPass_cmd
        Dim MM_cPass
        MM_updateSuccess = "prof.asp"
        MM_updateSQL = MM_updateSQL & " UPDATE stafftbl SET base_am_workmin = ?, base_pm_workmin = ? WHERE personalcode = ? "
        Set MM_cPass_cmd = Server.CreateObject ("ADODB.Command")
        MM_cPass_cmd.ActiveConnection = MM_workdbms_STRING
        MM_cPass_cmd.CommandText = MM_updateSQL
        MM_cPass_cmd.Prepared = true
        MM_cPass_cmd.Parameters.Append MM_cPass_cmd.CreateParameter(,3,,, time2Min(Request.Form("base_am_workmin")))
        MM_cPass_cmd.Parameters.Append MM_cPass_cmd.CreateParameter(,3,,, time2Min(Request.Form("base_pm_workmin")))
        MM_cPass_cmd.Parameters.Append MM_cPass_cmd.CreateParameter(,129,,5, Session("MM_Username"))
        Set MM_cPass = MM_cPass_cmd.Execute
        Response.Redirect(MM_updateSuccess)
    Else
        ' 入力エラーの場合の処理
        msg = "入力された時間が不正です"
    End If
Else
    ' stafftbl を読み、労働時間を取得
    Dim Rs_stafftbl
    Dim Rs_stafftbl_cmd
    Dim Rs_stafftbl_numRows
    Set Rs_stafftbl_cmd = Server.CreateObject ("ADODB.Command")
    Rs_stafftbl_cmd.ActiveConnection = MM_workdbms_STRING
    Rs_stafftbl_cmd.CommandText = "SELECT * FROM stafftbl WHERE personalcode = ?"
    Rs_stafftbl_cmd.Prepared = true
    Rs_stafftbl_cmd.Parameters.Append Rs_stafftbl_cmd.CreateParameter(,129,,5, Session("MM_Username"))
    Set Rs_stafftbl = Rs_stafftbl_cmd.Execute
    Rs_stafftbl_numRows = 0
    base_am_workmin = min2Time(Trim(Rs_stafftbl.Fields.Item("base_am_workmin").Value))
    base_pm_workmin = min2Time(Trim(Rs_stafftbl.Fields.Item("base_pm_workmin").Value))
    Rs_stafftbl.Close()
    Set Rs_stafftbl = Nothing
End If

%>

<!DOCTYPE HTML>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>勤務表管理システム</title>
<link href="css/default.css" rel="stylesheet" type="text/css">
<script language="vbscript">


</script>
</head>

<body>
    <div id="container">
        <!-- #include file="inc/header.source" -->

        <div id="contents">
            <form name="form1" method="post" action="">
                <br />
                <table class="data">
                    <tr>
                        <th width="65">個人CD</th>
                        <td width="65px" class="disabled" align="center"><%=Session("MM_Username")%></td>
                        <th width="65">氏名</th>
                        <td width="269px" class="disabled"><%=Session("MM_staffname")%></td>
                        <th width="65px">所属</th>
                        <td width="289px" class="disabled"><%=Session("MM_orgname")%></td>
                    </tr>
                </table>
                <br />
                <p>労働時間を変更します。<br>
                午前と午後の労働時間をHH:MMの形式で入力してください。</p>
                <table>
                    <tr>
                        <td align="right" nowrap>午前労働時間：</td>
                        <td>
                            <input type="text" name="base_am_workmin" 
                                id="base_am_workmin" class="<%=e_color%>"
                                value="<%=base_am_workmin%>">
                        </td>
                    </tr>
                    <tr>
                        <td align="right" nowrap>午後労働時間：</td>
                        <td>
                            <input type="text" name="base_pm_workmin" 
                                id="base_pm_workmin" class="<%=e_color%>"
                                value="<%=base_pm_workmin%>">
                        </td>
                    </tr>
                    <tr>
                        <td>&nbsp;</td>
                        <td>
                            <input type="hidden" name="MM_update" value="form1">
                            <input type="submit" name="submit" id="submit" value="設定">
                        </td>
                    </tr>
                </table>
                <p><%=msg%></p>
            </form>
            <p>※ご自身の雇入れ通知書をご確認ください。<br>
            入力例： 始業8:30　～　終業17:10　の場合<br>
            ＡＭ：　08:30～12:00　⇒　3時間30分　⇒　3:30<br>
            ＰＭ：　13:00～17:10　⇒　4時間10分　⇒　4:10</p>
            <p>なお、ここで入力いただいた労働時間を基に、個人勤務表の一日あたりの労働時間を表記しますので、<br>
            ご入力に誤りがないようご注意ください。</p>
        </div>
        <!-- #include file="inc/footer.source" -->
    </div>
</body>
</html>
