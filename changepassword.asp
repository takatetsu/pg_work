<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' #######################################
' プログラム仕様書
' #######################################
'
' ## プログラム概要 ##
' 個人パスワードの変更を行う
'
' ## 出力項目 ##
'
' ## 入力チェック ##
' 二か所のパスワード入力エリアに違う文字列が入力された場合はエラー表示
' パスワード入力文字数が2文字未満の場合はエラー表示
'
' ## 注意事項 ##
'
'
'
%>
<!-- #include file="Connections/workdbms.asp" -->
<!-- #include file="inc/RestrictAccess.asp" -->
<%

Dim messege
Dim e_cod
e_cod = 0
' フォームの入力内容のチェック
If Request.Form("c_password") = "c_password" Then
   If Request.Form("password") <> Request.Form("password2") Then
      messege = "新しいパスワードと、パスワード確認用が一致していません"
	  e_cod = 1
   End If
   If Len(Request.Form("password")) < 2 Then
      messege = "パスワードには2文字以上を入力して下さい"
	  e_cod = 1
   End If

   If e_cod = 1 Then
      e_color = "errorcolor"
   Else

    Dim MM_updateSuccess
    Dim MM_updateSQL
    Dim MM_cPass_cmd
    Dim MM_cPass
    MM_updateSuccess = "index.asp"

    MM_updateSQL = MM_updateSQL & " UPDATE stafftbl SET stafftbl.password = digest(?, 'sha1') WHERE stafftbl.personalcode = ? "
    Set MM_cPass_cmd = Server.CreateObject ("ADODB.Command")
    MM_cPass_cmd.ActiveConnection = MM_workdbms_STRING
    MM_cPass_cmd.CommandText = MM_updateSQL
    MM_cPass_cmd.Prepared = true
    MM_cPass_cmd.Parameters.Append MM_cPass_cmd.CreateParameter("param1", 200, 1, -1, Request.Form("password"))
    MM_cPass_cmd.Parameters.Append MM_cPass_cmd.CreateParameter("param2", 200, 1, 5, Session("MM_Username"))

    Set MM_cPass = MM_cPass_cmd.Execute
    Response.Redirect(MM_updateSuccess)
   End If
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
    <form name="c_password" method="post" action="">
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
        <p>パスワードを変更します。<br>
            新しいパスワードと確認用と同じものを入力してください。</p>
        <p>パスワードを変更したら、一度ログアウトされます。<br>
            続けてシステムを使用する場合、新しいパスワードでログインし直してからお使いください。</p>
        <table>
            <tr>
                <td align="right" nowrap>新しいパスワード：</td>
                <td><input type="password" name="password" id="password" class="<%=e_color%>"></td>
                <td class="fontred"><%=messege%></td>
            </tr>
            <tr>
                <td align="right" nowrap>パスワード確認用：</td>
                <td><input type="password" name="password2" id="password2" class="<%=e_color%>"></td>
                <td>&nbsp;</td>
            </tr>
            <tr>
                <td>&nbsp;</td>
                <td><input type="submit" name="c_password_submit" id="c_password_submit" value="変更"></td>
                <td>&nbsp;</td>
            </tr>
    </table>
<input type="hidden" name="c_password" id="c_password" value="c_password">
</form>
        <p>&nbsp;</p>

</div>

<!-- #include file="inc/footer.source" -->
</div>
</body>
</html>
