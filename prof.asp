<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="Connections/workdbms.asp" -->
<!-- #include file="inc/RestrictAccess.asp" -->
<!DOCTYPE HTML>
<html lang="ja">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>勤務表管理システム</title>
<link href="css/default.css" rel="stylesheet" type="text/css">
</head>

<body>
<div id="container">
<!-- #include file="inc/header.source" -->

<div id="contents">
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
    <p>▶<a href="changepassword.asp">パスワードの変更</a></p>
    <p>▶<a href="changeworktime.asp">労働時間設定</a></p>
    <p>&nbsp;</p>
</div>
<!-- #include file="inc/footer.source" -->
</div>
</body>
</html>
