<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' #######################################
' プログラム仕様書
' #######################################
'
' ## プログラム概要 ##
' データの登録・更新が完了したことを表示するためのページです。
'
' ## 出力項目 ##
' ありません。
'
' ## 入力チェック ##
' ありません。
'
' ## 注意事項 ##
' ありません。
'
'
%>
<!-- #include file="Connections/workdbms.asp" -->
<!-- #include file="inc/RestrictAccess.asp" -->
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

<div id="contents" style="height:600px;">
<p>&nbsp;</p>
<p align="center" style="font-size:14pt;">データは更新されました。</p>
<p align="center"><input type="button" value="元のページに戻る" onclick="goback()" /></p>
</div>

<!-- #include file="inc/footer.source" -->
</div>
</body>
<script type="text/javascript">
function goback() {
    document.location = '<%=Request.ServerVariables("HTTP_REFERER")%>';
}
</script>
</html>
