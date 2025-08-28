<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' #######################################
' プログラム仕様書
' #######################################
'
' ## プログラム概要 ##
' ユーザ認証を行い、認証できたときセッションにユーザ情報を書き込む。
' また、ユーザのフラグによって遷移先画面を切替える。
'
' ## 出力項目 ##
'
' ## 入力チェック ##
' 入力した個人コードとパスワードをそれぞれ stafftbl(社員テーブル)の personalcode, password と比較し、認証処理を行う。
'
' ## 注意事項 ##
' 認証時パスワードは暗号化した結果で検証する必要がある。
' stafftbl(社員テーブル)のデータには無効(is_enable が '0')になっているものが存在する。無効データで認証は不可とする。
'
%>
<!--#include file="Connections/workdbms.asp" -->
<%
' *** Validate request to log in to this site.
Session.Contents.RemoveAll
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("txt_id"))
If MM_valUsername <> "" Then
    Dim MM_fldUserAuthorization
    Dim MM_redirectLoginSuccess
    Dim MM_redirectLoginFailed
    Dim MM_loginSQL
    Dim MM_rsUser
    Dim MM_rsUser_cmd

    MM_fldUserAuthorization = ""
    MM_redirectLoginSuccess = "inputwork.asp"
    MM_redirectLoginFailed = "index.asp"

    MM_loginSQL = "SELECT personalcode, staffname, stafftbl.orgcode, is_operator, is_input, " & _
                    "is_deduction, is_charge, is_superior, orgname, password, opentime, " & _
                    "closetime, is_unionexecutive, workshift, old_workshift, old_workshift_last_ymb"
    If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
    MM_loginSQL = MM_loginSQL & " FROM stafftbl LEFT JOIN orgnametbl ON " & _
                                "stafftbl.orgcode=orgnametbl.orgcode " & _
                                "WHERE stafftbl.personalcode = ? AND " & _
                                "stafftbl.password = encode(digest(?, 'sha1'), 'hex') AND is_enable = '1'"
    Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
    MM_rsUser_cmd.ActiveConnection = MM_workdbms_STRING
    MM_rsUser_cmd.CommandText = MM_loginSQL
    MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, -1, MM_valUsername)
    MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 42, Request.Form("txt_pass"))
    MM_rsUser_cmd.Prepared = true
    Set MM_rsUser = MM_rsUser_cmd.Execute

    If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then
        ' username and password match - this is a valid user
        Session("MM_Username")    = MM_valUsername                                '社員コード
        Session("MM_staffname")   = Trim(MM_rsUser.Fields.Item("staffname").Value)'社員名
        Session("MM_orgcode")     = Trim(MM_rsUser.Fields.Item("orgcode").Value)  '組織コード
        Session("MM_is_input")    = MM_rsUser.Fields.Item("is_input").Value       '勤務表の入力
        Session("MM_is_deduction")= MM_rsUser.Fields.Item("is_deduction").Value   '支店控除入力
        Session("MM_is_charge")   = MM_rsUser.Fields.Item("is_charge").Value      '全体入力
        Session("MM_is_superior") = MM_rsUser.Fields.Item("is_superior").Value    '上長確認入力
        Session("MM_orgname")     = Trim(MM_rsUser.Fields.Item("orgname").Value)  '組織名
        Session("MM_opentime")    = "08:30"                                       '始業時刻
        If (Not(IsNULL(MM_rsUser.Fields.Item("opentime").Value)) And _
            Trim(MM_rsUser.Fields.Item("opentime").Value) <> "") Then
            Session("MM_opentime")  = Left (Trim(MM_rsUser.Fields.Item("opentime" ).Value),2) & ":" & _
                                      Right(Trim(MM_rsUser.Fields.Item("opentime" ).Value),2)
        End If
        Session("MM_closetime") = "17:10"                                         '終業時刻
        If (Not(IsNULL(MM_rsUser.Fields.Item("closetime").Value)) And _
            Trim(MM_rsUser.Fields.Item("closetime").Value) <> "") Then
            Session("MM_closetime") = Left (Trim(MM_rsUser.Fields.Item("closetime").Value),2) & ":" & _
                                      Right(Trim(MM_rsUser.Fields.Item("closetime").Value),2)
        End If
        Session("MM_is_unionexecutive") = MM_rsUser.Fields.Item("is_unionexecutive").Value  '組合執行部フラグ
        
        If  MM_rsUser.Fields.Item("is_operator").Value = "1" Then
            Session("MM_workshift")     = "8"   '勤務体系
        Else
            Session("MM_workshift")     = MM_rsUser.Fields.Item("workshift").Value  '勤務体系
        End If
        
'        Session("MM_old_workshift")          = MM_rsUser.Fields.Item("old_workshift").Value          '旧勤務体系
'        Session("MM_old_workshift_last_ymb") = MM_rsUser.Fields.Item("old_workshift_last_ymb").Value '旧勤務体系
        
        ' 入力機能利用権限によって遷移先画面を切替える
        If (MM_rsUser.Fields.Item("is_input").Value = "1") Then
            MM_redirectLoginSuccess = "inputwork.asp"
        ElseIf (MM_rsUser.Fields.Item("is_superior" ).Value = "1") Then
            MM_redirectLoginSuccess = "checklist.asp"
        ElseIf (MM_rsUser.Fields.Item("is_charge"   ).Value = "1") Then
            MM_redirectLoginSuccess = "inputall.asp"
        ElseIf (MM_rsUser.Fields.Item("is_deduction").Value = "1") Then
            MM_redirectLoginSuccess = "inputdeduction.asp"
        End If

        If (MM_fldUserAuthorization <> "") Then
            Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
        Else
            Session("MM_UserAuthorization") = ""
        End If
        If CStr(Request.QueryString("accessdenied")) <> "" And false Then
            MM_redirectLoginSuccess = Request.QueryString("accessdenied")
        End If
        MM_rsUser.Close
        Response.Redirect(MM_redirectLoginSuccess)
    End If
    MM_rsUser.Close
    errorMessage = "個人コード、もしくはパスワードが間違っています。"
    'Response.Redirect(MM_redirectLoginFailed)
End If
%>
<!DOCTYPE HTML>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>勤務表管理システム</title>
<link href="css/default.css" rel="stylesheet" type="text/css">
</head>

<body onLoad="document.form1.txt_id.focus()">
<div style="padding:30px 50px;">
<h1>勤務表管理システム</h1>
<form name="form1" method="POST" action="<%=MM_LoginAction%>">
    <span style="display:inline-block;font-size:12pt;width:600px;">
        <p>あなたが参照しようとしているページはアクセス制限されており認証が必要です。<br /></p>
    </span>
    <table class="data" style="width: 400px">
        <tr>
            <th align="center" style="width: 150px;">個人コード</th>
            <td style="width: 100px;"><input name="txt_id" type="text" id="txt_id" autocomplete="off" style="width: 250px;ime-mode:disabled;" maxlength="10"/></td>
        </tr>
        <tr>
            <th align="center" style="width: 150px;">パスワード</th>
            <td style="width: 100px;"><input name="txt_pass" type="password" id="txt_pass" autocomplete="off" style="width: 250px" maxlength="10"/></td>
        </tr>

    </table>
    <table style="width: 400px; height: 50px;">
        <tr>
            <td style="width: 200px;" align="center"><input type="submit" value="ログイン" style="width: 100px" /></td>
            <td style="width: 200px;" align="center"><input type="button" value="キャンセル" style="width: 100px" /></td>
        </tr>
    </table>
    <p class="fontred"><%=errorMessage%>&nbsp;</p>
</form>
<p>&nbsp;</p>
</div>
</body>
</html>
