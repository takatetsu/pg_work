<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
Session.Contents.RemoveAll
MM_logoutRedirectPage = "index.asp"
' redirect with URL parameters (remove the "MM_Logoutnow" query param).
If (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
  MM_newQS = "?"
  For Each Item In Request.QueryString
    If (Item <> "MM_Logoutnow") Then
      If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
      MM_newQS = MM_newQS & Item & "=" & Server.URLencode(Request.QueryString(Item))
    End If
  Next
  If (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
End If
Response.Redirect(MM_logoutRedirectPage)
End If
%>

<%
' *** Restrict Access To Page: Grant or deny access to this page
' ログイン認証していないユーザのアクセスを制限する。
MM_authorizedUsers=""
MM_authFailedURL="index.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If

' タイムカードデータ取込み中はアクセスを制限する。
MM_SystemFailedURL = "sorry.html"
Dim Rs_controltbl
Dim Rs_controltbl_cmd
Set Rs_controltbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_controltbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_controltbl_cmd.CommandText = "SELECT COUNT(*) AS count FROM dbo.controltbl WHERE systemenable='1'"
Rs_controltbl_cmd.Prepared = true
Set Rs_controltbl = Rs_controltbl_cmd.Execute
If Rs_controltbl.EOF Or Rs_controltbl.Fields.Item("count").Value <= 0 Then
  Rs_controltbl.Close()
  Set Rs_controltbl = Nothing
  Response.Redirect(MM_SystemFailedURL)
End If
Rs_controltbl.Close()
Set Rs_controltbl = Nothing
%>