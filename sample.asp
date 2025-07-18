<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/workdbms.asp" -->
<%
Dim Rs_staff
Dim Rs_staff_cmd
Dim Rs_staff_numRows

Set Rs_staff_cmd = Server.CreateObject ("ADODB.Command")
Rs_staff_cmd.ActiveConnection = MM_workdbms_STRING
Rs_staff_cmd.CommandText = "SELECT * FROM dbo.stafftbl ORDER BY id ASC" 
Rs_staff_cmd.Prepared = true

Set Rs_staff = Rs_staff_cmd.Execute
Rs_staff_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
Rs_staff_numRows = Rs_staff_numRows + Repeat1__numRows
%>
<!DOCTYPE HTML>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Sample</title>
</head>

<body>
<p>Sample</p>
<p>&nbsp;</p>
<table border="1" cellpadding="3">
    <tr>
        <td>id</td>
        <td>updatetime</td>
        <td>personalcode</td>
        <td>staffname</td>
        <td>orgcode</td>
        <td>is_input</td>
        <td>is_charge</td>
        <td>is_superior</td>
        <td>is_enable</td>
        <td>password</td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT Rs_staff.EOF)) %>
        <tr>
            <td><%=(Rs_staff.Fields.Item("id").Value)%></td>
            <td><%=(Rs_staff.Fields.Item("updatetime").Value)%></td>
            <td><%=(Rs_staff.Fields.Item("personalcode").Value)%></td>
            <td><%=(Rs_staff.Fields.Item("staffname").Value)%></td>
            <td><%=(Rs_staff.Fields.Item("orgcode").Value)%></td>
            <td><%=(Rs_staff.Fields.Item("is_input").Value)%></td>
            <td><%=(Rs_staff.Fields.Item("is_charge").Value)%></td>
            <td><%=(Rs_staff.Fields.Item("is_superior").Value)%></td>
            <td><%=(Rs_staff.Fields.Item("is_enable").Value)%></td>
            <td><%=(Rs_staff.Fields.Item("password").Value)%></td>
        </tr>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Rs_staff.MoveNext()
Wend
%>
</table>
</body>
</html>
<%
Rs_staff.Close()
Set Rs_staff = Nothing
%>
