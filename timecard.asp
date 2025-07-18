<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="Connections/workdbms.asp" -->
<!-- #include file="inc/RestrictAccess.asp" -->
<%
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

' 個人コードがurlパラメタに存在するとき、個人コードをパラメタから設定し、
' urlパラメタに存在しないとき、セションから設定。
If (Request.QueryString("p")<>"" And (Session("MM_is_superior")="1" Or Session("MM_is_charge"   )="1")) Then
    ' 上長チェック画面の設定
    ' -----------------------------------------------------------------------------
    ' 組織、社員テーブル orgtbl,stafftbl 読込
    ' -----------------------------------------------------------------------------
    Dim Rs_managetbl
    Dim Rs_managetbl_cmd
    Dim Rs_managetbl_numRows

    Set Rs_managetbl_cmd = Server.CreateObject ("ADODB.Command")
    Rs_managetbl_cmd.ActiveConnection = MM_workdbms_STRING
    Rs_managetbl_cmd.CommandText = "SELECT dbo.orgtbl.manageclass FROM dbo.orgtbl " & _
        "INNER JOIN dbo.stafftbl ON dbo.orgtbl.orgcode=dbo.stafftbl.orgcode "       & _
        "WHERE dbo.orgtbl.personalcode=? AND dbo.stafftbl.personalcode=? AND "      & _
        "dbo.stafftbl.is_enable='1'"
    Rs_managetbl_cmd.Prepared = true
    Rs_managetbl_cmd.Parameters.Append Rs_managetbl_cmd.CreateParameter(_
        "param1", 200, 1, 5, Session("MM_Username"))
    Rs_managetbl_cmd.Parameters.Append Rs_managetbl_cmd.CreateParameter(_
        "param2", 200, 1, 5, Request.QueryString("p"))
    Set Rs_managetbl = Rs_managetbl_cmd.Execute
    Rs_managetbl_numRows = 0

    If Rs_managetbl.EOF And Rs_managetbl.BOF Then
        ' 上長チェック対象の職員がいないとき
        Rs_managetbl.Close()
        Set Rs_managetbl = Nothing
        personalcode = Session("MM_Username")
        checkUser()
    Else
        ' 上長チェック対象の職員がいるとき
        personalcode = Request.QueryString("p")
        manageclass  = Trim(Rs_managetbl.Fields.Item("manageclass").Value)
        Rs_managetbl.Close()
        Set Rs_managetbl = Nothing
    End If

Else
    personalcode = Session("MM_Username")
End If

' -----------------------------------------------------------------------------
' 社員テーブル stafftbl 読込
' -----------------------------------------------------------------------------
Dim Rs_stafftbl
Dim Rs_stafftbl_cmd
Dim Rs_stafftbl_numRows
Set Rs_stafftbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_stafftbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_stafftbl_cmd.CommandText = "SELECT personalcode, staffname, " & _
    "dbo.stafftbl.orgcode AS orgcode, is_operator, is_input, is_charge, " & _
    "is_superior, orgname, processed_ymb, holidaytype " & _
    "FROM dbo.stafftbl LEFT JOIN dbo.orgnametbl " & _
    "ON dbo.stafftbl.orgcode=dbo.orgnametbl.orgcode " & _
    "WHERE dbo.stafftbl.personalcode = ?"
Rs_stafftbl_cmd.Prepared = true
Rs_stafftbl_cmd.Parameters.Append Rs_stafftbl_cmd.CreateParameter(_
    "param1", 200, 1, 5, personalcode)
Set Rs_stafftbl = Rs_stafftbl_cmd.Execute
Rs_stafftbl_numRows = 0
name            = Trim(Rs_stafftbl.Fields.Item("staffname"    ).Value)
orgname         = Trim(Rs_stafftbl.Fields.Item("orgname"      ).Value)
proceseed_ymb   = Trim(Rs_stafftbl.Fields.Item("processed_ymb").Value)
holidaytype     = Trim(Rs_stafftbl.Fields.Item("holidaytype"  ).Value)
If (Trim(Rs_stafftbl.Fields.Item("is_operator").Value) = "1") Then
    is_operator = True
Else
    is_operator = False
End If
Rs_stafftbl.Close()
Set Rs_stafftbl = Nothing

' -----------------------------------------------------------------------------
' タイムカードテーブル timecardtbl 読込
' -----------------------------------------------------------------------------
Set Rs_tctbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_tctbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_tctbl_cmd.CommandText = "SELECT * FROM dbo.timecardtbl " & _
    "WHERE personalcode = ? AND punchdate LIKE ? " & _
    "ORDER BY punchdate, punchtime ASC"
Rs_tctbl_cmd.Prepared = true
Rs_tctbl_cmd.Parameters.Append Rs_tctbl_cmd.CreateParameter(_
    "param1", 200, 1, 10, Right("0000000000" & personalcode, 10))
Rs_tctbl_cmd.Parameters.Append Rs_tctbl_cmd.CreateParameter(_
    "param2", 200, 1, 7, ymb & "%")
Set Rs_tctbl = Rs_tctbl_cmd.Execute
Rs_tctbl_numRows = 0

' -----------------------------------------------------------------------------
' PC電源時刻テーブル pctimetbl 読込
' -----------------------------------------------------------------------------
Set Rs_pttbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_pttbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_pttbl_cmd.CommandText = "SELECT i.ipnumber ,p.pcdate ,p.pctime ,p.pcstatus " & _
    "FROM iptbl i LEFT JOIN pctimetbl p ON i.ipnumber=p.ipnumber AND i.begindate<=p.pcdate AND i.enddate>=p.pcdate " & _
    "WHERE i.personalcode = ? AND p.pcdate LIKE ? ORDER BY p.pcdate, p.pctime"
Rs_pttbl_cmd.Prepared = true
Rs_pttbl_cmd.Parameters.Append Rs_pttbl_cmd.CreateParameter(_
    "param1", 200, 1, 5, Right("00000" & personalcode, 5))
Rs_pttbl_cmd.Parameters.Append Rs_pttbl_cmd.CreateParameter(_
    "param2", 200, 1, 7, ymb & "%")
Set Rs_pttbl = Rs_pttbl_cmd.Execute
Rs_pttbl_numRows = 0

%>
<!DOCTYPE HTML>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <title>勤怠管理システム</title>
    <link href="css/default.css" rel="stylesheet" type="text/css">
</head>

<body>
<div id="container">
    <!-- #include file="inc/header.source" -->
    <div id="contents">
        <a href="timecard.asp?p=<%=personalcode%>&ymb=<%=lastYmb%>">&lt;&lt;&nbsp;</a>
        <%=Left(ymb, 4)%>年<%=Right(ymb, 2)%>月分&nbsp;
        <a href="timecard.asp?p=<%=personalcode%>&ymb=<%=nextYmb%>">&gt;&gt;</a><br />
        <table class="data">
            <tr>
                <th width="65px">個人CD</th>
                <td width="65px" class="disabled" align="center"><%=personalcode%></th>
                <th width="65px">氏名</th>
                <td width="200px" class="disabled"><%=name%></th>
                <th width="65px">所属</th>
                <td width="339px" class="disabled"><%=orgname%></th>
            </tr>
        </table>
        <div>
        <div class="left">
        <p>タイムカード打刻データ</p>
        <% If Not Rs_tctbl.EOF Then %>
            <table class="data">
                <tr>
                    <th width="100px">日付</td>
                    <th width="60px">時刻</td>
                    <th width="60px">区分</td>
                </tr>
                <% While (NOT Rs_tctbl.EOF) %>
                <tr>
                    <td style="text-align:center;">
                        <%
                        Response.Write( _
                            Left(Trim(Rs_tctbl.Fields.Item("punchdate").Value), 4) & "年" & _
                            Mid(Trim(Rs_tctbl.Fields.Item("punchdate").Value), 5, 2) & "月" & _
                            Right(Trim(Rs_tctbl.Fields.Item("punchdate").Value), 2) & "日")
                        %>
                    </td>
                    <td style="text-align:center;">
                        <%=editTime(Trim(Rs_tctbl.Fields.Item("punchtime").Value))%>
                    </td>
                    <td style="text-align:center;">
                        <%
                        Select Case Trim(Rs_tctbl.Fields.Item("attendanceclass").Value)
                            Case "01"
                                Response.Write("出勤")
                            Case "02"
                                Response.Write("退勤")
                            Case "03"
                                Response.Write("外出")
                            Case "04"
                                Response.Write("戻り")
                            Case Else
                                Response.Write("－")
                        End Select
                        %>
                    </td>
                </tr>
                <% Rs_tctbl.MoveNext() %>
                <% WEnd %>
            </table>
        <% Else %>
            <p>タイムカードデータがありません。</p>
        <% End If %>
        <p>&nbsp;</p>
        </div>
        <div>
        <div class="left" style="margin-left: 100px;">
        <p>PC電源オンオフデータ</p>
        <% If Not Rs_pttbl.EOF Then %>
            <table class="data">
                <tr>
                    <th width="100px">日付</td>
                    <th width="60px" >時刻</td>
                    <th width="60px" >区分</td>
                    <th width="100px">IPアドレス</td>
                </tr>
                <% While (NOT Rs_pttbl.EOF) %>
                <tr>
                    <td style="text-align:center;">
                        <%
                        Response.Write( _
                            Left(Trim(Rs_pttbl.Fields.Item("pcdate").Value), 4) & "年" & _
                            Mid(Trim(Rs_pttbl.Fields.Item("pcdate").Value), 5, 2) & "月" & _
                            Right(Trim(Rs_pttbl.Fields.Item("pcdate").Value), 2) & "日")
                        %>
                    </td>
                    <td style="text-align:center;">
                        <%=editTime(Trim(Rs_pttbl.Fields.Item("pctime").Value))%>
                    </td style="text-align:center;">
                    <td style="text-align:left;">
                        <%=Trim(Rs_pttbl.Fields.Item("pcstatus").Value)%>
                    </td>
                    <td>
                        <%=Trim(Rs_pttbl.Fields.Item("ipnumber").Value)%>
                    </td>
                </tr>
                <% Rs_pttbl.MoveNext() %>
                <% WEnd %>
            </table>
        <% Else %>
            <p>PC電源オンオフデータがありません。</p>
        <% End If %>
        <p>&nbsp;</p>
        </div>
        <div class="right"></div>
        </div>
        </div>
        <div class="clear"></div>
    </div>
    <!-- #include file="inc/footer.source" -->
</div>
</body>
</html>
<%
Rs_tctbl.Close()
Set Rs_tctbl = Nothing
Rs_pttbl.Close()
Set Rs_pttbl = Nothing
%>
<!-- #include file="inc/util.asp" -->
