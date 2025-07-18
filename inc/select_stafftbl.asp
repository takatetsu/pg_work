<%
' -----------------------------------------------------------------------------
' 社員テーブル stafftbl 読込
' -----------------------------------------------------------------------------
Dim Rs_stafftbl
Dim Rs_stafftbl_cmd
Dim Rs_stafftbl_numRows
Set Rs_stafftbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_stafftbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_stafftbl_cmd.CommandText = "SELECT personalcode, staffname, dbo.stafftbl.gradecode, " & _
    "dbo.stafftbl.orgcode AS orgcode, is_operator, is_input, is_charge, " & _
    "is_superior, orgname, processed_ymb, holidaytype, grantdate, workshift, " & _
    "old_workshift, old_workshift_last_ymb, base_am_workmin, base_pm_workmin " & _
    "FROM dbo.stafftbl LEFT JOIN dbo.orgnametbl " & _
    "ON dbo.stafftbl.orgcode=dbo.orgnametbl.orgcode " & _
    "WHERE dbo.stafftbl.personalcode = ?"
Rs_stafftbl_cmd.Prepared = true
Rs_stafftbl_cmd.Parameters.Append Rs_stafftbl_cmd.CreateParameter("param1", 200, 1, 5, target_personalcode)
Set Rs_stafftbl = Rs_stafftbl_cmd.Execute
Rs_stafftbl_numRows = 0
name            = Trim(Rs_stafftbl.Fields.Item("staffname"    ).Value)
orgname         = Trim(Rs_stafftbl.Fields.Item("orgname"      ).Value)
proceseed_ymb   = Trim(Rs_stafftbl.Fields.Item("processed_ymb").Value)
holidaytype     = Trim(Rs_stafftbl.Fields.Item("holidaytype"  ).Value)
gradecode       = Trim(Rs_stafftbl.Fields.Item("gradecode"    ).Value)
grantdate       = Trim(Rs_stafftbl.Fields.Item("grantdate"    ).Value)
If ymb > Trim(Rs_stafftbl.Fields.Item("old_workshift_last_ymb").Value) Then
    workshift = Trim(Rs_stafftbl.Fields.Item("workshift").Value)
Else
    workshift = Trim(Rs_stafftbl.Fields.Item("old_workshift").Value)
End If
If (Trim(Rs_stafftbl.Fields.Item("is_operator").Value) = "1") Then
    is_operator = True
Else
    is_operator = False
End If
base_am_workmin = Rs_stafftbl.Fields.Item("base_am_workmin").Value
base_pm_workmin = Rs_stafftbl.Fields.Item("base_pm_workmin").Value
Rs_stafftbl.Close()
Set Rs_stafftbl = Nothing
%>