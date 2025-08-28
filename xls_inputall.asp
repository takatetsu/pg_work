<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<!-- #include file="Connections/workdbms.asp" -->
<!-- #include file="inc/RestrictAccess.asp" -->
<%
' #######################################
' �v���O�����d�l��
' #######################################
'
' ## �v���O�����T�v ##
' �S�̓��͉�ʂœo�^�������e�� CSV �t�@�C���Ƃ��ă_�E�����[�h����
'
' ## �o�͍��� ##
' XLS �t�@�C���͑S�̓��͉�ʂɏ����Ă��܂��B
'
' ## ���̓`�F�b�N ##
'
' ## ���ӎ��� ##
' EXCEL �t�@�C���Ƃ��ă_�E�����[�h���\�����A�t�@�C�����J�����т�
' �j�����Ă��邩�m�F�c�̃��b�Z�[�W���\�������B
'
' -----------------------------------------------------------------------------
' ��������
' -----------------------------------------------------------------------------
' ���t�v�Z
Dim sysDate     ' �V�X�e�����t
Dim dispDate    ' �\���p���t
Dim dispYear    ' �\���p�N yyyy
Dim dispMonth   ' �\���p�� mm
Dim lastDay     ' �Ώ۔N������

Dim v_workdays

errorMsg = ""
If (Request.QueryString("ymb")<>"") Then
    dispDate = DateSerial(                             _
                Mid(Request.QueryString("ymb"), 1, 4), _
                Mid(Request.QueryString("ymb"), 5, 2), _
                1)
Else
    sysDate = Date
    '15���𒴂���ꍇ�͓����A15���܂ł͑O����\�����Ƃ���
    If Day(sysDate) > 15 Then
        dispDate = sysDate
    else
        dispDate = DateAdd("m", -1, sysDate)
    End If
End If
dispYear  = Year(dispDate)
dispMonth = Right("0" & Month(dispDate), 2)
lastDay   = right(DateSerial(dispYear, dispMonth + 1, 0), 2)

' -----------------------------------------------------------------------------
' �G�N�Z���o�͎w��
' -----------------------------------------------------------------------------
Response.BUFFER=TRUE
Response.ContentType = "application/vnd.ms-excel"
Response.Charset = "utf-8"
Response.AddHeader "Content-Disposition","attachment; filename=�Ζ��\_" & _
                    dispYear & "�N" & dispMonth & "����.xls"

' -----------------------------------------------------------------------------
' ���͉�ʕ\������
' -----------------------------------------------------------------------------

' -----------------------------------------------------------------------------
' ���^�S���҂��Ǘ�����Ώێ҈ꗗ�擾
' -----------------------------------------------------------------------------
Dim Rs_worktbl
Dim Rs_worktbl_cmd
Dim Rs_worktbl_numRows

Set Rs_worktbl_cmd = Server.CreateObject ("ADODB.Command")
Rs_worktbl_cmd.ActiveConnection = MM_workdbms_STRING
Rs_worktbl_cmd.CommandText = "SELECT * FROM "                                           & _
    "(SELECT orgcode FROM orgtbl "                                                  & _
    "WHERE personalcode='" & Session("MM_Username") & "' AND manageclass='1') ORG "     & _
    "LEFT JOIN "                                                                        & _
    "(SELECT personalcode AS pcode, staffname, orgcode AS org FROM stafftbl "       & _
    "WHERE is_enable='1') STAFF "                                                       & _
    "ON ORG.orgcode=STAFF.org "                                                         & _
    "LEFT JOIN "                                                                        & _
    "(SELECT * FROM dutyrostertbl WHERE ymb='" & dispYear & dispMonth & "') DUTY "  & _
    "ON STAFF.pcode=DUTY.personalcode "                                                 & _
    "ORDER BY STAFF.pcode ASC"
Rs_worktbl_cmd.Prepared = true

Set Rs_worktbl = Rs_worktbl_cmd.Execute
Rs_worktbl_numRows = 0
%>
<table border="1px">
    <tr>
        <td>�lCD</td>
        <td>����</td>
        <td>�o�Γ���</td>
        <td>��x����</td>
        <td>���Γ���</td>
        <td>�L������</td>
        <td>�ۑ��x�ɓ���</td>
        <td>���x����</td>
        <td>�x�o����</td>
        <td>���o�Γ���</td>
        <td>�x����</td>
        <td>�h��A��</td>
        <td>�h��B��</td>
        <td>�h��C��</td>
        <td>�h��D��</td>
        <td>�x������</td>
        <td>����</td>
        <td>��֍b��</td>
        <td>��։���</td>
        <td>��֕���</td>
        <td>���A��</td>
        <td>���B��</td>
        <td>�ďo�ʏ��</td>
        <td>�ďo�[���</td>
        <td>�N���N�n1230</td>
        <td>�N���N�n1231</td>
        <td>���ԑ�x</td>
        <td>�[�銄��</td>
        <td>���ԊO</td>
        <td>�x���o��</td>
        <td>�x�o���O</td>
        <td>�x�o�[��</td>
        <td>���O�[��</td>
        <td>�x�o���O�[��</td>
    </tr>
<%
While (NOT Rs_worktbl.EOF)
    personalcode = Trim(Rs_worktbl.Fields.Item("pcode"    ).Value)
    staffname    = Trim(Rs_worktbl.Fields.Item("staffname").Value)
    workdays                = Rs_worktbl.Fields.Item("workdays"                  ).Value
    workholidays            = Rs_worktbl.Fields.Item("workholidays"              ).Value
    absencedays             = Rs_worktbl.Fields.Item("absencedays"               ).Value
    paidvacations           = Rs_worktbl.Fields.Item("paidvacations"             ).Value
    preservevacations       = Rs_worktbl.Fields.Item("preservevacations"         ).Value
    specialvacations        = Rs_worktbl.Fields.Item("specialvacations"          ).Value
    holidayshifts           = Rs_worktbl.Fields.Item("holidayshifts"             ).Value
    realworkdays            = Rs_worktbl.Fields.Item("realworkdays"              ).Value
    shortdays               = Rs_worktbl.Fields.Item("shortdays"                 ).Value
    nightduty_a             = Rs_worktbl.Fields.Item("nightduty_a"               ).Value
    nightduty_b             = Rs_worktbl.Fields.Item("nightduty_b"               ).Value
    nightduty_c             = Rs_worktbl.Fields.Item("nightduty_c"               ).Value
    nightduty_d             = Rs_worktbl.Fields.Item("nightduty_d"               ).Value
    holidaypremium          = Rs_worktbl.Fields.Item("holidaypremium"            ).Value
    dayduty                 = Rs_worktbl.Fields.Item("dayduty"                   ).Value
    shiftwork_kou           = Rs_worktbl.Fields.Item("shiftwork_kou"             ).Value
    shiftwork_otsu          = Rs_worktbl.Fields.Item("shiftwork_otsu"            ).Value
    shiftwork_hei           = Rs_worktbl.Fields.Item("shiftwork_hei"             ).Value
    shiftwork_a             = Rs_worktbl.Fields.Item("shiftwork_a"               ).Value
    shiftwork_b             = Rs_worktbl.Fields.Item("shiftwork_b"               ).Value
    summons                 = Rs_worktbl.Fields.Item("summons"                   ).Value
    summonslate             = Rs_worktbl.Fields.Item("summonslate"               ).Value
    yearend1230             = Rs_worktbl.Fields.Item("yearend1230"               ).Value
    yearend1231             = Rs_worktbl.Fields.Item("yearend1231"               ).Value
    workholidaytime         = Rs_worktbl.Fields.Item("workholidaytime"           ).Value
    latepremium             = Rs_worktbl.Fields.Item("latepremium"               ).Value
    overtime                = Rs_worktbl.Fields.Item("overtime"                  ).Value
    holidayshifttime        = Rs_worktbl.Fields.Item("holidayshifttime"          ).Value
    holidayshiftovertime    = Rs_worktbl.Fields.Item("holidayshiftovertime"      ).Value
    holidayshiftlate        = Rs_worktbl.Fields.Item("holidayshiftlate"          ).Value
    overtimelate            = Rs_worktbl.Fields.Item("overtimelate"              ).Value
    holidayshiftovertimelate= Rs_worktbl.Fields.Item("holidayshiftovertimelate"  ).Value
%>
    <tr>
        <td><%="=""" & personalcode & """"%></td>
        <td><%=staffname%></td>
        <td><%=workdays%></td>
        <td><%=workholidays%></td>
        <td><%=absencedays%></td>
        <td><%=paidvacations%></td>
        <td><%=preservevacations%></td>
        <td><%=specialvacations%></td>
        <td><%=holidayshifts%></td>
        <td><%=realworkdays%></td>
        <td><%=shortdays%></td>
        <td><%=nightduty_a%></td>
        <td><%=nightduty_b%></td>
        <td><%=nightduty_c%></td>
        <td><%=nightduty_d%></td>
        <td><%=holidaypremium%></td>
        <td><%=dayduty%></td>
        <td><%=shiftwork_kou%></td>
        <td><%=shiftwork_otsu%></td>
        <td><%=shiftwork_hei%></td>
        <td><%=shiftwork_a%></td>
        <td><%=shiftwork_b%></td>
        <td><%=summons%></td>
        <td><%=summonslate%></td>
        <td><%=yearend1230%></td>
        <td><%=yearend1231%></td>
        <td><%=workholidaytime%></td>
        <td><%=latepremium%></td>
        <td><%=overtime%></td>
        <td><%=holidayshifttime%></td>
        <td><%=holidayshiftovertime%></td>
        <td><%=holidayshiftlate%></td>
        <td><%=overtimelate%></td>
        <td><%=holidayshiftovertimelate%></td>
    </tr>
<%
    Rs_worktbl.MoveNext()
Wend
%>
</table>
<%
Rs_worktbl.Close()
Set Rs_worktbl = Nothing
Response.Flush
Response.End
Response.Redirect("inputall.asp")
%>
