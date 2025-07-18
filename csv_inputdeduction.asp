<%@LANGUAGE="VBSCRIPT" CODEPAGE="932"%>
<!-- #include file="Connections/workdbms.asp" -->
<%
' #######################################
' �v���O�����d�l��
' #######################################
'
' ## �v���O�����T�v ##
' �x�X�T�����͉�ʂœo�^�������e�� CSV �t�@�C���Ƃ��ă_�E�����[�h����
'
' ## �o�͍��� ##
' CSV �t�@�C���͎x�X�T�����͉�ʂɏ����Ă��܂��B
'
' ## ���̓`�F�b�N ##
'
' ## ���ӎ��� ##
' EXCEL �t�@�C���Ƃ��ă_�E�����[�h���\�����A�t�@�C�����J�����т�
' �j�����Ă��邩�m�F�c�̃��b�Z�[�W���\�������̂ŁA CSV �t�@�C��
' �Ƃ��ă_�E�����[�h����B

' ���t�̌v�Z
Dim sysDate     '�V�X�e�����t
Dim dispDate    '�\���p���t
Dim dispYear    '�\���p�N yyyy
Dim dispMonth   '�\���p�� mm
Dim i           '�J��Ԃ��p���t

If (Request.QueryString("ymb")<>"") Then
    dispDate = DateSerial(Mid(Request.QueryString("ymb"), 1, 4), Mid(Request.QueryString("ymb"), 5, 2), 1)
Else
    dispDate = Date
End If
dispYear    = Year(dispDate)
dispMonth   = Right("0" & Month(dispDate), 2)

' -----------------------------------------------------------------------------
' �G�N�Z���o�͎w��
' -----------------------------------------------------------------------------
Response.BUFFER=TRUE
Response.ContentType = "application/vnd.ms-excel"
Response.Charset = "utf-8"
Response.AddHeader "Content-Disposition","attachment; filename=�x�X�T��_" & _
                    dispYear & "�N" & dispMonth & "����.csv"

' deductiontbl���A�S���҂�ymd�ŐV���R�[�h���擾(���o�^�A�o�^�ρ@����p)
Dim Rs_dedu_ck
Dim Rs_dedu_ck_cmd
Set Rs_dedu_ck_cmd = Server.CreateObject ("ADODB.Command")
Rs_dedu_ck_cmd.ActiveConnection = MM_workdbms_STRING
Rs_dedu_ck_cmd.CommandText = "SELECT * FROM deductiontbl WHERE ymb = (SELECT MAX(ymb) FROM deductiontbl WHERE personalcode = ? ) AND personalcode = ?"
Rs_dedu_ck_cmd.Prepared = true
Rs_dedu_ck_cmd.Parameters.Append Rs_dedu_ck_cmd.CreateParameter("param1", 200, 1, -1, Session("MM_Username") )
Rs_dedu_ck_cmd.Parameters.Append Rs_dedu_ck_cmd.CreateParameter("param1", 200, 1, -1, Session("MM_Username") )
Set Rs_dedu_ck = Rs_dedu_ck_cmd.Execute
If Not Rs_dedu_ck.EOF Or Not Rs_dedu_ck.BOF Then
    other1Name = Trim(Rs_dedu_ck.Fields.Item("amount08ncr"))
    other2Name = Trim(Rs_dedu_ck.Fields.Item("amount09ncr"))
    other3Name = Trim(Rs_dedu_ck.Fields.Item("amount10ncr"))
End If
'�S���҂�ymd�ŐV���R�[�h���N���[�Y
Rs_dedu_ck.Close()
Set Rs_dedu_ck = Nothing

Response.Write("�lCD,����,�T���z���v,�΍Ћ���,��ʍЊQ,���ԏ��," & _
    "�Z��v��,������,���i�j��,�x����(�g��),���̑��P:" & other1Name & _
    ",���̑��Q:" & other2Name & ",���̑��R:" & other3Name & vbNewLine)

' stafftbl���A�\���E���ꗗ���擾
Dim Rs_staff
Dim Rs_staff_cmd
Set Rs_staff_cmd = Server.CreateObject ("ADODB.Command")
Rs_staff_cmd.ActiveConnection = MM_workdbms_STRING
Rs_staff_cmd.CommandText = "SELECT stafftbl.personalcode ,stafftbl.staffname FROM orgtbl " & _
        "RIGHT OUTER JOIN stafftbl stafftbl ON orgtbl.orgcode = stafftbl.orgcode " & _
        "WHERE stafftbl.is_enable = '1' AND orgtbl.personalcode = ?  AND " & _
        "orgtbl.manageclass = '0' ORDER BY stafftbl.orgcode, stafftbl.gradecode DESC, stafftbl.personalcode"
Rs_staff_cmd.Prepared = true
Rs_staff_cmd.Parameters.Append Rs_staff_cmd.CreateParameter("param1", 200, 1, -1, Session("MM_Username") )
'�\���E���̈ꗗ�擾SQL�̔��s
Set Rs_staff = Rs_staff_cmd.Execute

'�\���E���̐������A�f�[�^�̍X�V���s��
Dim perCod
While (NOT Rs_staff.EOF)
    perCod = Rs_staff.Fields.Item("personalcode")

    'deductiontbl���A�ŐV���̎x�X�T���ǂݍ���
    Dim Rs_dedu
    Dim Rs_dedu_cmd
    Set Rs_dedu_cmd = Server.CreateObject ("ADODB.Command")
    Rs_dedu_cmd.ActiveConnection = MM_workdbms_STRING
    Rs_dedu_cmd.CommandText = "SELECT * FROM deductiontbl WHERE ymb = ? AND personalcode = ?"
    Rs_dedu_cmd.Prepared = true
    Rs_dedu_cmd.Parameters.Append Rs_dedu_cmd.CreateParameter("param1", 200, 1, -1, Request.QueryString("ymb") )
    Rs_dedu_cmd.Parameters.Append Rs_dedu_cmd.CreateParameter("param1", 200, 1, -1, perCod )
    Set Rs_dedu = Rs_dedu_cmd.Execute

    sumAmount = Rs_dedu.Fields.Item("amount01") + _
                Rs_dedu.Fields.Item("amount02") + _
                Rs_dedu.Fields.Item("amount03") + _
                Rs_dedu.Fields.Item("amount04") + _
                Rs_dedu.Fields.Item("amount05") + _
                Rs_dedu.Fields.Item("amount06") + _
                Rs_dedu.Fields.Item("amount07") + _
                Rs_dedu.Fields.Item("amount08") + _
                Rs_dedu.Fields.Item("amount09") + _
                Rs_dedu.Fields.Item("amount10")

    Response.Write("=""" & perCod & """," & _
        Trim(Rs_staff.Fields.Item("staffname")) & _
        "," & sumAmount                         & _
        "," & Rs_dedu.Fields.Item("amount01")   & _
        "," & Rs_dedu.Fields.Item("amount02")   & _
        "," & Rs_dedu.Fields.Item("amount03")   & _
        "," & Rs_dedu.Fields.Item("amount04")   & _
        "," & Rs_dedu.Fields.Item("amount05")   & _
        "," & Rs_dedu.Fields.Item("amount06")   & _
        "," & Rs_dedu.Fields.Item("amount07")   & _
        "," & Rs_dedu.Fields.Item("amount08")   & _
        "," & Rs_dedu.Fields.Item("amount09")   & _
        "," & Rs_dedu.Fields.Item("amount10") & vbNewLine)
    Rs_staff.MoveNext()
Wend
'�\���E���ꗗ���擾����SQL���N���[�Y
Rs_staff.Close()
Set Rs_staff = Nothing
Response.Flush
Response.End
Response.Redirect("inputdeduction.asp")
%>
