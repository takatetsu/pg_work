<%
' -------------------------------------------------------------------------
' �J�����ԓK�����`�F�b�N
' �����Fv_operator      �I�y���[�^
'       v_morningwork   �o�΋敪(�ߑO)
'       v_afternoonwork �o�΋敪(�ߌ�)
'       cometime        �o�Ύ���
'       leavetime       �ދΎ���
'       pc_ontime       PC�N������
'       pc_offtime      PC�I������
'       dayduty         ����
'       nightduty       �h��
'       nightduty2      �O���h��
'       overtime_begin  ���ԊO�J�n����
'       overtime_end    ���ԊO�I������
'       memo2           �����Q
'       opentime        �n�Ǝ���
'       closetime       �I�Ǝ���
'       is_unionexecutive �g�����s����
'       v_operator      �O����֋Ζ�
' �ߒl�F0     �G���[�Ȃ�
'       0�ȊO �G���[
' -------------------------------------------------------------------------
Function workTimeCheck(v_operator, v_morningwork, v_afternoonwork, cometime, _
                       leavetime, pc_ontime, pc_offtime, dayduty, nightduty, _
                       nightduty2, overtime_begin, overtime_end, memo2, _
                       opentime, closetime, is_unionexecutive, v_operator2)
    ' Function���ł̐��l�ύX�ɔ����A�ʕϐ��ŏ������s��
    c_operator          = v_operator
    c_morningwork       = v_morningwork
    c_afternoonwork     = v_afternoonwork
    c_cometime          = cometime
    c_leavetime         = leavetime
    c_pc_ontime         = pc_ontime
    c_pc_offtime        = pc_offtime
    c_dayduty           = dayduty
    c_nightduty         = nightduty
    c_nightduty2        = nightduty2
    c_overtime_begin    = overtime_begin
    c_overtime_end      = overtime_end
    c_memo2             = memo2
    c_is_unionexecutive = is_unionexecutive
    ' �y�o�Ώ󋵂ɉ������`�F�b�N�p�̊������ݒ肵�A�K�v�ȃ`�F�b�N���ڂ̃t���O��1�ɂ���z
    workTimeCheck  = "0"        ' ����ߒl
    ref_starttime  = opentime   ' ��A�ƊJ�n����
    ref_endtime    = closetime  ' ��A�ƏI������
    res_timeDev    = "0"        ' �o�ދΎ����APC�N����~���������`�F�b�N����
    res_comeCheck  = "0"        ' �J�n�����`�F�b�N����
    res_outCheck   = "0"        ' �I�������`�F�b�N����
    flg_checkStart = "0"        ' �J�n���ԃ`�F�b�N���邩�̃t���O 0:�`�F�b�N���Ȃ� 1:�`�F�b�N����
    flg_checkEnd   = "0"        ' �I�����ԃ`�F�b�N���邩�̃t���O 0:�`�F�b�N���Ȃ� 1:�`�F�b�N����
    dif_startTime  = 30         ' �o�Ў����`�F�b�N�P�\����(��)
    dif_endTime    = 50         ' �ގЎ����`�F�b�N�P�\����(��)

    ' ����2����
    If c_memo2 = "1" Or c_memo2 = "2" Then
        ' �ʋΏa�؉���A�d�Ԏ��ԓs���̂Ƃ��o�Ύ����`�F�b�N�̓`�F�b�N�ΏۊO�Ƃ���
        c_cometime = ""
    End If
    If c_memo2 = "2" Or c_memo2 = "4" Then
        ' �d�Ԏ��ԓs���A���e��ԑ҂̂Ƃ��ގЎ����̓`�F�b�N�ΏۊO�Ƃ���
        c_leavetime  = ""
    End If
    If c_memo2 = "3" Then
        ' �g�������̂Ƃ��ގЎ����̓`�F�b�N�ΏۊO�Ƃ���
        c_leavetime  = ""
        If c_is_unionexecutive = "1" Then
            ' �g�����s���̂Ƃ�PC�I���������`�F�b�N�ΏۊO�Ƃ���
            c_pc_offtime = ""
        End If
    End If
    If c_memo2 = "5" Then
        ' PC�����Y��̂Ƃ�PC�I�������̓`�F�b�N�ΏۊO�Ƃ���
        c_pc_offtime = ""
    End If

    dec_starttime  = setTime(c_cometime,  c_pc_ontime,  "0") ' ����p�J�n����
    dec_endtime    = setTime(c_leavetime, c_pc_offtime, "1") ' ����p�I������

    If c_operator = "0" Then
        ' ���Ζ��ȊO(��ʋΖ�)
        If c_morningwork = "1" Or _
           c_morningwork = "4" Or _
           c_morningwork = "5" Or _
           c_morningwork = "9" Then
            ' �ߑO�o��(�U�֏o�΁A�o��(�o��)�A�o��(�U�֏o��)�A�o��)
            If c_afternoonwork = "1" Or _
               c_afternoonwork = "4" Or _
               c_afternoonwork = "5" Or _
               c_afternoonwork = "9" Then
                ' �ߑO�o�΁E�ߌ�o��(�U�֏o�΁A�o��(�o��)�A�o��(�U�֏o��)�A�o��)
                ref_starttime  = opentime
                ref_endtime    = closetime
                flg_checkStart = "1"
                flg_checkEnd   = "1"
            Else
                If (v_afterwork = "2" Or v_afterwork = "3" Or v_afterwork = "6") Then
                    ' �ߑO�o�΁E�ߌ�x�o�A�x�o(��������)�A�o��(�x�o)
                    ref_endtime    = c_overtime_end
                Else
                    ' �ߑO�o�΁E�ߌ�o�΂���
                    ref_endtime    = "12:00"
                End If
                ref_starttime  = opentime
                flg_checkStart = "1"
                flg_checkEnd   = "1"
                dif_endTime    = 30
            End If
        Else
             If (c_morningwork = "2" Or c_morningwork = "3" or c_morningwork = "6") Then
                ' �ߑO�x�o�A�x�o(��������)�A�o��(�x�o)
                If c_afternoonwork = "1" Or _
                   c_afternoonwork = "4" Or _
                   c_afternoonwork = "5" Or _
                   c_afternoonwork = "9" Then
                    ' �ߑO�x�o�A�x�o(��������)�A�o��(�x�o)�E�ߌ�o��(�U�֏o�΁A�o��(�o��)�A�o��(�U�֏o��)�A�o��)
                    ref_starttime  = c_overtime_begin
                    ref_endtime    = closetime
                    flg_checkStart = "1"
                    flg_checkEnd   = "1"
                Else
                    ' �ߑO�x�o�A�x�o(��������)�A�o��(�x�o)�E�ߌ�x�o�A�x�o(��������)�A�o��(�x�o) �������� �ߌ�o�΂��� �̂Ƃ�
                    ref_starttime  = c_overtime_begin
                    ref_endtime    = c_overtime_end
                    flg_checkStart = "1"
                    flg_checkEnd   = "1"
                End If
             Else
                ' �ߑO�o�΂���
                If c_afternoonwork = "1" Or _
                   c_afternoonwork = "4" Or _
                   c_afternoonwork = "5" Or _
                   c_afternoonwork = "9" Then
                    ' �ߑO�o�΂����E�ߌ�o��(�U�֏o�΁A�o��(�o��)�A�o��(�U�֏o��)�A�o��)
                    ref_starttime  = "13:00"
                    ref_endtime    = closetime
                    flg_checkStart = "1"
                    flg_checkEnd   = "1"
                Else
                    If (v_afterwork = "2" Or v_afterwork = "3" Or v_afterwork = "6") Then
                        ' �ߑO�o�΂����E�ߌ�x�o�A�x�o(��������)�A�o��(�x�o)
                        ref_starttime  = c_overtime_begin
                        ref_endtime    = c_overtime_end
                        flg_checkStart = "1"
                        flg_checkEnd   = "1"
                    Else
                        ' �ߑO�o�΂����E�ߌ�o�΂���
                        If c_cometime <> "" Or c_pc_ontime <> "" Then
                            If c_morningwork = "0" And c_afternoonwork = "0" And c_dayduty = "0" And c_nightduty = "0" Then
                                ' �ߑO�ߌ�Ƃ��o�΋敪�̓��͂������A�����h���ł��Ȃ��Ƃ��G���[
                                res_comeCheck = "1"
                                res_outCheck  = "1"
                            Else
                                If c_dayduty <> "0" And c_nightduty =  "0" Then
                                    ' �����ŏh���łȂ��Ƃ�
                                    flg_checkStart = "1"
                                    flg_checkEnd   = "1"
                                End If
                                If c_dayduty =  "0" And c_nightduty <> "0" Then
                                    ' �����łȂ��h���̂Ƃ�
                                    flg_checkStart = "1"
                                    ref_starttime  = "17:10"
                                End If
                                If c_dayduty <> "0" And c_nightduty <> "0" Then
                                    ' �����h���̂Ƃ�
                                    flg_checkStart = "1"
                                End If
                            End If
                        End If
                        If c_leavetime <> "" Or c_pc_offtime <> "" Then
                            If c_morningwork = "0" And c_afternoonwork = "0" And c_dayduty = "0" And c_nightduty = "0" And c_nightduty2 = "0" Then
                                ' �ߑO�ߌ�Ƃ��o�΋敪�̓��͂������A�����h���łȂ��A�O���h���ł��Ȃ��Ƃ��G���[
                                res_comeCheck = "1"
                                res_outCheck  = "1"
                            Else
                                If c_dayduty <> "0" And c_nightduty = "0" Then
                                    ' �����ŏh���łȂ��Ƃ�
                                    flg_checkEnd   = "1"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        ' ���Ζ��ɊY��
        If c_operator = "1" or _
           c_operator = "3" or _
           c_operator = "5" Then
           ' �b�ԋΖ�(�b�ԁA���΍b�A���K(�b))
            ref_starttime  = opentime
            ref_endtime    = "20:30"
            dif_endTime    = 30
            flg_checkStart = "1"
            flg_checkEnd   = "1"
        End If
        If c_operator = "2" or _
           c_operator = "6" Then
           ' ���ԋΖ�(���ԁA���K(��))
            ref_starttime  = "20:30"
            flg_checkStart = "1"
        End If
    End If
    ' ���ԊO��������
    If c_overtime_end > opentime Then
        ' ���ԊO�����͂���Ă��ďI���������n�Ǝ����ȍ~�̂Ƃ��A�ގЎ����`�F�b�N�P�\���Ԃ�30���Ƃ���
        dif_endTime    = 30
    End If
    ref_starttime  = setTime(ref_starttime,  c_overtime_begin,  "0") ' ����p�J�n����
    ref_endtime    = setTime(ref_endtime,    c_overtime_end,    "1") ' ����p�I������
    
    ' �h������
    If (c_nightduty2 = "1" Or c_nightduty2 = "2") Then ' �O���h��
        ' �O���h���̂Ƃ��J�n�����`�F�b�N�͖�������
        res_comeCheck  = "0"
        flg_checkStart = "0"
        If c_morningwork = "0" And c_afternoonwork = "0" And c_dayduty = "0" And c_nightduty = "0" Then
            ' �O���h���œ����o�΂Ȃ��A�����h���Ȃ��̂Ƃ��A�I�������`�F�b�N�����08:30�A�P�\30���ōs��
            flg_checkEnd   = "1"
            ref_endtime    = "08:30"
            dif_endTime    = 30
        Else
            ' �O���h���œ����o�΂���̂Ƃ��APC�N���������N���A
            c_pc_ontime = ""
        End If
    End If
    If (c_nightduty  = "1" Or c_nightduty  = "2") Then ' �����h��
        flg_checkEnd   = "0"
    End If
    
    ' ��֋Ζ�����
    If (v_operator2 = "2" Or v_operator2 = "4" Or v_operator2 = "6") Then ' �O������
        ' �O�����Ԃ̂Ƃ��J�n�����`�F�b�N�͖�������
        res_comeCheck  = "0"
        flg_checkStart = "0"
        If c_morningwork = "0" And c_afternoonwork = "0" Then
            ' �O�����Ԃœ����o�΂Ȃ��̂Ƃ��A�I�������`�F�b�N�����08:30�A�P�\30���ōs��
            flg_checkEnd   = "1"
            ref_endtime    = "08:30"
            dif_endTime    = 30
        Else
            ' �O�����Ԃœ����o�΂���̂Ƃ��APC�N���������N���A
            c_pc_ontime = ""
        End If
    End If
    
    ' �y�����`�F�b�N�z
    If flg_checkStart = "1" And dec_starttime <> "" Then
        ' �J�n�����`�F�b�N 40���O�܂�OK
        res_comeCheck = checkTimeInterval(dec_starttime, ref_starttime, dif_startTime)
 'response.write("<br />START/kijyun:" & ref_starttime & " hantei:" & dec_starttime & " res_comeCheck=" & res_comeCheck & " res_outCheck=" & res_outCheck)
    End If
    If flg_checkEnd   = "1" And dec_endtime   <> "" Then
        ' �I�������`�F�b�N
        res_outCheck  = checkTimeInterval(ref_endtime,   dec_endtime,   dif_endTime)
 'response.write("<br />E N D/kijyun:" & ref_endtime & " hantei:" & dec_endtime & " yuuyo:" & dif_endTime & " res_outCheck=" & res_outCheck & " res_comeCheck=" & res_comeCheck & "<br />")
    End If

    ' ���ʃR�[�h�ݒ�
    If res_comeCheck <> "0" Or res_outCheck <> "0" Then
        workTimeCheck = "1"
    End If
End Function

' -----------------------------------------------------------------------------
' �����Ԋu�`�F�b�N
' �����Ft1 (�`�F�b�N�J�n����)
'       t2 (�`�F�b�N�I������),
'       m  (���e����)
' �ߒl�F0 �`�F�b�NOK
'       1 t1��t2��m�������Ԋu���J���Ă��邽�߃G���[
' -----------------------------------------------------------------------------
Function checkTimeInterval(t1, t2, m)
    checkTimeInterval = "1"
    If t1 <> "" And t2 <> "" Then
        If t1 <= t2 Then
            If minDifIV(editTime(t1), editTime(t2)) <= m Then
                checkTimeInterval = "0"
            End If
        Else
            checkTimeInterval = "0"
        End If
    End If
End Function

' -----------------------------------------------------------------------------
' �`�F�b�N�p�����ݒ�
' �����Ft1  (�`�F�b�N�Ώێ���1)
'       t2  (�`�F�b�N�Ώێ���2),
'       j   (����t���O 0:���������擾�A0�ȊO:�x�������擾),
' �ߒl�Ft   t1��t2�̂���j�Őݒ肳�ꂽ����̌��ʂ�Ԃ�
'           �ǂ��炩�ɒl�������Ƃ��͒l���������Ԃ�
' -----------------------------------------------------------------------------
Function setTime(t1, t2, j)
    setTime = ""
    If t1 <> "" Then
        If t2 <> "" Then
            If t1 <= t2 Then
                setSmallTime = t1
                setBigTime   = t2
            Else
                setSmallTime = t2
                setBigTime   = t1
            End If
        Else
            setSmallTime = t1
            setBigTime   = t1
        End If
    Else
        If t2 <> "" then
            setSmallTime = t2
            setBigTime   = t2
        End If
    End If
    If j = "0" Then
        setTime = setSmallTime
    Else
        setTime = setBigTime
    End If
End Function

' -----------------------------------------------------------------------------
' �O����֋Ζ���ݒ�
' �����Fpersonalcode �l�R�[�h
'       ymb yyyymmdd�̃t�H�[�}�b�g�œ��t��ݒ�
' �ߒl�Ft   �����œn���ꂽ���t�̑O���̌�֋Ζ���Ԃ�
'           �ǂ��炩�ɒl�������Ƃ��͒l���������Ԃ�
' -----------------------------------------------------------------------------
Function setPreOp(personalcode, ymb)
    setPreOp = ""
    ' �O�����t�Z�o
    predate = DateAdd("d", -1, CDate(Left(ymb,4) & "/" & Mid(ymb,5,2) & "/" & Right(ymb,2)))
    predate = Left(predate,4) & Mid(predate,6,2) & Right(predate,2)
    ' �O����֋Ζ���ǂݍ��ݐݒ肷��
    Dim Rs_previous_worktbl
    Dim Rs_previous_worktbl_cmd
    Dim Rs_previous_worktbl_numRows
    Set Rs_previous_worktbl_cmd = Server.CreateObject ("ADODB.Command")
    Rs_previous_worktbl_cmd.ActiveConnection = MM_workdbms_STRING
    Rs_previous_worktbl_cmd.CommandText = "SELECT operator FROM dbo.worktbl WHERE personalcode = ? AND workingdate = ?"
    Rs_previous_worktbl_cmd.Prepared = true
    Rs_previous_worktbl_cmd.Parameters.Append Rs_previous_worktbl_cmd.CreateParameter("param1", 200, 1, 5, personalcode)
    Rs_previous_worktbl_cmd.Parameters.Append Rs_previous_worktbl_cmd.CreateParameter("param2", 200, 1, 8, predate)
    Set Rs_previous_worktbl = Rs_previous_worktbl_cmd.Execute
    Rs_previous_worktbl_numRows = 0
    If Rs_previous_worktbl.EOF And Rs_previous_worktbl.BOF Then
        setPreOp = ""
    Else
        setPreOp = Rs_previous_worktbl.Fields.Item("operator").Value
    End If
    Rs_previous_worktbl.Close()
    Set Rs_previous_worktbl = Nothing
End Function

%>
