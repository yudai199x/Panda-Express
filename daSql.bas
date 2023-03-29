Attribute VB_Name = "daSql"
Option Explicit
Option Base 1
'********************************************************************************************
'  ��͈˗��V�X�e�� - �Ɩ����ʏ������W���[��
'               Copyright 2015, XXXX All Rights Reserved.
'  2015-05-14 �V�K�쐬
'********************************************************************************************

' ********
' �萔��`
' ********
' �I�����ڃ}�X�^ <�˗��V�[�g�쐬> �u���ށv�擾�L�[
Public Const cItemKeyIraiBunrui As String = "REQCATEG"

' �I�����ڃ}�X�^ <�˗��V�[�g�쐬> �u��������v�擾�L�[
Public Const cItemKeyIraiHachu As String = "REQSEC"

' �I�����ڃ}�X�^ <�˗��V�[�g�쐬> �u�������e �i��v�擾�L�[
Public Const cItemKeyIraiHinshu As String = "SAMPKIND"

' �ݒ�l�擾 �˗��󋵎擾 �ő�擾����
Public Const cSetupIraiStatusMax As String = "�˗��󋵍ő�\����"

' �۔F���A���F���ɕ\�����镶�����`�i2015/10/15�ǉ��j
Public Const cNotApproved As String = "�۔F"

' �۔F���A���F���ɕ\�����镶�����`�i2016/4/11�ǉ��j
Public Const cReEstimate As String = "�Č��ς���"

' ���ѕς��L�[���[�h�����i2016/03/31�ǉ��j
Public Const cOrderByAsc As String = "asc"

' ���ѕς��L�[���[�h�~���i2016/03/31�ǉ��j
Public Const cOrderByDesc As String = "desc"


' **********
' �\���̒�`
' **********
' ���m�点��� �\����
Public Type SModifyHistory
    dtModifyDate    As Date         ' �X�V���t (�N����+�����b)
    strMessage      As String       ' ���b�Z�[�W
    sLink           As String       ' �����N��
End Type

'**********************************************************************
' @(f)
' �@�\      : �u���m�点�v�擾����
'
' �Ԃ�l    : Long : �擾����&���� (0�ȏ�=�擾�����A����=�G���[)
'
' ������    : Integer iTarget �擾�Ώ� (0:TSB(�˗���)�����A1:TBA����)
'           : SModifyHistory mhInfo() �擾�����u���m�点�v���
'
' �@�\����  : DB �X�V�����e�[�u������A�����̎擾�Ώۂ́u���m�点�v��
'           : �ŐV����4���擾���A�����̍\���̂Ɋi�[����B
'
' ���l      : ���m�点�����݂��Ȃ��ꍇ�́AmhInfo �� Empty ���ݒ肳���B
'
'**********************************************************************
Public Function GetModifyHistory(iTarget As Integer, mhInfo() As SModifyHistory) As Long
    Dim iCnt As Long
    Dim i As Long
    
    Erase mhInfo
    
    Dim wSql As SSqlSet
    ' �� SQL����������B
    '    5:�u���m�点�v�擾
    wSql = n_InitSql(5)
    
    ' [SQL] �N�G�����ɐݒ肷��l(�ύX�l�A������)��ݒ肷��B
    wSql.rep.Add "\Target", str(iTarget)
    wSql.rep.Add "\MaxCount", "4"
    
    ' �� SQL�����s����B
    iCnt = n_DoSql(wSql)

    If iCnt > 0 Then
        ' ���o�������ʂ��o�͗p�z��ɐݒ肷��B
        ReDim mhInfo(1 To iCnt)     ' �z��������ɍ��킹�Ċg������B
        For i = 1 To iCnt
            mhInfo(i).dtModifyDate = CDate(wSql.psData(1, i))         ' TEM_INPDAT �o�^����
            mhInfo(i).strMessage = DbResumeNewLine(wSql.psData(2, i)) ' TEM_VALUE ���m�点���e
            mhInfo(i).sLink = DbResumeNewLine(wSql.psData(3, i))      ' TEM_LINK  �����N��
        Next
    End If
    
    ' �I��
    GetModifyHistory = iCnt
End Function

