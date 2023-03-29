Attribute VB_Name = "SystemCommon"
Option Explicit
Option Base 1


Public ErNm             As Variant      'Err.Number���i�[
Public sEr              As Variant      'Err.Description���i�[
Public iRet             As Integer      '�߂�l�󂯎��

Public Const cSheetMain As String = "Main"

Private Const TNS_NAME As String = "MACSDB5A1"

'********************************************************************************************
'  ��͈˗��V�X�e�� - �V�X�e�����ʃ��W���[��
'               Copyright 2015, XXXX All Rights Reserved.
'  2015-05-14 �V�K�쐬
'********************************************************************************************

'**********************************************************************
'     Win32API�i2015/10/15�ǉ��j
'**********************************************************************
Public Declare Function SetWindowPos Lib "user32" ( _
                            ByVal hWnd As Long, _
                            ByVal hWndInsertAfter As Long, _
                            ByVal x As Long, _
                            ByVal y As Long, _
                            ByVal cx As Long, _
                            ByVal cy As Long, _
                            ByVal wFlags As Long _
                            ) As Long

Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" ( _
                            ByVal lpClassName As String, _
                            ByVal lpWindowName As String _
                            ) As Long

Public Const HWND_TOPMOST As Long = -1
Public Const SWP_NOSIZE As Long = &H1&
Public Const SWP_NOMOVE As Long = &H2&
'**********************************************************************



' ********
' �萔��`
' ********
' SQL �N�G�����s���ʊi�[�p�t�@�C����
Public Const cSqlResFilenameP As String = "TEMSqlRes#.txt"
' SQL �N�G�����s���ʊi�[�p�t�@�C�� �o�͐�f�B���N�g�����擾�p���ϐ�
Public Const cSqlResOutDirEnv As String = "WORKDIR"

'�V�[�g��(����)
Public Const cSheetSql       As String = "SQL"
Public Const cSheetSetup     As String = "SETUP"
Public Const cSheetTsbMenu   As String = "�˗���Top���j���["
Public Const cSheetTnaMenu   As String = "TNA Top���j���["
Public Const cSheetAdminMenu   As String = "�Ǘ��҃��j���["
Public Const cSheetSplash As String = "��͈˗��V�X�e��"



' ���Ə���
Public Const cTsb As String = "TSB"
Public Const cTna As String = "TNA"
Public Const cAim As String = "AIM"

' ���ꕔ��
Public Const cPg1 As String = "�iP�Z��j"
Public Const cPg2 As String = "�iP�Z��j"

' �������
Public Const cPermitTsbGeneral As String = "0"      ' TSB ���
Public Const cPermitTsbTheme As String = "1"        ' TSB �e�[�}��(�Q���A�喱)
Public Const cPermitTsbGroup As String = "2"        ' TSB �O���[�v��(�ے�)
Public Const cPermitTsbAim As String = "3"          ' TSB ��͋Z��(AIM)
Public Const cPermitTnaGeneral As String = "0"      ' TNA ���
Public Const cPermitTnaSection As String = "1"      ' TNA �ے�

' ���F�Ҏ�ʖ�
Public Const cApproveTheme As String = "�e�[�}��"
Public Const cApproveGroup As String = "�O���[�v��"
Public Const cApproveTna As String = "�㒷"

' ��Ԗ�
Public Const cStatusCreateNew As String = "(CREATE_NEW)"            '  -:(�V�K)
Public Const cStatusThmApproveWait As String = "THM_APPROVE_WAIT"   '  1:�e�[�}�����F�҂�
Public Const cStatusEstimateReqWait As String = "ESTIMATE_REQ_WAIT" '  2:���ς���˗��҂�
Public Const cStatusEstimating As String = "ESTIMATING"             '  3:�˗����ς��蒆
Public Const cStatusTnaApproveWait As String = "TNA_APPROVE_WAIT"   '  4:TNA�㒷���F�҂�
Public Const cStatusTsbApproveWait As String = "TSB_APPROVE_WAIT"   '  5:TSB�㒷���F�҂�
Public Const cStatusAimApproveWait As String = "AIM_APPROVE_WAIT"   '  6:AIM���F�҂�
Public Const cStatusReqReceptWait As String = "REQ_RECEPT_WAIT"     '  7:�˗���t�҂�
Public Const cStatusChkResultWait As String = "CHK_RESULT_WAIT"     '  8:�ώ@���ʊm�F�҂�
Public Const cStatusWorkEndWait As String = "WORK_END_WAIT"         '  9:��Ɗ����҂�
Public Const cStatusDone As String = "DONE"                         ' 10:��Ɗ���
Public Const cStatusTsbCanceled As String = "TSB_CANCELED"          ' 11:TSB�������ς�
Public Const cStatusTnaCanceled As String = "TNA_CANCELED"          ' 12:TNA�������ς�
Public Const cStatusAimCanceled As String = "AIM_CANCELED"          ' 13:AIM�������ς�

Public Const cRNewLine As Integer = 28  ' DB�o�^�����s�R�[�h�u������

' �t���O
Public Const cFlgFalse As String = "0"
Public Const cFlgTrue As String = "1"

' **********
' �񋓑̒�`
' **********

' �N�����[�U�[���
Public Enum EExecuteUserMode
    eumTsbMenu = 0          ' TSB�p���j���[
    eumTnaMenu = 1          ' TNA�p���j���[
    eumAdminMenu = 2        ' �Ǘ��җp���j���[
    eumThemeApprove = 3     ' �e�[�}�����F���
    eumSectionApprove = 4   ' �ے����F���
    eumAimApprove = 5       ' AIM���F���
    enmTnaApprove = 6       ' TNA�㒷�F�؉��
    enmOrderConfirm = 7     ' ���B�����m�F���
End Enum

' ************
' �L��ϐ���`
' ************


Public gsWorkDir    As String
Public gsFileNM     As String
Public Acol()       As String
Public Bcol()       As String

Public MAINBOOK     As String
' ## Public KenSu        As Integer             '�������ʐ�
Public psData()     As String

Public gExecuteUserMode As EExecuteUserMode ' �N�����[�U�[���

Public gbLogin As Boolean        ' ���O�C������ True=�����AFalse=���s
Public gsLoginUserId As String   ' ���O�C���������[�U�[ID
Public gsLoginUserName As String ' ���O�C���������[�U�[���O
Public gsLoginUserDiv As String  ' ���O�C���������[�U�[�̎��Ə�
Public gsLoginUserCls As String  ' ���O�C���������[�U�[�̖�E
Public gbAdminFlg As Boolean     ' �Ǘ��҃t���O True=ON�AFalse=OFF

Public previousSheetName As String ' �O��\��������ʖ�(=���j���[)

Public Function GetStatusRank(status As String) As Long
    Dim obj As Object
    
    Set obj = CreateObject("Scripting.Dictionary")
    
    Call obj.Add(cStatusCreateNew, 0)
    Call obj.Add(cStatusThmApproveWait, 1)
    Call obj.Add(cStatusEstimateReqWait, 2)
    Call obj.Add(cStatusEstimating, 3)
    Call obj.Add(cStatusTnaApproveWait, 4)
    Call obj.Add(cStatusTsbApproveWait, 5)
    Call obj.Add(cStatusAimApproveWait, 6)
    Call obj.Add(cStatusReqReceptWait, 7)
    Call obj.Add(cStatusChkResultWait, 8)
    Call obj.Add(cStatusWorkEndWait, 9)
    Call obj.Add(cStatusDone, 10)
    Call obj.Add(cStatusTsbCanceled, 11)
    Call obj.Add(cStatusTnaCanceled, 12)
    Call obj.Add(cStatusAimCanceled, 13)
    
    GetStatusRank = obj(status)
End Function

'**********************************************************************
' @(f)
' �@�\      : �ݒ�V�[�g(SETUP)����A�ݒ�l���擾����B
'
' �Ԃ�l    : Variant : �擾�����l
'
' ������    : String    strName         �擾����ݒ�l��
'
' �@�\����  :
'
' ���l      :
'
'**********************************************************************
Public Function GetSetup(strName As String) As Variant
    
    Dim wksheet As Worksheet
    Dim rngFind As Range
    On Error GoTo EH:
    
    ' �ݒ�V�[�g
    Set wksheet = ThisWorkbook.Worksheets(cSheetSetup)
    
    ' 2��ڂ̐ݒ�l���񂩂�A�����̐ݒ�l���Ɠ����l�̃Z������������B
    Set rngFind = wksheet.Columns(2).Find(What:=strName, LookAt:=xlWhole, MatchCase:=True)
    
    ' �������Z���̉E�̒l��Ԃ��B
    GetSetup = wksheet.Cells(rngFind.row, 3).Value

    Exit Function
EH:
    ' �G���[�����������ꍇ
        ' ��v���Ȃ��ꍇ
        MsgBox "�ݒ�V�[�g (" & cSheetSetup & ") ����ݒ�l " & strName & "���擾�ł��܂���B", vbCritical + vbOKOnly, "�G���["
        GetSetup = ""
End Function

'**********************************************************************
' @(f)
' �@�\      : �ݒ�V�[�g(SETUP)����A�ݒ�l�̃Z�����擾����B
'
' �Ԃ�l    : Range : �擾�����Z��(Range)
'
' ������    : String    strName         �擾����ݒ�l��
'
' �@�\����  :
'
' ���l      :
'
'**********************************************************************
Public Function GetSetupCell(strName As String) As Range
    
    Dim wksheet As Worksheet
    Dim rngFind As Range
    On Error GoTo EH:
    
    ' �ݒ�V�[�g
    Set wksheet = ThisWorkbook.Worksheets(cSheetSetup)
    
    ' 2��ڂ̐ݒ�l���񂩂�A�����̐ݒ�l���Ɠ����l�̃Z������������B
    Set rngFind = wksheet.Columns(2).Find(What:=strName, LookAt:=xlWhole, MatchCase:=True)
    
    ' �������Z���̉E�̃Z���̎Q�Ƃ�Ԃ�
    Set GetSetupCell = wksheet.Cells(rngFind.row, 3)

    Exit Function
EH:
    ' �G���[�����������ꍇ
        ' ��v���Ȃ��ꍇ
        MsgBox "�ݒ�V�[�g (" & cSheetSetup & ") ����ݒ�l " & strName & "���擾�ł��܂���B", vbCritical + vbOKOnly, "�G���["
        Set GetSetupCell = Nothing
End Function


'**********************************************************************
' @(f)
' �@�\      : �ݒ�V�[�g(SETUP)�̎w�荀�ڂ̐ݒ�l��ݒ肷��B
'
' �Ԃ�l    : Boolean : ��������(True=����/False=�G���[)
'
' ������    : String    strName         �擾����ݒ�l��
'             Variant   value           �ݒ肷��l
'
' �@�\����  :
'
' ���l      :
'
'**********************************************************************
Public Function SetSetup(strName As String, Value As Variant) As Boolean
    
    Dim wksheet As Worksheet
    Dim rngFind As Range
    
    ' �ݒ�V�[�g
    Set wksheet = ThisWorkbook.Worksheets(cSheetSetup)
    
    ' 2��ڂ̐ݒ�l���񂩂�A�����̐ݒ�l���Ɠ����l�̃Z������������B
    Set rngFind = wksheet.Columns(2).Find(What:=strName, LookAt:=xlWhole, MatchCase:=True)
    
    If rngFind <> Empty Then
        ' �������Z���̒l���X�V����B
        wksheet.Cells(rngFind.row, 3).Value = Value
    Else
        ' ��v���Ȃ��ꍇ
        MsgBox "�ݒ�V�[�g (" & cSheetSetup & ") �̐ݒ�l " & strName & "���ݒ�ł��܂���B", vbCritical + vbOKOnly, "�G���["
        
        SetSetup = False
    End If

End Function






'��*************************************************************��
' �@�\      : SQL���s ���O����
'             �o�̓e�L�X�g�t�@�C�����A�����ϐ�
' �Ԃ�l    : �o�̓t�@�C�����i�t���p�X�j
' ������    : �����ϐ��̔z�񐔁A�o�̓t�H���_�̐ݒ薼�A�o�̓t�@�C����
' �@�\����  :
' ���l      :
'��*************************************************************��
Public Function f_SqlInit(iWhere As Integer, sDirNM As String, sFileNM As String) As String
    
    MAINBOOK = ThisWorkbook.Name
    
    Erase Acol()
    Erase Bcol()
    
    ReDim Acol(iWhere)
    ReDim Bcol(iWhere)
    
    gsWorkDir = GetEnv(gcEnvFile, sDirNM)
    f_SqlInit = gsWorkDir & sFileNM
    
End Function

'��*************************************************************��
' �@�\      : SQL���s��̃e�L�X�g�t�@�C���f�[�^��ϐ��Ɋi�[
' �Ԃ�l    : �f�[�^�����i�f�[�^�j
' ������    : �i�[�ϐ��A�f�[�^�t�@�C����
' �@�\����  :
' ���l      : �i�[�ϐ�(0)�ɂ̓f�[�^�͓��炸�A(1)����f�[�^�Z�b�g����
'��*************************************************************��
Public Function f_GetData(ByRef sData() As String, sFileNM As String) As Long
    Dim iFp             As Integer
    Dim iColumn(100)    As Integer
    Dim getData(100)    As String
    Dim iAns            As Integer
    Dim lCnt            As Long
    Dim iCol            As Integer
    
    Erase sData()
    
    lCnt = 0
    iFp = FreeFile(1)
    Open sFileNM For Input As #iFp
    iAns = Check_Column(iFp, iColumn())
    
    Do While Get_Column(iFp, iColumn(), getData())
        lCnt = lCnt + 1
        ReDim Preserve sData(iAns, lCnt)
        For iCol = 1 To iAns
            sData(iCol, lCnt) = getData(iCol)
        Next iCol
        
    Loop
    Close #iFp
    
    f_GetData = lCnt
    
End Function






'-------���V�[�g�N���A��-------
'�@�����@�@�@�@�@�FcName  �ΏۃV�[�g���@iStRow �J�nRow�@iStcol �J�nCol
'�@�����i�ȗ��j�FiMode�@���[�h�i0 Delete�@1 ClearContents�j
'�@�@�@�@�@�@�@�@�FsUpLeft�@Delete���̃V�t�g���� �ȗ�����Up,Up�ȊO�̉������炪�����Ă��Left
'�@�@�@�@�@�@�@�@�FiEnRow �I��Row       iEnCol �I��Col
'�@�Ԃ�l�@�@�@�@�FInteger�^�@0�@���s�@ 1�@�����@�@-1�@�����f�[�^��

Public Function f_SheetClear(cName As String, iStRow As Integer, iStCol As Integer, Optional iMode As Integer = 0, _
                             Optional sUpLeft As String = "Up", Optional lEnRow As Long = 0, Optional lEnCol As Long = 0) As Integer

    On Error GoTo Errtrap
    
    With ThisWorkbook.Worksheets(cName)
    
        '�����邱�Ƃ�����̂ŕی����
        .Unprotect
        
        '�����ȗ����͎n�_�ȍ~�J�E���g
        If lEnRow <= 0 Then
            lEnRow = .Cells(Rows.Count, iStCol).End(xlUp).row
        End If
        
        If lEnCol <= 0 Then
            lEnCol = .Cells(iStRow, Columns.Count).End(xlToLeft).Column
        End If
        
        If iStRow > lEnRow Or iStCol > lEnCol Then
            f_SheetClear = -1 '�����f�[�^��
            Exit Function
        End If
        
        'sUpLeft �ȗ�����Up�ɂ���
        If sUpLeft = "Up" Then
            iRet = xlUp
        Else
            iRet = xlToLeft
        End If
        
        If iMode = 1 Then
            .Range(.Cells(iStRow, iStCol).Address(False, False), _
                   .Cells(lEnRow, lEnCol).Address(False, False)).ClearContents
        Else
            .Range(.Cells(iStRow, iStCol).Address(False, False), _
                   .Cells(lEnRow, lEnCol).Address(False, False)).Delete Shift:=iRet
        End If
        
        f_SheetClear = 1 '����
        
        iRet = .UsedRange.Resize(1, 1).Rows.Count
        
        Exit Function
        
Errtrap:
        
    f_SheetClear = 0 '�G���[�i���s�j
        
    End With
End Function


'-------��Null�l�𕶎���ɕϊ���-------
'�@�����@�FTarget �Ώۂ�Value�@RepStr �u����̕�����i�ȗ��j
'�@�Ԃ�l�FString�^�@�u���ォ���Ƃ̕������Ԃ�

Public Function NVL(Target As Variant, Optional RepStr As String = "") As String

    If IsNull(Target) = True Then
        NVL = RepStr
    Else
        NVL = CStr(Target)
    End If

End Function


'-------���r����������-------
'�@�����@�@�@�@�@�FcName �@�@ �ΏۃV�[�g��
'�@�@�@�@�@�@�@�@�FiBArea     ���������ꏊ�@(0 �S�́@1 �O�g�̂݁@2 �����c�@3�������@4 �����S��)
'�@�@�@�@�@�@�@�@�FiLpat    �@���̎�ށi0 ���Ȃ��@1 Hairline�@2 Thin�@3 Medium�j
'�@�@�@�@�@�@�@�@�FiStRow     �n�_Row�@�@�@iStCol�@�@�n�_Col
'�@�����i�ȗ��j�FlEnRow�@   �I�_Row�@�@�@lEnCol�@�@�I�_Col�@�@�@lColor�@�@���̐F�i�ȗ������j
'�@�Ԃ�l�@�@�@�@�FInteger�^�@0�@���s�@ 1�@����

Public Function BorderLiner(cName As String, iBArea As Integer, iLpat As Integer, iStRow As Integer, iStCol As Integer, _
                            Optional lEnRow As Long = 0, Optional lEnCol As Long = 0, Optional lColor As Long = 0)
    Dim iLine   As Integer
    Dim oRange  As Range
    
    On Error GoTo Errtrap
    
    With ThisWorkbook.Worksheets(cName)
    
    '�����ȗ����͎n�_�ȍ~�J�E���g
    If lEnRow = 0 Then
        lEnRow = .Cells(Rows.Count, iStCol).End(xlUp).row
    End If
    
    If lEnCol = 0 Then
        lEnCol = .Cells(iStRow, Columns.Count).End(xlToLeft).Column
    End If
    
    '���̐F���m��i�ȗ����͍��j
    If lColor = 0 Then
        lColor = RGB(0, 0, 0)
    End If
    
    
    '���̎�ނ��m��
    Select Case iLpat
        Case Is = 0
            iLine = xlNone
        Case Is = 1
            iLine = xlHairline
        Case Is = 2
            iLine = xlThin
        Case Is = 3
            iLine = xlMedium
        Case Else
            iLine = xlThin
    End Select
    
    '�����͈͂��m��
    Set oRange = .Range(.Cells(iStRow, iStCol).Address(False, False), _
                        .Cells(lEnRow, lEnCol).Address(False, False))
    
    '�r��
    Select Case iBArea
        Case Is = 0
            oRange.Borders.Color = lColor
            oRange.Borders.Weight = iLine
        Case Is = 1
            oRange.BorderAround Weight:=iLine, Color:=lColor
        Case Is = 2
            oRange.Borders.Item(xlInsideVertical).Color = lColor
            oRange.Borders.Item(xlInsideVertical).Weight = iLine
        Case Is = 3
            oRange.Borders.Item(xlInsideHorizontal).Color = lColor
            oRange.Borders.Item(xlInsideHorizontal).Weight = iLine
        Case Is = 4
            oRange.Borders.Item(xlInsideVertical).Color = lColor
            oRange.Borders.Item(xlInsideHorizontal).Color = lColor
            oRange.Borders.Item(xlInsideVertical).Weight = iLine
            oRange.Borders.Item(xlInsideHorizontal).Weight = iLine
    End Select
    
    End With
    
    BorderLiner = 1
    Exit Function
    
Errtrap:
        BorderLiner = 0
End Function

'-------���X�e�[�^�X���u����-------

Public Function Status2Name(sStatus As String) As String
        Select Case sStatus
            Case Is = ""
                Status2Name = "�����"
            Case Is = "REQ_CANCEL"
                Status2Name = "�˗�������"
            Case Is = "REQ_ACCEPT"
                Status2Name = "�˗����"
            Case Is = "SEM_WORKING"
                Status2Name = "SEM��ƒ�"
            Case Is = "SEM_CHECKING"
                Status2Name = "SEM�ώ@�m�F�҂�"
            Case Is = "SEM_MEAS_WAIT"
                Status2Name = "�����҂�"
            Case Is = "SEM_MEASURING"
                Status2Name = "������"
            Case Is = "SEM_REBUILD_WAIT"
                Status2Name = "�����҂�"
            Case Is = "SEM_REBUILD_CHECKING"
                Status2Name = "�����m�F�҂�"
            Case Is = "SEM_RECOVER_WAIT"
                Status2Name = "�E�F�n����҂�"
            Case Is = "SEM_DROP_PREPARE"
                Status2Name = "�E�F�n�p������҂�"
            Case Is = "SEM_DROP_WAIT"
                Status2Name = "�E�F�n�p���҂�"
            Case Is = "COMPLETE"
                Status2Name = "����"
            Case Else
                Status2Name = ""
        End Select
End Function

'��*************************************************************��
' �@�\      : �e�L�X�g�{�b�N�X ��/���F�؂�ւ�
' �Ԃ�l    :
' ������    :
' �@�\����  :
' ���l      : ���F�FH0099FFFF�i=99FFFF�j
'��*************************************************************��
Public Sub YellowWhite(myObject As Object)
    Application.EnableEvents = False
    If myObject.Text <> "" Then
        myObject.BackColor = RGB(255, 255, 255)
    Else
        myObject.BackColor = RGB(255, 255, 153)
    End If
    Application.EnableEvents = True
End Sub

'��*************************************************************��
' �@�\      : �˗��ҏ��擾
' �Ԃ�l    : �f�[�^�����i�f�[�^�j
' ������    : �˗�No
' �@�\����  :
' ���l      :
'��*************************************************************��
Public Function f_GetClientInfo(sSemReqno As String) As Long
    Dim lCnt    As Long
    
    '���r�p�k���s�O�ɕK�����s���邱��
    gsFileNM = f_SqlInit(5, "WORKDIR", "SqlData.txt")
    
    Bcol(1) = "\SemReqno"
    Acol(1) = "'" & sSemReqno & "'"
    
    '���r�p�k���s
    If CallAdoSql("SQL", 20, gsFileNM, Bcol(), Acol(), "", "") = False Then
        MsgBox "�G���[�������������߁A���������f����܂����B" & Chr(10) & Chr(10) & _
               ErNm & "�F" & sEr, vbCritical + vbOKOnly, "�G���["
        Set ErNm = Nothing
        Set sEr = Nothing
        Exit Function
    End If
    
    '���擾�f�[�^
    lCnt = f_GetData(psData(), gsFileNM)
    
    f_GetClientInfo = lCnt
End Function

'��*************************************************************��
' �@�\      : ���[�U���擾
' �Ԃ�l    : �f�[�^�����i�f�[�^�j
' ������    : ���ꃆ�[�U�[ID
' �@�\����  :
' ���l      :
'��*************************************************************��
Public Function f_GetUsrName(sSemUsrid As String) As Long
    Dim lCnt    As Long
    
    '���r�p�k���s�O�ɕK�����s���邱��
    gsFileNM = f_SqlInit(5, "WORKDIR", "SqlData.txt")
    
    Bcol(1) = "\SemUsrid"
    Acol(1) = "'" & sSemUsrid & "'"
    
    '���r�p�k���s
    If CallAdoSql("SQL", 21, gsFileNM, Bcol(), Acol(), "", "") = False Then
        MsgBox "�G���[�������������߁A���������f����܂����B" & Chr(10) & Chr(10) & _
               ErNm & "�F" & sEr, vbCritical + vbOKOnly, "�G���["
        Set ErNm = Nothing
        Set sEr = Nothing
        Exit Function
    End If
    
    '���擾�f�[�^
    lCnt = f_GetData(psData(), gsFileNM)
    
    f_GetUsrName = lCnt
End Function

'��*************************************************************��
' �@�\      : ���X�e�[�^�X�擾
' �Ԃ�l    : �f�[�^�����i�f�[�^�j
' ������    : �˗�No
' �@�\����  :
' ���l      :
'��*************************************************************��
Public Function f_GetStatus(sSemReqno As String) As String
    Dim lCnt    As Long
    
    '���r�p�k���s�O�ɕK�����s���邱��
    gsFileNM = f_SqlInit(5, "WORKDIR", "SqlData.txt")
    
    Bcol(1) = "\SemReqno"
    Acol(1) = "'" & sSemReqno & "'"
    
    '���r�p�k���s
    If CallAdoSql("SQL", 19, gsFileNM, Bcol(), Acol(), "", "") = False Then
        MsgBox "�G���[�������������߁A���������f����܂����B" & Chr(10) & Chr(10) & _
               ErNm & "�F" & sEr, vbCritical + vbOKOnly, "�G���["
        Set ErNm = Nothing
        Set sEr = Nothing
        Exit Function
    End If
    
    '���擾�f�[�^
    lCnt = f_GetData(psData(), gsFileNM)
    
    f_GetStatus = lCnt
End Function

'��*************************************************************��
' �@�\      : WF�̃X�e�[�^�X��1�i�߂�
' �Ԃ�l    : 1:����I���A-1:�G���[
' ������    : �˗�No�A�i���󋵁A���͎҃��[�UID�A���͎҃��[�UID2�A
' ������    : �o�^�����A�X�V�����A��蒼���񐔁A��߂蔭���t���O�A�R�����g
' �@�\����  :
' ���l      :
'��*************************************************************��
Public Function f_ProcApprovalWf(sSemReqnoVal As String, sSemStatusVal As String, _
sSemInpusr1Val As String, sSemInpusr2Val As String, sSemRegdatVal As Date, _
sSemUpddatVal As Date, iSemRepeatVal As Integer, iSemDropflgVal As Integer, _
sSemCommentVal As String) As Long
    Dim lCnt    As Long
    
    '���r�p�k���s�O�ɕK�����s���邱��
    gsFileNM = f_SqlInit(10, "WORKDIR", "SqlData.txt")
    
    Bcol(1) = "\SEM_REQNO"
    Acol(1) = "'" & sSemReqnoVal & "'"
    Bcol(2) = "\SEM_STATUS"
    Acol(2) = "'" & sSemStatusVal & "'"
    Bcol(3) = "\SEM_INPUSR1"
    Acol(3) = "'" & sSemInpusr1Val & "'"
    Bcol(4) = "\SEM_INPUSR2"
    Acol(4) = "'" & sSemInpusr2Val & "'"
    Bcol(5) = "\SEM_REGDAT"
    Acol(5) = "'" & sSemRegdatVal & "'"
    Bcol(6) = "\SEM_UPDDAT"
    Acol(6) = "'" & sSemUpddatVal & "'"
    Bcol(7) = "\SEM_REPEAT"
    Acol(7) = iSemRepeatVal
    Bcol(8) = "\SEM_DROPFLG"
    Acol(8) = iSemDropflgVal
    Bcol(9) = "\SEM_COMMENT"
    Acol(9) = "'" & sSemCommentVal & "'"
    
    '���r�p�k���s
    If CallAdoSql("SQL", 22, gsFileNM, Bcol(), Acol(), "", "") = False Then
        MsgBox "�G���[�������������߁A���������f����܂����B" & Chr(10) & Chr(10) & _
               ErNm & "�F" & sEr, vbCritical + vbOKOnly, "�G���["
        Set ErNm = Nothing
        Set sEr = Nothing
        Exit Function
    End If
    
    '���擾�f�[�^
    lCnt = f_GetData(psData(), gsFileNM)
    
    f_ProcApprovalWf = lCnt
End Function

'��*************************************************************��
' �@�\      : WF�̃X�e�[�^�X��1�߂�
' �Ԃ�l    : 1:����I���A-1:�G���[
' ������    : �˗�No
' �@�\����  :
' ���l      :
'��*************************************************************��
Public Function f_ProcDropWf(sSemReqnoVal As String) As Long
    Dim lCnt    As Long
    
    '���r�p�k���s�O�ɕK�����s���邱��
    gsFileNM = f_SqlInit(5, "WORKDIR", "SqlData.txt")
    
    Bcol(1) = "\SEM_REQNO"
    Acol(1) = "'" & sSemReqnoVal & "'"
    
    '���r�p�k���s
    If CallAdoSql("SQL", 23, gsFileNM, Bcol(), Acol(), "", "") = False Then
        MsgBox "�G���[�������������߁A���������f����܂����B" & Chr(10) & Chr(10) & _
               ErNm & "�F" & sEr, vbCritical + vbOKOnly, "�G���["
        Set ErNm = Nothing
        Set sEr = Nothing
        Exit Function
    End If
    
    '���擾�f�[�^
    lCnt = f_GetData(psData(), gsFileNM)
    
    f_ProcDropWf = lCnt
End Function

'��*************************************************************��
' �@�\      : �˗��ԍ�����f�[�^�擾
' �Ԃ�l    : 0�F����A-1�F�G���[
' ������    : �˗�No
' �@�\����  :
' ���l      :
'��*************************************************************��
Public Function f_GetSemReqtblData(sSemReqno As String) As String
    Dim lCnt    As Long
    
    '���r�p�k���s�O�ɕK�����s���邱��
    gsFileNM = f_SqlInit(5, "WORKDIR", "SqlData.txt")
    
    Bcol(1) = "\SemReqno"
    Acol(1) = "'" & sSemReqno & "'"
    
    '���r�p�k���s
    If CallAdoSql("SQL", 25, gsFileNM, Bcol(), Acol(), "", "") = False Then
        MsgBox "�G���[�������������߁A���������f����܂����B" & Chr(10) & Chr(10) & _
               ErNm & "�F" & sEr, vbCritical + vbOKOnly, "�G���["
        Set ErNm = Nothing
        Set sEr = Nothing
        f_GetSemReqtblData = -1
        Exit Function
    End If
    
    '���擾�f�[�^
    lCnt = f_GetData(psData(), gsFileNM)
    
    f_GetSemReqtblData = 1
End Function



'��********************************************************************************��
' �@�\      : �p�����Ǝw�蕶���`�F�b�N
' �Ԃ�l    : 1�F�`�F�b�NOK�A-1�F�G���[�L
' ������    : sStr     �c String�^    �`�F�b�N���镶����
'�@�@�@�@�@ : fAlpha   �c Boolean�^   �p����ʉ߂����邩
'           : fNumeric �c Boolean�^�@ ������ʉ߂����邩
'�@�@�@�@�@ : fSymbol  �c Boolean�^   �w�蕶����ʉ߂����邩
'�@�@�@�@�@ : sSymbol  �c String�^  �@�w�肷�镶��
'           : sDel     �c String�^�@  �f���~�^
' �@�\����  :
' ���l      : ���Ɏw�����Ȃ��ꍇ�A�ʉ߂����镶���̓n�C�t���A�f���~�^�̓A���p�T���h
'��********************************************************************************��
Public Function f_Almerics(sStr As String, fAlpha As Boolean, fNumeric As Boolean _
                          , fSymbol As Boolean, Optional sSymbol As String = "-" _
                                              , Optional sDel As String = "&") As Integer

    Dim i     As Integer
    Dim vSyms As Variant
    
    '���w�肷��L�����f���~�^�ŋ�؂�A�ϐ��֊i�[
    vSyms = Split(sSymbol, sDel)
    
    '���������Ԃ�̃��[�v
    For i = 1 To Len(sStr)
    
        '���܂��̓G���[�l�Ƃ��ĔF��
        f_Almerics = -1
    
        '���p�����ǂ����`�F�b�N
        If fAlpha = True Then
            If Mid(LCase(sStr), i, 1) Like "[a-z]" Then
                f_Almerics = 1
            End If
        End If
        
        '���������ǂ����`�F�b�N
        If fNumeric = True Then
            If Mid(sStr, i, 1) Like "[0-9]" Then
                f_Almerics = 1
            End If
        End If
        
        '���w��L�����ǂ����`�F�b�N
        If fSymbol = True Then
            If UBound(Filter(vSyms, Mid(sStr, i, 1))) > -1 Then
                f_Almerics = 1
            End If
        End If
        
        '���܂��G���[��������I���A�߂�l -1
        If f_Almerics = -1 Then
            Exit Function
        End If

    Next i

End Function

'��********************************************************************************��
' �@�\      : ActiveX�R���g���[���̗L����������H
' �Ԃ�l    : 1�F�L���A-1�F����
' �@�\����  :
' ���l      : �킴�ƃG���[���N����
'��********************************************************************************��
Public Function ActiveX_Chk() As Integer

    On Error GoTo Errtrap

    With ThisWorkbook.Worksheets(cSheetMain)
        .Visible = False
    End With
    
    ActiveX_Chk = 1
    
    Exit Function
    
Errtrap:
    ActiveX_Chk = -1

End Function

'-------���o�^������̃`�F�b�N��-------

Public Function CmntChk(ChkVal As String) As Integer
        CmntChk = 0

    If InStr(ChkVal, """") > 0 Then
        CmntChk = -1
        Exit Function
        
    ElseIf InStr(ChkVal, "'") > 0 Then
        CmntChk = -1
        Exit Function
        
    ElseIf InStr(ChkVal, ".") > 0 Then
        CmntChk = -1
        Exit Function
        
'    ElseIf InStr(ChkVal, ",") > 0 Then
'        CmntChk = -1
'        Exit Function
        
    ElseIf InStr(ChkVal, "&") > 0 Then
        CmntChk = -1
        Exit Function

    ElseIf InStr(ChkVal, "%") > 0 Then
        CmntChk = -1
        Exit Function
        
'    ElseIf InStr(ChkVal, "/") > 0 Then
'        CmntChk = -1
'        Exit Function
        
    ElseIf InStr(ChkVal, ":") > 0 Then
        CmntChk = -1
        Exit Function
        
    ElseIf InStr(ChkVal, ";") > 0 Then
        CmntChk = -1
        Exit Function
        
    ElseIf InStr(ChkVal, "<") > 0 Then
        CmntChk = -1
        Exit Function
        
    ElseIf InStr(ChkVal, ">") > 0 Then
        CmntChk = -1
        Exit Function
        
    ElseIf InStr(ChkVal, "*") > 0 Then
        CmntChk = -1
        Exit Function
        
    End If

    CmntChk = 1
    
End Function

'��********************************************************************************��
' �@�\      : �w��̃V�[�g�ȊO��\���ɂ���B
' ������    : sSheet    �c String�^    �\�����郏�[�N�V�[�g��
' �Ԃ�l    : �w��V�[�g�ւ̎Q��(Worksheet)
' �@�\����  :
' ���l      :
'��********************************************************************************��
Public Function ViewWorkSheet(sSheet As String) As Worksheet

    Dim wksht As Worksheet
    
    ' �����Ɠ����V�[�g���̏ꍇ�́A�\������B
    Set wksht = ThisWorkbook.Worksheets(sSheet)
    wksht.Visible = xlSheetVisible
    Set ViewWorkSheet = wksht
            
    For Each wksht In ThisWorkbook.Worksheets
        If sSheet <> wksht.Name Then
            ' ��L�ȊO�͔�\���Ƃ���B
            wksht.Visible = xlSheetHidden
        End If
    Next

End Function


'��********************************************************************************��
' �@�\      : �󂫔z�񔻒�
' ������    : target        �c Variant�^    �Ώۂ̔z��
' �Ԃ�l    : Boolean : True=�L���Ȕz��AFalse=�󂫂̔z��(�z��ȊO���܂�)
' �@�\����  :
' ���l      :
'��********************************************************************************��
Public Function IsEnableArray(Target() As Variant) As Boolean
    On Error GoTo EH:
    
    ' �z��v�f�ɃA�N�Z�X�����݂�B
    If Abs(UBound(Target) - LBound(Target)) > 0 Then
        IsEnableArray = True
    Else
        IsEnableArray = False
    End If
    
    Exit Function
EH:
    ' ��O�����������ꍇ�A�󂫔z��Ƃ���B
    IsEnableArray = False
    
End Function

'��********************************************************************************��
' �@�\      : ���t�t�H�[�}�b�g����
' ������    : target        �c String�^    �Ώۂ̕�����
'           : isPaat        �c Boolean�^   �V�X�e�������ȑO�̓��t�������邩�w�肷��t���O�B�ȗ���(�f�t�H���g=[False])
'           : sMsg          �c String�^    �s���ȕ�������w�肳�ꂽ�ꍇ�ɔ��藝�R�̃��b�Z�[�W�����͂����B�ȗ���
' �Ԃ�l    : String : �␳��̓��t������
' �@�\����  : �����̕����񂪓��t�̏����ł��邩����B�␳�\�ȏꍇ�͕␳����B(����
'           : �̕�����𒼐ڕύX����B)
' ���l      :
'��********************************************************************************��
Public Function CheckDateFormat(Target As String, Optional isPaat As Boolean = False, Optional ByRef sMsg As Variant) As String
    On Error GoTo EH:
    
    If IsMissing(sMsg) = False Then
        sMsg = Empty
    End If
    
    ' ���͂��ꂽ���������t�ɕϊ����Ă݂�B
    CheckDateFormat = Format(CDate(Target), "YYYY/MM/DD")
    
    If isPaat = False And CheckDateFormat < Date Then
        ' �ߋ����t���֎~�ŕϊ����ꂽ���t���V�X�e�����t���ߋ��̓��t�̏ꍇ�̓V�X�e�����t��Ԃ�
        CheckDateFormat = Format(Date, "YYYY/MM/DD")
        If IsMissing(sMsg) = False Then
            sMsg = "�{���ȍ~�̓��t����͂��Ă��������B" ' �V�X�e�����t��Ԃ����R�Ƃ��ă��b�Z�[�W��ݒ�
        End If
    End If

    ' �G���[�ɂȂ�Ȃ������ꍇ��OK�Ƃ���B
    On Error GoTo 0
    
    Exit Function
EH:
    ' ��O�����������ꍇ�̓G���[�Ƃ���B
    On Error GoTo 0
    If IsMissing(sMsg) = False Then
        sMsg = "�L���ȓ��t�ł͂���܂���B"
    End If
    CheckDateFormat = Empty
    
End Function

'��********************************************************************************��
' �@�\      : �t�@�C�����̂ݎ擾
' ������    : target        �c String�^    �Ώۂ̃t�@�C�� �p�X������
' �Ԃ�l    : String : �t�@�C����
' �@�\����  : �����̃t�@�C�� �p�X����t�@�C�����̕����݂̂��擾���Ԃ��B
' ���l      :
'��********************************************************************************��
Public Function GetFilename(Target As String)
    ' ������̏I�[����ŏ��Ɍ��ꂽ"\"�܂ł̕������Ԃ��B
    '(���݂��Ȃ��ꍇ�́A�S�̂�Ԃ��B)
    Dim pos As Long
    
    pos = InStrRev(Target, "\")
    If pos = 0 Then
        GetFilename = Target
    Else
        GetFilename = Mid(Target, pos + 1)
    End If
End Function

'��********************************************************************************��
' �@�\      : �f�B���N�g���������̂ݎ擾
' ������    : target        �c String�^    �Ώۂ̃t�@�C�� �p�X������
' �Ԃ�l    : String : �t�@�C����
' �@�\����  : �����̃t�@�C�� �p�X����f�B���N�g�����̕����݂̂��擾���Ԃ��B
' ���l      :
'��********************************************************************************��
Public Function GetDirectoryName(Target As String)
    ' ������̏I�[����ŏ��Ɍ��ꂽ"\"�܂ł̕������Ԃ��B
    '(���݂��Ȃ��ꍇ�́A�S�̂�Ԃ��B)
    Dim pos As Long
    
    pos = InStrRev(Target, "\")
    If pos = 0 Then
        GetDirectoryName = ""
    Else
        GetDirectoryName = Left(Target, pos)
    End If
End Function

'��********************************************************************************��
' �@�\      : �V�X�e���I��
' ������    : �Ȃ�
' �Ԃ�l    : �Ȃ�
' �@�\����  : ���[�U�[�Ɋm�F��A�V�X�e�����I������B(���[�N�u�b�N�����)
' ���l      :
'��********************************************************************************��
Public Sub CloseSystem()
    Dim res As Integer
    
    res = MsgBox("TNA��͈˗��V�X�e�����I�����܂���?", vbQuestion + vbYesNo + vbDefaultButton2, _
        "�m�F")
    If res = vbYes Then
        ' [�͂�]���I�����ꂽ�ꍇ�V�X�e�����I������B
        If Application.Workbooks.Count > 1 Then
            ' ���Ƀ��[�N�u�b�N���J����Ă���ꍇ�́A�t�@�C�������B
            ThisWorkbook.Close False        ' �t�@�C���ۑ��Ȃ��ŕ���B
        Else
            ' �ق��Ƀ��[�N�u�b�N���J���Ă��Ȃ��ꍇ�́AExcel �A�v���P�[�V�������I������B
            ThisWorkbook.Saved = True
            Application.Quit
        End If
    End If

End Sub

'��********************************************************************************��
' �@�\      : SQL�N�G�����G�X�P�[�v�����u��
' ������    : target        �c String�^ �Ώۂ̕�����
' �Ԃ�l    : String : �G�X�P�[�v������u����������̕�����
' �@�\����  : SQL�N�G�������ŕ�����������l(�u'�v(�V���O�� �N�E�H�[�g�ň͂܂ꂽ)�ɐݒ�
'           : ���镶����p�ɃG�X�P�[�v�����ɒu��������B
' ���l      : �ȉ��̕������u��������B
'           : Tab (�^�u)         �� CHR(9)
'           : CR  (�L�����b�W ���^�[��) �� CHR(13)
'           : LF  (�s����)              �� CHR(10)
'           : ' (�V���O�� �N�E�H�[�g)   �� CHR(39)
'��********************************************************************************��
Public Function ReplaceSqlEsc(sSrc As String) As String
    
    ' SQL�N�G���� �G�X�P�[�v�����K�v�̂��镶��
    Static caEscapedChars As Variant
    
    Dim iPos As Long
    Dim iLen As Long        ' �Ώە�����̑S�̂̒���
    Dim sCChr As String
    Dim v As Variant
    Dim iFoundPos As Long    ' �������G�X�P�[�v����镶���̈ʒu
    Dim cFound As String    ' �������G�X�P�[�v����镶��
    Dim iWPos As Long
    
    Dim sOut As String
    sOut = ""
    
    ' �G�X�P�[�v���K�v�ȕ����̔z���p�ӂ���B(����̂�)
    If IsArray(caEscapedChars) <> True Then
        caEscapedChars = Array(Chr(9), Chr(13), Chr(10), Chr(39))
    End If
    
    
    ' �Ώۂ̑S�̂̕��������擾����B
    iLen = Len(sSrc)
    
    ' �Ώۂ̐擪��������G�X�P�[�v����镶���̌������J�n����B
    iPos = 1
    
    ' �Ώۂ̕��������J��Ԃ��B
    Do While iPos <= iLen
            
        ' ���݂̈ʒu��������Ƃ��߂��G�X�P�[�v�K�v�ȕ�������������B
        iFoundPos = iLen
        cFound = Chr(0)
        
        For Each v In caEscapedChars
            iWPos = InStr(iPos, sSrc, CStr(v))
            If iWPos > 0 Then
                ' �������ʒu���A���̃G�X�P�[�v������菬�����ꍇ�A�ێ��p�ϐ�������������B
                If iFoundPos > iPos Then
                    iFoundPos = iWPos
                    cFound = CStr(v)
                End If
            End If
        Next
        
        ' �G�X�P�[�v�������������ꍇ�A�u��������B
        If Asc(cFound) <> 0 Then
            
            ' �擪�ȊO�Ō������G�X�P�[�v�����̑O�̕����܂ŏo�͂���B
            sOut = sOut & Mid(sSrc, iPos, iFoundPos - iPos) + "' "
            
            ' �G�X�P�[�v�������������J��Ԃ��B
            Do
                ' �G�X�P�[�v�ɒu��������B
                sOut = sOut & "|| CHR(" & CStr(AscB(cFound)) & ") "
                
                ' ���̕�����
                iPos = iFoundPos + 1
                
                ' ���̕������G�X�P�[�v�K�v�ȕ����ł��邩���肷��B
                sCChr = Mid(sSrc, iPos, 1)
                cFound = Chr(0)
                For Each v In caEscapedChars
                    If sCChr = CStr(v) Then
                        ' �G�X�P�[�v�K�v�ȕ����ƈ�v�����ꍇ�A
                        cFound = sCChr
                        iFoundPos = iPos
                        Exit For
                    End If
                Next
                
                ' �G�X�P�[�v�K�v�ȕ����ł͂Ȃ������ꍇ
                If Asc(cFound) = 0 Then
                    ' ������ɖ߂��āA�Ō�̕������o�͂���B
                    sOut = sOut & "|| '" & sCChr
                    iPos = iPos + 1
                    Exit Do         ' ���[�v�𔲂��āA���̕�����
                End If
            Loop
           
        Else
            ' ���݂̈ʒu����G�X�P�[�v�Ώۂ̕�����������Ȃ������ꍇ�́A�c��̕�����
            ' �o�͂��ďI������B
            sOut = sOut & Mid(sSrc, iPos)
            iPos = iLen + 1
        End If
    
    
        ' ���̕�����
    Loop
    
    ReplaceSqlEsc = sOut

End Function


'��********************************************************************************��
' �@�\      : DB������o�̓T�C�Y�擾
' ������    : string sTarget : �Ώۂ̕�����
' �Ԃ�l    : Long DB�ɏo�͂���T�C�Y(�P��: �o�C�g)
' �@�\����  : �����̕������DB�ɏo�͂������̃o�C�g����Ԃ��B
' ���l      : 20160406 DB�T�[�o�̕����R�[�h��UTF8�ɕύX�̂��ߕ����R�[�h�؂�ւ��\�ɕύX
'             (�p�~)DB�ɂ� Microsoft JIS (�V�t�gJIS) �R�[�h�ŏo�͂���B
'��********************************************************************************��
Function GetLebDb(sTarget As String) As Long
    Dim charSet As String
    
    GetLebDb = 0
    If Len(sTarget) > 0 Then
        charSet = GetSetupCell("ORACLE�����R�[�h")
        
        If charSet = "UTF8" Then
            Dim UTF8 As Object
            Set UTF8 = CreateObject("System.Text.UTF8Encoding")
            GetLebDb = UTF8.GetByteCount_2(sTarget)
        Else
            GetLebDb = LenB(StrConv(sTarget, vbFromUnicode))
        End If
    End If
End Function

'��********************************************************************************��
' �@�\      : PDF�t�@�C���`�F�b�N
' ������    : string sTarget : �`�F�b�N�Ώۂ̃t�@�C�� �p�X������
' �Ԃ�l    : Boolean : True=OK(PDF�t�@�C��)�AFalse=�G���[(PDF�t�@�C���ȊO)
' �@�\����  : �����Ŏw�肵���t�@�C�� �p�X�̃t�@�C����PDF�t�@�C���ł��邩�`�F�b�N����B
' ���l      :
'��********************************************************************************��
Public Function CheckPDF(sTarget As String) As Boolean

    On Error GoTo EH:
    
    Dim iFid As Integer
    Dim sRead As String
    iFid = FreeFile
    
    ' �Ώۂ̃t�@�C�����J��
    Open sTarget For Binary As iFid
    
    ' �t�@�C���̐擪���� 6 �o�C�g���̃f�[�^��ǂ݂��ށB
    sRead = String(6, " ")
    Get iFid, , sRead
    
    ' �t�@�C�������B
    Close iFid
    
    ' �ǂݍ��񂾓��e��PDF�̃w�b�_�[�ƈ�v���邩���肷��B
    ' "%PDF-?" ? �͐��l
    If (sRead Like "%PDF-#") = True Then
        ' PDF�̃p�^�[���ƈ�v�����ꍇ�B
        CheckPDF = True
    Else
        ' ��v���Ȃ������ꍇ�̓G���[�Ƃ���B
        CheckPDF = False
    End If
    
    Exit Function
EH:
    ' �t�@�C���Ǎ��Ɏ��s�����ꍇ���`�F�b�N �G���[�Ƃ���B
    CheckPDF = False
    
End Function

'��********************************************************************************��
' �@�\      : �őO�ʕ\��
' ������    : string sCaption : ��ʃL���v�V����
' �Ԃ�l    : Boolean : �Ȃ�
' �@�\����  : �����Ŏw�肵����ʂ��őO�ʕ\���ɂ���
' ���l      :
'��********************************************************************************��
Public Sub SetForeground(sCaption As String)

    Dim hWnd As Long
    hWnd = FindWindow(vbNullString, sCaption)
    Call SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

End Sub

'��********************************************************************************��
' �@�\      : Oracle Client Version �m�F
' ������    : �Ȃ�
' �Ԃ�l    : Boolean : True=OK(9����11)�AFalse=�G���[(9����11�ȊO)
' �@�\����  : Oracle Client Version ��9����11���`�F�b�N����
' ���l      :
'��********************************************************************************��
Function chkOracleClientVer() As Boolean
    Dim WSH, wExec, sCmd As String, Result As String
    Dim ans As Boolean
    Dim tmp() As String
    Dim strPath As String
    Set WSH = CreateObject("WScript.Shell")
    
    sCmd = "tnsping " & TNS_NAME
    Set wExec = WSH.Exec("%ComSpec% /c " & sCmd)
    Do While wExec.status = 0
        DoEvents
    Loop
    Result = wExec.StdOut.ReadAll
    Set wExec = Nothing
    Set WSH = Nothing
    tmp = Split(Result, vbCrLf)
    
    Dim ver As String
    Dim i  As Long
    If Left(UCase(tmp(1)), 16) = "TNS PING UTILITY" Then
        strPath = tmp(6)
        tmp = Split(tmp(1), " ")
        For i = LBound(tmp) To UBound(tmp)
            If tmp(i) = "Version" Then
                ver = tmp(i + 1)
                Exit For
            End If
        Next i
        If InStr(1, strPath, "network") > 0 Then
            strPath = Left(strPath, InStr(1, strPath, "network") - 1)
        End If
        
    Else
        chkOracleClientVer = False
        Exit Function
    End If
    If Len(ver) >= 3 Then
        If Left(ver, 2) = "9." Or Left(ver, 3) = "11." Then
        
'
            Dim environmentString As String
            Dim j As Long
            Dim spStr() As String
            
            i = 1
            Do
                environmentString = Environ(i)
                If (Left(UCase(environmentString), 5) = "PATH=") Then
                    spStr = Split(environmentString, ";")
'                    If InStr(1, spStr(0), "C:\Oralcs6i\jdk\bin") = 0 And UCase(Left(strPath, 2)) = "H:" Then
'                        chkOracleClientVer = True
'                        Exit Function
'                    End If
                    
                    For j = LBound(spStr) To UBound(spStr)
                        If spStr(j) = "C:\Oralcs6i\jdk\bin" Then
                            chkOracleClientVer = False
                            Exit Function
                        End If
                        If InStr(1, spStr(j), strPath) > 0 Then
                            chkOracleClientVer = True
                            Exit Function
                        End If
                    Next j
                End If
                
                i = i + 1
            Loop Until Environ(i) = ""

        End If
    Else
        chkOracleClientVer = False
    End If
End Function
