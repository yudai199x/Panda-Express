Attribute VB_Name = "bizCommon"
Option Explicit

'**********************************************************************
' @(f)
' �@�\      : �u���m�点�v�o��
'
' �Ԃ�l    : �Ȃ�
'
' ������    : Integer iTarget �擾�Ώ� (0:TSB(�˗���)�����A1:TBA����)
'           : Range rngDest �o�͐�
'
' �@�\����  : rngDest �Ŏw�肵���Z���͈͂� iTarget �Ŏw�肵�����m�点��
'           : �o�͂���B
'
' ���l      : 1�`4���(�����Z��) ����(YYYY/MM/DD HH:NN)
'             5�`���(�����Z��) �R�����g(�ő�2�s)
'
'**********************************************************************
Public Sub ViewModifyHistory(iTarget As Integer, rngDest As Range)
    Dim iRow As Long
    Dim iRowCount As Long
    Dim mhInfo() As SModifyHistory
    Dim sDate As String
    Dim msg1 As String
    Dim msg2 As String
    Dim npos As Long
    Dim iCnt As Long
    Dim wkCell As Range
    
    Dim i As Long
    
    Application.EnableEvents = False: Application.ScreenUpdating = False
    ' **********************
    ' �o�͗̈���N���A����B
    ' **********************
    For Each wkCell In rngDest
        wkCell.Value = ""
    Next
    rngDest.Hyperlinks.Delete
    rngDest.Font.Size = 12
    rngDest.Font.Underline = xlUnderlineStyleNone
    rngDest.Font.ThemeColor = xlThemeColorLight1
    
    ' ***********************************************
    ' DB �X�V�����e�[�u������u���m�点�v���擾����B
    ' ***********************************************
    iCnt = GetModifyHistory(iTarget, mhInfo)
    If iCnt < 0 Then
        ' �G���[�����������ꍇ�͉������Ȃ��B(�G���[�o�͍ς�)
        Application.EnableEvents = True: Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' ********************
    ' �擾���ʂ��o�͂���B
    ' ********************
    If iCnt > 0 Then
        ' �擾���� (�ő� 4 ��) ���擾����B
        iRowCount = UBound(mhInfo) - LBound(mhInfo) + 1
        
        '���[�N�V�[�g�ɏo�͂���B
        For i = 1 To iRowCount
            
            iRow = i + i - 1
            
            ' �X�V�������o�͗p�̕�����ɕϊ�����B
            sDate = Format(mhInfo(i).dtModifyDate, "YYYY/MM/DD HH:NN")
            rngDest.Cells(iRow, 1).Value = sDate    ' ���t
                
            ' ���m�点���e�ɉ��s�R�[�h���܂܂�邩�m�F����B
            npos = InStr(mhInfo(i).strMessage, vbCrLf)
            
            If npos = 0 Then
                ' ���s�R�[�h���܂܂Ȃ��ꍇ
                msg1 = mhInfo(i).strMessage
            Else
                ' ���s�R�[�h���܂ޏꍇ
                ' ���s�R�[�h�ʒu�ŕ�������B
                msg1 = Left(mhInfo(i).strMessage, npos - 1) ' 1 �s��
                ' 2 �s�� (3 �s�ڈȍ~��2�s�ڂɑ����ďo�͂���B)
                msg2 = Replace(Mid(mhInfo(i).strMessage, npos + 2), vbCrLf, "�@")
                
                ' ����
                msg1 = msg1 + vbCrLf + msg2
            End If
                        
            ' �Z���ɏo�͂���B
            If Len(Trim(mhInfo(i).sLink)) = 0 Then
                ' �ʏ탁�b�Z�[�W
                rngDest.Cells(iRow, 4).Value = msg1
            Else
                ' �����N�t�����b�Z�[�W
                rngDest.Hyperlinks.Add Anchor:=rngDest.Cells(iRow, 4), _
                                       Address:=mhInfo(i).sLink, _
                                       TextToDisplay:=msg1
                rngDest.VerticalAlignment = xlTop
                rngDest.Font.Size = 12
            End If
            ' ���̂��m�点��
        Next
    Else
        ' ���m�点�����݂��Ȃ��ꍇ�́u�Ȃ��v���o�͂���B
        rngDest.Cells(1, 4).Value = "���m�点�͂���܂���B"
    End If
    
    Application.EnableEvents = True: Application.ScreenUpdating = True
    ' �I��
End Sub



