Attribute VB_Name = "libDBAccess"
Option Explicit
'********************************************************************************************
'  ��͈˗��V�X�e�� - DB �A�N�Z�X ���ʃ��C�u���� ���W���[��
'               Copyright 2015, XXXX All Rights Reserved.
'  2015-05-14 �V�K�쐬 (M.NAKAI)
'********************************************************************************************

' ********
' �萔��`
' ********
' SQL �N�G�����s���ʊi�[�p�t�@�C����
Public Const cSqlResFilenameP As String = "TEMSqlRes#.txt"
' SQL �N�G�����s���ʊi�[�p�t�@�C�� �o�͐�f�B���N�g�����擾�p���ϐ�
Public Const cSqlResOutDirEnv As String = "WORKDIR"

' ********
' �L��ϐ�
' ********
' SQL���s���O �t�@�C�� �p�X
Public gSqlExecLogFilepath As String

' **********
' �\���̒�`
' **********

' SQL���s�ϐ��Z�b�g
Public Type SSqlSet
    sExecSqlSheet As String   ' ���sSQL�V�[�g��
    iExecSqlNo As Long        ' ���sSQL�ԍ�(SQL�V�[�g�̗�ԍ�)
    sResultFile As String     ' �N�G�����s���ʎ擾�p�t�@�C��
    rep As Object             ' SQL �u���Ώ۔z�� (Scripting.Dictionary�^)
    psData() As String        ' �N�G�����ʊi�[
End Type

'**********************************************************************
' @(f)
' �@�\      : ADO�𗘗p�����f�[�^����
'
' �Ԃ�l    : True  �F  ����I��
' �@�@�@      False �F  �ُ�I��
'
' ������    : String    SqlSheetName    SQL�����i�[���ꂽ�V�[�g��
' �@�@�@      Integer   Sqlc            SQL�i�[�J����
' �@�@�@      String    Setfile         �f�[�^�o�̓t�@�C����
' �@�@�@      String    Acol()          SQL�ϊ��㕶����
' �@�@�@      String    Bcol()          SQL�ϊ��O������
' �@�@�@      String    MES             �o�̓��b�Z�[�W���e
'
' �@�\����  :
'
' ���l      :
'
'**********************************************************************
Function CallAdoSql(SqlSheetName As String, sqlc As Integer, setfile As String, _
                    bb() As String, aa() As String, MES As String, Timeout As String) As Boolean
    Dim ans     As String
    Dim fp      As Integer
    Dim strhost As String
    Dim usename As String
    Dim strpassword As String
    Dim strsql  As String
    Dim Sqlstr  As String
    Dim ansb    As Boolean
    Dim lCnt    As Integer
    Dim dc      As Integer
    Dim SqlCom  As String
    Dim i       As Integer
    Dim j       As Integer
    Dim n       As Integer
    Dim m       As Integer
    Dim Sqltmp  As String
    Dim Sqladd  As String
    Dim Ctr     As Integer
    Dim inData  As String
    Dim firstflg As Integer
    
    On Error GoTo Err_Exit
    
'    Application.StatusBar = MES & " Start at " & Format(Now(), "YYYY/MM/DD HH:MM:SS")
    
    SqlCom = ""
    i = 2   ' 2�s�ڂ���ǂݍ���
    Do While Workbooks(MAINBOOK).Sheets(SqlSheetName).Cells(i, sqlc).Value <> ""
        Sqlstr = Workbooks(MAINBOOK).Sheets(SqlSheetName).Cells(i, sqlc).Value ''��s�ǂݍ���
        Select Case Left(Sqlstr, 1)
        Case "#"
        Case Else
        If InStr(1, UCase(Sqlstr), "EXIT") > 0 Then
        Else
            j = LBound(bb)
            Do While (j <= UBound(bb)) And (j <= UBound(aa))
                If bb(j) <> "" Then
                    n = InStr(1, Sqlstr, bb(j))  ''������͂��邩
                Else
                    n = 0
                End If
                If n > 0 Then
                    Select Case Left(bb(j), 1)
                    Case "\"
                        ans = Replace(Sqlstr, bb(j), aa(j))
                        Sqlstr = ans
                    Case "%"
                        Sqltmp = Left(Sqlstr, n - 1)    ''�擪�̕���
                        Sqladd = Mid(Sqlstr, n + Len(bb(j)), 256)   ''�c��̕���
                        Sqlstr = ""
                        Ctr = 1
                        fp = FreeFile(1)
                        Open aa(j) For Input As #fp
                        Do While Not EOF(fp)
                            Line Input #fp, inData
                            inData = Trim(inData)
                            If inData <> "" And _
                               Left(inData, 1) <> "-" Then
                                If Ctr > 1 Then
                                    Sqlstr = Sqlstr & "or " & Sqltmp & inData & Sqladd & _
                                                        " " & Chr(13) & Chr(10)
                                Else
                                    Sqlstr = Sqlstr & "(" & Sqltmp & inData & Sqladd & _
                                                      " " & Chr(13) & Chr(10)
                                End If
                                Ctr = Ctr + 1
                            End If
                        Loop
                        Close #fp
                        Sqlstr = Sqlstr & ")"
                    Case "$"
                        firstflg = 0
                        Sqltmp = Left(Sqlstr, n - 1)    ''�擪�̕���
                        Sqladd = Mid(Sqlstr, n + Len(bb(j)), 256)   ''�c��̕���
                        Sqlstr = ""
                        Ctr = 1
                        fp = FreeFile(1)
                        Open aa(j) For Input As #fp
                        Do While Not EOF(fp)
                            Line Input #fp, inData
                            inData = Trim(inData)
                            If inData <> "" And _
                               Left(inData, 1) <> "-" Then
                                If Ctr >= 240 Then
                                    Sqlstr = Sqlstr & "') or" & Chr(13) & Chr(10) & "  " & Sqltmp & "'" & inData
                                    Ctr = 1
                                    lCnt = 1
                                Else
                                    If Ctr > 1 Then
                                        If lCnt >= 8 Then
                                            Sqlstr = Sqlstr & "'," & Chr(13) & Chr(10) & "  " & Space(n - 1) & "'" & inData
                                            lCnt = 1
                                        Else
                                            If firstflg = 0 Then
                                                Sqlstr = Sqlstr & "'" & inData
                                                firstflg = 1
                                            End If
                                            Sqlstr = Sqlstr & "','" & inData
                                            lCnt = lCnt + 1
                                        End If
                                    Else
                                        Sqlstr = Sqlstr & " (" & Sqltmp
'                                        Sqlstr = Sqlstr & " (" & Sqltmp & "'" & Indata
                                        lCnt = 1
                                    End If
                                End If
                                Ctr = Ctr + 1
                            End If
                        Loop
                        Close #fp
                        Sqlstr = Sqlstr & "'" & Sqladd & ")"
                    Case "&"
                    Case "@"
                        ans = Replace(Sqlstr, bb(j), ReplaceSqlEsc(aa(j)))
                        Sqlstr = ans
                    End Select
                End If
                j = j + 1
            Loop
            If Right(Sqlstr, 1) = ";" Then Sqlstr = Left(Sqlstr, Len(Sqlstr) - 1)
            If Left(Sqlstr, 2) = "!!" Then
                n = InStr(1, Sqlstr, "/")
                m = InStr(1, Sqlstr, "@")
                strhost = Mid(Sqlstr, m + 1, 256)
                usename = Mid(Sqlstr, 3, n - 3)
                strpassword = Mid(Sqlstr, n + 1, m - n - 1)
            Else
                'If InStr(1, UCase(Sqlstr), "SET ") <= 0 And _
                '   InStr(1, UCase(Sqlstr), "COL ") <= 0 And _
                '   InStr(1, UCase(Sqlstr), "COLUMN ") <= 0 And _
                '   InStr(1, UCase(Sqlstr), "ALTER ") <= 0 Then
                If InStr(1, UCase(Sqlstr), "COL ") <= 0 And _
                   InStr(1, UCase(Sqlstr), "COLUMN ") <= 0 And _
                   InStr(1, UCase(Sqlstr), "ALTER ") <= 0 Then
                   SqlCom = SqlCom & " " & Sqlstr & Chr(13) & Chr(10)
                End If
            End If
        End If
        End Select
        i = i + 1
    Loop
    strsql = SqlCom
    
    ' ***********************
    ' SQL���s���O���o�͂���B
    ' ***********************
    ' �o�͐�͐ݒ�V�[�g����擾����B
    If Trim(gSqlExecLogFilepath) = "" Then
        ' ���擾�̏ꍇ�̂ݎ擾����B
        gSqlExecLogFilepath = "C:\Temp\SqlExec.log"
    End If
    
    ' �t�@�C�� �p�X���ݒ肳��Ă���ꍇ�̂ݏo�͂���B
    If Trim(gSqlExecLogFilepath) <> "" Then
        fp = FreeFile(1)
        Open gSqlExecLogFilepath For Output As #fp
        Print #fp, strsql
        Close #fp
    End If
    
    '' �Z���N�g���C��(���ۂ̃f�[�^����)
    dc = GetAdoSql(strhost, usename, strpassword, strsql, setfile)
    If dc < 0 Then GoTo Err_Exit

' ##       KenSu = dc
'        Application.StatusBar = MES & " Finish at " & Format(Now(), "YYYY/MM/DD HH:MM:SS") & " Return = True"
        CallAdoSql = True
    
    Exit Function

Err_Exit:
    
'    Application.StatusBar = "CallAdoSql Finish at " & Format(Now(), "YYYY/MM/DD HH:MM:SS") & " Return = False"
    CallAdoSql = False

End Function

'**********************************************************************
' @(f)
' �@�\      : �Z���N�g���C��
'
' �Ԃ�l    : 0�@   ����I��
'    �@�@�@ : -1    �ُ�I��
'
' ������    : String    Strhost         �ڑ���z�X�g��
'    �@�@�@   String    Username        �ڑ����[�U
'    �@�@�@   String    Strpassword     �ڑ��p�X���[�h
'    �@�@�@   String    Strsql          �����r�p�k
'    �@�@�@   String    Setfile         �o�̓t�@�C����
'
' �@�\����  :
'
' ���l      :
'
'**********************************************************************
Public Function GetAdoSql(strhost As String, usename As String, _
                          strpassword As String, strsql As String, _
                          setfile As String) As Integer

    Dim oraConnection  As ADODB.Connection
    Dim rcdSetWork     As ADODB.Recordset

    Dim cn      As Integer
    Dim sqldat  As String
    Dim fn(256) As String
    Dim fl(256) As Integer
    Dim fp      As Integer
    Dim i       As Integer
    Dim j       As Integer
    
    '<<< �R�l�N�V�����@�I�[�v�� >>>
    Dim bStatus As Boolean
    bStatus = Application.EnableEvents
    Application.EnableEvents = False
    
    Set oraConnection = ConnectOraS(strhost, usename, strpassword)
    '<<< ���t�f�[�^�擾���� >>>
    Set rcdSetWork = DataReadS(oraConnection, strsql)
    On Error GoTo err_exit2
    cn = rcdSetWork.Fields.Count
    For i = 1 To cn
        fn(i) = rcdSetWork.Fields(i - 1).Name
        fl(i) = rcdSetWork.Fields(i - 1).DefinedSize
    Next i

    fp = FreeFile(1)
    Open setfile For Output As #fp
    On Error GoTo err_exit1
    For i = 1 To cn
        If i > 1 Then Print #fp, " ";
        Print #fp, StrConv(LeftB(StrConv(fn(i) & Space(fl(i)), vbFromUnicode), fl(i)), vbUnicode);
    Next i
    Print #fp, ""
    For i = 1 To cn
        If i > 1 Then Print #fp, " ";
        Print #fp, String(fl(i), "-");
    Next i
    Print #fp, ""

    If cn > 0 Then
        i = 1
        Do While Not rcdSetWork.EOF
            DoEvents
            For j = 1 To cn
                If j > 1 Then Print #fp, " ";
                If IsNull(rcdSetWork.Fields(j - 1).Value) Then
                    sqldat = ""
                Else
                    sqldat = rcdSetWork.Fields(j - 1).Value
                End If
                Print #fp, StrConv(LeftB(StrConv(sqldat & Space(fl(j)), vbFromUnicode), fl(j)), vbUnicode);
            Next j
            Print #fp, ""
            rcdSetWork.MoveNext
            i = i + 1
        Loop
        rcdSetWork.Close
    End If

    Close #fp
    '<<< �R�l�N�V�����@�N���[�Y >>>
    oraConnection.Close
    Application.EnableEvents = bStatus

    If i > 1 Then
        GetAdoSql = i - 1
    Else
        GetAdoSql = 0
    End If

    Exit Function

err_exit1:
    Close #fp
err_exit2:
    GetAdoSql = -1
    Application.EnableEvents = bStatus

End Function

'**********************************************************************
' @(f)
' �@�\      : ORACLE�ɃR�l�N�g
'
' �Ԃ�l    : �Ȃ�
'
' ������    : �Ȃ�
'
' �@�\����  :
'
' ���l      : �ڑ��G���[�񕜏������܂��B(5SecX60=5Min)
'             ���ۂɂ͐ڑ����s�̎��Ԃ�����܂��̂ŁA�T���ȏ�ŃG���[
'             �I���ƂȂ�܂��B
'
'**********************************************************************
Public Function ConnectOraS(StrSource As String, StrUID As String, _
                           StrPwd As String) As ADODB.Connection
    Dim cnnOpen
    Dim connectEnv   As String
    Dim errLoop      As ADODB.Error

    '' �f�[�^�x�[�X�ɐڑ���
    Set ConnectOraS = New ADODB.Connection
    On Error GoTo Err_Execute
    connectEnv = "Provider=MSDAORA;" _
                    & "Data Source = " & StrSource & ";" _
                    & "User ID = " & StrUID & ";" _
                    & "Password = " & StrPwd & ";"
    
    With ConnectOraS
        .CommandTimeout = 60
        .ConnectionTimeout = 45
        .ConnectionString = connectEnv
        .mode = adModeReadWrite
        .Open
    End With
Exit Function

'*** Error Process ***
Err_Execute:
    Dim rc      As Integer
    If ConnectOraS.Errors.Count > 0 Then
'        For Each errLoop In ConnectOraS.Errors
'            '�G���[���O�������}�N�����L�������A�폜
'        Next
        ErNm = Err.Number
        sEr = Err.Description
    End If
    If Not cnnOpen Is Nothing Then
        Set cnnOpen = Nothing
    End If

End Function

'**********************************************************************
' @(s)
' �@�\      : �f�[�^�̓ǂݍ���
'
' �Ԃ�l    : �Ȃ�
'
' ������    : �Ȃ�
'
' �@�\����  :
'
' ���l      :
'
'**********************************************************************
Public Function DataReadS(cnn As ADODB.Connection, strsql As String) _
                                                         As ADODB.Recordset
    Dim cmdChange    As ADODB.Command
    Dim errLoop      As ADODB.Error
    
    On Error GoTo Err_Execute
    
    '' �f�[�^�x�[�X��������
    Set cmdChange = New ADODB.Command
    With cmdChange
        .ActiveConnection = cnn
        .CommandType = adCmdText
        .CommandText = strsql
        Set DataReadS = .Execute
    End With
    
Exit Function
    
'*** Error Process ***
Err_Execute:
    Dim rc      As Integer
    If cnn.Errors.Count > 0 Then
        Debug.Print Err.Description
        ErNm = Err.Number
        sEr = Err.Description
        For Each errLoop In cnn.Errors
            '�G���[���O�������}�N�����L�������A�폜
        Next
    End If
    cnn.Close
    Set cnn = Nothing

End Function

'**********************************************************************
' @(s)
' �@�\      : (��)SQL���s��������
'
' �Ԃ�l    : SSqlSet : SQL���s�f�[�^ �Z�b�g
'
' ������    : Long sExecSqlNo : ���s����SQL�N�G�����ԍ�
'           : String sSqlSheetName : SQL�V�[�g�̖��O(�����l�uSQL�v)
'
' �@�\����  : SQL���s�ɕK�v�ȃf�[�^��p�ӂ���B���̊֐��ŏo�͂���SQL���s
'           : �f�[�^ �Z�b�g�́uSQL �u���Ώ۔z��(rep)�v��ݒ��ASQL���s
'           : �����֐�(n_DoSql)�����s����B
'
' ���l      : SQL���s�f�[�^ �Z�b�g�̓��e�͈ȉ��̒ʂ�B
'                sExecSqlSheet As String   ' ���sSQL�V�[�g�� ��
'                iExecSqlNo As Long        ' ���sSQL�ԍ�(SQL�V�[�g�̗�ԍ�) ��
'                sResultFile As String     ' �N�G�����s���ʎ擾�p�t�@�C�� ��
'                rep As Object             ' SQL �u���Ώ۔z�� (Scripting.Dictionary�^)
'                psData() As String        ' �N�G�����ʊi�[
'            ������̍��ڂ͖{�֐�(n_InitSql)���l��ݒ肷�邽�߁A�g�p�҂��ݒ肷��K�v
'              �͂Ȃ��B
'            SQL�u���Ώ۔z��(rep)��SQL�����s����O�ɁASQL�V�[�g�ɋL�ڂ��ꂽSQL�N�G����
'            �ɑ΂��āA�u�������镶������w�肷��Brep ��Dictionary�^�ɂȂ��Ă���A�ȉ�
'            �̒ʂ�ɂ��Ă�����B
'             ��)
'               Dim wSql As SSqlSet
'               wSql = m_InitSql(1)
'               wSq.rep.Add "�u���O", "�u����" ' ���s����SQL�N�G���́u�u���O�v�Ƃ�����������u�u����v�ɒu�������āASQL�N�G�������s����B
'**********************************************************************
Public Function n_InitSql(sExecSqlNo As Long, Optional sSqlSheetName As String = "SQL") As SSqlSet
    Dim ssetWk As SSqlSet       ' SQL���s�f�[�^ �Z�b�g
    Dim wkFname As String
    
    ' ***********************
    ' ���sSQL����ݒ肷��B
    ' ***********************
    MAINBOOK = ThisWorkbook.Name

    ssetWk.sExecSqlSheet = sSqlSheetName
    ssetWk.iExecSqlNo = sExecSqlNo
    
    ' ************************************
    ' �N�G�����s���ʃt�@�C�����𐶐�����B
    ' ************************************
    ' ���t�@�C������o�͐�f�B���N�g�������擾����B
    gsWorkDir = GetEnv(gcEnvFile, cSqlResOutDirEnv)
    
    ' �t�@�C�����𐶐�����B
    wkFname = Replace(cSqlResFilenameP, "#", CStr(sExecSqlNo))
    
    ' �f�B���N�g�����ƃt�@�C��������������B
    ssetWk.sResultFile = gsWorkDir & wkFname
    
    ' ********************************
    ' SQL �N�G�����u���z�����������B
    ' ********************************
    Set ssetWk.rep = CreateObject("Scripting.Dictionary")
    ssetWk.rep.RemoveAll

    n_InitSql = ssetWk
    
End Function

'**********************************************************************
' @(s)
' �@�\      : (��)SQL���s����
'
' �Ԃ�l    : Long : ���s���� �܂��� �擾����
'                SELECT �N�G�����s���͎擾�����B�����̏ꍇ�̓G���[
'                INSERT,UPDATE,DELETE���̍X�V�N�G���̏ꍇ�́A0�̏ꍇ��
'                �����A�����̏ꍇ�̓G���[�������B
'
' ������    : SSqlSet sqlDataSet SQL���s�Z�b�g
'
' �@�\����  : SQL���s���������֐�(n_InitSql)�ŏ�������SQL���s�Z�b�g��
'           : �N�G�������s����B
' ���l      : SELECT�N�G���̏ꍇ�ASQL���s�Z�b�g�̃N�G�����ʊi�[(psData)
'           : �ɕ������2�����z��Œ��o���e���i�[�����B
'                e.g.) wSql.psData(col, row)
'                       'col'�͗�(1�`)�A'row'�͍s(1�`)�������B
Public Function n_DoSql(sqlDataSet As SSqlSet) As Long
    Dim wkAcol() As String
    Dim wkBcol() As String
    Dim iCnt As Long
    Dim i As Long
    Dim key As Variant
    Dim iRes As Long
    
    On Error GoTo EH:
    
    ' �u��������z��ɕϊ�����B
    iCnt = sqlDataSet.rep.Count
    
    If iCnt > 0 Then
        ReDim wkAcol(1 To iCnt)
        ReDim wkBcol(1 To iCnt)
        
        i = 1
        For Each key In sqlDataSet.rep
            wkBcol(i) = CStr(key)
            wkAcol(i) = sqlDataSet.rep(key)
            i = i + 1
        Next
    Else
        ' �u���������Ȃ��ꍇ�́A1�v�f�̋�v�f��n���B
        ReDim wkAcol(1)
        ReDim wkBcol(1)
        wkAcol(1) = ""
        wkBcol(1) = ""
        
    End If
    
    ' SQL�����s����B
    If CallAdoSql(sqlDataSet.sExecSqlSheet, CInt(sqlDataSet.iExecSqlNo), sqlDataSet.sResultFile, _
            wkBcol(), wkAcol(), "", "") = False Then
        MsgBox "�G���[�������������߁A���������f����܂����B" & vbCrLf & _
               ErNm & "�F" & sEr, vbCritical + vbOKOnly, "�G���["
        Set ErNm = Nothing
        Set sEr = Nothing
        
        n_DoSql = -1
        On Error GoTo 0
        Exit Function
    End If
    
    ' ���s���ʂ��擾����B
    iRes = f_GetData(sqlDataSet.psData, sqlDataSet.sResultFile)

#If DEBUG_MODE <> 1 Then
    ' ���s���ʃt�@�C�����폜����B
    Kill sqlDataSet.sResultFile
#End If

    n_DoSql = iRes
    Exit Function
    
EH:
        MsgBox "�G���[�������������߁A���������f����܂����B" & vbCrLf & _
               "[" & Err.Number & "]" & Err.Description, vbCritical + vbOKOnly, "�G���["
        Set ErNm = Nothing
        Set sEr = Nothing
    
End Function

'**********************************************************************
' @(s)
' �@�\      : ���s���������S�ȕ����u��
'
' �Ԃ�l    : String: �u����̕�����
'
' ������    : String sSrc : �u���Ώۂ̕���
'
' �@�\����  : ���s�R�[�h��SQL�ɐݒ�ł��镶���ɒu��������B
'
' ���l      : DB����ǂݏo�����s�R�[�h�ɕ�������ꍇ�� DbResumeNewLine()
'             ���g�p����B
'**********************************************************************
Public Function DbEscapeNewLine(sSrc) As String
    ' ���s�R�[�h�� &H7F �ɒu��������B
    DbEscapeNewLine = Replace(sSrc, vbCrLf, Chr(cRNewLine))
    
End Function

'**********************************************************************
' @(s)
' �@�\      : ���s��������
'
' �Ԃ�l    : String: �u����̕�����
'
' ������    : String sSrc : �u���Ώۂ̕���
'
' �@�\����  : DbEscapeNewLine()�Œu�����������s�R�[�h�𕜌�����B
'
' ���l      :
'
'**********************************************************************
Public Function DbResumeNewLine(sSrc) As String
    ' ���s�R�[�h�� &H7F �ɒu��������B
    DbResumeNewLine = Replace(sSrc, Chr(cRNewLine), vbCrLf)
    
End Function

