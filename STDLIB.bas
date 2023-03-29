Attribute VB_Name = "STDLIB"
'********************************************************************************************
'   EA�ƭ� ��p�T�u���[�`���W
'               Copyright �i���j���Ł@�l���s�H��@���Y�Z�p���i�� All Rights Reserved.
'
'   96-03-06    �����������݊J�n
'               �t�@�C���I�[�v���̎��A256�`511�͈̔͂Ŏw��i���̃A�v���Ŏg�p�\�Ƃ��邽�߁j
'�@             ����ɔ����Afopen,fclose,fcloseall�̔p�~
'               Run���\�b�h�Ŏg�p�ł���悤�ɃT�u���[�`����ǉ�
'   96-03-15    Run���\�b�h�p�T�u���[�`���͔p�~
'               get_column�̕�����i�[����������
'   96-03-28    get_sql�̎��ԕ\��������ύX
'               get_sql�̃��[�h��ǉ�
'   96-06-05    get_sql�̌����҂��ɂ�DoEvents�ɂ��OS�ɐ����n���悤�ɂ���
'   96-06-07    DoEvents�ɂ��OS�ɐ����n���̂���߂�
'   96-06-12    ����ς�߂�(^^;
'               wait_sql�ɂāA�t�@�C���̑傫����������ƕ\������悤�ɂ���
'   96-07-08    count_sqlfile�̖߂�l��Long�ɂ���
'   96-08-30    get_column�ɂāA�������ʃt�@�C����1�s�̃f�[�^�����R�[�h�������̒�����
'               �����Ȃ������ꍇ�ɃG���[�ƂȂ��Ă����̂��C��
'   96-12-09    file_merge�ɂāA�t�@�C���T�C�Y���O�̎��ɃG���[�Ŏ~�܂��Ă��܂��̂��C��
'   96-12-24    MidB,LeftB,InstrB���g�p���邱�ƂŁA�t�@�C�����Ɋ����������Ă���ꍇ�̕s����C��
'   97-02-18    Write_Log���[�`���ǉ�
'   97-05-22    get_sql�ɂăt�@�C���������݂Ɏ��s�����ꍇ�ɂ�����x�������J�n����
'   97-06-09    get_sql�ɂăt�@�C���������݂Ɏ��s�����ꍇ�ɖ߂�l��False�ɂ���
'               get_sql�ɂē���"#!"�̏ꍇ�V�F���X�N���v�g���̍s���Ƃ݂Ȃ��ď������ށB
'   97-06-10    get_cmd�ǉ��B
'               get_sql�ł͓���"#!"�̏ꍇ����������B
'   97-09-11    make_sqlfile�ɂāA�R�����g��30�s�𒴂����ꍇ����ȍ~�̍s��ǂ܂Ȃ��s����C��
'   99-02-24    make_sqlfile�ɂāA$,%,&��p�����u���ɐ旧���Ēʏ�̒u�����s�Ȃ��悤�ύX
' 2000-01-25    Office2000�ɑΉ�
'********************************************************************************************

Public Const gcEnvFile As String = "C:\ENVIRON.DAT"

'**********************************************************************
'NAME       : check_column(fp,column())
'FUNCTION   : �t�@�C�������݈ʒu����ǂݍ���ōs�̓���-�̍s������������A
'             ���̍s�𓪂̕������璲�ׂăX�y�[�X�̃J�����ʒu��z��ɃZ�b�g����B
'NOTE       : ����-�̍s��1000�s�ȓ��ɂȂ���΃G���[
'           : column()�̏��index�ԍ����J�����ʒu�̐��ɖ����Ȃ��ꍇ��
'           : ���̃J�����͖��������B
'           : �^�u���܂܂�Ă����ꍇ�A����ȃJ�����ʒu��ۏ؂ł��Ȃ�
'           : ���SQL�������ʃt�@�C���Ɏg�p
'INPUT      : integer fp       �ΏۂƂȂ�t�@�C���̃t�@�C���ԍ�
'           : integer column() �J�����ʒu�i�[�p�z��
'RETURN     : ���� : �J������
'           : ���� : -1 --- ���R�[�h��������������Ȃ�
'           :        -2 --- col()�z�񐔂����R�[�h�������̕����ӏ���菭�Ȃ�
'**********************************************************************
'
Function Check_Column(fp As Integer, Column() As Integer) As Integer
    Dim n, i As Integer
    Dim Column_Num As Integer
    Dim st As String
    Dim inData As String
    Dim UB As Integer

    UB = UBound(Column)
    Column_Num = 0
    Erase Column

    For i = 1 To 1000                       '�m�F�񐔂�1000��
        On Error Resume Next
        Line Input #fp, inData
        If EOF(fp) Then Exit For
        flg = True
        For n = 1 To Len(inData)
            If Mid(inData, n, 1) <> "-" And Mid(inData, n, 1) <> " " Then Exit For
        Next n
        If n > Len(inData) And inData <> "" Then    '�S��'-'�̍s��������
            n = 0
            Do
                Do                                  '-���T�[�`����
                    n = n + 1
                    st = Mid(inData, n, 1)
                Loop Until st = "-" Or st = ""
                If st <> "" Then
                    Column_Num = Column_Num + 1
                    If Column_Num > UB Then GoTo array_error
                    Column(Column_Num) = n
                Else
                    Exit Do
                End If
                Do                                  '-�ȊO���T�[�`����
                    n = n + 1
                    st = Mid(inData, n, 1)
                Loop Until st <> "-" Or st = ""
            Loop Until st = ""
            Exit For
        End If
    Next i
    If i = 1001 Or EOF(fp) Then GoTo search_error
    If Column_Num < UB Then Column(Column_Num + 1) = 0
    Check_Column = Column_Num
    Exit Function
search_error:
    Check_Column = -1
    Exit Function
array_error:
    Check_Column = -2
End Function

'**********************************************************************
' NAME     : get_column(fp,col(),data())
' FUNCTION : �t�@�C�������݈ʒu����ǂݍ����col()�z��̈ʒu�ŕ�������
'          : �z��ɕ�����Ƃ��ăZ�b�g����
' NOTE     : �ϐ��̔z��͂P�����̂݁B�������ł͓���ۏ؂ł��Ȃ��B
'          : col()�ɂ͕����ʒu���Z�b�g���Ă����K�v����B�܂��Acol()=0�ȍ~��
'          : �f�[�^�͓ǂݍ��܂Ȃ��B��{�I��check_column()���s��Ɏg�p���邱�ƁB
'          : NULL�l�͖������ANULL�l��100��A��������ُ�I��
'          : �t�@�C���G���h�ł���ȏ�ǂݍ��߂Ȃ��ꍇ�ُ�I��
' INPUT    : integer fp �ΏۂƂȂ�t�@�C���̃t�@�C���ԍ�
'          : integer col() �J�����\�����Z�b�g�����z��
'          : string data() �Z�b�g����z��
' RETURN   : TRUE:����I���@FALSE:�ُ�I��
'**********************************************************************
'
Function Get_Column(fp As Integer, col() As Integer, Data() As String) As Boolean
    Dim i As Integer
    Dim inData As String
    Dim UB As Integer

    UB = UBound(Data)
    If UB > UBound(col) Then UB = UBound(col)
    On Error Resume Next
    For i = 1 To 100
        If EOF(fp) Then
            i = 101
            Exit For
        End If
        Line Input #fp, inData
        If Trim(inData) <> "" Then Exit For
    Next i
    On Error GoTo 0
    If i = 101 Then GoTo Err

    Erase Data
    inData = StrConv(inData, vbFromUnicode)
    i = 1
    Do
        If i = UB Then
            If LenB(inData) <= col(i) - 1 Then
                Data(i) = ""
            Else
                Data(i) = Trim(StrConv(MidB(inData, col(i), LenB(inData) - col(i) + 1), vbUnicode))
            End If
        ElseIf col(i + 1) = 0 Then
            If LenB(inData) <= col(i) - 1 Then
                Data(i) = ""
            Else
                Data(i) = Trim(StrConv(MidB(inData, col(i), LenB(inData) - col(i) + 1), vbUnicode))
            End If
        Else
'            data(i) = Trim(MidB(indata, col(i), col(i + 1) - col(i)))
'           '96-3-15 �ύX
            Data(i) = Trim(StrConv(MidB(inData, col(i), col(i + 1) - col(i) - 1), vbUnicode))
        End If
        i = i + 1
        If i > UB Then Exit Do
    Loop Until col(i) = 0

    Get_Column = True
    Exit Function

Err:
    Get_Column = False
End Function

'**********************************************************************
' NAME      : get_sql(sqlsheet,sqlcolumn,outfile,B_column,A_column,comment,mode)
' FUNCTION  : SQL�����s����
' NOTE      : �V�[�g��SQL���́A"#"�����ɂ���Ƃ����������B
'           : �V�[�g��SQL���P�s�ɁA�����̕ϐ������Ă��悢�B�A���A$,%,&�Ŏn�܂�ϐ�������
'           : �s�ɂ͑��̕ϐ������Ȃ����ƁB
'           : B_column()=""�ȍ~�̃f�[�^�͓ǂ܂Ȃ��B
'           : B_column()�̕ϐ��̓���$�̏ꍇ�AA_column����̧�ٖ����J����get_column()�ɂ����
'           : �f�[�^��ǂݍ��݁A------ in ('******', ��� �̌`�ɕϊ��i̧�ٖ��̓t���p�X�w��̂��Ɓj�B
'           :   ex) hed_lotnum in ($LOTNUM.DAT) �� hed_lotnum in ('111111','111112','111113'...
'           : B_column()�̕ϐ��̓���%�̏ꍇ�AA_column����̧�ٖ����J����get_column()�ɂ����
'           : �f�[�^��ǂݍ��݁A------ like '******%' or �܂��� ------ = '******' or �̌`�ɕϊ�
'           : �i̧�ٖ��̓t���p�X�w��̂��Ɓj�B
'           :   ex) hed_kndnam like '%KNDNAM.DAT%' �� hed_kndnam like 'T5W33%' or ...
'           : B_column()�̕ϐ��̓���&�̏ꍇ�AA_column����̧�ٖ����J����get_column()�ɂ����
'           : �f�[�^��ǂݍ��݁A------ like '******%' and �܂��� ------ = '******' and �̌`�ɕϊ�
'           : �i̧�ٖ��̓t���p�X�w��̂��Ɓj�B
'           :   ex) hed_kndnam like '%KNDNAM.DAT%' �� hed_kndnam like 'T5W33%' and ...
'           : �z�񐔂́A�Œ���K�v�ȃJ�������������K�v�B
'           : comment�́A�ȉ��̓���L�����g�p�ł���B
'           :   %T% : �������Ԃ�\��(Disp Time)
'           :   %A% : �������̃t�@�C���̐���\��(Disp Access file)
'           : mode�͕����̕��тŎ����B�����I���\�B
'           :    L  : 'c:\'�ɂ�SQL�t�@�C������������(Local mode)
'           :    D  : �����͍s��Ȃ�(Debug mode)
'           :    P  : SQL�t�@�C�����M��A���ʂ�҂����ɏI��(Pass mode)
'           :    W  : SQL�t�@�C���͑��M�����A���ʂ̂ݑ҂�(Wait mode)
' INPUT     : string sqlsheet   SQL�V�[�g���iNULL����"SQL"�j
'           : integer sqlcolumn SQL���̃J�����ʒu
'           : string outfile    �o�̓t�@�C����
'           : string B_column() �u�������O�̃J�������z��
'           : string A_column() �u��������̃J�������z��
'           : string comment    �������R�����g
'           : string mode       �e�탂�[�h
' RETURN    : TRUE ����I�� FALSE ���ϐ�HOSTNAME�̖��o�^
'**********************************************************************
'
'Function Get_Sql1(SqlSheet As String, SqlColumn As Integer, OutFile As String, _
'                 B_Column() As String, A_Column() As String, comment As String, _
'                 MODE As String) As Boolean
'    Dim hostName As String
'    Dim Lmode As Boolean
'    Dim i As Long
'    Dim SqlCount As Long
'    Dim SqlRcd As String
'    Dim FstRcd As String
'
'    on error goto 0
'    If InStr(1, MODE, "D", 1) <> 0 Then
'        Get_Sql = True
'        Exit Function
'    End If
'    If InStr(1, MODE, "L", 1) <> 0 Then
'        Lmode = True
'    Else
'        Lmode = False
'    End If
'    If InStr(1, MODE, "W", 1) = 0 Then
'        hostName = Make_SQLFile(SqlSheet, SqlColumn, B_Column(), A_Column(), Lmode)
'        If hostName = "" Then
'            Get_Sql = False
'            Exit Function
'        End If
'    End If
'
'    If InStr(1, MODE, "P", 1) = 0 Then
'        i = Wait_SQL(hostName & ".txt", comment)
'        on error goto exit_failure
'        If Dir(OutFile) <> "" Then Kill OutFile
'        FileCopy "o:\" & hostName & ".txt", OutFile
'        on error goto 0
'        If Dir(OutFile) <> "" Then Kill "o:\" & hostName & ".txt"
'    End If
'
'    SqlCount = Count_SqlFile(OutFile)
'    If SqlCount < 1 Then
'        Get_Sql = False
'    Else
'        Get_Sql = True
'    End If
'    Get_Sql = True
'    Exit Function
'exit_failure:
'    Get_Sql = False
'    Exit Function
'End Function

'**********************************************************************
' NAME      : get_cmd(sqlsheet,sqlcolumn,B_column,A_column,mode)
' FUNCTION  : �V�F���X�N���v�g�����s����
' NOTE      : �V�[�g�̃V�F���X�N���v�g���́A"#"�����ɂ���Ƃ����������B�A���A"#!"�̏ꍇ�̓V�F��
'           : �X�N���v�g���̍s���Ƃ݂Ȃ��Ă��̂܂܏������ށB
'           : �V�[�g��SQL���P�s�ɁA�����̕ϐ������Ă��悢�B
'           : B_column()=""�ȍ~�̃f�[�^�͓ǂ܂Ȃ��B
'           : �z�񐔂́A�Œ���K�v�ȃJ�������������K�v�B
'           : comment�́A�ȉ��̓���L�����g�p�ł���B
'           :   %T% : �������Ԃ�\��(Disp Time)
'           :   %A% : �������̃t�@�C���̐���\��(Disp Access file)
'           : mode�͕����̕��тŎ����B�����I���\�B
'           :    L  : 'c:\'�ɂ��V�F���X�N���v�g����������(Local mode)
'           :    D  : �����s��Ȃ�(Debug mode)
' INPUT     : string sqlsheet   �V�F���X�N���v�g�V�[�g���iNULL����"CMD"�j
'           : integer sqlcolumn �V�F���X�N���v�g�̃J�����ʒu
'           : string B_column() �u�������O�̃J�������z��
'           : string A_column() �u��������̃J�������z��
'           : string mode       �e�탂�[�h
' RETURN    : TRUE ����I�� FALSE ���ϐ�HOSTNAME�̖��o�^
'**********************************************************************
'
Function Get_Cmd(CmdSheet As String, CmdColumn As Integer, _
                 B_Column() As String, A_Column() As String, _
                 mode As String) As Boolean
    Dim hostName As String
    Dim Lmode As Boolean
    Dim i As Long

    On Error GoTo 0
    If InStr(1, mode, "D", 1) <> 0 Then
        Get_Cmd = True
        Exit Function
    End If
    If InStr(1, mode, "L", 1) <> 0 Then
        Lmode = True
    Else
        Lmode = False
    End If
    hostName = Make_CMDFile(CmdSheet, CmdColumn, B_Column(), A_Column(), Lmode)
    If hostName = "" Then
        Get_Cmd = False
        Exit Function
    End If

    Get_Cmd = True
    Exit Function
exit_failure:
    Get_Cmd = False
    Exit Function
End Function

'**********************************************************************
' NAME      : make_sqlfile
' FUNCTION  : SQL�����t�@�C�����쐬����
' NOTE      : get_sql()��p
' INPUT     : string sht        SQL�V�[�g��
'           : integer col       SQL���̃J�����ʒu
'           : string B_col()    �u�������O�̃J�������z��
'           : string A_col()    �u��������̃J�������z��
'           : string mode       ���[�J���Ƀt�@�C�����������ރt���O
' RETURN    : TRUE ����I�� FALSE �ُ�I��
' HISTORY   : 96-04-23 �u���O������̈ꕔ�����̕�����ƈ�v�����ꍇ�ɐ�����s��̑Ώ�
'           :          �u���O�����񂪈�v�����ꍇ�ُ�I��
'**********************************************************************
'
Function Make_SQLFile(sht As String, col As Integer, B_Column() As String, _
                      A_Column() As String, mode As Boolean, SqlRcd As String, FstRcd As String) As String
    Dim LineData As String
    Dim host As String
    Dim LineCtr As Integer      ' �V�[�g�̍s�ԍ�
    Dim A_UB, B_UB As Integer   ' b_col,a_col�̔z��
    Dim MAX_UB As Integer       ' b_col,a_col�̔z��
    Dim a_col() As String, b_col() As String
    Dim fp As Integer
    Dim fpd As Integer          ' �f�o�b�O�p
    Dim i, j As Integer
    Dim st As String

    '�z��̐����m�F
    B_UB = UBound(B_Column)
    A_UB = UBound(A_Column)
    For i = 0 To B_UB - 1
        If B_Column(i + 1) = "" Then
            B_UB = i
            Exit For
        End If
    Next i
    If B_UB > A_UB Then
        Make_SQLFile = ""
        Exit Function
    End If
    '�u���O������̏d�����Ȃ����m�F
    If B_UB > 1 Then
        For j = 1 To B_UB - 1
            For i = j + 1 To B_UB
                If B_Column(j) = B_Column(i) Then
                    Make_SQLFile = ""
                    Exit Function
                End If
            Next i
        Next j
    End If
    '�u���O�������̊m�F�ƒu��������̑��
    If B_UB > 0 Then
        ReDim b_col(B_UB)
        ReDim a_col(B_UB)
        For i = 1 To B_UB
            b_col(i) = StrConv(B_Column(i), vbFromUnicode)
            a_col(i) = StrConv(A_Column(i), vbFromUnicode)
        Next i
        '�u��������𕶎����̑傫�����ɕ��בւ���B
        i = 1
        Do Until i > B_UB - 1
            If LenB(b_col(i)) >= LenB(b_col(i + 1)) Then
                i = i + 1
            Else
                st = b_col(i)
                b_col(i) = b_col(i + 1)
                b_col(i + 1) = st
                st = a_col(i)
                a_col(i) = a_col(i + 1)
                a_col(i + 1) = st
                If i > 1 Then i = i - 1
            End If
        Loop
    End If
    host = GetEnv("c:\environ.dat", "HOSTNAME")
    If host = "" Then
        Make_SQLFile = ""
        Exit Function
    End If
    On Error Resume Next
    If Dir("o:\" & host & ".txt") <> "" Then Kill "o:\" & host & ".txt"
    On Error GoTo 0
    If mode Then
        fpd = FreeFile(1)
        Open "c:\" & host & ".sql" For Output As #fpd ' for Debug
    End If
'    fp = FreeFile(1)
'    Open "o:\" & Host & ".sql" For Output As #fp
    LineCtr = 1
    Do While 1
        For i = 1 To 30
            LineData = StrConv(Trim(Sheets(sht).Cells(LineCtr, col).Text), vbFromUnicode)
            If LineData <> "" And LeftB(LineData, 1) <> StrConv("#", vbFromUnicode) Then Exit For
            If LeftB(LineData, 1) = StrConv("#", vbFromUnicode) Then i = i - 1
            LineCtr = LineCtr + 1
        Next i
        LineCtr = LineCtr + 1
        If i = 31 Then Exit Do
        If B_UB > 0 Then
            For i = 1 To B_UB
                If InStrB(LineData, b_col(i)) <> 0 Then   ' b_col(i)���܂܂�Ă�����
                    If LeftB(b_col(i), 1) <> StrConv("$", vbFromUnicode) And _
                     LeftB(b_col(i), 1) <> StrConv("%", vbFromUnicode) And _
                     LeftB(b_col(i), 1) <> StrConv("&", vbFromUnicode) Then
                        LineData = Replace(LineData, b_col(i), a_col(i))
                    End If
                End If
            Next i
            For i = 1 To B_UB
                If InStrB(LineData, b_col(i)) <> 0 Then   ' b_col(i)���܂܂�Ă�����
                    If LeftB(b_col(i), 1) = StrConv("$", vbFromUnicode) Then      ' $��� �̏ꍇ
                        If mode Then
                            j = Replace_In(LineData, b_col(i), a_col(i), fpd)
                            If j < 0 Then GoTo exit_failure
                        End If
                        j = Replace_In(LineData, b_col(i), a_col(i), fp)
                        If j < 0 Then GoTo exit_failure
                        LineData = ""
                        Exit For
                    ElseIf LeftB(b_col(i), 1) = StrConv("%", vbFromUnicode) Then      ' %��� �̏ꍇ
                        If mode Then
                            j = Replace_File(LineData, b_col(i), a_col(i), fpd, "or")
                            If j < 0 Then GoTo exit_failure
                        End If
                        j = Replace_File(LineData, b_col(i), a_col(i), fp, "or")
                        If j < 0 Then GoTo exit_failure
                        LineData = ""
                        Exit For
                    ElseIf LeftB(b_col(i), 1) = StrConv("&", vbFromUnicode) Then      ' &��� �̏ꍇ
                        If mode Then
                            j = Replace_File(LineData, b_col(i), a_col(i), fpd, "and")
                            If j < 0 Then GoTo exit_failure
                        End If
                        j = Replace_File(LineData, b_col(i), a_col(i), fp, "and")
                        If j < 0 Then GoTo exit_failure
                        LineData = ""
                        Exit For
                    End If
                End If
            Next i
        End If
            If fno = 2 Then SqlRcd = SqlRcd & Space(1) & StrConv(LineData, vbUnicode)
            If fno = 1 Then
                SqlRcd = StrConv(LineData, vbUnicode)
                fno = 2
            End If
            If fno = 0 Then
                SqlRcd = StrConv(LineData, vbUnicode)
                FstRcd = SqlRcd
                SqlRcd = ""
                fno = 1
            End If
'            If MODE Then Print #Fpd, StrConv(LineData, vbUnicode)
'            Print #fp, StrConv(LineData, vbUnicode)
    Loop
    If mode Then Print #fpd, FstRcd
    If mode Then Print #fpd, SqlRcd
    Make_SQLFile = host
    If mode Then Close #fpd
    Close #fp
    Exit Function
exit_failure:
    Make_SQLFile = ""
    If mode Then Close #fpd
    Close #fp
End Function
    
'**********************************************************************
' NAME      : make_cmdfile
' FUNCTION  : �V�F���X�N���v�g���쐬����
' NOTE      : get_cmd()��p
' INPUT     : string sht        SQL�V�[�g��
'           : integer col       SQL���̃J�����ʒu
'           : string B_col()    �u�������O�̃J�������z��
'           : string A_col()    �u��������̃J�������z��
'           : string mode       ���[�J���Ƀt�@�C�����������ރt���O
' RETURN    : TRUE ����I�� FALSE �ُ�I��
' HISTORY   : 96-04-23 �u���O������̈ꕔ�����̕�����ƈ�v�����ꍇ�ɐ�����s��̑Ώ�
'           :          �u���O�����񂪈�v�����ꍇ�ُ�I��
'**********************************************************************
'
Function Make_CMDFile(sht As String, col As Integer, B_Column() As String, _
                      A_Column() As String, mode As Boolean) As String
    Dim LineData As String
    Dim host As String
    Dim LineCtr As Integer      ' �V�[�g�̍s�ԍ�
    Dim A_UB, B_UB As Integer   ' b_col,a_col�̔z��
    Dim MAX_UB As Integer       ' b_col,a_col�̔z��
    Dim a_col() As String, b_col() As String
    Dim fp As Integer
    Dim fpd As Integer          ' �f�o�b�O�p
    Dim i, j As Integer
    Dim st As String
    
    '�z��̐����m�F
    B_UB = UBound(B_Column)
    A_UB = UBound(A_Column)
    For i = 0 To B_UB - 1
        If B_Column(i + 1) = "" Then
            B_UB = i
            Exit For
        End If
    Next i
    If B_UB > A_UB Then
        Make_CMDFile = ""
        Exit Function
    End If
    '�u���O������̏d�����Ȃ����m�F
    If B_UB > 1 Then
        For j = 1 To B_UB - 1
            For i = j + 1 To B_UB
                If B_Column(j) = B_Column(i) Then
                    Make_CMDFile = ""
                    Exit Function
                End If
            Next i
        Next j
    End If
    '�u���O�������̊m�F�ƒu��������̑��
    If B_UB > 0 Then
        ReDim b_col(B_UB)
        ReDim a_col(B_UB)
        For i = 1 To B_UB
            b_col(i) = StrConv(B_Column(i), vbFromUnicode)
            a_col(i) = StrConv(A_Column(i), vbFromUnicode)
        Next i
        '�u��������𕶎����̑傫�����ɕ��בւ���B
        i = 1
        Do Until i > B_UB - 1
            If LenB(b_col(i)) >= LenB(b_col(i + 1)) Then
                i = i + 1
            Else
                st = b_col(i)
                b_col(i) = b_col(i + 1)
                b_col(i + 1) = st
                st = a_col(i)
                a_col(i) = a_col(i + 1)
                a_col(i + 1) = st
                If i > 1 Then i = i - 1
            End If
        Loop
    End If
    host = GetEnv("c:\environ.dat", "HOSTNAME")
    If host = "" Then
        Make_CMDFile = ""
        Exit Function
    End If
    If mode Then
        fpd = FreeFile(1)
        Open "c:\" & host & ".cmd" For Output As #fpd ' for Debug
    End If
    fp = FreeFile(1)
    Open "o:\" & host & ".cmd" For Output As #fp
    LineCtr = 1
    Do While 1
        For i = 1 To 30
            LineData = StrConv(Trim(Sheets(sht).Cells(LineCtr, col).Text), vbFromUnicode)
            If LineData <> "" And (LeftB(LineData, 1) <> StrConv("#", vbFromUnicode) Or _
             Left(LineData, 2) = StrConv("#!", vbFromUnicode)) Then Exit For
            LineCtr = LineCtr + 1
        Next i
        LineCtr = LineCtr + 1
        If i = 31 Then Exit Do
        If B_UB > 0 Then
            For i = 1 To B_UB
                If InStrB(LineData, b_col(i)) <> 0 Then   ' b_col(i)���܂܂�Ă�����
                    LineData = Replace(LineData, b_col(i), a_col(i))
                End If
            Next i
        End If
        If LineData <> "" Then
            If mode Then Print #fpd, StrConv(LineData, vbUnicode)
            Print #fp, StrConv(LineData, vbUnicode)
        End If
    Loop
    Make_CMDFile = host
    If mode Then Close #fpd
    Close #fp
    Exit Function
exit_failure:
    Make_CMDFile = ""
    If mode Then Close #fpd
    Close #fp
End Function
    
'**********************************************************************
' NAME      : wait_sql
' FUNCTION  : SQL�����̏I����҂�
' NOTE      : get_sql()��p
'           : �ꉞget_cmd()���s��Ƀt�@�C�����`�F�b�N���邱�Ƃ��\
' INPUT     : string fname      �t�@�C����
'           : string comment    �R�����g
' RETURN    : Long fsize        ̧�ٻ���(Bytes)
' HISTORY   : 97-06-10 ̧�ٖ��Ɋg���q���t����
'           : 97-07-25 ������Excel��Active�ɂ���̂���߂�
'**********************************************************************
'
Function Wait_SQL(fname As String, Comment As String) As Long
    Dim StartTime As Date
    Dim ToTime As Date
    Dim Combuf As String
    Dim fl As Long, Fb As Long
    Dim Awin As String
    
    Awin = Application.ActiveWindow.Caption
    StartTime = Now()
    Do Until Dir("o:\" & fname) <> ""
        ToTime = Now()
        Combuf = StrConv(Replace(StrConv(Comment, vbFromUnicode), StrConv("%T%", vbFromUnicode), StrConv(Format(ToTime - StartTime, "hh:mm:ss"), vbFromUnicode)), vbUnicode)
        i = 0
        If Dir("o:\x*.sql") <> "" Then
            Do
                i = i + 1
            Loop Until Dir() = ""
        End If
        Combuf = StrConv(Replace(StrConv(Combuf, vbFromUnicode), StrConv("%A%", vbFromUnicode), StrConv(Trim(str(i)), vbFromUnicode)), vbUnicode)
        Do
            'DoEvents
        Loop Until Format(ToTime, "hh:mm:ss") <> Format(Now(), "hh:mm:ss")
        Application.StatusBar = Combuf
    Loop
    'AppActivate "Microsoft Excel"
    fl = FileLen("o:\" & fname)
    Fb = 0
    If fl = 0 Then
        Fb = 1
    End If
    Do Until fl = Fb
        Fb = fl
        ToTime = Now() + 1
        Do
            Application.StatusBar = "�������ʃt�@�C�����쐬���Ă��܂�(" & fl & ")"
        Loop Until Format(ToTime, "hh:mm:ss") < Format(Now(), "hh:mm:ss")
        fl = FileLen("o:\" & fname)
    Loop
    fl = FileLen("o:\" & fname)
    Application.StatusBar = "�������ʃt�@�C�����쐬���Ă��܂�(" & fl & ")"
    Wait_SQL = fl
End Function
    
'**********************************************************************
' NAME      : count_sqlfile
' FUNCTION  : SQL�t�@�C���̐����J�E���g���A�t�@�C���̑̍ق𐮂���B
' NOTE      : �r���ɂ���^�C�g���A���R�[�h�������̓J�b�g����B
'             �s�������ɕʂ�Ă���ꍇ�A1�s�ɂ܂Ƃ߂�B�A���A10�s�ȏ��
'           �@�ʂ�Ă���ꍇ�̓G���[�B
' INPUT     : string filename �t�@�C����
' RETURN    : �����F�f�[�^��
'              -1 �F���R�[�h��������������Ȃ��B
'              -2 �F�Ώۃt�@�C����������Ȃ��B
'              -3 �F�s����10�𒴂���
'**********************************************************************
'
Function Count_SqlFile(fname As String) As Long
    Dim i As Integer, j As Integer, k As Integer
    Dim TempFile As String
    Dim Fpt As Integer
    Dim fp As Integer
    Dim inData As String
    Dim Title(11) As String
    Dim Separater(11) As String
    Dim Count As Long
    Dim Tcount As Long          '�^�C�g���̍s��
    Dim ErrFlg As Boolean

    If Dir(fname) = "" Then GoTo nofileerr
    fp = FreeFile(1)
    Open fname For Input As #fp
    'FName��TempFile�Ɠ����������ꍇ�̑Ώ��t��
    Do
        TempFile = Env("WORKDIR") & Format(Now(), "YYYYMMDDHHMMSS") & ".TMP"
    Loop While fname = TempFile
    If Dir(TempFile) <> "" Then Kill TempFile
    Fpt = FreeFile(1)
    Open TempFile For Output As #Fpt

    '�^�C�g���s�̎擾
    Tcount = 1
    On Error Resume Next
    For i = 1 To 1000                                   '�m�F�񐔂�1000��
        Title(Tcount) = StrConv(inData, vbFromUnicode)
        Line Input #fp, inData
        If EOF(fp) Then Exit For
        If Left(Trim(inData), 1) = "-" Then Exit For            '-�s��������
    Next i
    If i = 1001 Or EOF(fp) Then GoTo searcherr
    Separater(Tcount) = StrConv(inData, vbFromUnicode)
    Do
        Line Input #fp, inData
        Title(Tcount + 1) = StrConv(inData, vbFromUnicode)
        If EOF(fp) Then Exit Do
        Line Input #fp, inData
        Separater(Tcount + 1) = StrConv(inData, vbFromUnicode)
        If Left(Trim(inData), 1) <> "-" Then Exit Do            '-�s�łȂ�
        Tcount = Tcount + 1
    Loop Until Tcount > 11
    If Tcount > 11 Then GoTo titleerr       '10�s�𒴂���
    'Tcount = Tcount - 1
    Close #fp
    
    '�ēx�t�@�C�����J���A�^�C�g���s��ǂݔ�΂�
    fp = FreeFile(1)
    Open fname For Input As #fp
    For j = 1 To i
        Line Input #fp, inData
    Next j
    If Tcount > 1 Then
        For i = 1 To (Tcount - 1) * 2
            Line Input #fp, inData
        Next i
    End If
    '�^�C�g���A���R�[�h������
    inData = ""
    For i = 1 To Tcount
        inData = inData & Title(i) & StrConv(" ", vbFromUnicode)
    Next i
    Print #Fpt, StrConv(LeftB(inData, LenB(inData) - 1), vbUnicode)
    inData = ""
    For i = 1 To Tcount
        inData = inData & Separater(i) & StrConv(" ", vbFromUnicode)
    Next i
    Print #Fpt, StrConv(LeftB(inData, LenB(inData) - 1), vbUnicode)
    '�f�[�^��荞��
    Count = 0
    Do Until EOF(fp)
        Line Input #fp, inData
        inData = StrConv(inData, vbFromUnicode)
        If Trim(inData) <> "" Then
            If inData = Title(Tcount) Then
                If Tcount > 1 Then
                    For i = 1 To (Tcount - 1) * 2 - 1
                        Line Input #fp, inData
                    Next i
                Else
                    Line Input #fp, inData
                End If
            Else
                If LenB(inData) < LenB(Separater(1)) Then _
                 inData = inData & StrConv(Space(LenB(Separater(1)) - LenB(inData)), vbFromUnicode)
                Print #Fpt, StrConv(inData, vbUnicode);
                If Tcount > 1 Then
                    For i = 1 To Tcount - 1
                        Line Input #fp, inData
                        inData = StrConv(inData, vbFromUnicode)
                        If LenB(inData) < LenB(Separater(i + 1)) Then _
                         inData = inData & StrConv(Space(LenB(Separater(i + 1)) - LenB(inData)), vbFromUnicode)
                        Print #Fpt, " " & StrConv(inData, vbUnicode);
                    Next i
                End If
                Print #Fpt, ""
                Count = Count + 1
            End If
        End If
    Loop
    On Error GoTo 0
    Close #fp, #Fpt
    Kill fname
    FileCopy TempFile, fname
    Kill TempFile
    Count_SqlFile = Count
    Exit Function
titleerr:
    Close #fp, #Fpt
    Kill TempFile
    Count_SqlFile = -3
    Exit Function
searcherr:
    Close #fp, #Fpt
    Kill TempFile
    Count_SqlFile = -1
    Exit Function
nofileerr:
    Count_SqlFile = -2
End Function
    
'**********************************************************************
' NAME      : GetEnv, Env
' FUNCTION  : ���t�@�C��������ϐ�����荞��
' NOTE      : Env�̏ꍇ�A�����ɂ���Ĉȉ��̗l�Ɏd�l���ς��
'           :  �����Ȃ����ENVFILE����HOSTNAME�Œ�`���ꂽ���ϐ�����荞��
'           :  �����P�¥��ENVFILE��������Ŏw�肵�����ϐ�����荞��
'           :  �����Q�¥��GetEnv�Ɠ��l�̎d�l
' INPUT     : String EnvFileName ���t�@�C����
'           : String EnvStr      ���ϐ���
' RETURN    : ���ϐ�(�Ȃ��ꍇ��NULL)
'**********************************************************************
'
Function Env(Optional ByVal EnvFileName As Variant, Optional ByVal EnvStr As Variant) As String
    
    If IsMissing(EnvFileName) Then
        EnvFileName = envfile
        EnvStr = "HOSTNAME"
    ElseIf IsMissing(EnvStr) Then
        EnvStr = EnvFileName
        EnvFileName = envfile
    End If
        
    Env = GetEnv(EnvFileName, EnvStr)
End Function

Function GetEnv(ByVal EnvFileName As String, ByVal EnvStr As String) As String
    Dim n As Integer        ' ���ݒ��ڂ��Ă���v�f
    Dim Data As String      ' �t�@�C������ǂݍ��񂾂P�s
    Dim fp As Integer       ' file number
    Dim st As String        ' �ėp������

    GetEnv = ""
    On Error GoTo Err
    fp = FreeFile(1)
    Open EnvFileName For Input As #fp       '���t�@�C���̃I�[�v��
    On Error GoTo 0

    Do Until EOF(fp)
        Line Input #fp, Data
        If InStr(Data, EnvStr) <> 0 Then                '���ϐ�����
            n = 0
            Do                                          '��߰�or��ނ�T��
                n = n + 1
                st = Mid(Data, n, 1)
            Loop Until st = " " Or st = Chr(9) Or st = Chr(13) Or st = ""
            If st = Chr(13) Or st = "" Then GoTo Err
            Do                                          '��߰�or��ނ��X�L�b�v
                n = n + 1
                st = Mid(Data, n, 1)
            Loop Until st <> " " And st <> Chr(9) And st <> Chr(13) And st <> ""
            If st = Chr(13) Or st = "" Then GoTo Err
            Do                                          '��߰�or��ނ��X�L�b�v
                st = Mid(Data, n, 1)
                If st <> " " And st <> Chr(9) And st <> Chr(13) And st <> "" Then
                    GetEnv = GetEnv & st
                    n = n + 1
                End If
            Loop Until st = " " Or st = Chr(9) Or st = Chr(13) Or st = ""
        End If
    Loop
    Close #fp
Err:
    Exit Function
End Function

'**********************************************************************
' NAME      : replace
' FUNCTION  : ������̎w�蕔����u��
' NOTE      :
' INPUT     : string arg1�F�Ώۂ̕�����(SJIS)
'           : string arg2�F�u���O�̕�����(SJIS)
'           : string arg3�F�u����̕�����(SJIS)
' RETURN    : �u��������������i�����������Ȃ������ꍇ�ɂ͂��̂܂܁j
'**********************************************************************
'
Function Replace(ByVal Arg1 As String, Arg2 As String, Arg3 As String) As String
    Dim n As Integer
    Dim Arg_Left As String
    Dim Arg_Right As String
    Dim Arg_Ret As String
    Dim Len_Arg1 As Integer
    Dim Len_Arg2 As Integer

    If Arg2 = Arg3 Then         '�u�������񂪓����Ȃ�
        Replace = Arg1          '�߂�l�͕ω����Ȃ�
        Exit Function
    End If

    If Arg1 = Arg2 Then         '�Ώە�����ƒu���O�̕����񂪓����Ȃ�
        Replace = Arg3          '�߂�l��arg3
        Exit Function
    End If

    Len_Arg1 = LenB(Arg1)
    Len_Arg2 = LenB(Arg2)

    Do
        n = InStrB(1, Arg1, Arg2, 1)
        If n > 0 Then
            If n = 1 Then
                Arg_Left = ""
            Else
                Arg_Left = LeftB(Arg1, n - 1)
            End If
'            If n + Len_Arg2 = Len_Arg1 + 1 Then
'                Arg_Right = ""
'            Else
'                Arg_Right = MidB(Arg1, n + Len_Arg2)
'            End If
            Arg_Right = MidB(Arg1, n + Len_Arg2)
            Arg1 = Arg_Left & Arg3 & Arg_Right
        End If
    Loop Until n = 0

    Replace = Arg1

End Function

'**********************************************************************
' NAME      : replace_in(arg,b_arg,fname,fp)
' FUNCTION  : ������arg�Ɋ܂܂�镶����b_arg���A�t�@�C��fname���當��������
'           : ����"'"�ň͂�� ","�ŋ�؂���������ɒu�����āA�t�@�C���|�C���^
'           : fp�ŊJ�����t�@�C���ɏ������ށB
'           : ������ in (------) �`����SQL���Ɏg�p
' NOTE      : fname�̓t���p�X�Ŏw�肷��
' INPUT     : string Arg    �u���Ώە�����(SJIS)
'           : string B_Arg  �u���O������(SJIS)
'           : string FName  ������ۑ��t�@�C����(SJIS)
'           : integer FFp   �������݃t�@�C���|�C���^
' RETURN    : ���� ����I�� ���� �ُ�I��
'**********************************************************************
'
Function Replace_In(Arg As String, B_Arg As String, fname As String, FFp As Integer) As Integer
    Dim RFp As Integer      ' �t�@�C���|�C���^
    Dim Arg_Point As Integer
    Dim col(2) As Integer
    Dim Data(2) As String
    Dim i As Integer, j As Integer
    Dim Count As Integer
    Dim Chr_Flg As Integer
    Dim st As String

    RFp = FreeFile(1)
    On Error GoTo open_error
    Open StrConv(fname, vbUnicode) For Input As #RFp
    On Error GoTo 0

    Arg_Point = InStrB(Arg, B_Arg)                       '�u��������O�̕�������������
    If Arg_Point = 0 Then GoTo arg_error
    
    i = Check_Column(RFp, col())
    If i = -1 Then GoTo column_error
    Count = 0
    j = 0
    Chr_Flg = 0
    st = ""
    On Error GoTo write_error
    Do While Get_Column(RFp, col(), Data())
        If Count = 0 Then
            Print #FFp, "(" & StrConv(LeftB(Arg, Arg_Point - 1), vbUnicode)
        End If
        Count = Count + 1
        Chr_Flg = Chr_Flg + 1
        If Chr_Flg > 10 Then
            j = j + 1
            Chr_Flg = 1
            Print #FFp, st
            st = "'" & Data(1) & "',"
            If j > 9 Then
                j = 0
                Print #FFp, "'') or"
                Print #FFp, StrConv(LeftB(Arg, Arg_Point - 1), vbUnicode)
            End If
        Else
            st = st & "'" & Data(1) & "',"
        End If
    Loop
            
    If Chr_Flg > 0 Then Print #FFp, st; "''"
    If Arg_Point + LenB(B_Arg) <= LenB(Arg) Then Print #FFp, StrConv(MidB(Arg, Arg_Point + LenB(B_Arg)), vbUnicode); ")"

    On Error GoTo 0
    Replace_In = Count
    Close #RFp
    Exit Function
open_error:
    Close #RFp
    Replace_In = -1
    Exit Function
write_error:
    Close #RFp
    Replace_In = -2
    Exit Function
arg_error:
    Close #RFp
    Replace_In = -3
    Exit Function
column_error:
    Close #RFp
    Replace_In = -4
End Function

'**********************************************************************
' NAME      : replace_file(arg,b_arg,fname,fp,and_or)
' FUNCTION  : ������arg�Ɋ܂܂�镶����b_arg���t�@�C��fname���當��������
'           : ���݁Aor�ŕ��ׂ�fp�ŊJ�����t�@�C���ɏ������ށB
'           : ������ like '------%' ; ������ = '------' �Ȃǂ̌`���Ɏg�p
' NOTE      : fname�̓t���p�X�Ŏw�肷��
' INPUT     : string Arg    �u���Ώە�����(SJIS)
'           : string B_Arg  �u���O������(SJIS)
'           : string FName  ������ۑ��t�@�C����(SJIS)
'           : integer FFp   �������݃t�@�C���|�C���^
'           : string And_Or �ڑ���(and �� or)
' RETURN    : ���� ����I�� ���� �ُ�I��
'**********************************************************************
'
Function Replace_File(Arg As String, B_Arg As String, fname As String, FFp As Integer, And_Or As String) As Integer
    Dim RFp As Integer      ' �t�@�C���|�C���^
    Dim Arg_Point As Integer
    Dim col(2) As Integer
    Dim Data(2) As String
    Dim i As Integer
    Dim Count As Integer

    RFp = FreeFile(1)
    On Error GoTo open_error
    Open StrConv(fname, vbUnicode) For Input As #RFp
    On Error GoTo 0

    Arg_Point = InStrB(Arg, B_Arg)                       '�u��������O�̕�������������
    If Arg_Point = 0 Then GoTo arg_error
    
    i = Check_Column(RFp, col())
    If i = -1 Then GoTo column_error
    Count = 0
    On Error GoTo write_error
    Do While Get_Column(RFp, col(), Data())
        If Count > 0 Then
            Print #FFp, And_Or; " ";
        Else
            Print #FFp, "(  ";
        End If
        Count = Count + 1
        Print #FFp, StrConv(LeftB(Arg, Arg_Point - 1), vbUnicode); Data(1);
        If Arg_Point + LenB(B_Arg) <= LenB(Arg) Then Print #FFp, StrConv(MidB(Arg, Arg_Point + LenB(B_Arg)), vbUnicode)
    Loop

    Print #FFp, ")"
    On Error GoTo 0
    Replace_File = Count
    Close #RFp
    Exit Function
open_error:
    Close #RFp
    Replace_File = -1
    Exit Function
write_error:
    Close #RFp
    Replace_File = -2
    Exit Function
arg_error:
    Close #RFp
    Replace_File = -3
    Exit Function
column_error:
    Close #RFp
    Replace_File = -4
End Function

'**********************************************************************
' NAME      : file_merge(fname1,key1,fname2,key2,newfname)
' FUNCTION  : file1�Afile2�̂Q�t�@�C���ɂ��āA�J�����ԍ�key1�Akey2���L�[��
'�@�@�@�@�@ : ���ă}�[�W���s���A�t�@�C��newfname�Ƃ��ĕۑ�����B
' NOTE      : fname1,fname2,newfname�̓t���p�X�Ŏw�肷��B
'           : �L�[�����͂��炩���ߏ����\�[�g����Ă���K�v������B
'           : ��ɂr�p�k�������ʃt�@�C���Ŏg�p���邪�A�y�[�W�T�C�Y�ȏ��
'           : �f�[�^�ʂ�����ꍇ��؂肪�ł��邽�߂ɐ���ȓ��삪�ۏ؂���
'�@�@�@�@�@ : �Ȃ��̂Œ��ӁB
' INPUT     : string fname1    �}�[�W�Ώۃt�@�C���P
'           : integer key1     �t�@�C���P�̃L�[�J�����ԍ�
'           : string fname2    �}�[�W�Ώۃt�@�C���Q
'           : integer key2     �t�@�C���Q�̃L�[�J�����ԍ�
'           : string newfname  �}�[�W��̕ۑ��t�@�C����
' RETURN    : ���� �}�[�W�����f�[�^�� ���� �ُ�I��
'**********************************************************************
'
Function File_Merge(fname1 As String, Key1 As Integer, _
                    fname2 As String, Key2 As Integer, NewFName As String) As Integer
    Dim fp1 As Integer, Fp2 As Integer, NewFp As Integer    'File Pointer
    Dim BFR1 As String, AFT1 As String
    Dim BFR2 As String, AFT2 As String
    Dim LENBFR1 As Integer                                  '������1�����̕�����
    Dim LENBFR2 As Integer                                  '������2�����̕�����
    Dim LENARG1 As Integer                                  '�L�[������1�̕�����
    Dim LENARG2 As Integer                                  '�L�[������2�̕�����
    Dim LENAFT As Integer                                   '������1�㕔�̕�����
    Dim LENARGm As Integer                                  '�L�[������̃}�[�W��̕�����
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim St1 As String
    Dim St2 As String
    Dim LOst1 As Integer, LOst2 As Integer                  '�͂��߂̕�����̒���

    'File check,open
    If Dir(NewFName) <> "" Then Kill NewFName
    If Dir(fname1) = "" Or Dir(fname2) = "" Then GoTo file_err
    If FileLen(fname1) = 0 Or FileLen(fname2) = 0 Then GoTo file_err
    fp1 = FreeFile(1)
    Open fname1 For Input As #fp1
    Fp2 = FreeFile(1)
    Open fname2 For Input As #Fp2
    
    'FName1 Column check
    Do
        Line Input #fp1, St1
        For i = 1 To Len(St1)
            If Mid(St1, i, 1) <> "-" Or Mid(St1, i, 1) <> " " Then Exit For
        Next i
    Loop Until EOF(fp1) Or i > Len(St1)
    If EOF(fp1) And i <= Len(St1) Then
        Close #fp1, #Fp2
        GoTo file_err
    End If
    St1 = StrConv(St1, vbFromUnicode)
    LENBFR1 = 0
    If Key1 > 1 Then
        For i = 1 To Key1 - 1
            LENBFR1 = InStrB(LENBFR1 + 1, St1, " ")
            If LENBFR1 = 0 Then
                Close #fp1, #Fp2
                GoTo par_err
            End If
        Next i
    End If
    i = InStrB(LENBFR1 + 1, St1, " ")
    If i > 0 Then
        LENARG1 = i - LENBFR1
    Else
        LENARG1 = LenB(St1) - LENBFR1
    End If
    LENAFT = LenB(St1) + 1 - LENBFR1 - LENARG1
    
    'FName2 Column check
    Do
        Line Input #Fp2, St2
        For i = 1 To Len(St2)
            If Mid(St2, i, 1) <> "-" Or Mid(St2, i, 1) <> " " Then Exit For
        Next i
    Loop Until EOF(Fp2) Or i > Len(St2)
    If EOF(Fp2) And i <= Len(St2) Then
        Close #fp1, #Fp2
        GoTo file_err
    End If
    St2 = StrConv(St2, vbFromUnicode)
    LENBFR2 = 0
    If Key2 > 1 Then
        For i = 1 To Key2 - 1
            LENBFR2 = InStrB(LENBFR2 + 1, St2, " ")
            If LENBFR2 = 0 Then
                Close #fp1, Fp2
                GoTo par_err
            End If
        Next i
    End If
    i = InStrB(LENBFR2 + 1, St2, " ")
    If i > 0 Then
        LENARG2 = i - LENBFR2
    Else
        LENARG2 = LenB(St2) - LENBFR2
    End If
    LENARGm = LENARG1
    If LENARGm < LENARG2 Then LENARGm = LENARG2
    '�}�[�W�t�@�C������
    NewFp = FreeFile(1)
    Open NewFName For Output As #NewFp
    Print #NewFp, String(LENARGm - 1, "-") & " ";
    Print #NewFp, StrConv(LeftB(St1, LENBFR1), vbUnicode);
    Print #NewFp, StrConv(MidB(St1, LENBFR1 + LENARG1 + 1), vbUnicode); " ";
    Print #NewFp, StrConv(LeftB(St2, LENBFR2), vbUnicode);
    Print #NewFp, StrConv(MidB(St2, LENBFR2 + LENARG2 + 1), vbUnicode)
    '�}�[�W�J�n
    St1 = Linput(fp1)                                                '�͂��߂̕�������擾
    St2 = Linput(Fp2)
    LOst1 = LenB(St1)
    LOst2 = LenB(St2)
    File_Merge = 0
    Do While 1
        If St1 = "" Then                                             '�t�@�C���P�͏I���
            Do Until St2 = ""
                Merge_Str Trim(MidB(St2, LENBFR2 + 1, LENARG2)), LENARGm, _
                 StrConv(Space(LOst1), vbFromUnicode), LENBFR1, LENARG1, LENAFT, _
                 St2, LENBFR2, LENARG2, NewFp
                File_Merge = File_Merge + 1
                St2 = Linput(Fp2)
            Loop
            Exit Do
        End If
        If St2 = "" Then                                             '�t�@�C���Q�͏I���
            Do Until St1 = ""
                Merge_Str Trim(MidB(St1, LENBFR1 + 1, LENARG1)), LENARGm, _
                 St1, LENBFR1, LENARG1, LENAFT, StrConv(Space(LOst2), vbFromUnicode), _
                 LENBFR2, LENARG2, NewFp
                File_Merge = File_Merge + 1
                St1 = Linput(fp1)
            Loop
            Exit Do
        End If
        
        If Trim(MidB(St1, LENBFR1 + 1, LENARG1)) = _
         Trim(MidB(St2, LENBFR2 + 1, LENARG2)) Then                 '��v����!!
            Merge_Str Trim(MidB(St1, LENBFR1 + 1, LENARG1)), LENARGm, _
             St1, LENBFR1, LENARG1, LENAFT, St2, LENBFR2, LENARG2, NewFp
            St1 = Linput(fp1)
            St2 = Linput(Fp2)
        ElseIf Trim(MidB(St1, LENBFR1 + 1, LENARG1)) > _
         Trim(MidB(St2, LENBFR2 + 1, LENARG2)) Then                 '������P���ǂ��z����
            Merge_Str Trim(MidB(St2, LENBFR2 + 1, LENARG2)), LENARGm, _
             StrConv(Space(LenB(St1)), vbFromUnicode), LENBFR1, LENARG1, _
             LENAFT, St2, LENBFR2, LENARG2, NewFp
            St2 = Linput(Fp2)
        Else                                                        '������Q���ǂ��z����
            Merge_Str Trim(MidB(St1, LENBFR1 + 1, LENARG1)), LENARGm, _
             St1, LENBFR1, LENARG1, LENAFT, StrConv(Space(LenB(St2)), vbFromUnicode), _
             LENBFR2, LENARG2, NewFp
            St1 = Linput(fp1)
        End If
        File_Merge = File_Merge + 1
    Loop
    Close #fp1, #Fp2, #NewFp
    Exit Function
file_err:
par_err:
End Function

'**********************************************************************
' NAME      : merge_str(Marg,Mlen,st1,lenbfr1,lenarg1,st2,lenbfr2,lenarg2,fp)
' FUNCTION  : st1�Ast2�̃}�[�W
' NOTE      : file_merge()��p
' INPUT     : string Marg       �}�[�W�Ώە�����(SJIS)
'           : integer Mlen      �}�[�W�Ώە�����̒���
'           : string st1        �A��������P(SJIS)
'           : integer lenbfr1   st1�̃}�[�W������܂ł̒���
'           : integer lenarg1   st1���̃}�[�W�Ώە�����
'           : string st2        �A��������Q(SJIS)
'           : integer lenbfr2   st2�̃}�[�W������܂ł̒���
'           : integer lenarg2   st2���̃}�[�W�Ώە�����
'           : string fp         �������݃t�@�C���|�C���^
' RETURN    : �Ȃ�
'**********************************************************************
'
Sub Merge_Str(Marg As String, Mlen As Integer, St1 As String, _
              LB1 As Integer, LA1 As Integer, LAFT As Integer, _
              St2 As String, LB2 As Integer, LA2 As Integer, fp As Integer)
    Print #fp, StrConv(Marg, vbUnicode) & Space(Mlen - LenB(Marg));
    Print #fp, StrConv(LeftB(St1, LB1), vbUnicode);
    Print #fp, StrConv(MidB(St1, LB1 + LA1 + 1), vbUnicode);
    Print #fp, Space(LAFT - LenB(MidB(St1, LB1 + LA1 + 1)));
    Print #fp, StrConv(LeftB(St2, LB2), vbUnicode);
    Print #fp, StrConv(MidB(St2, LB2 + LA2 + 1), vbUnicode)
End Sub

'**********************************************************************
' NAME      : Linput(fp)
' FUNCTION  : �r�p�k�������ʃt�@�C���̂P�s��荞��
' NOTE      : ���R�[�h�������ANULL�l�͖���
' INPUT     : integer fp         �������݃t�@�C���|�C���^
' RETURN    : ����I���F�ǂݍ��񂾕�����(SJIS)�@EOF���FNULL�l
'**********************************************************************
'
Function Linput(Lin_fp As Integer) As String
    Dim st As String

    st = ""
    Do Until EOF(Lin_fp)
        Line Input #Lin_fp, st
        If st <> "" Then Exit Do
    Loop
    Linput = StrConv(st, vbFromUnicode)
End Function

'**********************************************************************
' NAME      : History
' FUNCTION  : �g�p�����̕ۑ�
' NOTE      :
' INPUT     : String FName      �����t�@�C��
'           : String Comment    �R�����g
' RETURN    : �Ȃ�
'**********************************************************************
'
Sub History(ByVal message As String, Optional ByVal fname As Variant)
    Const LogDir As String = "N:\DSP\HIS\"
    Const logfile As String = "HISTORY1.LOG"
    Const Max_Size_Of_LogFile As Long = 262144  '���O�t�@�C���̍ő�o�C�g��(=256K)
    
    Dim fp As Integer
    Dim hostName As String
    Dim BackUpLogFile As String

    '���O�t�@�C�����w�肳��Ă��Ȃ�������A�f�t�H���g�̃��O�t�@�C�����w�肷��
    If IsMissing(fname) Then fname = LogDir & logfile
    '�z�X�g���擾
    hostName = Env()
    If hostName = "" Then Exit Sub
    
    ' ���t���P���ɂȂ�����t�@�C���͐V�K�X�V
    BackUpLogFile = fname & "." & Format(Now(), "YYYY_MM")
    If Format(Now(), "DD") = "01" And Dir(fname) <> "" _
     And Dir(BackUpLogFile) = "" Then Name fname As BackUpLogFile
    
    '���O���L��
    On Error Resume Next
    fp = FreeFile(1)
    If Dir(fname) = "" Then
        Open fname For Output As #fp
    Else
        Open fname For Append As #fp
    End If
    Print #fp, Format(Now(), "YYYY/MM/DD HH:MM:SS"); " "; hostName; " "; message
    Close #fp
End Sub




