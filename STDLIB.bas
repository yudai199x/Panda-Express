Attribute VB_Name = "STDLIB"
'********************************************************************************************
'   EAﾒﾆｭｰ 専用サブルーチン集
'               Copyright （株）東芝　四日市工場　生産技術推進課 All Rights Reserved.
'
'   96-03-06    履歴書き込み開始
'               ファイルオープンの時、256～511の範囲で指定（他のアプリで使用可能とするため）
'　             それに伴い、fopen,fclose,fcloseallの廃止
'               Runメソッドで使用できるようにサブルーチンを追加
'   96-03-15    Runメソッド用サブルーチンは廃止
'               get_columnの文字列格納部分を改良
'   96-03-28    get_sqlの時間表示部分を変更
'               get_sqlのモードを追加
'   96-06-05    get_sqlの検索待ちにてDoEventsによりOSに制御を渡すようにした
'   96-06-07    DoEventsによりOSに制御を渡すのをやめる
'   96-06-12    やっぱり戻す(^^;
'               wait_sqlにて、ファイルの大きさをきちんと表示するようにする
'   96-07-08    count_sqlfileの戻り値をLongにする
'   96-08-30    get_columnにて、検索結果ファイルの1行のデータがレコード分離符の長さに
'               満たなかった場合にエラーとなっていたのを修正
'   96-12-09    file_mergeにて、ファイルサイズが０の時にエラーで止まってしまうのを修正
'   96-12-24    MidB,LeftB,InstrBを使用することで、ファイル内に漢字が入っている場合の不具合を修正
'   97-02-18    Write_Logルーチン追加
'   97-05-22    get_sqlにてファイル書き込みに失敗した場合にもう一度検索を開始する
'   97-06-09    get_sqlにてファイル書き込みに失敗した場合に戻り値をFalseにする
'               get_sqlにて頭が"#!"の場合シェルスクリプト文の行頭とみなして書き込む。
'   97-06-10    get_cmd追加。
'               get_sqlでは頭が"#!"の場合も無視する。
'   97-09-11    make_sqlfileにて、コメントが30行を超えた場合それ以降の行を読まない不具合を修正
'   99-02-24    make_sqlfileにて、$,%,&を用いた置換に先立って通常の置換を行なうよう変更
' 2000-01-25    Office2000に対応
'********************************************************************************************

Public Const gcEnvFile As String = "C:\ENVIRON.DAT"

'**********************************************************************
'NAME       : check_column(fp,column())
'FUNCTION   : ファイルを現在位置から読み込んで行の頭が-の行が見つかったら、
'             その行を頭の文字から調べてスペースのカラム位置を配列にセットする。
'NOTE       : 頭が-の行が1000行以内になければエラー
'           : column()の上限index番号がカラム位置の数に満たない場合は
'           : 後ろのカラムは無視される。
'           : タブが含まれていた場合、正常なカラム位置を保証できない
'           : 主にSQL検索結果ファイルに使用
'INPUT      : integer fp       対象となるファイルのファイル番号
'           : integer column() カラム位置格納用配列
'RETURN     : 正数 : カラム数
'           : 負数 : -1 --- レコード分離符が見つからない
'           :        -2 --- col()配列数がレコード分離符の分離箇所より少ない
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

    For i = 1 To 1000                       '確認回数は1000回
        On Error Resume Next
        Line Input #fp, inData
        If EOF(fp) Then Exit For
        flg = True
        For n = 1 To Len(inData)
            If Mid(inData, n, 1) <> "-" And Mid(inData, n, 1) <> " " Then Exit For
        Next n
        If n > Len(inData) And inData <> "" Then    '全て'-'の行があった
            n = 0
            Do
                Do                                  '-をサーチする
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
                Do                                  '-以外をサーチする
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
' FUNCTION : ファイルを現在位置から読み込んでcol()配列の位置で分割して
'          : 配列に文字列としてセットする
' NOTE     : 変数の配列は１次元のみ。多次元では動作保証できない。
'          : col()には分割位置をセットしておく必要あり。また、col()=0以降の
'          : データは読み込まない。基本的にcheck_column()実行後に使用すること。
'          : NULL値は無視し、NULL値が100回連続したら異常終了
'          : ファイルエンドでこれ以上読み込めない場合異常終了
' INPUT    : integer fp 対象となるファイルのファイル番号
'          : integer col() カラム構成をセットした配列
'          : string data() セットする配列
' RETURN   : TRUE:正常終了　FALSE:異常終了
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
'           '96-3-15 変更
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
' FUNCTION  : SQLを実行する
' NOTE      : シートのSQL文は、"#"が頭にあるとき無視される。
'           : シートのSQL文１行に、複数の変数を入れてもよい。但し、$,%,&で始まる変数がある
'           : 行には他の変数を入れないこと。
'           : B_column()=""以降のデータは読まない。
'           : B_column()の変数の頭が$の場合、A_column内のﾌｧｲﾙ名を開いてget_column()によって
'           : データを読み込み、------ in ('******', ･･･ の形に変換（ﾌｧｲﾙ名はフルパス指定のこと）。
'           :   ex) hed_lotnum in ($LOTNUM.DAT) → hed_lotnum in ('111111','111112','111113'...
'           : B_column()の変数の頭が%の場合、A_column内のﾌｧｲﾙ名を開いてget_column()によって
'           : データを読み込み、------ like '******%' or または ------ = '******' or の形に変換
'           : （ﾌｧｲﾙ名はフルパス指定のこと）。
'           :   ex) hed_kndnam like '%KNDNAM.DAT%' → hed_kndnam like 'T5W33%' or ...
'           : B_column()の変数の頭が&の場合、A_column内のﾌｧｲﾙ名を開いてget_column()によって
'           : データを読み込み、------ like '******%' and または ------ = '******' and の形に変換
'           : （ﾌｧｲﾙ名はフルパス指定のこと）。
'           :   ex) hed_kndnam like '%KNDNAM.DAT%' → hed_kndnam like 'T5W33%' and ...
'           : 配列数は、最低限必要なカラム名数だけ必要。
'           : commentは、以下の特殊記号が使用できる。
'           :   %T% : 検索時間を表示(Disp Time)
'           :   %A% : 検索中のファイルの数を表示(Disp Access file)
'           : modeは文字の並びで示す。複数選択可能。
'           :    L  : 'c:\'にもSQLファイルを書き込む(Local mode)
'           :    D  : 検索は行わない(Debug mode)
'           :    P  : SQLファイル送信後、結果を待たずに終了(Pass mode)
'           :    W  : SQLファイルは送信せず、結果のみ待つ(Wait mode)
' INPUT     : string sqlsheet   SQLシート名（NULL時は"SQL"）
'           : integer sqlcolumn SQL分のカラム位置
'           : string outfile    出力ファイル名
'           : string B_column() 置き換え前のカラム名配列
'           : string A_column() 置き換え後のカラム名配列
'           : string comment    検索中コメント
'           : string mode       各種モード
' RETURN    : TRUE 正常終了 FALSE 環境変数HOSTNAMEの未登録
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
' FUNCTION  : シェルスクリプトを実行する
' NOTE      : シートのシェルスクリプト文は、"#"が頭にあるとき無視される。但し、"#!"の場合はシェル
'           : スクリプト文の行頭とみなしてそのまま書き込む。
'           : シートのSQL文１行に、複数の変数を入れてもよい。
'           : B_column()=""以降のデータは読まない。
'           : 配列数は、最低限必要なカラム名数だけ必要。
'           : commentは、以下の特殊記号が使用できる。
'           :   %T% : 検索時間を表示(Disp Time)
'           :   %A% : 検索中のファイルの数を表示(Disp Access file)
'           : modeは文字の並びで示す。複数選択可能。
'           :    L  : 'c:\'にもシェルスクリプトを書き込む(Local mode)
'           :    D  : 何も行わない(Debug mode)
' INPUT     : string sqlsheet   シェルスクリプトシート名（NULL時は"CMD"）
'           : integer sqlcolumn シェルスクリプトのカラム位置
'           : string B_column() 置き換え前のカラム名配列
'           : string A_column() 置き換え後のカラム名配列
'           : string mode       各種モード
' RETURN    : TRUE 正常終了 FALSE 環境変数HOSTNAMEの未登録
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
' FUNCTION  : SQL検索ファイルを作成する
' NOTE      : get_sql()専用
' INPUT     : string sht        SQLシート名
'           : integer col       SQL文のカラム位置
'           : string B_col()    置き換え前のカラム名配列
'           : string A_col()    置き換え後のカラム名配列
'           : string mode       ローカルにファイルを書き込むフラグ
' RETURN    : TRUE 正常終了 FALSE 異常終了
' HISTORY   : 96-04-23 置換前文字列の一部が他の文字列と一致した場合に生じる不具合の対処
'           :          置換前文字列が一致した場合異常終了
'**********************************************************************
'
Function Make_SQLFile(sht As String, col As Integer, B_Column() As String, _
                      A_Column() As String, mode As Boolean, SqlRcd As String, FstRcd As String) As String
    Dim LineData As String
    Dim host As String
    Dim LineCtr As Integer      ' シートの行番号
    Dim A_UB, B_UB As Integer   ' b_col,a_colの配列数
    Dim MAX_UB As Integer       ' b_col,a_colの配列数
    Dim a_col() As String, b_col() As String
    Dim fp As Integer
    Dim fpd As Integer          ' デバッグ用
    Dim i, j As Integer
    Dim st As String

    '配列の数を確認
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
    '置換前文字列の重複がないか確認
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
    '置換前文字数の確認と置換文字列の代入
    If B_UB > 0 Then
        ReDim b_col(B_UB)
        ReDim a_col(B_UB)
        For i = 1 To B_UB
            b_col(i) = StrConv(B_Column(i), vbFromUnicode)
            a_col(i) = StrConv(A_Column(i), vbFromUnicode)
        Next i
        '置換文字列を文字数の大きい順に並べ替える。
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
                If InStrB(LineData, b_col(i)) <> 0 Then   ' b_col(i)が含まれていたら
                    If LeftB(b_col(i), 1) <> StrConv("$", vbFromUnicode) And _
                     LeftB(b_col(i), 1) <> StrConv("%", vbFromUnicode) And _
                     LeftB(b_col(i), 1) <> StrConv("&", vbFromUnicode) Then
                        LineData = Replace(LineData, b_col(i), a_col(i))
                    End If
                End If
            Next i
            For i = 1 To B_UB
                If InStrB(LineData, b_col(i)) <> 0 Then   ' b_col(i)が含まれていたら
                    If LeftB(b_col(i), 1) = StrConv("$", vbFromUnicode) Then      ' $･･･ の場合
                        If mode Then
                            j = Replace_In(LineData, b_col(i), a_col(i), fpd)
                            If j < 0 Then GoTo exit_failure
                        End If
                        j = Replace_In(LineData, b_col(i), a_col(i), fp)
                        If j < 0 Then GoTo exit_failure
                        LineData = ""
                        Exit For
                    ElseIf LeftB(b_col(i), 1) = StrConv("%", vbFromUnicode) Then      ' %･･･ の場合
                        If mode Then
                            j = Replace_File(LineData, b_col(i), a_col(i), fpd, "or")
                            If j < 0 Then GoTo exit_failure
                        End If
                        j = Replace_File(LineData, b_col(i), a_col(i), fp, "or")
                        If j < 0 Then GoTo exit_failure
                        LineData = ""
                        Exit For
                    ElseIf LeftB(b_col(i), 1) = StrConv("&", vbFromUnicode) Then      ' &･･･ の場合
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
' FUNCTION  : シェルスクリプトを作成する
' NOTE      : get_cmd()専用
' INPUT     : string sht        SQLシート名
'           : integer col       SQL文のカラム位置
'           : string B_col()    置き換え前のカラム名配列
'           : string A_col()    置き換え後のカラム名配列
'           : string mode       ローカルにファイルを書き込むフラグ
' RETURN    : TRUE 正常終了 FALSE 異常終了
' HISTORY   : 96-04-23 置換前文字列の一部が他の文字列と一致した場合に生じる不具合の対処
'           :          置換前文字列が一致した場合異常終了
'**********************************************************************
'
Function Make_CMDFile(sht As String, col As Integer, B_Column() As String, _
                      A_Column() As String, mode As Boolean) As String
    Dim LineData As String
    Dim host As String
    Dim LineCtr As Integer      ' シートの行番号
    Dim A_UB, B_UB As Integer   ' b_col,a_colの配列数
    Dim MAX_UB As Integer       ' b_col,a_colの配列数
    Dim a_col() As String, b_col() As String
    Dim fp As Integer
    Dim fpd As Integer          ' デバッグ用
    Dim i, j As Integer
    Dim st As String
    
    '配列の数を確認
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
    '置換前文字列の重複がないか確認
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
    '置換前文字数の確認と置換文字列の代入
    If B_UB > 0 Then
        ReDim b_col(B_UB)
        ReDim a_col(B_UB)
        For i = 1 To B_UB
            b_col(i) = StrConv(B_Column(i), vbFromUnicode)
            a_col(i) = StrConv(A_Column(i), vbFromUnicode)
        Next i
        '置換文字列を文字数の大きい順に並べ替える。
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
                If InStrB(LineData, b_col(i)) <> 0 Then   ' b_col(i)が含まれていたら
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
' FUNCTION  : SQL検索の終了を待つ
' NOTE      : get_sql()専用
'           : 一応get_cmd()実行後にファイルをチェックすることも可能
' INPUT     : string fname      ファイル名
'           : string comment    コメント
' RETURN    : Long fsize        ﾌｧｲﾙｻｲｽﾞ(Bytes)
' HISTORY   : 97-06-10 ﾌｧｲﾙ名に拡張子も付ける
'           : 97-07-25 検索後ExcelをActiveにするのをやめる
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
            Application.StatusBar = "検索結果ファイルを作成しています(" & fl & ")"
        Loop Until Format(ToTime, "hh:mm:ss") < Format(Now(), "hh:mm:ss")
        fl = FileLen("o:\" & fname)
    Loop
    fl = FileLen("o:\" & fname)
    Application.StatusBar = "検索結果ファイルを作成しています(" & fl & ")"
    Wait_SQL = fl
End Function
    
'**********************************************************************
' NAME      : count_sqlfile
' FUNCTION  : SQLファイルの数をカウントし、ファイルの体裁を整える。
' NOTE      : 途中にあるタイトル、レコード分離符はカットする。
'             行が複数に別れている場合、1行にまとめる。但し、10行以上に
'           　別れている場合はエラー。
' INPUT     : string filename ファイル名
' RETURN    : 正数：データ数
'              -1 ：レコード分離符が見つからない。
'              -2 ：対象ファイルが見つからない。
'              -3 ：行数が10を超えた
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
    Dim Tcount As Long          'タイトルの行数
    Dim ErrFlg As Boolean

    If Dir(fname) = "" Then GoTo nofileerr
    fp = FreeFile(1)
    Open fname For Input As #fp
    'FNameがTempFileと同名だった場合の対処付き
    Do
        TempFile = Env("WORKDIR") & Format(Now(), "YYYYMMDDHHMMSS") & ".TMP"
    Loop While fname = TempFile
    If Dir(TempFile) <> "" Then Kill TempFile
    Fpt = FreeFile(1)
    Open TempFile For Output As #Fpt

    'タイトル行の取得
    Tcount = 1
    On Error Resume Next
    For i = 1 To 1000                                   '確認回数は1000回
        Title(Tcount) = StrConv(inData, vbFromUnicode)
        Line Input #fp, inData
        If EOF(fp) Then Exit For
        If Left(Trim(inData), 1) = "-" Then Exit For            '-行があった
    Next i
    If i = 1001 Or EOF(fp) Then GoTo searcherr
    Separater(Tcount) = StrConv(inData, vbFromUnicode)
    Do
        Line Input #fp, inData
        Title(Tcount + 1) = StrConv(inData, vbFromUnicode)
        If EOF(fp) Then Exit Do
        Line Input #fp, inData
        Separater(Tcount + 1) = StrConv(inData, vbFromUnicode)
        If Left(Trim(inData), 1) <> "-" Then Exit Do            '-行でない
        Tcount = Tcount + 1
    Loop Until Tcount > 11
    If Tcount > 11 Then GoTo titleerr       '10行を超えた
    'Tcount = Tcount - 1
    Close #fp
    
    '再度ファイルを開き、タイトル行を読み飛ばす
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
    'タイトル、レコード分離符
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
    'データ取り込み
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
' FUNCTION  : 環境ファイルから環境変数を取り込む
' NOTE      : Envの場合、引数によって以下の様に仕様が変わる
'           :  引数なし･･･ENVFILEからHOSTNAMEで定義された環境変数を取り込む
'           :  引数１つ･･･ENVFILEから引数で指定した環境変数を取り込む
'           :  引数２つ･･･GetEnvと同様の仕様
' INPUT     : String EnvFileName 環境ファイル名
'           : String EnvStr      環境変数名
' RETURN    : 環境変数(ない場合はNULL)
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
    Dim n As Integer        ' 現在注目している要素
    Dim Data As String      ' ファイルから読み込んだ１行
    Dim fp As Integer       ' file number
    Dim st As String        ' 汎用文字列

    GetEnv = ""
    On Error GoTo Err
    fp = FreeFile(1)
    Open EnvFileName For Input As #fp       '環境ファイルのオープン
    On Error GoTo 0

    Do Until EOF(fp)
        Line Input #fp, Data
        If InStr(Data, EnvStr) <> 0 Then                '環境変数検索
            n = 0
            Do                                          'ｽﾍﾟｰｽorﾀﾌﾞを探す
                n = n + 1
                st = Mid(Data, n, 1)
            Loop Until st = " " Or st = Chr(9) Or st = Chr(13) Or st = ""
            If st = Chr(13) Or st = "" Then GoTo Err
            Do                                          'ｽﾍﾟｰｽorﾀﾌﾞをスキップ
                n = n + 1
                st = Mid(Data, n, 1)
            Loop Until st <> " " And st <> Chr(9) And st <> Chr(13) And st <> ""
            If st = Chr(13) Or st = "" Then GoTo Err
            Do                                          'ｽﾍﾟｰｽorﾀﾌﾞをスキップ
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
' FUNCTION  : 文字列の指定部分を置換
' NOTE      :
' INPUT     : string arg1：対象の文字列(SJIS)
'           : string arg2：置換前の文字列(SJIS)
'           : string arg3：置換後の文字列(SJIS)
' RETURN    : 置き換えた文字列（検索文字がなかった場合にはそのまま）
'**********************************************************************
'
Function Replace(ByVal Arg1 As String, Arg2 As String, Arg3 As String) As String
    Dim n As Integer
    Dim Arg_Left As String
    Dim Arg_Right As String
    Dim Arg_Ret As String
    Dim Len_Arg1 As Integer
    Dim Len_Arg2 As Integer

    If Arg2 = Arg3 Then         '置換文字列が同じなら
        Replace = Arg1          '戻り値は変化しない
        Exit Function
    End If

    If Arg1 = Arg2 Then         '対象文字列と置換前の文字列が同じなら
        Replace = Arg3          '戻り値はarg3
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
' FUNCTION  : 文字列argに含まれる文字列b_argを、ファイルfnameから文字列を取り
'           : 込み"'"で囲んで ","で区切った文字列に置換して、ファイルポインタ
'           : fpで開いたファイルに書き込む。
'           : ･･････ in (------) 形式のSQL文に使用
' NOTE      : fnameはフルパスで指定する
' INPUT     : string Arg    置換対象文字列(SJIS)
'           : string B_Arg  置換前文字列(SJIS)
'           : string FName  文字列保存ファイル名(SJIS)
'           : integer FFp   書き込みファイルポインタ
' RETURN    : 正数 正常終了 負数 異常終了
'**********************************************************************
'
Function Replace_In(Arg As String, B_Arg As String, fname As String, FFp As Integer) As Integer
    Dim RFp As Integer      ' ファイルポインタ
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

    Arg_Point = InStrB(Arg, B_Arg)                       '置換文字列前の部分を書き込む
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
' FUNCTION  : 文字列argに含まれる文字列b_argをファイルfnameから文字列を取り
'           : 込み、orで並べてfpで開いたファイルに書き込む。
'           : ･･････ like '------%' ; ･･････ = '------' などの形式に使用
' NOTE      : fnameはフルパスで指定する
' INPUT     : string Arg    置換対象文字列(SJIS)
'           : string B_Arg  置換前文字列(SJIS)
'           : string FName  文字列保存ファイル名(SJIS)
'           : integer FFp   書き込みファイルポインタ
'           : string And_Or 接続詞(and か or)
' RETURN    : 正数 正常終了 負数 異常終了
'**********************************************************************
'
Function Replace_File(Arg As String, B_Arg As String, fname As String, FFp As Integer, And_Or As String) As Integer
    Dim RFp As Integer      ' ファイルポインタ
    Dim Arg_Point As Integer
    Dim col(2) As Integer
    Dim Data(2) As String
    Dim i As Integer
    Dim Count As Integer

    RFp = FreeFile(1)
    On Error GoTo open_error
    Open StrConv(fname, vbUnicode) For Input As #RFp
    On Error GoTo 0

    Arg_Point = InStrB(Arg, B_Arg)                       '置換文字列前の部分を書き込む
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
' FUNCTION  : file1、file2の２ファイルについて、カラム番号key1、key2をキーと
'　　　　　 : してマージを行い、ファイルnewfnameとして保存する。
' NOTE      : fname1,fname2,newfnameはフルパスで指定する。
'           : キー部分はあらかじめ昇順ソートされている必要がある。
'           : 主にＳＱＬ検索結果ファイルで使用するが、ページサイズ以上の
'           : データ量がある場合区切りができるために正常な動作が保証され
'　　　　　 : ないので注意。
' INPUT     : string fname1    マージ対象ファイル１
'           : integer key1     ファイル１のキーカラム番号
'           : string fname2    マージ対象ファイル２
'           : integer key2     ファイル２のキーカラム番号
'           : string newfname  マージ後の保存ファイル名
' RETURN    : 正数 マージしたデータ数 負数 異常終了
'**********************************************************************
'
Function File_Merge(fname1 As String, Key1 As Integer, _
                    fname2 As String, Key2 As Integer, NewFName As String) As Integer
    Dim fp1 As Integer, Fp2 As Integer, NewFp As Integer    'File Pointer
    Dim BFR1 As String, AFT1 As String
    Dim BFR2 As String, AFT2 As String
    Dim LENBFR1 As Integer                                  '文字列1頭部の文字数
    Dim LENBFR2 As Integer                                  '文字列2頭部の文字数
    Dim LENARG1 As Integer                                  'キー文字列1の文字数
    Dim LENARG2 As Integer                                  'キー文字列2の文字数
    Dim LENAFT As Integer                                   '文字列1後部の文字数
    Dim LENARGm As Integer                                  'キー文字列のマージ後の文字数
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim St1 As String
    Dim St2 As String
    Dim LOst1 As Integer, LOst2 As Integer                  'はじめの文字列の長さ

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
    'マージファイル準備
    NewFp = FreeFile(1)
    Open NewFName For Output As #NewFp
    Print #NewFp, String(LENARGm - 1, "-") & " ";
    Print #NewFp, StrConv(LeftB(St1, LENBFR1), vbUnicode);
    Print #NewFp, StrConv(MidB(St1, LENBFR1 + LENARG1 + 1), vbUnicode); " ";
    Print #NewFp, StrConv(LeftB(St2, LENBFR2), vbUnicode);
    Print #NewFp, StrConv(MidB(St2, LENBFR2 + LENARG2 + 1), vbUnicode)
    'マージ開始
    St1 = Linput(fp1)                                                'はじめの文字列を取得
    St2 = Linput(Fp2)
    LOst1 = LenB(St1)
    LOst2 = LenB(St2)
    File_Merge = 0
    Do While 1
        If St1 = "" Then                                             'ファイル１は終わり
            Do Until St2 = ""
                Merge_Str Trim(MidB(St2, LENBFR2 + 1, LENARG2)), LENARGm, _
                 StrConv(Space(LOst1), vbFromUnicode), LENBFR1, LENARG1, LENAFT, _
                 St2, LENBFR2, LENARG2, NewFp
                File_Merge = File_Merge + 1
                St2 = Linput(Fp2)
            Loop
            Exit Do
        End If
        If St2 = "" Then                                             'ファイル２は終わり
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
         Trim(MidB(St2, LENBFR2 + 1, LENARG2)) Then                 '一致した!!
            Merge_Str Trim(MidB(St1, LENBFR1 + 1, LENARG1)), LENARGm, _
             St1, LENBFR1, LENARG1, LENAFT, St2, LENBFR2, LENARG2, NewFp
            St1 = Linput(fp1)
            St2 = Linput(Fp2)
        ElseIf Trim(MidB(St1, LENBFR1 + 1, LENARG1)) > _
         Trim(MidB(St2, LENBFR2 + 1, LENARG2)) Then                 '文字列１が追い越した
            Merge_Str Trim(MidB(St2, LENBFR2 + 1, LENARG2)), LENARGm, _
             StrConv(Space(LenB(St1)), vbFromUnicode), LENBFR1, LENARG1, _
             LENAFT, St2, LENBFR2, LENARG2, NewFp
            St2 = Linput(Fp2)
        Else                                                        '文字列２が追い越した
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
' FUNCTION  : st1、st2のマージ
' NOTE      : file_merge()専用
' INPUT     : string Marg       マージ対象文字列(SJIS)
'           : integer Mlen      マージ対象文字列の長さ
'           : string st1        連結文字列１(SJIS)
'           : integer lenbfr1   st1のマージ文字列までの長さ
'           : integer lenarg1   st1内のマージ対象文字列長
'           : string st2        連結文字列２(SJIS)
'           : integer lenbfr2   st2のマージ文字列までの長さ
'           : integer lenarg2   st2内のマージ対象文字列長
'           : string fp         書き込みファイルポインタ
' RETURN    : なし
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
' FUNCTION  : ＳＱＬ検索結果ファイルの１行取り込み
' NOTE      : レコード分離符、NULL値は無視
' INPUT     : integer fp         書き込みファイルポインタ
' RETURN    : 正常終了：読み込んだ文字列(SJIS)　EOF時：NULL値
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
' FUNCTION  : 使用履歴の保存
' NOTE      :
' INPUT     : String FName      履歴ファイル
'           : String Comment    コメント
' RETURN    : なし
'**********************************************************************
'
Sub History(ByVal message As String, Optional ByVal fname As Variant)
    Const LogDir As String = "N:\DSP\HIS\"
    Const logfile As String = "HISTORY1.LOG"
    Const Max_Size_Of_LogFile As Long = 262144  'ログファイルの最大バイト数(=256K)
    
    Dim fp As Integer
    Dim hostName As String
    Dim BackUpLogFile As String

    'ログファイルが指定されていなかったら、デフォルトのログファイルを指定する
    If IsMissing(fname) Then fname = LogDir & logfile
    'ホスト名取得
    hostName = Env()
    If hostName = "" Then Exit Sub
    
    ' 日付が１日になったらファイルは新規更新
    BackUpLogFile = fname & "." & Format(Now(), "YYYY_MM")
    If Format(Now(), "DD") = "01" And Dir(fname) <> "" _
     And Dir(BackUpLogFile) = "" Then Name fname As BackUpLogFile
    
    'ログを記入
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




