Attribute VB_Name = "libDBAccess"
Option Explicit
'********************************************************************************************
'  解析依頼システム - DB アクセス 共通ライブラリ モジュール
'               Copyright 2015, XXXX All Rights Reserved.
'  2015-05-14 新規作成 (M.NAKAI)
'********************************************************************************************

' ********
' 定数定義
' ********
' SQL クエリ実行結果格納用ファイル名
Public Const cSqlResFilenameP As String = "TEMSqlRes#.txt"
' SQL クエリ実行結果格納用ファイル 出力先ディレクトリ名取得用環境変数
Public Const cSqlResOutDirEnv As String = "WORKDIR"

' ********
' 広域変数
' ********
' SQL実行ログ ファイル パス
Public gSqlExecLogFilepath As String

' **********
' 構造体定義
' **********

' SQL実行変数セット
Public Type SSqlSet
    sExecSqlSheet As String   ' 実行SQLシート名
    iExecSqlNo As Long        ' 実行SQL番号(SQLシートの列番号)
    sResultFile As String     ' クエリ実行結果取得用ファイル
    rep As Object             ' SQL 置換対象配列 (Scripting.Dictionary型)
    psData() As String        ' クエリ結果格納
End Type

'**********************************************************************
' @(f)
' 機能      : ADOを利用したデータ検索
'
' 返り値    : True  ：  正常終了
' 　　　      False ：  異常終了
'
' 引き数    : String    SqlSheetName    SQL文が格納されたシート名
' 　　　      Integer   Sqlc            SQL格納カラム
' 　　　      String    Setfile         データ出力ファイル名
' 　　　      String    Acol()          SQL変換後文字列
' 　　　      String    Bcol()          SQL変換前文字列
' 　　　      String    MES             出力メッセージ内容
'
' 機能説明  :
'
' 備考      :
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
    i = 2   ' 2行目から読み込み
    Do While Workbooks(MAINBOOK).Sheets(SqlSheetName).Cells(i, sqlc).Value <> ""
        Sqlstr = Workbooks(MAINBOOK).Sheets(SqlSheetName).Cells(i, sqlc).Value ''一行読み込み
        Select Case Left(Sqlstr, 1)
        Case "#"
        Case Else
        If InStr(1, UCase(Sqlstr), "EXIT") > 0 Then
        Else
            j = LBound(bb)
            Do While (j <= UBound(bb)) And (j <= UBound(aa))
                If bb(j) <> "" Then
                    n = InStr(1, Sqlstr, bb(j))  ''文字列はあるか
                Else
                    n = 0
                End If
                If n > 0 Then
                    Select Case Left(bb(j), 1)
                    Case "\"
                        ans = Replace(Sqlstr, bb(j), aa(j))
                        Sqlstr = ans
                    Case "%"
                        Sqltmp = Left(Sqlstr, n - 1)    ''先頭の文字
                        Sqladd = Mid(Sqlstr, n + Len(bb(j)), 256)   ''残りの文字
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
                        Sqltmp = Left(Sqlstr, n - 1)    ''先頭の文字
                        Sqladd = Mid(Sqlstr, n + Len(bb(j)), 256)   ''残りの文字
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
    ' SQL実行ログを出力する。
    ' ***********************
    ' 出力先は設定シートから取得する。
    If Trim(gSqlExecLogFilepath) = "" Then
        ' 未取得の場合のみ取得する。
        gSqlExecLogFilepath = "C:\Temp\SqlExec.log"
    End If
    
    ' ファイル パスが設定されている場合のみ出力する。
    If Trim(gSqlExecLogFilepath) <> "" Then
        fp = FreeFile(1)
        Open gSqlExecLogFilepath For Output As #fp
        Print #fp, strsql
        Close #fp
    End If
    
    '' セレクトメイン(実際のデータ検索)
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
' 機能      : セレクトメイン
'
' 返り値    : 0　   正常終了
'    　　　 : -1    異常終了
'
' 引き数    : String    Strhost         接続先ホスト名
'    　　　   String    Username        接続ユーザ
'    　　　   String    Strpassword     接続パスワード
'    　　　   String    Strsql          検索ＳＱＬ
'    　　　   String    Setfile         出力ファイル名
'
' 機能説明  :
'
' 備考      :
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
    
    '<<< コネクション　オープン >>>
    Dim bStatus As Boolean
    bStatus = Application.EnableEvents
    Application.EnableEvents = False
    
    Set oraConnection = ConnectOraS(strhost, usename, strpassword)
    '<<< 日付データ取得処理 >>>
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
    '<<< コネクション　クローズ >>>
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
' 機能      : ORACLEにコネクト
'
' 返り値    : なし
'
' 引き数    : なし
'
' 機能説明  :
'
' 備考      : 接続エラー回復処理します。(5SecX60=5Min)
'             実際には接続実行の時間がありますので、５分以上でエラー
'             終結となります。
'
'**********************************************************************
Public Function ConnectOraS(StrSource As String, StrUID As String, _
                           StrPwd As String) As ADODB.Connection
    Dim cnnOpen
    Dim connectEnv   As String
    Dim errLoop      As ADODB.Error

    '' データベースに接続中
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
'            'エラーログを書くマクロが有ったが、削除
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
' 機能      : データの読み込み
'
' 返り値    : なし
'
' 引き数    : なし
'
' 機能説明  :
'
' 備考      :
'
'**********************************************************************
Public Function DataReadS(cnn As ADODB.Connection, strsql As String) _
                                                         As ADODB.Recordset
    Dim cmdChange    As ADODB.Command
    Dim errLoop      As ADODB.Error
    
    On Error GoTo Err_Execute
    
    '' データベースを検索中
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
            'エラーログを書くマクロが有ったが、削除
        Next
    End If
    cnn.Close
    Set cnn = Nothing

End Function

'**********************************************************************
' @(s)
' 機能      : (改)SQL実行準備処理
'
' 返り値    : SSqlSet : SQL実行データ セット
'
' 引き数    : Long sExecSqlNo : 実行するSQLクエリ文番号
'           : String sSqlSheetName : SQLシートの名前(初期値「SQL」)
'
' 機能説明  : SQL実行に必要なデータを用意する。この関数で出力したSQL実行
'           : データ セットの「SQL 置換対象配列(rep)」を設定後、SQL実行
'           : 処理関数(n_DoSql)を実行する。
'
' 備考      : SQL実行データ セットの内容は以下の通り。
'                sExecSqlSheet As String   ' 実行SQLシート名 ※
'                iExecSqlNo As Long        ' 実行SQL番号(SQLシートの列番号) ※
'                sResultFile As String     ' クエリ実行結果取得用ファイル ※
'                rep As Object             ' SQL 置換対象配列 (Scripting.Dictionary型)
'                psData() As String        ' クエリ結果格納
'            ※印つきの項目は本関数(n_InitSql)が値を設定するため、使用者が設定する必要
'              はない。
'            SQL置換対象配列(rep)はSQLを実行する前に、SQLシートに記載されたSQLクエリ文
'            に対して、置き換える文字列を指定する。rep はDictionary型になっており、以下
'            の通りにしていする。
'             例)
'               Dim wSql As SSqlSet
'               wSql = m_InitSql(1)
'               wSq.rep.Add "置換前", "置換後" ' 実行するSQLクエリの「置換前」という文字列を「置換後」に置き換えて、SQLクエリを実行する。
'**********************************************************************
Public Function n_InitSql(sExecSqlNo As Long, Optional sSqlSheetName As String = "SQL") As SSqlSet
    Dim ssetWk As SSqlSet       ' SQL実行データ セット
    Dim wkFname As String
    
    ' ***********************
    ' 実行SQL情報を設定する。
    ' ***********************
    MAINBOOK = ThisWorkbook.Name

    ssetWk.sExecSqlSheet = sSqlSheetName
    ssetWk.iExecSqlNo = sExecSqlNo
    
    ' ************************************
    ' クエリ実行結果ファイル名を生成する。
    ' ************************************
    ' 環境ファイルから出力先ディレクトリ名を取得する。
    gsWorkDir = GetEnv(gcEnvFile, cSqlResOutDirEnv)
    
    ' ファイル名を生成する。
    wkFname = Replace(cSqlResFilenameP, "#", CStr(sExecSqlNo))
    
    ' ディレクトリ名とファイル名を結合する。
    ssetWk.sResultFile = gsWorkDir & wkFname
    
    ' ********************************
    ' SQL クエリ分置換配列を準備する。
    ' ********************************
    Set ssetWk.rep = CreateObject("Scripting.Dictionary")
    ssetWk.rep.RemoveAll

    n_InitSql = ssetWk
    
End Function

'**********************************************************************
' @(s)
' 機能      : (改)SQL実行処理
'
' 返り値    : Long : 実行結果 または 取得件数
'                SELECT クエリ実行時は取得件数。負数の場合はエラー
'                INSERT,UPDATE,DELETE等の更新クエリの場合は、0の場合は
'                成功、負数の場合はエラーを示す。
'
' 引き数    : SSqlSet sqlDataSet SQL実行セット
'
' 機能説明  : SQL実行準備処理関数(n_InitSql)で準備したSQL実行セットの
'           : クエリを実行する。
' 備考      : SELECTクエリの場合、SQL実行セットのクエリ結果格納(psData)
'           : に文字列の2次元配列で抽出内容が格納される。
'                e.g.) wSql.psData(col, row)
'                       'col'は列(1～)、'row'は行(1～)を示す。
Public Function n_DoSql(sqlDataSet As SSqlSet) As Long
    Dim wkAcol() As String
    Dim wkBcol() As String
    Dim iCnt As Long
    Dim i As Long
    Dim key As Variant
    Dim iRes As Long
    
    On Error GoTo EH:
    
    ' 置換文字を配列に変換する。
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
        ' 置換文字がない場合は、1要素の空要素を渡す。
        ReDim wkAcol(1)
        ReDim wkBcol(1)
        wkAcol(1) = ""
        wkBcol(1) = ""
        
    End If
    
    ' SQLを実行する。
    If CallAdoSql(sqlDataSet.sExecSqlSheet, CInt(sqlDataSet.iExecSqlNo), sqlDataSet.sResultFile, _
            wkBcol(), wkAcol(), "", "") = False Then
        MsgBox "エラーが発生したため、処理が中断されました。" & vbCrLf & _
               ErNm & "：" & sEr, vbCritical + vbOKOnly, "エラー"
        Set ErNm = Nothing
        Set sEr = Nothing
        
        n_DoSql = -1
        On Error GoTo 0
        Exit Function
    End If
    
    ' 実行結果を取得する。
    iRes = f_GetData(sqlDataSet.psData, sqlDataSet.sResultFile)

#If DEBUG_MODE <> 1 Then
    ' 実行結果ファイルを削除する。
    Kill sqlDataSet.sResultFile
#End If

    n_DoSql = iRes
    Exit Function
    
EH:
        MsgBox "エラーが発生したため、処理が中断されました。" & vbCrLf & _
               "[" & Err.Number & "]" & Err.Description, vbCritical + vbOKOnly, "エラー"
        Set ErNm = Nothing
        Set sEr = Nothing
    
End Function

'**********************************************************************
' @(s)
' 機能      : 改行文字→安全な文字置換
'
' 返り値    : String: 置換後の文字列
'
' 引き数    : String sSrc : 置換対象の文字
'
' 機能説明  : 改行コードをSQLに設定できる文字に置き換える。
'
' 備考      : DBから読み出し改行コードに復元する場合は DbResumeNewLine()
'             を使用する。
'**********************************************************************
Public Function DbEscapeNewLine(sSrc) As String
    ' 改行コードを &H7F に置き換える。
    DbEscapeNewLine = Replace(sSrc, vbCrLf, Chr(cRNewLine))
    
End Function

'**********************************************************************
' @(s)
' 機能      : 改行文字復元
'
' 返り値    : String: 置換後の文字列
'
' 引き数    : String sSrc : 置換対象の文字
'
' 機能説明  : DbEscapeNewLine()で置き換えた改行コードを復元する。
'
' 備考      :
'
'**********************************************************************
Public Function DbResumeNewLine(sSrc) As String
    ' 改行コードを &H7F に置き換える。
    DbResumeNewLine = Replace(sSrc, Chr(cRNewLine), vbCrLf)
    
End Function

