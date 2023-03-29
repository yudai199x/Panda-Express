Attribute VB_Name = "SystemCommon"
Option Explicit
Option Base 1


Public ErNm             As Variant      'Err.Numberを格納
Public sEr              As Variant      'Err.Descriptionを格納
Public iRet             As Integer      '戻り値受け取り

Public Const cSheetMain As String = "Main"

Private Const TNS_NAME As String = "MACSDB5A1"

'********************************************************************************************
'  解析依頼システム - システム共通モジュール
'               Copyright 2015, XXXX All Rights Reserved.
'  2015-05-14 新規作成
'********************************************************************************************

'**********************************************************************
'     Win32API（2015/10/15追加）
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
' 定数定義
' ********
' SQL クエリ実行結果格納用ファイル名
Public Const cSqlResFilenameP As String = "TEMSqlRes#.txt"
' SQL クエリ実行結果格納用ファイル 出力先ディレクトリ名取得用環境変数
Public Const cSqlResOutDirEnv As String = "WORKDIR"

'シート名(共通)
Public Const cSheetSql       As String = "SQL"
Public Const cSheetSetup     As String = "SETUP"
Public Const cSheetTsbMenu   As String = "依頼者Topメニュー"
Public Const cSheetTnaMenu   As String = "TNA Topメニュー"
Public Const cSheetAdminMenu   As String = "管理者メニュー"
Public Const cSheetSplash As String = "解析依頼システム"



' 事業所名
Public Const cTsb As String = "TSB"
Public Const cTna As String = "TNA"
Public Const cAim As String = "AIM"

' 特殊部課
Public Const cPg1 As String = "（P技一）"
Public Const cPg2 As String = "（P技二）"

' 権限種別
Public Const cPermitTsbGeneral As String = "0"      ' TSB 一般
Public Const cPermitTsbTheme As String = "1"        ' TSB テーマ長(参事、主務)
Public Const cPermitTsbGroup As String = "2"        ' TSB グループ長(課長)
Public Const cPermitTsbAim As String = "3"          ' TSB 解析技長(AIM)
Public Const cPermitTnaGeneral As String = "0"      ' TNA 一般
Public Const cPermitTnaSection As String = "1"      ' TNA 課長

' 承認者種別名
Public Const cApproveTheme As String = "テーマ長"
Public Const cApproveGroup As String = "グループ長"
Public Const cApproveTna As String = "上長"

' 状態名
Public Const cStatusCreateNew As String = "(CREATE_NEW)"            '  -:(新規)
Public Const cStatusThmApproveWait As String = "THM_APPROVE_WAIT"   '  1:テーマ長承認待ち
Public Const cStatusEstimateReqWait As String = "ESTIMATE_REQ_WAIT" '  2:見積もり依頼待ち
Public Const cStatusEstimating As String = "ESTIMATING"             '  3:依頼見積もり中
Public Const cStatusTnaApproveWait As String = "TNA_APPROVE_WAIT"   '  4:TNA上長承認待ち
Public Const cStatusTsbApproveWait As String = "TSB_APPROVE_WAIT"   '  5:TSB上長承認待ち
Public Const cStatusAimApproveWait As String = "AIM_APPROVE_WAIT"   '  6:AIM承認待ち
Public Const cStatusReqReceptWait As String = "REQ_RECEPT_WAIT"     '  7:依頼受付待ち
Public Const cStatusChkResultWait As String = "CHK_RESULT_WAIT"     '  8:観察結果確認待ち
Public Const cStatusWorkEndWait As String = "WORK_END_WAIT"         '  9:作業完了待ち
Public Const cStatusDone As String = "DONE"                         ' 10:作業完了
Public Const cStatusTsbCanceled As String = "TSB_CANCELED"          ' 11:TSB取り消し済み
Public Const cStatusTnaCanceled As String = "TNA_CANCELED"          ' 12:TNA取り消し済み
Public Const cStatusAimCanceled As String = "AIM_CANCELED"          ' 13:AIM取り消し済み

Public Const cRNewLine As Integer = 28  ' DB登録時改行コード置換文字

' フラグ
Public Const cFlgFalse As String = "0"
Public Const cFlgTrue As String = "1"

' **********
' 列挙体定義
' **********

' 起動ユーザー種別
Public Enum EExecuteUserMode
    eumTsbMenu = 0          ' TSB用メニュー
    eumTnaMenu = 1          ' TNA用メニュー
    eumAdminMenu = 2        ' 管理者用メニュー
    eumThemeApprove = 3     ' テーマ長承認画面
    eumSectionApprove = 4   ' 課長承認画面
    eumAimApprove = 5       ' AIM承認画面
    enmTnaApprove = 6       ' TNA上長認証画面
    enmOrderConfirm = 7     ' 調達発注確認画面
End Enum

' ************
' 広域変数定義
' ************


Public gsWorkDir    As String
Public gsFileNM     As String
Public Acol()       As String
Public Bcol()       As String

Public MAINBOOK     As String
' ## Public KenSu        As Integer             '検索結果数
Public psData()     As String

Public gExecuteUserMode As EExecuteUserMode ' 起動ユーザー種別

Public gbLogin As Boolean        ' ログイン結果 True=成功、False=失敗
Public gsLoginUserId As String   ' ログインしたユーザーID
Public gsLoginUserName As String ' ログインしたユーザー名前
Public gsLoginUserDiv As String  ' ログインしたユーザーの事業所
Public gsLoginUserCls As String  ' ログインしたユーザーの役職
Public gbAdminFlg As Boolean     ' 管理者フラグ True=ON、False=OFF

Public previousSheetName As String ' 前回表示した画面名(=メニュー)

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
' 機能      : 設定シート(SETUP)から、設定値を取得する。
'
' 返り値    : Variant : 取得した値
'
' 引き数    : String    strName         取得する設定値名
'
' 機能説明  :
'
' 備考      :
'
'**********************************************************************
Public Function GetSetup(strName As String) As Variant
    
    Dim wksheet As Worksheet
    Dim rngFind As Range
    On Error GoTo EH:
    
    ' 設定シート
    Set wksheet = ThisWorkbook.Worksheets(cSheetSetup)
    
    ' 2列目の設定値名列から、引数の設定値名と同じ値のセルを検索する。
    Set rngFind = wksheet.Columns(2).Find(What:=strName, LookAt:=xlWhole, MatchCase:=True)
    
    ' 見つけたセルの右の値を返す。
    GetSetup = wksheet.Cells(rngFind.row, 3).Value

    Exit Function
EH:
    ' エラーが発生した場合
        ' 一致しない場合
        MsgBox "設定シート (" & cSheetSetup & ") から設定値 " & strName & "が取得できません。", vbCritical + vbOKOnly, "エラー"
        GetSetup = ""
End Function

'**********************************************************************
' @(f)
' 機能      : 設定シート(SETUP)から、設定値のセルを取得する。
'
' 返り値    : Range : 取得したセル(Range)
'
' 引き数    : String    strName         取得する設定値名
'
' 機能説明  :
'
' 備考      :
'
'**********************************************************************
Public Function GetSetupCell(strName As String) As Range
    
    Dim wksheet As Worksheet
    Dim rngFind As Range
    On Error GoTo EH:
    
    ' 設定シート
    Set wksheet = ThisWorkbook.Worksheets(cSheetSetup)
    
    ' 2列目の設定値名列から、引数の設定値名と同じ値のセルを検索する。
    Set rngFind = wksheet.Columns(2).Find(What:=strName, LookAt:=xlWhole, MatchCase:=True)
    
    ' 見つけたセルの右のセルの参照を返す
    Set GetSetupCell = wksheet.Cells(rngFind.row, 3)

    Exit Function
EH:
    ' エラーが発生した場合
        ' 一致しない場合
        MsgBox "設定シート (" & cSheetSetup & ") から設定値 " & strName & "が取得できません。", vbCritical + vbOKOnly, "エラー"
        Set GetSetupCell = Nothing
End Function


'**********************************************************************
' @(f)
' 機能      : 設定シート(SETUP)の指定項目の設定値を設定する。
'
' 返り値    : Boolean : 処理結果(True=成功/False=エラー)
'
' 引き数    : String    strName         取得する設定値名
'             Variant   value           設定する値
'
' 機能説明  :
'
' 備考      :
'
'**********************************************************************
Public Function SetSetup(strName As String, Value As Variant) As Boolean
    
    Dim wksheet As Worksheet
    Dim rngFind As Range
    
    ' 設定シート
    Set wksheet = ThisWorkbook.Worksheets(cSheetSetup)
    
    ' 2列目の設定値名列から、引数の設定値名と同じ値のセルを検索する。
    Set rngFind = wksheet.Columns(2).Find(What:=strName, LookAt:=xlWhole, MatchCase:=True)
    
    If rngFind <> Empty Then
        ' 見つけたセルの値を更新する。
        wksheet.Cells(rngFind.row, 3).Value = Value
    Else
        ' 一致しない場合
        MsgBox "設定シート (" & cSheetSetup & ") の設定値 " & strName & "が設定できません。", vbCritical + vbOKOnly, "エラー"
        
        SetSetup = False
    End If

End Function






'▼*************************************************************▼
' 機能      : SQL実行 事前処理
'             出力テキストファイル名、条件変数
' 返り値    : 出力ファイル名（フルパス）
' 引き数    : 条件変数の配列数、出力フォルダの設定名、出力ファイル名
' 機能説明  :
' 備考      :
'▼*************************************************************▼
Public Function f_SqlInit(iWhere As Integer, sDirNM As String, sFileNM As String) As String
    
    MAINBOOK = ThisWorkbook.Name
    
    Erase Acol()
    Erase Bcol()
    
    ReDim Acol(iWhere)
    ReDim Bcol(iWhere)
    
    gsWorkDir = GetEnv(gcEnvFile, sDirNM)
    f_SqlInit = gsWorkDir & sFileNM
    
End Function

'▼*************************************************************▼
' 機能      : SQL実行後のテキストファイルデータを変数に格納
' 返り値    : データ件数（データ）
' 引き数    : 格納変数、データファイル名
' 機能説明  :
' 備考      : 格納変数(0)にはデータは入らず、(1)からデータセットする
'▼*************************************************************▼
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






'-------▼シートクリア▼-------
'　引数　　　　　：cName  対象シート名　iStRow 開始Row　iStcol 開始Col
'　引数（省略可）：iMode　モード（0 Delete　1 ClearContents）
'　　　　　　　　：sUpLeft　Delete時のシフト方向 省略時はUp,Up以外の何かしらが入ってるとLeft
'　　　　　　　　：iEnRow 終了Row       iEnCol 終了Col
'　返り値　　　　：Integer型　0　失敗　 1　成功　　-1　消去データ無

Public Function f_SheetClear(cName As String, iStRow As Integer, iStCol As Integer, Optional iMode As Integer = 0, _
                             Optional sUpLeft As String = "Up", Optional lEnRow As Long = 0, Optional lEnCol As Long = 0) As Integer

    On Error GoTo Errtrap
    
    With ThisWorkbook.Worksheets(cName)
    
        '落ちることがあるので保護解除
        .Unprotect
        
        '引数省略時は始点以降カウント
        If lEnRow <= 0 Then
            lEnRow = .Cells(Rows.Count, iStCol).End(xlUp).row
        End If
        
        If lEnCol <= 0 Then
            lEnCol = .Cells(iStRow, Columns.Count).End(xlToLeft).Column
        End If
        
        If iStRow > lEnRow Or iStCol > lEnCol Then
            f_SheetClear = -1 '消去データ無
            Exit Function
        End If
        
        'sUpLeft 省略時はUpにする
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
        
        f_SheetClear = 1 '成功
        
        iRet = .UsedRange.Resize(1, 1).Rows.Count
        
        Exit Function
        
Errtrap:
        
    f_SheetClear = 0 'エラー（失敗）
        
    End With
End Function


'-------▼Null値を文字列に変換▼-------
'　引数　：Target 対象のValue　RepStr 置換後の文字列（省略可）
'　返り値：String型　置換後かもとの文字列を返す

Public Function NVL(Target As Variant, Optional RepStr As String = "") As String

    If IsNull(Target) = True Then
        NVL = RepStr
    Else
        NVL = CStr(Target)
    End If

End Function


'-------▼罫線を引く▼-------
'　引数　　　　　：cName 　　 対象シート名
'　　　　　　　　：iBArea     線を引く場所　(0 全体　1 外枠のみ　2 内側縦　3内側横　4 内側全体)
'　　　　　　　　：iLpat    　線の種類（0 線なし　1 Hairline　2 Thin　3 Medium）
'　　　　　　　　：iStRow     始点Row　　　iStCol　　始点Col
'　引数（省略可）：lEnRow　   終点Row　　　lEnCol　　終点Col　　　lColor　　線の色（省略時黒）
'　返り値　　　　：Integer型　0　失敗　 1　成功

Public Function BorderLiner(cName As String, iBArea As Integer, iLpat As Integer, iStRow As Integer, iStCol As Integer, _
                            Optional lEnRow As Long = 0, Optional lEnCol As Long = 0, Optional lColor As Long = 0)
    Dim iLine   As Integer
    Dim oRange  As Range
    
    On Error GoTo Errtrap
    
    With ThisWorkbook.Worksheets(cName)
    
    '引数省略時は始点以降カウント
    If lEnRow = 0 Then
        lEnRow = .Cells(Rows.Count, iStCol).End(xlUp).row
    End If
    
    If lEnCol = 0 Then
        lEnCol = .Cells(iStRow, Columns.Count).End(xlToLeft).Column
    End If
    
    '線の色を確定（省略時は黒）
    If lColor = 0 Then
        lColor = RGB(0, 0, 0)
    End If
    
    
    '線の種類を確定
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
    
    '処理範囲を確定
    Set oRange = .Range(.Cells(iStRow, iStCol).Address(False, False), _
                        .Cells(lEnRow, lEnCol).Address(False, False))
    
    '罫線
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

'-------▼ステータス名置換▼-------

Public Function Status2Name(sStatus As String) As String
        Select Case sStatus
            Case Is = ""
                Status2Name = "未受入"
            Case Is = "REQ_CANCEL"
                Status2Name = "依頼取り消し"
            Case Is = "REQ_ACCEPT"
                Status2Name = "依頼受入"
            Case Is = "SEM_WORKING"
                Status2Name = "SEM作業中"
            Case Is = "SEM_CHECKING"
                Status2Name = "SEM観察確認待ち"
            Case Is = "SEM_MEAS_WAIT"
                Status2Name = "測長待ち"
            Case Is = "SEM_MEASURING"
                Status2Name = "測長中"
            Case Is = "SEM_REBUILD_WAIT"
                Status2Name = "復元待ち"
            Case Is = "SEM_REBUILD_CHECKING"
                Status2Name = "復元確認待ち"
            Case Is = "SEM_RECOVER_WAIT"
                Status2Name = "ウェハ回収待ち"
            Case Is = "SEM_DROP_PREPARE"
                Status2Name = "ウェハ廃棄受入待ち"
            Case Is = "SEM_DROP_WAIT"
                Status2Name = "ウェハ廃棄待ち"
            Case Is = "COMPLETE"
                Status2Name = "完了"
            Case Else
                Status2Name = ""
        End Select
End Function

'▼*************************************************************▼
' 機能      : テキストボックス 白/黄色切り替え
' 返り値    :
' 引き数    :
' 機能説明  :
' 備考      : 黄色：H0099FFFF（=99FFFF）
'▼*************************************************************▼
Public Sub YellowWhite(myObject As Object)
    Application.EnableEvents = False
    If myObject.Text <> "" Then
        myObject.BackColor = RGB(255, 255, 255)
    Else
        myObject.BackColor = RGB(255, 255, 153)
    End If
    Application.EnableEvents = True
End Sub

'▼*************************************************************▼
' 機能      : 依頼者情報取得
' 返り値    : データ件数（データ）
' 引き数    : 依頼No
' 機能説明  :
' 備考      :
'▼*************************************************************▼
Public Function f_GetClientInfo(sSemReqno As String) As Long
    Dim lCnt    As Long
    
    '▼ＳＱＬ実行前に必ず実行すること
    gsFileNM = f_SqlInit(5, "WORKDIR", "SqlData.txt")
    
    Bcol(1) = "\SemReqno"
    Acol(1) = "'" & sSemReqno & "'"
    
    '▼ＳＱＬ実行
    If CallAdoSql("SQL", 20, gsFileNM, Bcol(), Acol(), "", "") = False Then
        MsgBox "エラーが発生したため、処理が中断されました。" & Chr(10) & Chr(10) & _
               ErNm & "：" & sEr, vbCritical + vbOKOnly, "エラー"
        Set ErNm = Nothing
        Set sEr = Nothing
        Exit Function
    End If
    
    '▼取得データ
    lCnt = f_GetData(psData(), gsFileNM)
    
    f_GetClientInfo = lCnt
End Function

'▼*************************************************************▼
' 機能      : ユーザ情報取得
' 返り値    : データ件数（データ）
' 引き数    : 統一ユーザーID
' 機能説明  :
' 備考      :
'▼*************************************************************▼
Public Function f_GetUsrName(sSemUsrid As String) As Long
    Dim lCnt    As Long
    
    '▼ＳＱＬ実行前に必ず実行すること
    gsFileNM = f_SqlInit(5, "WORKDIR", "SqlData.txt")
    
    Bcol(1) = "\SemUsrid"
    Acol(1) = "'" & sSemUsrid & "'"
    
    '▼ＳＱＬ実行
    If CallAdoSql("SQL", 21, gsFileNM, Bcol(), Acol(), "", "") = False Then
        MsgBox "エラーが発生したため、処理が中断されました。" & Chr(10) & Chr(10) & _
               ErNm & "：" & sEr, vbCritical + vbOKOnly, "エラー"
        Set ErNm = Nothing
        Set sEr = Nothing
        Exit Function
    End If
    
    '▼取得データ
    lCnt = f_GetData(psData(), gsFileNM)
    
    f_GetUsrName = lCnt
End Function

'▼*************************************************************▼
' 機能      : 現ステータス取得
' 返り値    : データ件数（データ）
' 引き数    : 依頼No
' 機能説明  :
' 備考      :
'▼*************************************************************▼
Public Function f_GetStatus(sSemReqno As String) As String
    Dim lCnt    As Long
    
    '▼ＳＱＬ実行前に必ず実行すること
    gsFileNM = f_SqlInit(5, "WORKDIR", "SqlData.txt")
    
    Bcol(1) = "\SemReqno"
    Acol(1) = "'" & sSemReqno & "'"
    
    '▼ＳＱＬ実行
    If CallAdoSql("SQL", 19, gsFileNM, Bcol(), Acol(), "", "") = False Then
        MsgBox "エラーが発生したため、処理が中断されました。" & Chr(10) & Chr(10) & _
               ErNm & "：" & sEr, vbCritical + vbOKOnly, "エラー"
        Set ErNm = Nothing
        Set sEr = Nothing
        Exit Function
    End If
    
    '▼取得データ
    lCnt = f_GetData(psData(), gsFileNM)
    
    f_GetStatus = lCnt
End Function

'▼*************************************************************▼
' 機能      : WFのステータスを1つ進める
' 返り値    : 1:正常終了、-1:エラー
' 引き数    : 依頼No、進捗状況、入力者ユーザID、入力者ユーザID2、
' 引き数    : 登録日時、更新日時、やり直し回数、後戻り発生フラグ、コメント
' 機能説明  :
' 備考      :
'▼*************************************************************▼
Public Function f_ProcApprovalWf(sSemReqnoVal As String, sSemStatusVal As String, _
sSemInpusr1Val As String, sSemInpusr2Val As String, sSemRegdatVal As Date, _
sSemUpddatVal As Date, iSemRepeatVal As Integer, iSemDropflgVal As Integer, _
sSemCommentVal As String) As Long
    Dim lCnt    As Long
    
    '▼ＳＱＬ実行前に必ず実行すること
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
    
    '▼ＳＱＬ実行
    If CallAdoSql("SQL", 22, gsFileNM, Bcol(), Acol(), "", "") = False Then
        MsgBox "エラーが発生したため、処理が中断されました。" & Chr(10) & Chr(10) & _
               ErNm & "：" & sEr, vbCritical + vbOKOnly, "エラー"
        Set ErNm = Nothing
        Set sEr = Nothing
        Exit Function
    End If
    
    '▼取得データ
    lCnt = f_GetData(psData(), gsFileNM)
    
    f_ProcApprovalWf = lCnt
End Function

'▼*************************************************************▼
' 機能      : WFのステータスを1つ戻す
' 返り値    : 1:正常終了、-1:エラー
' 引き数    : 依頼No
' 機能説明  :
' 備考      :
'▼*************************************************************▼
Public Function f_ProcDropWf(sSemReqnoVal As String) As Long
    Dim lCnt    As Long
    
    '▼ＳＱＬ実行前に必ず実行すること
    gsFileNM = f_SqlInit(5, "WORKDIR", "SqlData.txt")
    
    Bcol(1) = "\SEM_REQNO"
    Acol(1) = "'" & sSemReqnoVal & "'"
    
    '▼ＳＱＬ実行
    If CallAdoSql("SQL", 23, gsFileNM, Bcol(), Acol(), "", "") = False Then
        MsgBox "エラーが発生したため、処理が中断されました。" & Chr(10) & Chr(10) & _
               ErNm & "：" & sEr, vbCritical + vbOKOnly, "エラー"
        Set ErNm = Nothing
        Set sEr = Nothing
        Exit Function
    End If
    
    '▼取得データ
    lCnt = f_GetData(psData(), gsFileNM)
    
    f_ProcDropWf = lCnt
End Function

'▼*************************************************************▼
' 機能      : 依頼番号からデータ取得
' 返り値    : 0：正常、-1：エラー
' 引き数    : 依頼No
' 機能説明  :
' 備考      :
'▼*************************************************************▼
Public Function f_GetSemReqtblData(sSemReqno As String) As String
    Dim lCnt    As Long
    
    '▼ＳＱＬ実行前に必ず実行すること
    gsFileNM = f_SqlInit(5, "WORKDIR", "SqlData.txt")
    
    Bcol(1) = "\SemReqno"
    Acol(1) = "'" & sSemReqno & "'"
    
    '▼ＳＱＬ実行
    If CallAdoSql("SQL", 25, gsFileNM, Bcol(), Acol(), "", "") = False Then
        MsgBox "エラーが発生したため、処理が中断されました。" & Chr(10) & Chr(10) & _
               ErNm & "：" & sEr, vbCritical + vbOKOnly, "エラー"
        Set ErNm = Nothing
        Set sEr = Nothing
        f_GetSemReqtblData = -1
        Exit Function
    End If
    
    '▼取得データ
    lCnt = f_GetData(psData(), gsFileNM)
    
    f_GetSemReqtblData = 1
End Function



'▼********************************************************************************▼
' 機能      : 英数字と指定文字チェック
' 返り値    : 1：チェックOK、-1：エラー有
' 引き数    : sStr     … String型    チェックする文字列
'　　　　　 : fAlpha   … Boolean型   英字を通過させるか
'           : fNumeric … Boolean型　 数字を通過させるか
'　　　　　 : fSymbol  … Boolean型   指定文字を通過させるか
'　　　　　 : sSymbol  … String型  　指定する文字
'           : sDel     … String型　  デリミタ
' 機能説明  :
' 備考      : 特に指示がない場合、通過させる文字はハイフン、デリミタはアンパサンド
'▼********************************************************************************▼
Public Function f_Almerics(sStr As String, fAlpha As Boolean, fNumeric As Boolean _
                          , fSymbol As Boolean, Optional sSymbol As String = "-" _
                                              , Optional sDel As String = "&") As Integer

    Dim i     As Integer
    Dim vSyms As Variant
    
    '▼指定する記号をデリミタで区切り、変数へ格納
    vSyms = Split(sSymbol, sDel)
    
    '▼文字数ぶんのループ
    For i = 1 To Len(sStr)
    
        '▽まずはエラー値として認識
        f_Almerics = -1
    
        '▽英字かどうかチェック
        If fAlpha = True Then
            If Mid(LCase(sStr), i, 1) Like "[a-z]" Then
                f_Almerics = 1
            End If
        End If
        
        '▽数字かどうかチェック
        If fNumeric = True Then
            If Mid(sStr, i, 1) Like "[0-9]" Then
                f_Almerics = 1
            End If
        End If
        
        '▽指定記号かどうかチェック
        If fSymbol = True Then
            If UBound(Filter(vSyms, Mid(sStr, i, 1))) > -1 Then
                f_Almerics = 1
            End If
        End If
        
        '▽まだエラーだったら終了、戻り値 -1
        If f_Almerics = -1 Then
            Exit Function
        End If

    Next i

End Function

'▼********************************************************************************▼
' 機能      : ActiveXコントロールの有効無効判定？
' 返り値    : 1：有効、-1：無効
' 機能説明  :
' 備考      : わざとエラーを起こす
'▼********************************************************************************▼
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

'-------▼登録文字列のチェック▼-------

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

'▼********************************************************************************▼
' 機能      : 指定のシート以外非表示にする。
' 引き数    : sSheet    … String型    表示するワークシート名
' 返り値    : 指定シートへの参照(Worksheet)
' 機能説明  :
' 備考      :
'▼********************************************************************************▼
Public Function ViewWorkSheet(sSheet As String) As Worksheet

    Dim wksht As Worksheet
    
    ' 引数と同じシート名の場合は、表示する。
    Set wksht = ThisWorkbook.Worksheets(sSheet)
    wksht.Visible = xlSheetVisible
    Set ViewWorkSheet = wksht
            
    For Each wksht In ThisWorkbook.Worksheets
        If sSheet <> wksht.Name Then
            ' 上記以外は非表示とする。
            wksht.Visible = xlSheetHidden
        End If
    Next

End Function


'▼********************************************************************************▼
' 機能      : 空き配列判定
' 引き数    : target        … Variant型    対象の配列
' 返り値    : Boolean : True=有効な配列、False=空きの配列(配列以外も含む)
' 機能説明  :
' 備考      :
'▼********************************************************************************▼
Public Function IsEnableArray(Target() As Variant) As Boolean
    On Error GoTo EH:
    
    ' 配列要素にアクセスを試みる。
    If Abs(UBound(Target) - LBound(Target)) > 0 Then
        IsEnableArray = True
    Else
        IsEnableArray = False
    End If
    
    Exit Function
EH:
    ' 例外が発生した場合、空き配列とする。
    IsEnableArray = False
    
End Function

'▼********************************************************************************▼
' 機能      : 日付フォーマット判定
' 引き数    : target        … String型    対象の文字列
'           : isPaat        … Boolean型   システム日時以前の日付を許可するか指定するフラグ。省略可(デフォルト=[False])
'           : sMsg          … String型    不正な文字列を指定された場合に判定理由のメッセージが入力される。省略可
' 返り値    : String : 補正後の日付文字列
' 機能説明  : 引数の文字列が日付の書式であるか判定。補正可能な場合は補正する。(引数
'           : の文字列を直接変更する。)
' 備考      :
'▼********************************************************************************▼
Public Function CheckDateFormat(Target As String, Optional isPaat As Boolean = False, Optional ByRef sMsg As Variant) As String
    On Error GoTo EH:
    
    If IsMissing(sMsg) = False Then
        sMsg = Empty
    End If
    
    ' 入力された文字列を日付に変換してみる。
    CheckDateFormat = Format(CDate(Target), "YYYY/MM/DD")
    
    If isPaat = False And CheckDateFormat < Date Then
        ' 過去日付が禁止で変換された日付がシステム日付より過去の日付の場合はシステム日付を返す
        CheckDateFormat = Format(Date, "YYYY/MM/DD")
        If IsMissing(sMsg) = False Then
            sMsg = "本日以降の日付を入力してください。" ' システム日付を返す理由としてメッセージを設定
        End If
    End If

    ' エラーにならなかった場合はOKとする。
    On Error GoTo 0
    
    Exit Function
EH:
    ' 例外が発生した場合はエラーとする。
    On Error GoTo 0
    If IsMissing(sMsg) = False Then
        sMsg = "有効な日付ではありません。"
    End If
    CheckDateFormat = Empty
    
End Function

'▼********************************************************************************▼
' 機能      : ファイル名のみ取得
' 引き数    : target        … String型    対象のファイル パス文字列
' 返り値    : String : ファイル名
' 機能説明  : 引数のファイル パスからファイル名の部分のみを取得し返す。
' 備考      :
'▼********************************************************************************▼
Public Function GetFilename(Target As String)
    ' 文字列の終端から最初に現れた"\"までの文字列を返す。
    '(存在しない場合は、全体を返す。)
    Dim pos As Long
    
    pos = InStrRev(Target, "\")
    If pos = 0 Then
        GetFilename = Target
    Else
        GetFilename = Mid(Target, pos + 1)
    End If
End Function

'▼********************************************************************************▼
' 機能      : ディレクトリ名部分のみ取得
' 引き数    : target        … String型    対象のファイル パス文字列
' 返り値    : String : ファイル名
' 機能説明  : 引数のファイル パスからディレクトリ名の部分のみを取得し返す。
' 備考      :
'▼********************************************************************************▼
Public Function GetDirectoryName(Target As String)
    ' 文字列の終端から最初に現れた"\"までの文字列を返す。
    '(存在しない場合は、全体を返す。)
    Dim pos As Long
    
    pos = InStrRev(Target, "\")
    If pos = 0 Then
        GetDirectoryName = ""
    Else
        GetDirectoryName = Left(Target, pos)
    End If
End Function

'▼********************************************************************************▼
' 機能      : システム終了
' 引き数    : なし
' 返り値    : なし
' 機能説明  : ユーザーに確認後、システムを終了する。(ワークブックを閉じる)
' 備考      :
'▼********************************************************************************▼
Public Sub CloseSystem()
    Dim res As Integer
    
    res = MsgBox("TNA解析依頼システムを終了しますか?", vbQuestion + vbYesNo + vbDefaultButton2, _
        "確認")
    If res = vbYes Then
        ' [はい]が選択された場合システムを終了する。
        If Application.Workbooks.Count > 1 Then
            ' 他にワークブックが開かれている場合は、ファイルを閉じる。
            ThisWorkbook.Close False        ' ファイル保存なしで閉じる。
        Else
            ' ほかにワークブックを開いていない場合は、Excel アプリケーションを終了する。
            ThisWorkbook.Saved = True
            Application.Quit
        End If
    End If

End Sub

'▼********************************************************************************▼
' 機能      : SQLクエリ文エスケープ文字置換
' 引き数    : target        … String型 対象の文字列
' 返り値    : String : エスケープ文字を置き換えた後の文字列
' 機能説明  : SQLクエリ文内で文字列を示す値(「'」(シングル クウォートで囲まれた)に設定
'           : する文字列用にエスケープ文字に置き換える。
' 備考      : 以下の文字列を置き換える。
'           : Tab (タブ)         → CHR(9)
'           : CR  (キャリッジ リターン) → CHR(13)
'           : LF  (行送り)              → CHR(10)
'           : ' (シングル クウォート)   → CHR(39)
'▼********************************************************************************▼
Public Function ReplaceSqlEsc(sSrc As String) As String
    
    ' SQLクエリ文 エスケープされる必要のある文字
    Static caEscapedChars As Variant
    
    Dim iPos As Long
    Dim iLen As Long        ' 対象文字列の全体の長さ
    Dim sCChr As String
    Dim v As Variant
    Dim iFoundPos As Long    ' 見つけたエスケープされる文字の位置
    Dim cFound As String    ' 見つけたエスケープされる文字
    Dim iWPos As Long
    
    Dim sOut As String
    sOut = ""
    
    ' エスケープが必要な文字の配列を用意する。(初回のみ)
    If IsArray(caEscapedChars) <> True Then
        caEscapedChars = Array(Chr(9), Chr(13), Chr(10), Chr(39))
    End If
    
    
    ' 対象の全体の文字数を取得する。
    iLen = Len(sSrc)
    
    ' 対象の先頭文字からエスケープされる文字の検索を開始する。
    iPos = 1
    
    ' 対象の文字数分繰り返す。
    Do While iPos <= iLen
            
        ' 現在の位置からもっとも近いエスケープ必要な文字を検索する。
        iFoundPos = iLen
        cFound = Chr(0)
        
        For Each v In caEscapedChars
            iWPos = InStr(iPos, sSrc, CStr(v))
            If iWPos > 0 Then
                ' 見つけた位置が、他のエスケープ文字より小さい場合、保持用変数を書き換える。
                If iFoundPos > iPos Then
                    iFoundPos = iWPos
                    cFound = CStr(v)
                End If
            End If
        Next
        
        ' エスケープ文字があった場合、置き換える。
        If Asc(cFound) <> 0 Then
            
            ' 先頭以外で見つけたエスケープ文字の前の文字まで出力する。
            sOut = sOut & Mid(sSrc, iPos, iFoundPos - iPos) + "' "
            
            ' エスケープ文字が続く分繰り返す。
            Do
                ' エスケープに置き換える。
                sOut = sOut & "|| CHR(" & CStr(AscB(cFound)) & ") "
                
                ' 次の文字へ
                iPos = iFoundPos + 1
                
                ' 次の文字がエスケープ必要な文字であるか判定する。
                sCChr = Mid(sSrc, iPos, 1)
                cFound = Chr(0)
                For Each v In caEscapedChars
                    If sCChr = CStr(v) Then
                        ' エスケープ必要な文字と一致した場合、
                        cFound = sCChr
                        iFoundPos = iPos
                        Exit For
                    End If
                Next
                
                ' エスケープ必要な文字ではなかった場合
                If Asc(cFound) = 0 Then
                    ' 文字列に戻して、最後の文字を出力する。
                    sOut = sOut & "|| '" & sCChr
                    iPos = iPos + 1
                    Exit Do         ' ループを抜けて、次の文字へ
                End If
            Loop
           
        Else
            ' 現在の位置からエスケープ対象の文字が見つからなかった場合は、残りの部分を
            ' 出力して終了する。
            sOut = sOut & Mid(sSrc, iPos)
            iPos = iLen + 1
        End If
    
    
        ' 次の文字へ
    Loop
    
    ReplaceSqlEsc = sOut

End Function


'▼********************************************************************************▼
' 機能      : DB文字列出力サイズ取得
' 引き数    : string sTarget : 対象の文字列
' 返り値    : Long DBに出力するサイズ(単位: バイト)
' 機能説明  : 引数の文字列をDBに出力した時のバイト数を返す。
' 備考      : 20160406 DBサーバの文字コードがUTF8に変更のため文字コード切り替え可能に変更
'             (廃止)DBには Microsoft JIS (シフトJIS) コードで出力する。
'▼********************************************************************************▼
Function GetLebDb(sTarget As String) As Long
    Dim charSet As String
    
    GetLebDb = 0
    If Len(sTarget) > 0 Then
        charSet = GetSetupCell("ORACLE文字コード")
        
        If charSet = "UTF8" Then
            Dim UTF8 As Object
            Set UTF8 = CreateObject("System.Text.UTF8Encoding")
            GetLebDb = UTF8.GetByteCount_2(sTarget)
        Else
            GetLebDb = LenB(StrConv(sTarget, vbFromUnicode))
        End If
    End If
End Function

'▼********************************************************************************▼
' 機能      : PDFファイルチェック
' 引き数    : string sTarget : チェック対象のファイル パス文字列
' 返り値    : Boolean : True=OK(PDFファイル)、False=エラー(PDFファイル以外)
' 機能説明  : 引数で指定したファイル パスのファイルがPDFファイルであるかチェックする。
' 備考      :
'▼********************************************************************************▼
Public Function CheckPDF(sTarget As String) As Boolean

    On Error GoTo EH:
    
    Dim iFid As Integer
    Dim sRead As String
    iFid = FreeFile
    
    ' 対象のファイルを開く
    Open sTarget For Binary As iFid
    
    ' ファイルの先頭から 6 バイト分のデータを読みこむ。
    sRead = String(6, " ")
    Get iFid, , sRead
    
    ' ファイルを閉じる。
    Close iFid
    
    ' 読み込んだ内容がPDFのヘッダーと一致するか判定する。
    ' "%PDF-?" ? は数値
    If (sRead Like "%PDF-#") = True Then
        ' PDFのパターンと一致した場合。
        CheckPDF = True
    Else
        ' 一致しなかった場合はエラーとする。
        CheckPDF = False
    End If
    
    Exit Function
EH:
    ' ファイル読込に失敗した場合もチェック エラーとする。
    CheckPDF = False
    
End Function

'▼********************************************************************************▼
' 機能      : 最前面表示
' 引き数    : string sCaption : 画面キャプション
' 返り値    : Boolean : なし
' 機能説明  : 引数で指定した画面を最前面表示にする
' 備考      :
'▼********************************************************************************▼
Public Sub SetForeground(sCaption As String)

    Dim hWnd As Long
    hWnd = FindWindow(vbNullString, sCaption)
    Call SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

End Sub

'▼********************************************************************************▼
' 機能      : Oracle Client Version 確認
' 引き数    : なし
' 返り値    : Boolean : True=OK(9又は11)、False=エラー(9又は11以外)
' 機能説明  : Oracle Client Version が9又は11かチェックする
' 備考      :
'▼********************************************************************************▼
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
