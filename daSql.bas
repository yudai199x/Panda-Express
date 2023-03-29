Attribute VB_Name = "daSql"
Option Explicit
Option Base 1
'********************************************************************************************
'  解析依頼システム - 業務共通処理モジュール
'               Copyright 2015, XXXX All Rights Reserved.
'  2015-05-14 新規作成
'********************************************************************************************

' ********
' 定数定義
' ********
' 選択項目マスタ <依頼シート作成> 「分類」取得キー
Public Const cItemKeyIraiBunrui As String = "REQCATEG"

' 選択項目マスタ <依頼シート作成> 「発注部門」取得キー
Public Const cItemKeyIraiHachu As String = "REQSEC"

' 選択項目マスタ <依頼シート作成> 「試料内容 品種」取得キー
Public Const cItemKeyIraiHinshu As String = "SAMPKIND"

' 設定値取得 依頼状況取得 最大取得件数
Public Const cSetupIraiStatusMax As String = "依頼状況最大表示数"

' 否認時、承認欄に表示する文字列定義（2015/10/15追加）
Public Const cNotApproved As String = "否認"

' 否認時、承認欄に表示する文字列定義（2016/4/11追加）
Public Const cReEstimate As String = "再見積もり"

' 並び変えキーワード昇順（2016/03/31追加）
Public Const cOrderByAsc As String = "asc"

' 並び変えキーワード降順（2016/03/31追加）
Public Const cOrderByDesc As String = "desc"


' **********
' 構造体定義
' **********
' お知らせ情報 構造体
Public Type SModifyHistory
    dtModifyDate    As Date         ' 更新日付 (年月日+時分秒)
    strMessage      As String       ' メッセージ
    sLink           As String       ' リンク先
End Type

'**********************************************************************
' @(f)
' 機能      : 「お知らせ」取得処理
'
' 返り値    : Long : 取得件数&結果 (0以上=取得件数、負数=エラー)
'
' 引き数    : Integer iTarget 取得対象 (0:TSB(依頼者)向け、1:TBA向け)
'           : SModifyHistory mhInfo() 取得した「お知らせ」情報
'
' 機能説明  : DB 更新履歴テーブルから、引数の取得対象の「お知らせ」を
'           : 最新から4件取得し、引数の構造体に格納する。
'
' 備考      : お知らせが存在しない場合は、mhInfo に Empty が設定される。
'
'**********************************************************************
Public Function GetModifyHistory(iTarget As Integer, mhInfo() As SModifyHistory) As Long
    Dim iCnt As Long
    Dim i As Long
    
    Erase mhInfo
    
    Dim wSql As SSqlSet
    ' ★ SQLを準備する。
    '    5:「お知らせ」取得
    wSql = n_InitSql(5)
    
    ' [SQL] クエリ文に設定する値(変更値、条件等)を設定する。
    wSql.rep.Add "\Target", str(iTarget)
    wSql.rep.Add "\MaxCount", "4"
    
    ' ★ SQLを実行する。
    iCnt = n_DoSql(wSql)

    If iCnt > 0 Then
        ' 抽出した結果を出力用配列に設定する。
        ReDim mhInfo(1 To iCnt)     ' 配列を件数に合わせて拡張する。
        For i = 1 To iCnt
            mhInfo(i).dtModifyDate = CDate(wSql.psData(1, i))         ' TEM_INPDAT 登録日時
            mhInfo(i).strMessage = DbResumeNewLine(wSql.psData(2, i)) ' TEM_VALUE お知らせ内容
            mhInfo(i).sLink = DbResumeNewLine(wSql.psData(3, i))      ' TEM_LINK  リンク先
        Next
    End If
    
    ' 終了
    GetModifyHistory = iCnt
End Function

