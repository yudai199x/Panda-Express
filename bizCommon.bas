Attribute VB_Name = "bizCommon"
Option Explicit

'**********************************************************************
' @(f)
' 機能      : 「お知らせ」出力
'
' 返り値    : なし
'
' 引き数    : Integer iTarget 取得対象 (0:TSB(依頼者)向け、1:TBA向け)
'           : Range rngDest 出力先
'
' 機能説明  : rngDest で指定したセル範囲に iTarget で指定したお知らせを
'           : 出力する。
'
' 備考      : 1～4列目(結合セル) 日時(YYYY/MM/DD HH:NN)
'             5～列目(結合セル) コメント(最大2行)
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
    ' 出力領域をクリアする。
    ' **********************
    For Each wkCell In rngDest
        wkCell.Value = ""
    Next
    rngDest.Hyperlinks.Delete
    rngDest.Font.Size = 12
    rngDest.Font.Underline = xlUnderlineStyleNone
    rngDest.Font.ThemeColor = xlThemeColorLight1
    
    ' ***********************************************
    ' DB 更新履歴テーブルから「お知らせ」を取得する。
    ' ***********************************************
    iCnt = GetModifyHistory(iTarget, mhInfo)
    If iCnt < 0 Then
        ' エラーが発生した場合は何もしない。(エラー出力済み)
        Application.EnableEvents = True: Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' ********************
    ' 取得結果を出力する。
    ' ********************
    If iCnt > 0 Then
        ' 取得件数 (最大 4 件) を取得する。
        iRowCount = UBound(mhInfo) - LBound(mhInfo) + 1
        
        'ワークシートに出力する。
        For i = 1 To iRowCount
            
            iRow = i + i - 1
            
            ' 更新日時を出力用の文字列に変換する。
            sDate = Format(mhInfo(i).dtModifyDate, "YYYY/MM/DD HH:NN")
            rngDest.Cells(iRow, 1).Value = sDate    ' 日付
                
            ' お知らせ内容に改行コードが含まれるか確認する。
            npos = InStr(mhInfo(i).strMessage, vbCrLf)
            
            If npos = 0 Then
                ' 改行コードを含まない場合
                msg1 = mhInfo(i).strMessage
            Else
                ' 改行コードを含む場合
                ' 改行コード位置で分割する。
                msg1 = Left(mhInfo(i).strMessage, npos - 1) ' 1 行目
                ' 2 行目 (3 行目以降は2行目に続けて出力する。)
                msg2 = Replace(Mid(mhInfo(i).strMessage, npos + 2), vbCrLf, "　")
                
                ' 結合
                msg1 = msg1 + vbCrLf + msg2
            End If
                        
            ' セルに出力する。
            If Len(Trim(mhInfo(i).sLink)) = 0 Then
                ' 通常メッセージ
                rngDest.Cells(iRow, 4).Value = msg1
            Else
                ' リンク付きメッセージ
                rngDest.Hyperlinks.Add Anchor:=rngDest.Cells(iRow, 4), _
                                       Address:=mhInfo(i).sLink, _
                                       TextToDisplay:=msg1
                rngDest.VerticalAlignment = xlTop
                rngDest.Font.Size = 12
            End If
            ' 次のお知らせへ
        Next
    Else
        ' お知らせが存在しない場合は「なし」を出力する。
        rngDest.Cells(1, 4).Value = "お知らせはありません。"
    End If
    
    Application.EnableEvents = True: Application.ScreenUpdating = True
    ' 終了
End Sub



