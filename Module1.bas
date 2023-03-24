Attribute VB_Name = "Module1"
Public MyBln As Boolean
Public mtbg As String
Public swOK As Long
Public ID As String
Public CHIP(32, 7) As String
Public product As String
Public SEDAI As String
Public pellet As String

Public stack As String
Public PKG As String
Public Ball As String
Public CTR As String


Sub チップ取り出し()
Dim awb As Workbook
Set awb = ActiveWorkbook

Call アクティブシートを初期化する

'Call 解析DBから製品名を検索し開発リストから開発Noを調べる
Call oracle_TNS確認
swOK = 0
MyBln = False
簡易ビューア.Show vbModeless
Do
    DoEvents
Loop Until MyBln

If swOK = 1 Then

awb.Activate

Call MtBg図から情報抽出
Call 貼り付け用シートを編集
End If


End Sub


Sub 解析DBから製品名を検索し開発リストから開発Noを調べる()
    On Error GoTo CatchErr

Dim kaihatsu_No As String
Dim sercMBfile As Range

Dim wb As Workbook
Set wb = ActiveWorkbook
 

Call oracle_TNS確認

 Workbooks.Open filename:="\\133.119.131.52\HinSho\02_品質改善Gr\01_解析関連\01_業務関連\00_品質保証\冶具整備\解析環境整備リスト2015-2016_Rev.7.xlsm", ReadOnly:=True
 wb.Activate


    ' Oracleのセッション
'    Dim OraSess As Object
    ' OracleDB
'    Dim OraDB As Object

    Dim OraSess As OraSession
    Dim OraDB As OraDatabase
    Dim objRS As OraDynaset
   
    ' Oracleのセッションクリエイト
    Set OraSess = CreateObject("OracleInProcServer.XOraSession")
    ' データベース名、ユーザID、パスワード
    Set OraDB = OraSess.OpenDatabase(Sheets("設定情報").Cells(2, 2).Value, _
                                     Sheets("設定情報").Cells(3, 2).Value & "/" & _
                                      Sheets("設定情報").Cells(4, 2).Value, 0&)

    ' SQLを変数としてセット
    Dim strSql As String
    
    strSql = "select * from T_DATA_1ST where " & """" & "CLIENT_KANRI_NO" & """" & " LIKE '%" & Cells(2, 1) & "%'" & _
             " ORDER BY CLIENT_KANRI_NO"
    
    Cells(1, 1) = strSql
    
    ' 検索実行
    Set objRS = OraDB.CreateDynaset(strSql, 0&)

    ' 検索結果0件チェック
    If objRS.EOF = False Then
    ' 検索結果が1件以上であれば、シートに結果を出力します。
        Dim i As Integer
        i = 3

        ' *************************************************
        ' カーソルのフィールドに値があれば、
        ' カーソルの内容をループしてシートに入力します。
        ' このサンプルコードでは、結果は1件です。
        ' *************************************************
        If Not IsNull(objRS.Fields("KANRI_NO").Value) Then
            'レコードの最後まで繰り返し
            Do While objRS.EOF = False
                Cells(i, "A").Value = objRS.Fields("CLIENT_KANRI_NO").Value
                Cells(i, "B").Value = objRS.Fields("KANRI_NO").Value
                Cells(i, "C").Value = objRS.Fields("PRODUCT_NAME").Value
                kaihatsu_No = ""
                Call 情報確認_マスタ(objRS.Fields("PRODUCT_NAME").Value, kaihatsu_No)
                wb.Activate
                Cells(i, "D").Value = kaihatsu_No
                If kaihatsu_No <> "" Then
                   Set sercMBfile = Sheets("MtBg対応表").Range("A:A").Find(kaihatsu_No, LookAt:=xlWhole)
                   If Not sercMBfile Is Nothing Then
                        Cells(i, "E").Value = Sheets("MtBg対応表").Cells(sercMBfile.Row, 2).Value
'                        Call MtBg図から情報抽出
                   End If
                End If
                i = i + 1
                'カーソルを次のレコードに移動
                objRS.MoveNext
            Loop
        End If
    End If

    'オブジェクト開放
    objRS.Close
    Set objRS = Nothing
    OraDB.Close
    Set OraDB = Nothing
    Set OraSess = Nothing

    Exit Sub

    '============================================
    'エラーハンドリング
    '============================================
CatchErr:
    If (OraSess.LastServerErr <> 0) Then     'OraSession でエラー発生
        MsgBox OraSess.LastServerErrText        'エラー内容の表示
        OraSess.LastServerErrReset         'エラーのクリア
        Set OraSession = Nothing                'オブジェクト開放
    ElseIf (OraDB.LastServerErr <> 0) Then
        MsgBox OraDB.LastServerErrText
        OraDB.LastServerErrReset
        Set objRS = Nothing
        Set OraDB = Nothing
        Set OraSess = Nothing
    Else
        MsgBox Err.Description
        Set objRS = Nothing
        Set OraDB = Nothing
        Set OraSess = Nothing
    End If


    Application.DisplayAlerts = False
    Workbooks("解析環境整備リスト2015-2016_Rev.7.xlsm").Close
    Application.DisplayAlerts = True

End Sub


Sub test()
    rootFolder = "C:\Documents and Settings\hogehoge\デスクトップ\test"
    filename = Dir(rootFolder & "\*.*", vbNormal)
    Filename2 = Dir()
    Do While filename <> ""
        If FileDateTime(rootFolder & "\" & Filename2) < FileDateTime(rootFolder & "\" & filename) Then
            Filename2 = filename
        End If
        filename = Dir()
    Loop
    MsgBox Filename2
End Sub

Sub 情報確認_マスタ(PROD As String, development_No As String)

flg = 0


    
 Windows("解析環境整備リスト2015-2016_Rev.7.xlsm").Activate
 Sheets("解析環境整備状況(Package)").Activate
   
    
 Dim FoundCell As Range    '完全一致検索するよー
 Set FoundCell = Range("F7").CurrentRegion.Find(What:=PROD, LookAt:=xlWhole)
 
 If FoundCell Is Nothing Then '駄目っぽいからあきらめる
     GoTo MSTLB
 Else
     flg = 1
     FoundCell.Activate
 End If
 
 If Left(ActiveCell.Offset(0, 11).Value, 2) = "N-" Then
    Call 数字だけ抜き出し(ActiveCell.Offset(0, 11).Value, development_No)
 End If
MSTLB:

End Sub



Sub MtBg図から情報抽出()

    Dim book1 As Workbook
    Dim str As String
    Dim wb As Workbook
    Dim MaxCol As Long
    
    Set wb = ActiveWorkbook
    Erase CHIP
    
    Workbooks.Open "\\10.23.7.43\hin-kaiseki\00_解析関連\05_NAND解析関連\00_仕様書関連\01_MtBg図\" & mtbg

    Set book1 = Workbooks(mtbg)
    maxf = 0
    copfl = 0
    aaa = book1.Sheets.Count
    bbb = book1.Sheets(aaa).Shapes.Count
    flag = 0
    cnt = 0
    MaxRow = book1.Sheets(aaa).Cells(Rows.Count, 3).End(xlUp).Row
    For i = 1 To bbb
        If book1.Sheets(aaa).Shapes(i).Type = 6 Then
            cnt = cnt + 1
            rectflag = 0
            For j = 1 To book1.Sheets(aaa).Shapes(i).GroupItems.Count
                 If InStr(book1.Sheets(aaa).Shapes(i).GroupItems(j).Name, "Resin") > 0 Then rectflag = 1
            Next j

            For j = 1 To book1.Sheets(aaa).Shapes(i).GroupItems.Count
                 If Left(book1.Sheets(aaa).Shapes(i).GroupItems(j).Name, 5) = "Line_" And _
                    InStr(book1.Sheets(aaa).Shapes(i).GroupItems(j).Name, "_Wire") > 0 And rectflag = 1 Then
                    flag = flag + 1
                 End If
            Next j
        
            For j = 1 To book1.Sheets(aaa).Shapes(i).GroupItems.Count
                 If Right(book1.Sheets(aaa).Shapes(i).GroupItems(j).Name, 11) = "DrawingArea" Then
                    book1.Sheets(aaa).Shapes(i).GroupItems(j).Delete
                    Exit For
                 End If
            Next j
        
        End If
        If flag > maxf Then
            maxf = flag
            flag = 0
            copfl = 1
            Set wiresp = book1.Sheets(aaa).Shapes(i)
        End If
        
    Next i
    If copfl = 1 Then
       If wiresp.Height > wiresp.Width Then
          wiresp.IncrementRotation 90
          wiresp.Height = 20 * 72 / 2.54
          wiresp.Width = 2.27 * 72 / 2.54
       Else
          wiresp.Height = 2.27 * 72 / 2.54
          wiresp.Width = 20 * 72 / 2.54
       End If
       wiresp.Copy
       wb.Worksheets("チップ取り出し").Activate
       Cells(2, 2).Select
       ActiveSheet.Pictures.Paste
    End If
    
    flag = 0
    stcount = 0
    
    book1.Sheets(aaa).Activate
    For i = 1 To MaxRow
       If flag = 1 Then
          cmcount = 0
          For j = 1 To MaxCol
              nodatact = 0
              For k = 3 To MaxCol
                  If book1.Sheets(aaa).Cells(i, k).Value <> "" Then nodatact = 1
              Next k
              If nodatact = 0 Then Exit For
              
              If j = 1 Or j = 2 Or j = 3 Then
                 CHIP(stcount, j - 1) = book1.Sheets(aaa).Cells(i, j).Value
                 cmcount = cmcount + 1
              Else
                 If book1.Sheets(aaa).Cells(i, j).Value <> "" Then
                    CHIP(stcount, cmcount) = book1.Sheets(aaa).Cells(i, j).Value
                    cmcount = cmcount + 1
                 End If
              End If
              If cmcount = 8 Then Exit For
          Next j
          stcount = stcount + 1
       End If
       If InStr(book1.Sheets(aaa).Cells(i, 3).Value, "チップ名称") > 0 Then
          MaxCol = book1.Sheets(aaa).Cells(i, Columns.Count).End(xlToLeft).Column
          flag = 1
       End If
       If InStr(book1.Sheets(aaa).Cells(i, 1).Value, "上段") > 0 Then flag = 0
       If book1.Sheets(aaa).Cells(i, 1).Value = "チップ名" Then flag = 0
    Next i
    If flag = 1 And rectflag = 1 Then
       MsgBox ("情報が拾えていない可能性があるためM'tBg図をそのまま開いておきます")
    Else
       Application.DisplayAlerts = False
       Workbooks(mtbg).Close
       Application.DisplayAlerts = True
    End If
    wb.Worksheets("チップ取り出し").Activate
    
End Sub

Sub 貼り付け用シートを編集()
    Dim wb As Workbook
    Dim MaxCol As Long
    Dim TxtSample As Shape
    
    Set wb = ActiveWorkbook

    Count = 0
    For i = 0 To 32
        If CHIP(i, 7) <> "" Then
           For j = 0 To 7
              wb.Worksheets("チップ取り出し").Cells(11, 1).Offset(Count, j) = CHIP(i, j)
              If j <> 0 Then
                 wb.Worksheets("チップ取り出し").Cells(11, 1).Offset(Count, j).BorderAround Weight:=xlThin
                 If j = 4 Then
                    wb.Worksheets("チップ取り出し").Cells(11, 1).Offset(Count, 8).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
                    wb.Worksheets("チップ取り出し").Cells(11, 1).Offset(Count, 8).Borders(xlEdgeRight).LineStyle = xlLineStyleNone
                 End If
                 If j = 1 And CHIP(i, j) = "" Then wb.Worksheets("チップ取り出し").Cells(11, 1).Offset(Count, 8).Borders(xlEdgeTop).LineStyle = xlLineStyleNone
              End If
           Next j
           wb.Worksheets("チップ取り出し").Cells(11, 1).Offset(Count, 8).BorderAround Weight:=xlThin
           wb.Worksheets("チップ取り出し").Cells(11, 1).Offset(Count, 9).BorderAround Weight:=xlThin
           wb.Worksheets("チップ取り出し").Cells(11, 1).Offset(Count, 10).BorderAround Weight:=xlThin
           Count = Count + 1
        End If
    Next i

    dansu = Val(Replace(stack, "X", ""))
    sedai_dandu = ""
    If Cells(12, 15) = "BiCS3" Then sedai_dandu = "Chip対応表_BiCS3"
    If Cells(12, 15) = "BiCS4" Then sedai_dandu = "Chip対応表_BiCS4"
    
    If sedai_dandu <> "" Then
       MaxRow = wb.Worksheets(sedai_dandu).Cells(Rows.Count, 1).End(xlUp).Row
    
       If stack = "X1" Then st = 3
       If stack = "X2" Then st = 6
       If stack = "X4" Then st = 9
       If stack = "X6" Then st = 12
       If stack = "X8" Then st = 15
       If stack = "X12" Then st = 18
       If stack = "X16" Then st = 21
    
       If PKG = "BGA" Then
          If Ball = "272" And stack = "X4" Then st = 24
          If Ball = "272" And stack = "X8" Then st = 27
          If Ball = "272" And stack = "X16" Then st = 30
          If InStr(CTR, "MIF") > 0 And stack = "X8" Then st = 33
          If InStr(CTR, "MIF") > 0 And stack = "X16" Then st = 36
       End If
    
    
       flag = 0
       For i = 1 To MaxRow
         If wb.Worksheets(sedai_dandu).Cells(i, 1) = PKG Then
            If InStr(CTR, "MIF") > 0 Then ofst = 2
            If PKG = "UFS_BGA" Then ofst = 1
            For j = 0 To dansu - 1
               wb.Worksheets("チップ取り出し").Cells(11 + j + ofst, 9) = wb.Worksheets(sedai_dandu).Cells(i + 1 + j, st)
               wb.Worksheets("チップ取り出し").Cells(11 + j + ofst, 10) = wb.Worksheets(sedai_dandu).Cells(i + 1 + j, st + 1)
               wb.Worksheets("チップ取り出し").Cells(11 + j + ofst, 11) = wb.Worksheets(sedai_dandu).Cells(i + 1 + j, st + 2)
               wb.Worksheets("チップ取り出し").Cells(11 + j + ofst, 9).BorderAround Weight:=xlThin
               wb.Worksheets("チップ取り出し").Cells(11 + j + ofst, 10).BorderAround Weight:=xlThin
               wb.Worksheets("チップ取り出し").Cells(11 + j + ofst, 11).BorderAround Weight:=xlThin
            Next j
         End If
       Next i
    Else
       MsgBox ("BiCS3,BiCS4以外は段数情報未対応です")
    End If

    Set TxtSample = wb.Worksheets("チップ取り出し").Shapes.AddTextbox _
                   (msoTextOrientationHorizontal, _
                    1.2 * 72 / 2.54, _
                    (5.5 + (0.51 * dansu)) * 72 / 2.54, _
                    25 * 72 / 2.54, _
                    0.6 * 72 / 2.54)
   
   'テキストボックスに文字追加
   TxtSample.TextFrame.Characters.Text = ID
   TxtSample.TextFrame.AutoSize = True

End Sub



 '// 引数1：対象文字列
'// 引数2：検索結果
Sub 数字だけ抜き出し(s, result)
    Dim reg             As New RegExp       '// 正規表現クラスオブジェクト
    Dim oMatches        As MatchCollection  '// RegExp.Execute結果
    Dim oMatch          As Match            '// 検索結果オブジェクト
    Dim i                                   '// ループカウンタ
    Dim iCount                              '// 検索一致件数
    
    '// 検索範囲＝文字列の最後まで検索
    reg.Global = True
    '// 検索条件＝数字を検索
    reg.Pattern = "[0-9]"
    
    '// 検索実行
    Set oMatches = reg.Execute(s)
    
    '// 検索一致件数を取得
    iCount = oMatches.Count
    
    result = ""
    
    '// 検索一致件数だけループ
    For i = 0 To iCount - 1
        '// コレクションの現ループオブジェクトを取得
        Set oMatch = oMatches.Item(i)
        
        '// 検索一致文字列
        result = result & oMatch.Value
    Next
End Sub

Sub PowerPointファイルを開く()
 
Dim ppApp As New PowerPoint.Application
Dim objPPT As Object
Dim PPT_P As Object, PPT_SD As Object, PPT_SP As Object, PPT_MT As Object, PPT_CL As Object
Dim x As Long, y As Long
Dim ppPrs As PowerPoint.Presentation 'プレゼンテーションオブジェクト
Dim ppH As Long
Dim ppW As Long
Dim i As Long
Dim ioffset As Long
Dim offcnt As Long
Dim PPT_col As Object
Dim PPT_row As Object
Dim PPT_cel As Object
Dim searchW As String
Dim searchS As Long
Dim entsplno As Long
Dim cntent As Long
Dim dlg As FileDialog
Dim slw As Long
Dim spw As Long
Dim filename As Variant
Dim ws As Worksheet

Dim chigh As Double
Dim ppShape As PowerPoint.Shape

Set ws = ThisWorkbook.Worksheets("チップ取り出し")
MaxRow = ws.Cells(Rows.Count, 2).End(xlUp).Row


ppApp.Visible = True
 
Set ppPrs = ppApp.Presentations.Open(ThisWorkbook.Path & "\Chip取り出し【KIOXIAフォーマット】.pptx") 'プレゼンテーションを開く
'MsgBox "PowerPointファイルが開きました"

'スライドマスタを編集する
'With ppPrs.SlideMaster
'  For Each PPT_MT In .CustomLayouts
'    For Each PPT_CL In PPT_MT.Shapes
'      If PPT_CL.HasTextFrame Then
'         Debug.Print PPT_CL.TextFrame.TextRange.Text
'         If PPT_CL.TextFrame.TextRange.Text = "YQR-xxxxx ver. 2" Then
'            PPT_CL.TextFrame.TextRange.Text = "YQR-" & YQR(0) & " ver. 2"
'            Debug.Print PPT_CL.TextFrame.TextRange.Text
'         End If
'      End If
'    Next PPT_CL
'  Next PPT_MT
'  slw = .Width
'End With




PRODUC = ws.Cells(6, 15) & "_" & Replace(ws.Cells(7, 15), "X", "") & "st-" & ws.Cells(8, 15)
If ws.Cells(10, 15) <> "" And ws.Cells(10, 15) <> "-" Then PRODUC = PRODUC & "_" & ws.Cells(10, 15)

'1枚目のスライドを編集
With ppPrs.Slides(1)
    For Each PPT_CL In .Shapes
      If PPT_CL.HasTextFrame Then
         If Left(PPT_CL.TextFrame.TextRange.Text, 14) = "(YQR-XXYYY#ZZ)" Then
            PPT_CL.TextFrame.TextRange.Text = Replace(PPT_CL.TextFrame.TextRange.Text, "XXYYY#ZZ", ws.Cells(2, 15) & "_#" & ws.Cells(11, 15))
         End If
         If Left(PPT_CL.TextFrame.TextRange.Text, 3) = "製品名" Then
            PPT_CL.TextFrame.TextRange.Text = Replace(PPT_CL.TextFrame.TextRange.Text, "PRODUCT*****", PRODUC)
            PPT_CL.TextFrame.TextRange.Text = Replace(PPT_CL.TextFrame.TextRange.Text, "20XX/YY/ZZ", Date)
         End If
      End If
    Next PPT_CL
End With









'2枚目のスライドを編集


With ppPrs.Slides(2)
    ws.Range(Cells(10, 1), Cells(MaxRow, 11)).CurrentRegion.Copy
    'PasteSpeciaでエラーが出るときは、ここに待ちを作ります。
    'スライド番号を指定
    .Shapes.PasteSpecial DataType:=ppPasteEnhancedMetafile, Link:=msoFalse
    Set ppShape = .Shapes(.Shapes.Count)
    '上位置
    ppShape.Top = 5.53 * 72 / 2.54
    '左位置
    ppShape.Left = 1.55 * 72 / 2.54
    '縦横比を固定
    ppShape.LockAspectRatio = msoTrue
    '横幅
    ppShape.Width = 30 * 72 / 2.54
    chigh = (5.53 * 72 / 2.54) + ppShape.Height
    Application.CutCopyMode = False

    For Each shp In ws.Shapes
        Debug.Print shp.Name
        If Left(shp.Name, 3) = "Pic" Then
           shp.Copy
           .Shapes.PasteSpecial DataType:=ppPasteEnhancedMetafile, Link:=msoFalse
           Set ppShape = .Shapes(.Shapes.Count)
           '上位置
           ppShape.Top = 2.74 * 72 / 2.54
           '左位置
           ppShape.Left = 6.34 * 72 / 2.54
           '縦横比を固定
           ppShape.LockAspectRatio = msoTrue
           '横幅
           ppShape.Width = 20 * 72 / 2.54
           Application.CutCopyMode = False
        End If
        If Left(shp.Name, 3) = "Tex" Then
           shp.Copy
           '.Shapes.PasteSpecial DataType:=ppPasteText, Link:=msoFalse
           .Shapes.PasteSpecial
           Set ppShape = .Shapes(.Shapes.Count)
           '上位置
           ppShape.Top = (0.1 * 72 / 2.54) + chigh
           '左位置
           ppShape.Left = 0.8 * 72 / 2.54
           ppShape.TextFrame.AutoSize = ppAutoSizeShapeToFitText
           ppShape.TextFrame.TextRange.Font.Name = "メイリオ"
           ppShape.TextFrame.TextRange.Font.Size = 10
           
           Application.CutCopyMode = False
        End If
    Next shp

    For Each PPT_CL In .Shapes
      If PPT_CL.HasTextFrame Then
         If Left(PPT_CL.TextFrame.TextRange.Text, 13) = "【Chip取り出し指示書】" Then
            PPT_CL.TextFrame.TextRange.Text = Replace(PPT_CL.TextFrame.TextRange.Text, "XXYYY#ZZ", ws.Cells(2, 15) & "_#" & ws.Cells(11, 15))
            PPT_CL.TextFrame.TextRange.Text = Replace(PPT_CL.TextFrame.TextRange.Text, "PRODUCT*****", PRODUC)
         End If
      End If
    Next PPT_CL
End With


    filename = _
        Application.GetSaveAsFilename( _
             InitialFileName:="【Chip取り出し依頼】_YQR-" & ws.Cells(2, 15) _
           , FileFilter:="パワーポイントプレゼンテーションファイル(*.pptx),*.pptx" _
           , FilterIndex:=1 _
           , Title:="保存先の指定" _
           )
    If filename <> False Then
       Application.DisplayAlerts = False
       ppPrs.SaveAs filename, ppSaveAsDefault
       Application.DisplayAlerts = True
    End If


Set ppApp = Nothing


End Sub




Sub oracle_TNS確認()
Dim buf As String

    If Dir("C:\Oracle\product\11.2.0\client_1\network\admin\TNSNames.ora") <> "" Then
       Open "C:\Oracle\product\11.2.0\client_1\network\admin\TNSNames.ora" For Input As #1
          Do Until EOF(1)
             Line Input #1, buf
             If InStr(buf, "DRBFM.WORLD") > 0 Then
                Close #1
                GoTo next1
             End If
          Loop
       Close #1
       FileCopy "C:\Oracle\product\11.2.0\client_1\network\admin\TNSNames.ora", _
                "C:\Oracle\product\11.2.0\client_1\network\admin\TNSNames_org.ora"
       Open "C:\Oracle\product\11.2.0\client_1\network\admin\TNSNames.ora" For Append As #1
            Print #1, "DRBFM.WORLD ="
            Print #1, "  (DESCRIPTION ="
            Print #1, "    (ADDRESS = (PROTOCOL = TCP)(HOST = 133.116.128.79)(PORT = 1521))"
            Print #1, "    (CONNECT_DATA ="
            Print #1, "      (SID = DRBFM)"
            Print #1, "    )"
            Print #1, "  )"
            Print #1, ""
            Print #1, ""
            Print #1, "FLRANA.WORLD ="
            Print #1, "  (DESCRIPTION ="
            Print #1, "    (ADDRESS = (PROTOCOL = TCP)(HOST = 133.116.128.79)(PORT = 1521))"
            Print #1, "    (CONNECT_DATA ="
            Print #1, "      (SID = FLRANA)"
            Print #1, "    )"
            Print #1, "  )"
            Print #1, ""
       Close #1
    End If

next1:
    If Dir("C:\Oracle\product\9.2.0\client\network\admin\TNSNames.ora") <> "" Then
       Open "C:\Oracle\product\9.2.0\client\network\admin\TNSNames.ora" For Input As #1
          Do Until EOF(1)
             Line Input #1, buf
             If InStr(buf, "DRBFM.WORLD") > 0 Then
                Close #1
                GoTo next2
             End If
          Loop
       Close #1
       FileCopy "C:\Oracle\product\9.2.0\client\network\admin\TNSNames.ora", _
                "C:\Oracle\product\9.2.0\client\network\admin\TNSNames_org.ora"
       Open "C:\Oracle\product\9.2.0\client\network\admin\TNSNames.ora" For Append As #1
            Print #1, "DRBFM.WORLD ="
            Print #1, "  (DESCRIPTION ="
            Print #1, "    (ADDRESS = (PROTOCOL = TCP)(HOST = 133.116.128.79)(PORT = 1521))"
            Print #1, "    (CONNECT_DATA ="
            Print #1, "      (SID = DRBFM)"
            Print #1, "    )"
            Print #1, "  )"
            Print #1, ""
            Print #1, ""
            Print #1, "FLRANA.WORLD ="
            Print #1, "  (DESCRIPTION ="
            Print #1, "    (ADDRESS = (PROTOCOL = TCP)(HOST = 133.116.128.79)(PORT = 1521))"
            Print #1, "    (CONNECT_DATA ="
            Print #1, "      (SID = FLRANA)"
            Print #1, "    )"
            Print #1, "  )"
            Print #1, ""
       Close #1
    End If

next2:
    

End Sub


Sub アクティブシートを初期化する()

  Dim shp As Shape

  For Each shp In ActiveSheet.Shapes
     If Left(shp.Name, 3) <> "But" Then shp.Delete
  Next shp

  Range(Cells(11, 1), Cells(42, 11)).Clear
  Range(Cells(2, 15), Cells(12, 15)).Clear

End Sub

