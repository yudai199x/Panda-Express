Attribute VB_Name = "Module1"
Option Explicit
Sub 製品名から開発No､PKGを調べる()

    Dim prName As String
    Dim dic1 As Object, dic2 As Object
    Dim prdc(5000) As String, PKG(5000) As String, num(5000) As String
    Dim i As Long, j As Long
    
    Workbooks.Open "C:\Users\z12063p0\Desktop\VBA\NAND製品生産計画(210630_Y2106S)_一品別.xlsx"
    For i = 5 To Cells(Rows.Count, 6).End(xlUp).Row
        prdc(j) = Cells(i, 6).Value
        PKG(j) = Cells(i, 11).Value
        num(j) = Cells(i, 17).Value
        j = j + 1
    Next
    Application.DisplayAlerts = False
    Workbooks("NAND製品生産計画(210630_Y2106S)_一品別.xlsx").Close
    Application.DisplayAlerts = True
    ThisWorkbook.Worksheets(1).Activate
    Range(Cells(2, 1), Cells(2 + UBound(prdc), 1)) = WorksheetFunction.Transpose(prdc)
    Range(Cells(2, 2), Cells(2 + UBound(num), 2)) = WorksheetFunction.Transpose(num)
    Range(Cells(2, 3), Cells(2 + UBound(PKG), 3)) = WorksheetFunction.Transpose(PKG)
    Range(Cells(2, 1), Cells(2 + UBound(prdc), 1)).Font.Color = vbBlack
    Range("A:C").RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
    prName = Range("D2")
    Set dic1 = CreateObject("Scripting.Dictionary")
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        dic1.Add Cells(i, 1).Value, Cells(i, 2).Value
    Next
    Range("E2") = dic1.Item(prName)
    Set dic2 = CreateObject("Scripting.Dictionary")
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        dic2.Add Cells(i, 1).Value, Cells(i, 3).Value
    Next
    Range("F2") = dic2.Item(prName)
    
End Sub
Sub 開発NoからMtBg対応表を開く()
    
    Dim mydic As Object
    Dim proName As String, fileName As String
    Dim i As Long, j As Long, k As Long
    Dim l As Long, m As Long, n As Long
    
    Dim cnt(50) As Integer
    Dim step(15) As String
    Dim chip(15) As String
    Dim thick(15) As String
    Dim flag As Double
    
    proName = Range("E2")
    Worksheets(2).Activate
    Set mydic = CreateObject("Scripting.Dictionary")
    With mydic
        For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
            .Add Cells(i, 1).Value, Cells(i, 2).Value
        Next
        fileName = .Item(proName)
    End With
    Workbooks.Open "\\10.23.7.43\HIN-kaiseki\00_解析関連\05_NAND解析関連\00_仕様書関連\01_MtBg図\" & fileName
    
    For j = 1 To Worksheets.Count
        For k = 1 To Worksheets(j).Shapes.Count
            If Left(Worksheets(j).Shapes(k).Name, 5) = "グループ化" Then
                cnt(l) = Worksheets(j).Shapes(k).GroupItems.Count
                l = l + 1
            End If
        Next
    Next
    For j = 1 To Worksheets.Count
        For k = 1 To Worksheets(j).Shapes.Count
            If Left(Worksheets(j).Shapes(k).Name, 5) = "グループ化" Then
                If Worksheets(j).Shapes(k).GroupItems.Count = WorksheetFunction.Max(cnt) Then
                    Worksheets(j).Shapes(k).Copy
                End If
            End If
        Next
    Next
    ThisWorkbook.Worksheets(1).Activate
    Range("D5").PasteSpecial xlPasteAll
    Workbooks(fileName).Activate
    flag = 0
    For j = 1 To Worksheets.Count
        For k = 1 To Cells(Rows.Count, 3).End(xlUp).Row
            For l = 1 To Cells(k, Columns.Count).End(xlToLeft).Column
                If InStr(Worksheets(j).Cells(k, l).Value, "チップ名称") > 0 Then
                    flag = 1
                    For m = k + 1 To Cells(k, l).End(xlDown).Row
                        step(n) = Worksheets(j).Cells(m, l - 1).Value
                        chip(n) = Worksheets(j).Cells(m, l).Value
                        n = n + 1
                    Next
                End If
                If flag = 1 Then Exit For
            Next
            If flag = 1 Then Exit For
        Next
        If flag = 1 Then Exit For
    Next
    flag = 0
    n = 0
    For j = 1 To Worksheets.Count
        For k = 1 To Cells(Rows.Count, 3).End(xlUp).Row
            For l = 1 To Cells(k, Columns.Count).End(xlToLeft).Column
                If InStr(Worksheets(j).Cells(k, l).Value, "チップ厚(um)") > 0 Then
                    flag = 1
                    For m = k + 1 To Cells(k, l).End(xlDown).Row
                        thick(n) = Worksheets(j).Cells(m, l).Value
                        n = n + 1
                    Next
                End If
                If flag = 1 Then Exit For
            Next
            If flag = 1 Then Exit For
        Next
        If flag = 1 Then Exit For
    Next
    Application.DisplayAlerts = False
    Workbooks(fileName).Close
    Application.DisplayAlerts = True
    ThisWorkbook.Worksheets(1).Activate
    Range(Cells(12, 6), Cells(12 + UBound(step), 6)) = WorksheetFunction.Transpose(step)
    Range(Cells(12, 7), Cells(12 + UBound(chip), 7)) = WorksheetFunction.Transpose(chip)
    Range(Cells(12, 8), Cells(12 + UBound(thick), 8)) = WorksheetFunction.Transpose(thick)
    
End Sub
Sub PowerPointファイルを開く()
    
    Dim cnt As Integer
    Dim stcnt As Integer
    Dim i As Long, j As Long, k As Long
    
    Dim ppApp As New PowerPoint.Application
    Dim ppPrs As PowerPoint.Presentation
    Dim ppshp As PowerPoint.Shape
    
    For i = 12 To Cells(12, 7).End(xlDown).Row
        If Left(Worksheets(1).Cells(i, 7).Value, 5) = "BiCS3" Then cnt = 3
        If Left(Worksheets(1).Cells(i, 7).Value, 5) = "BiCS4" Then cnt = 4
    Next
    If Worksheets(1).Cells(2, 6).Value = "TSOP" Then j = 5
    If Worksheets(1).Cells(2, 6).Value = "BGA" Then j = 24
    If Worksheets(1).Cells(2, 6).Value = "ExPBA" Then j = 43
    If Worksheets(1).Cells(2, 6).Value = "UFS_BGA" Then j = 62
    For i = 12 To Cells(12, 7).End(xlDown).Row
        If Left(Worksheets(1).Cells(i, 7).Value, 4) = "BiCS" Then
            stcnt = WorksheetFunction.CountIf(Range(Cells(i, 7), Cells(Cells(i, 7).End(xlDown).Row, 7)), Cells(i, 7))
            If stcnt = 1 Then k = 3
            If stcnt = 2 Then k = 6
            If stcnt = 4 Then k = 9
            If stcnt = 8 Then k = 15
            If stcnt = 16 Then k = 21
            Exit For
        End If
    Next
    For i = 12 To Cells(12, 7).End(xlDown).Row
        If Not Left(Worksheets(1).Cells(i, 7).Value, 4) = "BiCS" Then
            Cells(i, 9).Value = "-"
            Cells(i, 10).Value = "-"
            i = i + 1
        End If
        Worksheets(1).Cells(i, 10) = Worksheets(cnt).Cells(j, k)
        Worksheets(1).Cells(i, 9) = Worksheets(cnt).Cells(j, k + 2)
        j = j + 1
    Next
    Worksheets(1).Columns("G:H").AutoFit
    Worksheets(1).Range(Cells(11, 6), Cells(Range("F11").End(xlDown).Row, 10)).Font.Bold = False
    Worksheets(1).Range(Cells(11, 6), Cells(Range("F11").End(xlDown).Row, 10)).Borders.LineStyle = xlContinuous
    ppApp.Visible = True
    Set ppPrs = ppApp.Presentations.Open(ThisWorkbook.Path & "\Chip取り出し【KIOXIAフォーマット】.pptx")
    With ppPrs.Slides(2)
        ThisWorkbook.Worksheets(1).Shapes(1).Copy
        Set ppshp = .Shapes.Paste.PlaceholderFormat.Parent
        With ppshp
            .LockAspectRatio = msoTrue
            .Width = 950
            .Left = 5
            .Top = 80
        End With
        ThisWorkbook.Worksheets(1).Range(Cells(11, 6), Cells(Range("F11").End(xlDown).Row, 10)).CopyPicture Appearance:=xlScreen
        Set ppshp = .Shapes.Paste.PlaceholderFormat.Parent
        With ppshp
            .LockAspectRatio = msoTrue
            .Width = 700
            .Left = 130
            .Top = 185
        End With
    End With
    
End Sub
