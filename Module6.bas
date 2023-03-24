Attribute VB_Name = "Module3"
Option Explicit

Sub ppt貼り付け()

    Dim pptObj As PowerPoint.Application
    Dim pptPrs As PowerPoint.Presentation
    Dim pptShp As PowerPoint.Shape
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim m As Long
    

    '新規PowerPointアプリケーションを表示
    Set pptObj = New PowerPoint.Application
    pptObj.Visible = True
    
    
    'プレゼンテーションを開く
    Set pptPrs = pptObj.Presentations.Open(ThisWorkbook.Path & "\【KIOXIAフォーマット】電特結果.pptx")
    With pptPrs
        .Slides(2).Shapes(1).TextFrame.TextRange.Text = "#30"
    End With
    
    Worksheets("出力").Activate
    Range(Cells(22, 12), Cells(25, 20)).CopyPicture appearance:=xlPrinter
    Set pptShp = pptPrs.Slides(2).Shapes.Paste.PlaceholderFormat.Parent
    With pptShp
        .Left = 7
        .Top = 80
    End With
    For i = 4 To 11
        Cells(19, i).CopyPicture appearance:=xlPrinter
        Set pptShp = pptPrs.Slides(2).Shapes.Paste.PlaceholderFormat.Parent
        With pptShp
            Select Case i
                Case 4
                    .Left = 117
                    .Top = 160
                Case 5
                    .Left = 352
                    .Top = 160
                Case 6
                    .Left = 587
                    .Top = 160
                Case 7
                    .Left = 822
                    .Top = 160
                Case 8
                    .Left = 117
                    .Top = 318
                Case 9
                    .Left = 352
                    .Top = 318
                Case 10
                    .Left = 587
                    .Top = 318
                Case 11
                    .Left = 822
                    .Top = 318
            End Select
        End With
    Next
    
    Worksheets("グラフ").Activate
    For j = 6 To 139 Step 19
        k = j + 16
        Range(Cells(j, 2), Cells(k, 8)).CopyPicture appearance:=xlPrinter
        Set pptShp = pptPrs.Slides(2).Shapes.Paste.PlaceholderFormat.Parent
        With pptShp
            .LockAspectRatio = True
            .Height = .Height * 2 / 3
            Select Case j
                Case 6
                    .Left = 7
                    .Top = 170
                Case 25
                    .Left = 243
                    .Top = 170
                Case 44
                    .Left = 477
                    .Top = 170
                Case 63
                    .Left = 711
                    .Top = 170
                Case 82
                    .Left = 7
                    .Top = 328
                Case 101
                    .Left = 243
                    .Top = 328
                Case 120
                    .Left = 477
                    .Top = 328
                Case 139
                    .Left = 711
                    .Top = 328
            End Select
        End With
    Next
                    
End Sub
