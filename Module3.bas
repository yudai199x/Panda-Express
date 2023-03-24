Attribute VB_Name = "Module1"
Sub ŽžŠÔŒvŽZ()

         Dim da1 As String
         Dim da2 As String
         Dim da3 As String
         Dim da5 As String
         Dim sinya As String
         Dim vrng As Range
         Dim ans As Long
         Dim ansI As Long


         For Each vrng In Range("C11:C41").Rows
          If Not vrng = "" Then
             da1 = vrng.Value
             da2 = vrng.Offset(0, 2).Value
             da3 = vrng.Offset(0, 3).Value
             da5 = vrng.Offset(0, 5).Value
             sinya = 0
             ansI = 0
        
             ans = (da3 * 60 + da5) - (da1 * 60 + da2)
             
             If da3 = 12 And da5 > 30 Then ansI = ansI + (da5 - 30)
             If da3 = 13 And da5 <= 30 Then ansI = ansI + (da5 + 30)
             If da3 = 13 And da5 >= 45 Or da3 >= 14 Then ansI = ansI + 60
             
             If da3 = 19 And da5 = 30 Or da3 = 19 And da5 = 45 Then ansI = ansI + (da5 - 15)
             If da3 >= 20 Then ansI = ansI + 30
             If da3 = 23 And da5 <= 15 Then ansI = ansI + (da5 + 15)
             If da3 = 23 And da5 <= 15 Then sinya = da5 + 15
             If da3 = 23 And da5 >= 30 Or da3 >= 24 Then ansI = ansI + 30
             If da3 = 23 And da5 >= 30 Or da3 >= 24 Then sinya = sinya + 30
             If da3 = 27 And da5 > 15 Then ansI = ansI + (da5 - 15)
             If da3 = 27 And da5 > 15 Then sinya = sinya + (da5 - 15)
             If da3 = 28 And da5 <= 15 Then ansI = ansI + (da5 + 45)
             If da3 = 28 And da5 <= 15 Then sinya = sinya + (da5 + 45)
             If da3 = 28 And da5 >= 30 Or da3 >= 29 Then ansI = ansI + 60
             If da3 = 28 And da5 >= 30 Or da3 >= 29 Then sinya = sinya + 60
             vrng.Offset(0, 6).Value = ansI
             
             If ans - ansI >= 480 Then
                vrng.Offset(0, 7).Value = 8
                vrng.Offset(0, 8).Value = (ans - ansI - 480) / 60
             Else
                vrng.Offset(0, 7).Value = (ans - ansI) / 60
                vrng.Offset(0, 8).Value = ""
             End If
             
             If da3 >= 22 Then
                vrng.Offset(0, 10).Value = ((da3 - 22) * 60 + da5 - sinya) / 60
                If da3 >= 29 Then
                   vrng.Offset(0, 10).Value = ((da3 - 22) * 60 + da5 - sinya - ((da3 - 29) * 60 + da5)) / 60
                End If
             Else
                vrng.Offset(0, 10).Value = ""
             End If
             
          Else
             vrng.Value = ""
          End If
         Next
                   

End Sub

Sub ‹Î–±•\Ž©“®“\‚è•t‚¯()

    Dim i As Long, j As Long, cnt As Long
    Dim day(30) As String, day_of_the_week(30) As String
    Dim started_at(30) As Date, finished_at(30) As Date
    
    Filename = Application.GetOpenFilename()
    Workbooks.Open (Filename)
    If Filename = False Then Exit Sub
    For i = 12 To 42
        day(j) = Cells(i, 1).Value
        day_of_the_week(j) = Cells(i, 2).Value
        started_at(j) = Cells(i, 5).Value
        finished_at(j) = Cells(i, 6).Value
        j = j + 1
    Next
    ThisWorkbook.Activate
    i = 11
    For j = 0 To UBound(started_at)
        Worksheets(1).Cells(i, 1) = day(j)
        Worksheets(1).Cells(i, 2) = day_of_the_week(j)
        Worksheets(1).Cells(i, 3) = Hour(started_at(j))
        Worksheets(1).Cells(i, 5) = Minute(started_at(j))
        Worksheets(1).Cells(i, 6) = Hour(finished_at(j))
        Worksheets(1).Cells(i, 8) = Minute(finished_at(j))
        i = i + 1
    Next
    cnt = WorksheetFunction.CountIf(Range(Cells(11, 3), Cells(41, 3)), 0)
    Worksheets(1).Cells(44, 5).Value = 31 - cnt
    
End Sub




