Attribute VB_Name = "Module1"
Option Explicit
Sub é¿ê—Ç‹Ç∆Çﬂ()
    Dim i As Long, j As Long, k As Long, l As Long
    Dim m As Long, n As Long, o As Long, p As Long
    Dim q As Long, r As Long, s As Long
    
    Dim name As String
    Dim hour As Long
    Dim minute As Long
    Dim YM(7, 32, 100) As Long
    Dim LE(7, 32, 100) As Long
    Dim PE(7, 32, 100) As Long
    Dim HIN(7, 32, 100) As Long
    Dim YM_ALL(7, 32) As Long
    Dim LE_ALL(7, 32) As Long
    Dim PE_ALL(7, 32) As Long
    Dim HIN_ALL(7, 32) As Long
    
    For i = 20 To 22
        For j = 1 To 12
            If i = 20 Then
                k = j + 3
            Else
                k = j
            End If
            If k > 12 Then Exit For
            name = i & "îN" & Space(1) & k & "åé"
            Worksheets(name).Activate
            
            p = 0
            q = 0
            r = 0
            s = 0
            
            For l = 9 To Cells(Rows.Count, 10).End(xlUp).Row
            
                If Cells(l, 13) = "ã≥àÁ" Or Cells(l, 13) = "ÇªÇÃëº" Then
                    m = 7
                Else
                    If Cells(l, 14) = "T6XH8" Then
                        m = 1
                    ElseIf Cells(l, 14) = "T6XN6" Then
                        m = 2
                    ElseIf Cells(l, 14) = "T6XN5" Then
                        m = 3
                    ElseIf Cells(l, 14) = "T6XM9" Then
                        m = 4
                    ElseIf Cells(l, 14) = "T6XP6" Then
                        m = 5
                    ElseIf Cells(l, 14) = "T6XZ0" Then
                        m = 6
                    Else
                        m = 0
                    End If
                End If
                hour = Cells(l, 7) - Cells(l, 4)
                minute = Cells(l, 9) - Cells(l, 6)
                If hour < 0 Then
                    n = 24
                Else
                    n = 0
                End If
                If Cells(l, 24) = "YMãZ" Or Cells(l, 24) = "âêÕ2" Then
                    YM(m, o, p) = ((hour + n) * 60) + minute
                    p = p + 1
                ElseIf Cells(l, 24) = "êMãZ" Then
                    LE(m, o, q) = ((hour + n) * 60) + minute
                    q = q + 1
                ElseIf Cells(l, 24) = "PEãZ" Then
                    PE(m, o, r) = ((hour + n) * 60) + minute
                    r = r + 1
                ElseIf Cells(l, 24) = "ïièÿ" Then
                    HIN(m, o, s) = ((hour + n) * 60) + minute
                    s = s + 1
                End If
        
            Next
            
            For k = LBound(YM, 1) To UBound(YM, 1)
                For l = LBound(YM, 3) To UBound(YM, 3)
                    YM_ALL(k, o) = YM_ALL(k, o) + YM(k, o, l)
                Next
            Next
            For k = LBound(LE, 1) To UBound(LE, 1)
                For l = LBound(LE, 3) To UBound(LE, 3)
                    LE_ALL(k, o) = LE_ALL(k, o) + LE(k, o, l)
                Next
            Next
            For k = LBound(PE, 1) To UBound(PE, 1)
                For l = LBound(PE, 3) To UBound(PE, 3)
                    PE_ALL(k, o) = PE_ALL(k, o) + PE(k, o, l)
                Next
            Next
            For k = LBound(HIN, 1) To UBound(HIN, 1)
                For l = LBound(HIN, 3) To UBound(HIN, 3)
                    HIN_ALL(k, o) = HIN_ALL(k, o) + HIN(k, o, l)
                Next
            Next
            o = o + 1
        Next
    Next
    
    For i = LBound(YM_ALL, 1) To UBound(YM_ALL, 1)
        l = 9
        For j = LBound(YM_ALL, 2) To UBound(YM_ALL, 2)
        
            If i = 0 Then k = 4
            If i = 1 Then k = 5
            If i = 2 Then k = 6
            If i = 3 Then k = 7
            If i = 4 Then k = 8
            If i = 5 Then k = 9
            If i = 6 Then k = 11
            If i = 7 Then k = 12
            
            Worksheets("Sheet1").Cells(k, l) = YM_ALL(i, j)
            l = l + 1
        Next
    Next
    For i = LBound(LE_ALL, 1) To UBound(LE_ALL, 1)
        l = 9
        For j = LBound(LE_ALL, 2) To UBound(LE_ALL, 2)
        
            If i = 0 Then k = 16
            If i = 1 Then k = 17
            If i = 2 Then k = 18
            If i = 3 Then k = 19
            If i = 4 Then k = 20
            If i = 5 Then k = 21
            If i = 6 Then k = 23
            If i = 7 Then k = 24
            
            Worksheets("Sheet1").Cells(k, l) = LE_ALL(i, j)
            l = l + 1
        Next
    Next
    For i = LBound(PE_ALL, 1) To UBound(PE_ALL, 1)
        l = 9
        For j = LBound(PE_ALL, 2) To UBound(PE_ALL, 2)
        
            If i = 0 Then k = 28
            If i = 1 Then k = 29
            If i = 2 Then k = 30
            If i = 3 Then k = 31
            If i = 4 Then k = 32
            If i = 5 Then k = 33
            If i = 6 Then k = 35
            If i = 7 Then k = 36
            
            Worksheets("Sheet1").Cells(k, l) = PE_ALL(i, j)
            l = l + 1
        Next
    Next
    For i = LBound(HIN_ALL, 1) To UBound(HIN_ALL, 1)
        l = 9
        For j = LBound(HIN_ALL, 2) To UBound(HIN_ALL, 2)
        
            If i = 0 Then k = 40
            If i = 1 Then k = 41
            If i = 2 Then k = 42
            If i = 3 Then k = 43
            If i = 4 Then k = 44
            If i = 5 Then k = 45
            If i = 6 Then k = 47
            If i = 7 Then k = 48
            
            Worksheets("Sheet1").Cells(k, l) = HIN_ALL(i, j)
            l = l + 1
        Next
    Next
    
    Worksheets("Sheet1").Activate
    
    For i = 5 To 41 Step 12
        For j = 9 To 41
            Cells(i + 5, j) = Application.WorksheetFunction.Sum(Range(Cells(i, j), Cells(i + 4, j)))
            Cells(i + 8, j) = Application.WorksheetFunction.Sum(Range(Cells(i + 5, j), Cells(i + 7, j))) + Cells(i - 1, j)
        Next
    Next
   
End Sub
Sub é¿ê—Ç‹Ç∆Çﬂ_2()

    Dim i As Long, j As Long, k As Long, l As Long
    Dim m As Long, n As Long, o As Long, p As Long
    Dim q As Long, r As Long, s As Long
    
    Dim name As String
    Dim YM(7, 32, 100) As Long
    Dim LE(7, 32, 100) As Long
    Dim PE(7, 32, 100) As Long
    Dim HIN(7, 32, 100) As Long
    Dim YM_ALL(7, 32) As Long
    Dim LE_ALL(7, 32) As Long
    Dim PE_ALL(7, 32) As Long
    Dim HIN_ALL(7, 32) As Long
    
    For i = 20 To 22
        For j = 1 To 12
            If i = 20 Then
                k = j + 3
            Else
                k = j
            End If
            If k > 12 Then Exit For
            name = i & "îN" & Space(1) & k & "åé"
            Worksheets(name).Activate
            
            p = 0
            q = 0
            r = 0
            s = 0
            
            For l = 9 To Cells(Rows.Count, 10).End(xlUp).Row
            
                If Cells(l, 13) = "ã≥àÁ" Or Cells(l, 13) = "ÇªÇÃëº" Then
                    m = 7
                Else
                    If Cells(l, 14) = "T6XH8" Then
                        m = 1
                    ElseIf Cells(l, 14) = "T6XN6" Then
                        m = 2
                    ElseIf Cells(l, 14) = "T6XN5" Then
                        m = 3
                    ElseIf Cells(l, 14) = "T6XM9" Then
                        m = 4
                    ElseIf Cells(l, 14) = "T6XP6" Then
                        m = 5
                    ElseIf Cells(l, 14) = "T6XZ0" Then
                        m = 6
                    Else
                        m = 0
                    End If
                End If

                If Cells(l, 24) = "YMãZ" Or Cells(l, 24) = "âêÕ2" Then
                    YM(m, o, p) = Cells(l, 16).Value
                    p = p + 1
                ElseIf Cells(l, 24) = "êMãZ" Then
                    LE(m, o, q) = Cells(l, 16).Value
                    q = q + 1
                ElseIf Cells(l, 24) = "PEãZ" Then
                    PE(m, o, r) = Cells(l, 16).Value
                    r = r + 1
                ElseIf Cells(l, 24) = "ïièÿ" Then
                    HIN(m, o, s) = Cells(l, 16).Value
                    s = s + 1
                End If
        
            Next
            
            For k = LBound(YM, 1) To UBound(YM, 1)
                For l = LBound(YM, 3) To UBound(YM, 3)
                    YM_ALL(k, o) = YM_ALL(k, o) + YM(k, o, l)
                Next
            Next
            For k = LBound(LE, 1) To UBound(LE, 1)
                For l = LBound(LE, 3) To UBound(LE, 3)
                    LE_ALL(k, o) = LE_ALL(k, o) + LE(k, o, l)
                Next
            Next
            For k = LBound(PE, 1) To UBound(PE, 1)
                For l = LBound(PE, 3) To UBound(PE, 3)
                    PE_ALL(k, o) = PE_ALL(k, o) + PE(k, o, l)
                Next
            Next
            For k = LBound(HIN, 1) To UBound(HIN, 1)
                For l = LBound(HIN, 3) To UBound(HIN, 3)
                    HIN_ALL(k, o) = HIN_ALL(k, o) + HIN(k, o, l)
                Next
            Next
            o = o + 1
        Next
    Next
    
    For i = LBound(YM_ALL, 1) To UBound(YM_ALL, 1)
        l = 9
        For j = LBound(YM_ALL, 2) To UBound(YM_ALL, 2)
        
            If i = 0 Then k = 4
            If i = 1 Then k = 5
            If i = 2 Then k = 6
            If i = 3 Then k = 7
            If i = 4 Then k = 8
            If i = 5 Then k = 9
            If i = 6 Then k = 11
            If i = 7 Then k = 12
            
            Worksheets("Sheet2").Cells(k, l) = YM_ALL(i, j)
            l = l + 1
        Next
    Next
    For i = LBound(LE_ALL, 1) To UBound(LE_ALL, 1)
        l = 9
        For j = LBound(LE_ALL, 2) To UBound(LE_ALL, 2)
        
            If i = 0 Then k = 16
            If i = 1 Then k = 17
            If i = 2 Then k = 18
            If i = 3 Then k = 19
            If i = 4 Then k = 20
            If i = 5 Then k = 21
            If i = 6 Then k = 23
            If i = 7 Then k = 24
            
            Worksheets("Sheet2").Cells(k, l) = LE_ALL(i, j)
            l = l + 1
        Next
    Next
    For i = LBound(PE_ALL, 1) To UBound(PE_ALL, 1)
        l = 9
        For j = LBound(PE_ALL, 2) To UBound(PE_ALL, 2)
        
            If i = 0 Then k = 28
            If i = 1 Then k = 29
            If i = 2 Then k = 30
            If i = 3 Then k = 31
            If i = 4 Then k = 32
            If i = 5 Then k = 33
            If i = 6 Then k = 35
            If i = 7 Then k = 36
            
            Worksheets("Sheet2").Cells(k, l) = PE_ALL(i, j)
            l = l + 1
        Next
    Next
    For i = LBound(HIN_ALL, 1) To UBound(HIN_ALL, 1)
        l = 9
        For j = LBound(HIN_ALL, 2) To UBound(HIN_ALL, 2)
        
            If i = 0 Then k = 40
            If i = 1 Then k = 41
            If i = 2 Then k = 42
            If i = 3 Then k = 43
            If i = 4 Then k = 44
            If i = 5 Then k = 45
            If i = 6 Then k = 47
            If i = 7 Then k = 48
            
            Worksheets("Sheet2").Cells(k, l) = HIN_ALL(i, j)
            l = l + 1
        Next
    Next
    
    Worksheets("Sheet2").Activate
    
    For i = 5 To 41 Step 12
        For j = 9 To 41
            Cells(i + 5, j) = Application.WorksheetFunction.Sum(Range(Cells(i, j), Cells(i + 4, j)))
            Cells(i + 8, j) = Application.WorksheetFunction.Sum(Range(Cells(i + 5, j), Cells(i + 7, j))) + Cells(i - 1, j)
        Next
    Next
   
End Sub
