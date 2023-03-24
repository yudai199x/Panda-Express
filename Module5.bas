Attribute VB_Name = "Module2"
Option Explicit

Sub 貼り付け()

    Dim Data As String
    Dim FileName As Variant
    Dim FileNum As Long
    Dim myRange As Range
    
    Dim SheetObj As Worksheet
    Dim ICCS(35) As String
    Dim VDD(16, 3) As String
    Dim VDDA(27, 3) As String
    Dim VFOUR(15, 3) As String
    Dim VDDSAP0(15, 3) As String
    Dim VDDSAP1(15, 3) As String
    Dim IREF(15, 3) As String
    Dim VREF(15, 3) As String
    
    Dim Flag As Long
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim e As Long
    Dim f As Long
    Dim g As Long
    Dim h As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim m As Long
    Dim n As Long
    Dim o As Long
    Dim p As Long
    Dim q As Long
    Dim r As Long
    Dim s As Long
    Dim t As Long
    
    '入力シートにテキストファイルから行単位で取り込む
    
    Set SheetObj = ThisWorkbook.Worksheets("入力")
     
    
    'PTログを開くためのウインドウを開く
    
    FileName = Application.GetOpenFilename("ぜんぶ,*.*")
    
    
    'キャンセルやxで閉じたときは処理終了
    
    If FileName = False Then
    
    Exit Sub
       
    End If
    
    'ファイル番号の取得
    
    FileNum = FreeFile
    
    Open FileName For Input As #FileNum
    
    'Excelの何行目から出力するか
    
    i = 1
    
    'Do Until～Loopは条件式が真になるまで繰り返す
    
    Do Until EOF(FileNum)
    
    'データを1行ずつ読み込む
 
    Line Input #FileNum, Data
    
    '読み込んだデータをA列に出力
    SheetObj.Cells(i, 1) = Data
    i = i + 1
    
    
    '"DELAY 0x32(50)"だった場合フラグを立てる Iccs用
    If Left(Data, 14) = "DELAY 0x32(50)" Then Flag = 1
        If Flag = 1 Then
    
            If InStr(Data, " : ") > 0 Then
                Data = Mid(Data, 7)
                Data = Replace(Data, " uA", "")
                ICCS(j) = Data
                j = j + 1
            End If
            If InStr(Data, "Finish") Then
                Flag = 0
            End If
        End If
        
            
    '"ALL依存,CE=0,Chip=0"と完全一致したセルが見つかれば実行
    Set myRange = Worksheets("入力").Range("A:A").Find(What:="ALL依存,CE=,0", lookAt:=xlWhole)
    If Not myRange Is Nothing Then
        
        
        If Flag = 0 Then j = 0
   
        '"VDD依存"だった場合フラグを立てる
        If Left(Data, 5) = "VDD依存" Then Flag = 1
            If Flag = 1 And VDD(16, 0) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 16, 10)
                    VDD(j, 0) = Data
                    j = j + 1
                End If
                If j >= 17 Then
                    Flag = 0
                End If
            End If
    
        '"VDDA依存"だった場合フラグを立てる
        If Left(Data, 6) = "VDDA依存" Then Flag = 2
            If Flag = 2 And VDDA(27, 0) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 17, 10)
                    VDDA(j, 0) = Data
                    j = j + 1
                End If
                If j >= 28 Then
                    Flag = 0
                End If
            End If

        '"VREF依存"だった場合フラグを立てる
        If Left(Data, 6) = "VREF依存" Then Flag = 3
            If Flag = 3 And VREF(15, 0) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 17, 10)
                    VREF(j, 0) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If

        '"VFOUR依存"だった場合フラグを立てる
        If Left(Data, 7) = "VFOUR依存" Then Flag = 4
            If Flag = 4 And VFOUR(15, 0) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 18, 10)
                    VFOUR(j, 0) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If

        '"VDDSAP0依存"だった場合フラグを立てる
        If Left(Data, 11) = "VDDSA_PB0依存" Then Flag = 5
            If Flag = 5 And VDDSAP0(15, 0) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 22, 10)
                    VDDSAP0(j, 0) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If

        '"VDDSAP1依存"だった場合フラグを立てる
        If Left(Data, 11) = "VDDSA_PB1依存" Then Flag = 6
            If Flag = 6 And VDDSAP1(15, 0) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 22, 10)
                    VDDSAP1(j, 0) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If

        '"IREF依存"だった場合フラグを立てる
        If Left(Data, 6) = "IREF依存" Then Flag = 7
            If Flag = 7 And IREF(15, 0) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 17, 10)
                    IREF(j, 0) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If
    End If
    
    
    '"ALL依存,CE=0,Chip=1"と完全一致したセルが見つかれば実行
    Set myRange = Worksheets("入力").Range("A:A").Find(What:="ALL依存,CE=,1", lookAt:=xlWhole)
    If Not myRange Is Nothing Then
    
    
        If Flag = 0 Then j = 0
   
        '"VDD依存"だった場合フラグを立てる
        If Left(Data, 5) = "VDD依存" Then Flag = 1
            If Flag = 1 And VDD(16, 1) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 16, 10)
                    VDD(j, 1) = Data
                    j = j + 1
                End If
                If j >= 17 Then
                    Flag = 0
                End If
            End If
    
        '"VDDA依存"だった場合フラグを立てる
        If Left(Data, 6) = "VDDA依存" Then Flag = 2
            If Flag = 2 And VDDA(27, 1) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 17, 10)
                    VDDA(j, 1) = Data
                    j = j + 1
                End If
                If j >= 28 Then
                    Flag = 0
                End If
            End If

        '"VREF依存"だった場合フラグを立てる
        If Left(Data, 6) = "VREF依存" Then Flag = 3
            If Flag = 3 And VREF(15, 1) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 17, 10)
                    VREF(j, 1) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If

        '"VFOUR依存"だった場合フラグを立てる
        If Left(Data, 7) = "VFOUR依存" Then Flag = 4
            If Flag = 4 And VFOUR(15, 1) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 18, 10)
                    VFOUR(j, 1) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If

        '"VDDSAP0依存"だった場合フラグを立てる
        If Left(Data, 11) = "VDDSA_PB0依存" Then Flag = 5
            If Flag = 5 And VDDSAP0(15, 1) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 22, 10)
                    VDDSAP0(j, 1) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If

        '"VDDSAP1依存"だった場合フラグを立てる
        If Left(Data, 11) = "VDDSA_PB1依存" Then Flag = 6
            If Flag = 6 And VDDSAP1(15, 1) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 22, 10)
                    VDDSAP1(j, 1) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If

        '"IREF依存"だった場合フラグを立てる
        If Left(Data, 6) = "IREF依存" Then Flag = 7
            If Flag = 7 And IREF(15, 1) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 17, 10)
                    IREF(j, 1) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If
    End If
    
    
    '"ALL依存,CE=1,Chip=0"と完全一致したセルが見つかれば実行
    Set myRange = Worksheets("入力").Range("A:A").Find(What:="ALL依存,CE=,2", lookAt:=xlWhole)
    If Not myRange Is Nothing Then
    
    
        If Flag = 0 Then j = 0
   
        '"VDD依存"だった場合フラグを立てる
        If Left(Data, 5) = "VDD依存" Then Flag = 1
            If Flag = 1 And VDD(16, 2) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 16, 10)
                    VDD(j, 2) = Data
                    j = j + 1
                End If
                If j >= 17 Then
                    Flag = 0
                End If
            End If
    
        '"VDDA依存"だった場合フラグを立てる
        If Left(Data, 6) = "VDDA依存" Then Flag = 2
            If Flag = 2 And VDDA(27, 2) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 17, 10)
                    VDDA(j, 2) = Data
                    j = j + 1
                End If
                If j >= 28 Then
                    Flag = 0
                End If
            End If

        '"VREF依存"だった場合フラグを立てる
        If Left(Data, 6) = "VREF依存" Then Flag = 3
            If Flag = 3 And VREF(15, 2) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 17, 10)
                    VREF(j, 2) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If

        '"VFOUR依存"だった場合フラグを立てる
        If Left(Data, 7) = "VFOUR依存" Then Flag = 4
            If Flag = 4 And VFOUR(15, 2) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 18, 10)
                    VFOUR(j, 2) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If

        '"VDDSAP0依存"だった場合フラグを立てる
        If Left(Data, 11) = "VDDSA_PB0依存" Then Flag = 5
            If Flag = 5 And VDDSAP0(15, 2) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 22, 10)
                    VDDSAP0(j, 2) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If

        '"VDDSAP1依存"だった場合フラグを立てる
        If Left(Data, 11) = "VDDSA_PB1依存" Then Flag = 6
            If Flag = 6 And VDDSAP1(15, 2) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 22, 10)
                    VDDSAP1(j, 2) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If

        '"IREF依存"だった場合フラグを立てる
        If Left(Data, 6) = "IREF依存" Then Flag = 7
            If Flag = 7 And IREF(15, 2) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 17, 10)
                    IREF(j, 2) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If
    End If
    

    '"ALL依存,CE=1,Chip=1"と完全一致したセルが見つかれば実行
    Set myRange = Worksheets("入力").Range("A:A").Find(What:="ALL依存,CE=,3", lookAt:=xlWhole)
    If Not myRange Is Nothing Then
    
    
        If Flag = 0 Then j = 0
   
        '"VDD依存"だった場合フラグを立てる
        If Left(Data, 5) = "VDD依存" Then Flag = 1
            If Flag = 1 And VDD(16, 3) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 16, 10)
                    VDD(j, 3) = Data
                    j = j + 1
                End If
                If j >= 17 Then
                    Flag = 0
                End If
            End If
    
        '"VDDA依存"だった場合フラグを立てる
        If Left(Data, 6) = "VDDA依存" Then Flag = 2
            If Flag = 2 And VDDA(27, 3) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 17, 10)
                    VDDA(j, 3) = Data
                    j = j + 1
                End If
                If j >= 28 Then
                    Flag = 0
                End If
            End If

        '"VREF依存"だった場合フラグを立てる
        If Left(Data, 6) = "VREF依存" Then Flag = 3
            If Flag = 3 And VREF(15, 3) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 17, 10)
                    VREF(j, 3) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If

        '"VFOUR依存"だった場合フラグを立てる
        If Left(Data, 7) = "VFOUR依存" Then Flag = 4
            If Flag = 4 And VFOUR(15, 3) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 18, 10)
                    VFOUR(j, 3) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If

        '"VDDSAP0依存"だった場合フラグを立てる
        If Left(Data, 11) = "VDDSA_PB0依存" Then Flag = 5
            If Flag = 5 And VDDSAP0(15, 3) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 22, 10)
                    VDDSAP0(j, 3) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If

        '"VDDSAP1依存"だった場合フラグを立てる
        If Left(Data, 11) = "VDDSA_PB1依存" Then Flag = 6
            If Flag = 6 And VDDSAP1(15, 3) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 22, 10)
                    VDDSAP1(j, 3) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If

        '"IREF依存"だった場合フラグを立てる
        If Left(Data, 6) = "IREF依存" Then Flag = 7
            If Flag = 7 And IREF(15, 3) = "" Then
    
                If InStr(Data, "Iccs") > 0 Then
                    Data = Mid(Data, 17, 10)
                    IREF(j, 3) = Data
                    j = j + 1
                End If
                If j >= 16 Then
                    Flag = 0
                End If
            End If
    End If
    
    Loop
    
    Close #FileNum
     
    Columns(1).Font.Name = "ＭＳ ゴシック"
    
    '処理の高速化
    Application.ScreenUpdating = True
    
    'シートに出力
    Worksheets("Box").Activate
    

    i = 143
    
    For k = LBound(ICCS) To UBound(ICCS)

        If Not ICCS(k) = "" Then
            Cells(i, 2).Value = ICCS(k)
            i = i + 1
        End If
        
    Next
    
    a = 4: b = 4: c = 4: d = 4: e = 4: f = 4: g = 4: h = 4
    m = 4: n = 4: o = 4: p = 4: q = 4: r = 4: s = 4: t = 4
    
    
    For l = LBound(VDD, 2) To UBound(VDD, 2)
        For k = LBound(VDD, 1) To UBound(VDD, 1)
        
            If Not VDD(k, l) = "" Then
                If l = 0 Then
                    Cells(a, 2).Value = VDD(k, l)
                    a = a + 1
                        
                ElseIf l = 1 Then
                    Cells(b, 5).Value = VDD(k, l)
                    b = b + 1
                        
                ElseIf l = 2 Then
                    Cells(c, 8).Value = VDD(k, l)
                    c = c + 1
                        
                ElseIf l = 3 Then
                    Cells(d, 11).Value = VDD(k, l)
                    d = d + 1
                        
                ElseIf l = 4 Then
                    Cells(e, 14).Value = VDD(k, l)
                    e = e + 1
                        
                ElseIf l = 5 Then
                    Cells(f, 17).Value = VDD(k, l)
                    f = f + 1
                        
                ElseIf l = 6 Then
                    Cells(g, 20).Value = VDD(k, l)
                    g = g + 1
                        
                ElseIf l = 7 Then
                    Cells(h, 23).Value = VDD(k, l)
                    h = h + 1
                        
                ElseIf l = 8 Then
                    Cells(m, 26).Value = VDD(k, l)
                    m = m + 1
                        
                ElseIf l = 9 Then
                    Cells(n, 29).Value = VDD(k, l)
                    n = n + 1
                        
                ElseIf l = 10 Then
                    Cells(o, 32).Value = VDD(k, l)
                    o = o + 1
                        
                ElseIf l = 11 Then
                    Cells(p, 35).Value = VDD(k, l)
                    p = p + 1
                        
                ElseIf l = 12 Then
                    Cells(q, 38).Value = VDD(k, l)
                    q = q + 1
                        
                ElseIf l = 13 Then
                    Cells(r, 41).Value = VDD(k, l)
                    r = r + 1
                        
                ElseIf l = 14 Then
                    Cells(s, 44).Value = VDD(k, l)
                    s = s + 1
                        
                Else
                    Cells(t, 47).Value = VDD(k, l)
                    t = t + 1
                                           
                End If
            End If
            
        Next
    Next
    
    
    a = 23: b = 23: c = 23: d = 23: e = 23: f = 23: g = 23: h = 23
    m = 23: n = 23: o = 23: p = 23: q = 23: r = 23: s = 23: t = 23
    
    
    For l = LBound(VDDA, 2) To UBound(VDDA, 2)
        For k = LBound(VDDA, 1) To UBound(VDDA, 1)
        
            If Not VDDA(k, l) = "" Then
                If l = 0 Then
                    Cells(a, 2).Value = VDDA(k, l)
                    a = a + 1
                        
                ElseIf l = 1 Then
                    Cells(b, 5).Value = VDDA(k, l)
                    b = b + 1
                        
                ElseIf l = 2 Then
                    Cells(c, 8).Value = VDDA(k, l)
                    c = c + 1
                        
                ElseIf l = 3 Then
                    Cells(d, 11).Value = VDDA(k, l)
                    d = d + 1
                        
                ElseIf l = 4 Then
                    Cells(e, 14).Value = VDDA(k, l)
                    e = e + 1
                        
                ElseIf l = 5 Then
                    Cells(f, 17).Value = VDDA(k, l)
                    f = f + 1
                        
                ElseIf l = 6 Then
                    Cells(g, 20).Value = VDDA(k, l)
                    g = g + 1
                        
                ElseIf l = 7 Then
                    Cells(h, 23).Value = VDDA(k, l)
                    h = h + 1
                        
                ElseIf l = 8 Then
                    Cells(m, 26).Value = VDDA(k, l)
                    m = m + 1
                        
                ElseIf l = 9 Then
                    Cells(n, 29).Value = VDDA(k, l)
                    n = n + 1
                        
                ElseIf l = 10 Then
                    Cells(o, 32).Value = VDDA(k, l)
                    o = o + 1
                        
                ElseIf l = 11 Then
                    Cells(p, 35).Value = VDDA(k, l)
                    p = p + 1
                        
                ElseIf l = 12 Then
                    Cells(q, 38).Value = VDDA(k, l)
                    q = q + 1
                        
                ElseIf l = 13 Then
                    Cells(r, 41).Value = VDDA(k, l)
                    r = r + 1
                        
                ElseIf l = 14 Then
                    Cells(s, 44).Value = VDDA(k, l)
                    s = s + 1
                        
                Else
                    Cells(t, 47).Value = VDDA(k, l)
                    t = t + 1
                    
                End If
            End If
            
        Next
    Next
    
    
    a = 53: b = 53: c = 53: d = 53: e = 53: f = 53: g = 53: h = 53
    m = 53: n = 53: o = 53: p = 53: q = 53: r = 53: s = 53: t = 53
    
    
    For l = LBound(VREF, 2) To UBound(VREF, 2)
        For k = LBound(VREF, 1) To UBound(VREF, 1)
        
            If Not VREF(k, l) = "" Then
                If l = 0 Then
                   Cells(a, 2).Value = VREF(k, l)
                   a = a + 1
                   
                ElseIf l = 1 Then
                    Cells(b, 5).Value = VREF(k, l)
                    b = b + 1
                    
                ElseIf l = 2 Then
                    Cells(c, 8).Value = VREF(k, l)
                    c = c + 1
                        
                ElseIf l = 3 Then
                    Cells(d, 11).Value = VREF(k, l)
                    d = d + 1
                        
                ElseIf l = 4 Then
                    Cells(e, 14).Value = VREF(k, l)
                    e = e + 1
                        
                ElseIf l = 5 Then
                    Cells(f, 17).Value = VREF(k, l)
                    f = f + 1
                        
                ElseIf l = 6 Then
                    Cells(g, 20).Value = VREF(k, l)
                    g = g + 1
                        
                ElseIf l = 7 Then
                    Cells(h, 23).Value = VREF(k, l)
                    h = h + 1
                        
                ElseIf l = 8 Then
                    Cells(m, 26).Value = VREF(k, l)
                    m = m + 1
                        
                ElseIf l = 9 Then
                    Cells(n, 29).Value = VREF(k, l)
                    n = n + 1
                        
                ElseIf l = 10 Then
                    Cells(o, 32).Value = VREF(k, l)
                    o = o + 1
                        
                ElseIf l = 11 Then
                    Cells(p, 35).Value = VREF(k, l)
                    p = p + 1
                        
                ElseIf l = 12 Then
                    Cells(q, 38).Value = VREF(k, l)
                    q = q + 1
                        
                ElseIf l = 13 Then
                    Cells(r, 41).Value = VREF(k, l)
                    r = r + 1
                        
                ElseIf l = 14 Then
                    Cells(s, 44).Value = VREF(k, l)
                    s = s + 1
                        
                Else
                    Cells(t, 47).Value = VREF(k, l)
                    t = t + 1
                    
                End If
            End If
            
        Next
    Next


    a = 71: b = 71: c = 71: d = 71: e = 71: f = 71: g = 71: h = 71
    m = 71: n = 71: o = 71: p = 71: q = 71: r = 71: s = 71: t = 71
    
    
    For l = LBound(VFOUR, 2) To UBound(VFOUR, 2)
        For k = LBound(VFOUR, 1) To UBound(VFOUR, 1)
        
            If Not VFOUR(k, l) = "" Then
                If l = 0 Then
                    Cells(a, 2).Value = VFOUR(k, l)
                    a = a + 1
                    
                ElseIf l = 1 Then
                    Cells(b, 5).Value = VFOUR(k, l)
                    b = b + 1
                    
                ElseIf l = 2 Then
                    Cells(c, 8).Value = VFOUR(k, l)
                    c = c + 1
                        
                ElseIf l = 3 Then
                    Cells(d, 11).Value = VFOUR(k, l)
                    d = d + 1
                        
                ElseIf l = 4 Then
                    Cells(e, 14).Value = VFOUR(k, l)
                    e = e + 1
                        
                ElseIf l = 5 Then
                    Cells(f, 17).Value = VFOUR(k, l)
                    f = f + 1
                        
                ElseIf l = 6 Then
                    Cells(g, 20).Value = VFOUR(k, l)
                    g = g + 1
                        
                ElseIf l = 7 Then
                    Cells(h, 23).Value = VFOUR(k, l)
                    h = h + 1
                        
                ElseIf l = 8 Then
                    Cells(m, 26).Value = VFOUR(k, l)
                    m = m + 1
                        
                ElseIf l = 9 Then
                    Cells(n, 29).Value = VFOUR(k, l)
                    n = n + 1
                        
                ElseIf l = 10 Then
                    Cells(o, 32).Value = VFOUR(k, l)
                    o = o + 1
                        
                ElseIf l = 11 Then
                    Cells(p, 35).Value = VFOUR(k, l)
                    p = p + 1
                        
                ElseIf l = 12 Then
                    Cells(q, 38).Value = VFOUR(k, l)
                    q = q + 1
                        
                ElseIf l = 13 Then
                    Cells(r, 41).Value = VFOUR(k, l)
                    r = r + 1
                        
                ElseIf l = 14 Then
                    Cells(s, 44).Value = VFOUR(k, l)
                    s = s + 1
                        
                Else
                    Cells(t, 47).Value = VFOUR(k, l)
                    t = t + 1
                    
                End If
            End If
            
        Next
    Next
    
    
    a = 89: b = 89: c = 89: d = 89: e = 89: f = 89: g = 89: h = 89
    m = 89: n = 89: o = 89: p = 89: q = 89: r = 89: s = 89: t = 89
    
    
    For l = LBound(VDDSAP0, 2) To UBound(VDDSAP0, 2)
        For k = LBound(VDDSAP0, 1) To UBound(VDDSAP0, 1)
        
            If Not VDDSAP0(k, l) = "" Then
                If l = 0 Then
                    Cells(a, 2).Value = VDDSAP0(k, l)
                    a = a + 1
                    
                ElseIf l = 1 Then
                    Cells(b, 5).Value = VDDSAP0(k, l)
                    b = b + 1
                    
                ElseIf l = 2 Then
                    Cells(c, 8).Value = VDDSAP0(k, l)
                    c = c + 1
                        
                ElseIf l = 3 Then
                    Cells(d, 11).Value = VDDSAP0(k, l)
                    d = d + 1
                        
                ElseIf l = 4 Then
                    Cells(e, 14).Value = VDDSAP0(k, l)
                    e = e + 1
                        
                ElseIf l = 5 Then
                    Cells(f, 17).Value = VDDSAP0(k, l)
                    f = f + 1
                        
                ElseIf l = 6 Then
                    Cells(g, 20).Value = VDDSAP0(k, l)
                    g = g + 1
                        
                ElseIf l = 7 Then
                    Cells(h, 23).Value = VDDSAP0(k, l)
                    h = h + 1
                        
                ElseIf l = 8 Then
                    Cells(m, 26).Value = VDDSAP0(k, l)
                    m = m + 1
                        
                ElseIf l = 9 Then
                    Cells(n, 29).Value = VDDSAP0(k, l)
                    n = n + 1
                        
                ElseIf l = 10 Then
                    Cells(o, 32).Value = VDDSAP0(k, l)
                    o = o + 1
                        
                ElseIf l = 11 Then
                    Cells(p, 35).Value = VDDSAP0(k, l)
                    p = p + 1
                        
                ElseIf l = 12 Then
                    Cells(q, 38).Value = VDDSAP0(k, l)
                    q = q + 1
                        
                ElseIf l = 13 Then
                    Cells(r, 41).Value = VDDSAP0(k, l)
                    r = r + 1
                        
                ElseIf l = 14 Then
                    Cells(s, 44).Value = VDDSAP0(k, l)
                    s = s + 1
                        
                Else
                    Cells(t, 47).Value = VDDSAP0(k, l)
                    t = t + 1
                    
                End If
            End If
            
        Next
    Next


    a = 107: b = 107: c = 107: d = 107: e = 107: f = 107: g = 107: h = 107
    m = 107: n = 107: o = 107: p = 107: q = 107: r = 107: s = 107: t = 107

    
    For l = LBound(VDDSAP1, 2) To UBound(VDDSAP1, 2)
        For k = LBound(VDDSAP1, 1) To UBound(VDDSAP1, 1)

            If Not VDDSAP1(k, l) = "" Then
                If l = 0 Then
                    Cells(a, 2).Value = VDDSAP1(k, l)
                    a = a + 1
                    
                ElseIf l = 1 Then
                    Cells(b, 5).Value = VDDSAP1(k, l)
                    b = b + 1
                    
                ElseIf l = 2 Then
                    Cells(c, 8).Value = VDDSAP1(k, l)
                    c = c + 1
                        
                ElseIf l = 3 Then
                    Cells(d, 11).Value = VDDSAP1(k, l)
                    d = d + 1
                        
                ElseIf l = 4 Then
                    Cells(e, 14).Value = VDDSAP1(k, l)
                    e = e + 1
                        
                ElseIf l = 5 Then
                    Cells(f, 17).Value = VDDSAP1(k, l)
                    f = f + 1
                        
                ElseIf l = 6 Then
                    Cells(g, 20).Value = VDDSAP1(k, l)
                    g = g + 1
                        
                ElseIf l = 7 Then
                    Cells(h, 23).Value = VDDSAP1(k, l)
                    h = h + 1
                        
                ElseIf l = 8 Then
                    Cells(m, 26).Value = VDDSAP1(k, l)
                    m = m + 1
                        
                ElseIf l = 9 Then
                    Cells(n, 29).Value = VDDSAP1(k, l)
                    n = n + 1
                        
                ElseIf l = 10 Then
                    Cells(o, 32).Value = VDDSAP1(k, l)
                    o = o + 1
                        
                ElseIf l = 11 Then
                    Cells(p, 35).Value = VDDSAP1(k, l)
                    p = p + 1
                        
                ElseIf l = 12 Then
                    Cells(q, 38).Value = VDDSAP1(k, l)
                    q = q + 1
                        
                ElseIf l = 13 Then
                    Cells(r, 41).Value = VDDSAP1(k, l)
                    r = r + 1
                        
                ElseIf l = 14 Then
                    Cells(s, 44).Value = VDDSAP1(k, l)
                    s = s + 1
                        
                Else
                    Cells(t, 47).Value = VDDSAP1(k, l)
                    t = t + 1
                    
                End If
            End If
            
        Next
    Next


    a = 125: b = 125: c = 125: d = 125: e = 125: f = 125: g = 125: h = 125
    m = 125: n = 125: o = 125: p = 125: q = 125: r = 125: s = 125: t = 125

    
    For l = LBound(IREF, 2) To UBound(IREF, 2)
        For k = LBound(IREF, 1) To UBound(IREF, 1)
        
            If Not IREF(k, l) = "" Then
                If l = 0 Then
                    Cells(a, 2).Value = IREF(k, l)
                    a = a + 1
                    
                ElseIf l = 1 Then
                    Cells(b, 5).Value = IREF(k, l)
                    b = b + 1
                    
                ElseIf l = 2 Then
                    Cells(c, 8).Value = IREF(k, l)
                    c = c + 1
                        
                ElseIf l = 3 Then
                    Cells(d, 11).Value = IREF(k, l)
                    d = d + 1
                        
                ElseIf l = 4 Then
                    Cells(e, 14).Value = IREF(k, l)
                    e = e + 1
                        
                ElseIf l = 5 Then
                    Cells(f, 17).Value = IREF(k, l)
                    f = f + 1
                        
                ElseIf l = 6 Then
                    Cells(g, 20).Value = IREF(k, l)
                    g = g + 1
                        
                ElseIf l = 7 Then
                    Cells(h, 23).Value = IREF(k, l)
                    h = h + 1
                        
                ElseIf l = 8 Then
                    Cells(m, 26).Value = IREF(k, l)
                    m = m + 1
                        
                ElseIf l = 9 Then
                    Cells(n, 29).Value = IREF(k, l)
                    n = n + 1
                        
                ElseIf l = 10 Then
                    Cells(o, 32).Value = IREF(k, l)
                    o = o + 1
                        
                ElseIf l = 11 Then
                    Cells(p, 35).Value = IREF(k, l)
                    p = p + 1
                        
                ElseIf l = 12 Then
                    Cells(q, 38).Value = IREF(k, l)
                    q = q + 1
                        
                ElseIf l = 13 Then
                    Cells(r, 41).Value = IREF(k, l)
                    r = r + 1
                        
                ElseIf l = 14 Then
                    Cells(s, 44).Value = IREF(k, l)
                    s = s + 1
                        
                Else
                    Cells(t, 47).Value = IREF(k, l)
                    t = t + 1
                End If
            End If
            
        Next
    Next

End Sub





