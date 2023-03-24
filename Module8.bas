Attribute VB_Name = "Module1"
Option Explicit
Public L As Double, T As Double, W As Double, H As Double, i As Long, j As Long, d As Date
Public started_day As Date, finished_day As Date, flag As Byte, flag2 As Byte, bar As Shape
Public started_at As Date, finished_at As Date, ans As Integer, MaxRow As Long
Sub TAT()

    Dim seigou_started_at As Date
    Dim seigou_finished_at As Date
    
    Worksheets("場所特定_TAT管理台帳").Activate
    Columns(5).Validation.Delete
    For i = 3 To Cells(3, 7).End(xlDown).Row
        If Cells(i, 7) = "" And Cells(i, 8) = "" Then
            Cells(i, 7) = Format(Now)
            With Worksheets("ガントチャート")
                For j = 1 To 88 Step 3
                    .Cells(j, 1) = ""
                Next
            End With
            Exit For
        End If
        If Cells(i, 7) <> "" And Cells(i, 8) = "" Then
            If Cells(i, 2) <> "LE" Then
                Cells(i, 1) = "⑩場所特定"
                Cells(i, 2) = "(品証)場所特定"
                Cells(i, 3) = "作業時間"
                Cells(i, 5) = "入力不要"
                Cells(i, 8) = Format(Now)
                Cells(i + 1, 7) = Format(Now)
                
                seigou_started_at = CDate(Year(Cells(i, 8)) & "/" & Month(Cells(i, 8)) & "/" & Day(Cells(i, 8)) & Space(2) & "8:30:00")
                seigou_finished_at = CDate(Year(Cells(i, 8)) & "/" & Month(Cells(i, 8)) & "/" & Day(Cells(i, 8)) & Space(2) & "17:15:00")
                
                Select Case Cells(1, 1).Interior.ColorIndex
                    Case 41
                        flag = 48
                        Cells(i, 4) = "発光解析"
                        Cells(i, 1) = "⑨発光解析"
                    Case 46
                        flag = 53
                        Cells(i, 4) = "場所特定"
                    Case 44
                        flag = 51
                        Cells(i, 4) = "整合性確認"
                        Cells(i, 3) = "待機時間"
                        Cells(i, 8) = seigou_started_at
                        Cells(i + 1, 7) = seigou_started_at
                    Case 15
                        flag = 22
                        Cells(i, 4) = "PFA指示書"
                End Select
    
                If Cells(i - 1, 3) = "作業時間" Then
                    If Cells(i, 4) <> "整合性確認" Then
                        flag = 62
                    End If
                    Cells(i, 3) = "待機時間"
                End If
                If Day(Cells(i, 7)) <> Day(Cells(i, 8)) Then
                    flag = 50
                    Cells(i, 3) = "未作業時間(帰宅、休日)"
                End If
                If Cells(i, 7) = seigou_started_at Then
                    Cells(i, 8) = seigou_finished_at
                    Cells(i + 1, 7) = seigou_finished_at
                End If
            
                If Cells(i, 3) = "待機時間" And (Cells(i, 4) = "発光解析" Or Cells(i, 4) = "場所特定") Then
                    Cells(i, 5).Validation.Add Type:=xlValidateList, Formula1:="入力不要,指示待ち,方針検討,リソース不足(装置)"
                End If
            Else
                Cells(i, 8) = Format(Now)
                Cells(i + 1, 7) = Format(Now)
            End If
            started_day = CDate(Year(Cells(i, 7)) & "/" & Month(Cells(i, 7)) & "/" & Day(Cells(i, 7)))
            finished_day = CDate(Year(Cells(i, 8)) & "/" & Month(Cells(i, 8)) & "/" & Day(Cells(i, 8)))
            started_at = CDate(Hour(Cells(i, 7)) & ":" & Minute(Cells(i, 7)) & ":00")
            finished_at = CDate(Hour(Cells(i, 8)) & ":" & Minute(Cells(i, 8)) & ":00")
            
            Worksheets("ガントチャート").Activate
            d = DateAdd("d", 1, started_day - 1)
            For j = 1 To 88 Step 3
                If Cells(j, 1) = "" Then
                    Cells(j, 1) = d
                    d = d + 1
                End If
            Next
            For j = 1 To 88 Step 3
                flag2 = 0
                T = Cells(j + 2, Hour(started_at) + 1).Top
                H = Cells(j + 2, Hour(started_at) + 1).Height
                
                If started_day = finished_day Then
                    If Cells(j, 1) = started_day Then
                        L = Cells(j + 2, Hour(started_at) + 1).Left + Cells(j, 1).Width * (Minute(started_at) / 60)
                        W = Cells(j + 2, Hour(finished_at) + 1).Left + Cells(j, 1).Width * (Minute(finished_at) / 60) - L
                        flag2 = 1
                    End If
                Else
                    Select Case Cells(j, 1)
                        Case started_day
                            L = Cells(j + 2, Hour(started_at) + 1).Left + Cells(j, 1).Width * (Minute(started_at) / 60)
                            W = Cells(j + 2, 25).Left - L
                            flag2 = 1
                        Case started_day + 1 To finished_day - 1
                            L = Cells(j + 2, 1).Left
                            W = Cells(j + 2, 1).Width * 24
                            flag2 = 1
                        Case finished_day
                            L = Cells(j + 2, 1).Left
                            W = Cells(j + 2, Hour(finished_at) + 1).Left + Cells(j, 1).Width * (Minute(finished_at) / 60)
                            flag2 = 1
                    End Select
                End If
                If flag2 = 1 Then
                    Set bar = ActiveSheet.Shapes.AddShape(msoTextOrientationHorizontal, L, T, W, H)
                    bar.Fill.ForeColor.SchemeColor = flag
                End If
            Next
            Worksheets("場所特定_TAT管理台帳").Activate
            Exit For
        End If
    Next
    
End Sub
Sub 作業変更()
    ans = MsgBox("待機時間に変更しますか？", vbYesNo)
    If ans = vbYes Then
        Worksheets("場所特定_TAT管理台帳").Activate
        Columns(5).Validation.Delete
        MaxRow = Cells(3, 3).End(xlDown).Row
        For i = 3 To MaxRow
            If Cells(i, 3) = "作業時間" And Cells(i + 1, 3) = "" Then
                Cells(i, 3) = "待機時間"
                Select Case Cells(i, 4)
                    Case "発光解析", "場所特定"
                        Cells(i, 5).Validation.Add Type:=xlValidateList, Formula1:="入力不要,指示待ち,方針検討,リソース不足(装置)"
                End Select
                Exit For
            End If
        Next
        With Worksheets("ガントチャート")
            .Shapes(.Shapes.Count).Fill.ForeColor.SchemeColor = 62
        End With
    End If
End Sub
Sub TAT集計()

    Dim total_working_days As Byte
    Dim total_working_times As Date
    
    Range(Cells(13, 12), Cells(18, 12)).Clear
    MaxRow = Cells(3, 8).End(xlDown).Row
    For i = 3 To MaxRow
        If i = MaxRow Then
            Cells(i + 1, 7).Clear
        End If
        started_at = CDate(Cells(i, 7))
        finished_at = CDate(Cells(i, 8))
        Cells(i, 9) = finished_at - started_at
    Next
    For i = 3 To MaxRow
        If Cells(i, 3) = "作業時間" Then
            Select Case Cells(i, 4)
                Case "発光解析"
                   Cells(13, 12) = Cells(13, 12) + Cells(i, 9)
                Case "場所特定"
                    Cells(14, 12) = Cells(14, 12) + Cells(i, 9)
                Case "PFA指示書"
                    Cells(16, 12) = Cells(16, 12) + Cells(i, 9)
            End Select
        End If
            If Cells(i, 4) = "整合性確認" Then Cells(15, 12) = Cells(15, 12) + Cells(i, 9)
            If Cells(i, 3) = "待機時間" Then Cells(17, 12) = Cells(17, 12) + Cells(i, 9)
            If Cells(i, 3) = "未作業時間(帰宅、休日)" Then Cells(18, 12) = Cells(18, 12) + Cells(i, 9)
    Next
    For i = 13 To 18
        total_working_days = Int(Cells(i, 12))
        total_working_times = TimeValue(CDate(Cells(i, 12)))
        Cells(i, 12) = total_working_days & "日" & Space(2) & _
                       Application.WorksheetFunction.RoundDown _
                       ((Hour(total_working_times) * 60 + Minute(total_working_times)) / 60, 2)
    Next
    
End Sub
Sub 発光解析()
    ans = MsgBox("発光解析がんばってください！", vbYesNo)
    If ans = vbYes Then Range("A1").Interior.ThemeColor = xlThemeColorAccent1
End Sub
Sub 場所特定()
    ans = MsgBox("発光解析お疲れ様です！場所特定へ進めますか？", vbYesNo)
    If ans = vbYes Then Range("A1").Interior.ThemeColor = xlThemeColorAccent2
End Sub
Sub 整合性確認()
    ans = MsgBox("場所特定お疲れ様です！整合性確認へ進めますか？", vbYesNo)
    If ans = vbYes Then Range("A1").Interior.ThemeColor = xlThemeColorAccent3
End Sub
Sub PFA指示書()
    ans = MsgBox("ばっちしです！PFA指示書へ進めますか？", vbYesNo)
    If ans = vbYes Then Range("A1").Interior.ThemeColor = xlThemeColorAccent4
End Sub
Sub test()
End Sub
