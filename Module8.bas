Attribute VB_Name = "Module1"
Option Explicit
Public L As Double, T As Double, W As Double, H As Double, i As Long, j As Long, d As Date
Public started_day As Date, finished_day As Date, flag As Byte, flag2 As Byte, bar As Shape
Public started_at As Date, finished_at As Date, ans As Integer, MaxRow As Long
Sub TAT()

    Dim seigou_started_at As Date
    Dim seigou_finished_at As Date
    
    Worksheets("�ꏊ����_TAT�Ǘ��䒠").Activate
    Columns(5).Validation.Delete
    For i = 3 To Cells(3, 7).End(xlDown).Row
        If Cells(i, 7) = "" And Cells(i, 8) = "" Then
            Cells(i, 7) = Format(Now)
            With Worksheets("�K���g�`���[�g")
                For j = 1 To 88 Step 3
                    .Cells(j, 1) = ""
                Next
            End With
            Exit For
        End If
        If Cells(i, 7) <> "" And Cells(i, 8) = "" Then
            If Cells(i, 2) <> "LE" Then
                Cells(i, 1) = "�I�ꏊ����"
                Cells(i, 2) = "(�i��)�ꏊ����"
                Cells(i, 3) = "��Ǝ���"
                Cells(i, 5) = "���͕s�v"
                Cells(i, 8) = Format(Now)
                Cells(i + 1, 7) = Format(Now)
                
                seigou_started_at = CDate(Year(Cells(i, 8)) & "/" & Month(Cells(i, 8)) & "/" & Day(Cells(i, 8)) & Space(2) & "8:30:00")
                seigou_finished_at = CDate(Year(Cells(i, 8)) & "/" & Month(Cells(i, 8)) & "/" & Day(Cells(i, 8)) & Space(2) & "17:15:00")
                
                Select Case Cells(1, 1).Interior.ColorIndex
                    Case 41
                        flag = 48
                        Cells(i, 4) = "�������"
                        Cells(i, 1) = "�H�������"
                    Case 46
                        flag = 53
                        Cells(i, 4) = "�ꏊ����"
                    Case 44
                        flag = 51
                        Cells(i, 4) = "�������m�F"
                        Cells(i, 3) = "�ҋ@����"
                        Cells(i, 8) = seigou_started_at
                        Cells(i + 1, 7) = seigou_started_at
                    Case 15
                        flag = 22
                        Cells(i, 4) = "PFA�w����"
                End Select
    
                If Cells(i - 1, 3) = "��Ǝ���" Then
                    If Cells(i, 4) <> "�������m�F" Then
                        flag = 62
                    End If
                    Cells(i, 3) = "�ҋ@����"
                End If
                If Day(Cells(i, 7)) <> Day(Cells(i, 8)) Then
                    flag = 50
                    Cells(i, 3) = "����Ǝ���(�A��A�x��)"
                End If
                If Cells(i, 7) = seigou_started_at Then
                    Cells(i, 8) = seigou_finished_at
                    Cells(i + 1, 7) = seigou_finished_at
                End If
            
                If Cells(i, 3) = "�ҋ@����" And (Cells(i, 4) = "�������" Or Cells(i, 4) = "�ꏊ����") Then
                    Cells(i, 5).Validation.Add Type:=xlValidateList, Formula1:="���͕s�v,�w���҂�,���j����,���\�[�X�s��(���u)"
                End If
            Else
                Cells(i, 8) = Format(Now)
                Cells(i + 1, 7) = Format(Now)
            End If
            started_day = CDate(Year(Cells(i, 7)) & "/" & Month(Cells(i, 7)) & "/" & Day(Cells(i, 7)))
            finished_day = CDate(Year(Cells(i, 8)) & "/" & Month(Cells(i, 8)) & "/" & Day(Cells(i, 8)))
            started_at = CDate(Hour(Cells(i, 7)) & ":" & Minute(Cells(i, 7)) & ":00")
            finished_at = CDate(Hour(Cells(i, 8)) & ":" & Minute(Cells(i, 8)) & ":00")
            
            Worksheets("�K���g�`���[�g").Activate
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
            Worksheets("�ꏊ����_TAT�Ǘ��䒠").Activate
            Exit For
        End If
    Next
    
End Sub
Sub ��ƕύX()
    ans = MsgBox("�ҋ@���ԂɕύX���܂����H", vbYesNo)
    If ans = vbYes Then
        Worksheets("�ꏊ����_TAT�Ǘ��䒠").Activate
        Columns(5).Validation.Delete
        MaxRow = Cells(3, 3).End(xlDown).Row
        For i = 3 To MaxRow
            If Cells(i, 3) = "��Ǝ���" And Cells(i + 1, 3) = "" Then
                Cells(i, 3) = "�ҋ@����"
                Select Case Cells(i, 4)
                    Case "�������", "�ꏊ����"
                        Cells(i, 5).Validation.Add Type:=xlValidateList, Formula1:="���͕s�v,�w���҂�,���j����,���\�[�X�s��(���u)"
                End Select
                Exit For
            End If
        Next
        With Worksheets("�K���g�`���[�g")
            .Shapes(.Shapes.Count).Fill.ForeColor.SchemeColor = 62
        End With
    End If
End Sub
Sub TAT�W�v()

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
        If Cells(i, 3) = "��Ǝ���" Then
            Select Case Cells(i, 4)
                Case "�������"
                   Cells(13, 12) = Cells(13, 12) + Cells(i, 9)
                Case "�ꏊ����"
                    Cells(14, 12) = Cells(14, 12) + Cells(i, 9)
                Case "PFA�w����"
                    Cells(16, 12) = Cells(16, 12) + Cells(i, 9)
            End Select
        End If
            If Cells(i, 4) = "�������m�F" Then Cells(15, 12) = Cells(15, 12) + Cells(i, 9)
            If Cells(i, 3) = "�ҋ@����" Then Cells(17, 12) = Cells(17, 12) + Cells(i, 9)
            If Cells(i, 3) = "����Ǝ���(�A��A�x��)" Then Cells(18, 12) = Cells(18, 12) + Cells(i, 9)
    Next
    For i = 13 To 18
        total_working_days = Int(Cells(i, 12))
        total_working_times = TimeValue(CDate(Cells(i, 12)))
        Cells(i, 12) = total_working_days & "��" & Space(2) & _
                       Application.WorksheetFunction.RoundDown _
                       ((Hour(total_working_times) * 60 + Minute(total_working_times)) / 60, 2)
    Next
    
End Sub
Sub �������()
    ans = MsgBox("������͂���΂��Ă��������I", vbYesNo)
    If ans = vbYes Then Range("A1").Interior.ThemeColor = xlThemeColorAccent1
End Sub
Sub �ꏊ����()
    ans = MsgBox("������͂����l�ł��I�ꏊ����֐i�߂܂����H", vbYesNo)
    If ans = vbYes Then Range("A1").Interior.ThemeColor = xlThemeColorAccent2
End Sub
Sub �������m�F()
    ans = MsgBox("�ꏊ���肨���l�ł��I�������m�F�֐i�߂܂����H", vbYesNo)
    If ans = vbYes Then Range("A1").Interior.ThemeColor = xlThemeColorAccent3
End Sub
Sub PFA�w����()
    ans = MsgBox("�΂������ł��IPFA�w�����֐i�߂܂����H", vbYesNo)
    If ans = vbYes Then Range("A1").Interior.ThemeColor = xlThemeColorAccent4
End Sub
Sub test()
End Sub
