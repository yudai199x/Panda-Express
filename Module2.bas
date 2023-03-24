Attribute VB_Name = "Module1"
Public y1 As Integer
Public m1 As Integer
Public d1 As Integer
Public h1 As Integer
Public t1 As Integer
Public y2 As Integer
Public m2 As Integer
Public d2 As Integer
Public h2 As Integer
Public t2 As Integer
Public swich As Integer
Public datSTRD As Date
Public datENDD As Date
Public userNum As String
Public UsagItm As String
Public No As Integer



Sub 予約input()
 
Dim 予約テキスト As Object

Dim strd As Date
Dim endd As Date
Dim strt As Date
Dim endt As Date

 Call 初期化
 maxrow1 = Worksheets("予約表").Cells(Rows.Count, 2).End(xlUp).Row
 maxrow2 = Worksheets("予約表元データ").Cells(Rows.Count, 1).End(xlUp).Row
' For i = 4 To maxrow1
'     If CDate(Year(Cells(i, 2)) + Month(Cells(i, 2)) + Day(Cells(i, 2))) = CDate(Year(Now) + Month(Now) + Day(Now)) Then Cells(i - 3, 1).Show

' Next
  For k = 3 To maxrow2
    
    With Worksheets("予約表元データ")
         No = .Cells(k, 1).Value
         datSTRD = CDate(.Cells(k, 3).Value + .Cells(k, 4).Value)
         datENDD = CDate(.Cells(k, 5).Value + .Cells(k, 6).Value)
         userNum = .Cells(k, 2).Value
         UsagItm = .Cells(k, 7).Value
    End With
    
    strd = CDate(Year(datSTRD) & "/" & Month(datSTRD) & "/" & Day(datSTRD))
    endd = CDate(Year(datENDD) & "/" & Month(datENDD) & "/" & Day(datENDD))
    strt = CDate(Hour(datSTRD) & ":" & Minute(datSTRD) & ":00")
    endt = CDate(Hour(datENDD) & ":" & Minute(datENDD) & ":00")
    dayw = endd - strd
    
    flag = 0
    flag2 = 0
    For i = 7 To maxrow1
    
      If Cells(i, 2).Value = strd Then
         For j = 0 To dayw
              If flag = 1 Then
                  L = Cells(i + j, 4).Left
                  flag = 0
              Else
                  L = Cells(i + j, Hour(datSTRD) + 4).Left + Cells(i, 4).Width * (Minute(datSTRD) / 60)
              End If
              
              T = Cells(i + j, Hour(datSTRD) + 4).Top
              
              If (dayw - j) <> 0 Then
                 If flag2 = 1 Then
                    W = Cells(i, 4).Width * 24
                 Else
                    W = Cells(i, 4).Width * (24 - (Hour(datSTRD) + Minute(datSTRD) / 60))
                 End If
                 flag = 1
                 flag2 = 1
              Else
                 If flag2 = 1 Then
                    W = Cells(i, 4).Width * endt * 24
                 Else
                    W = Cells(i, 4).Width * (endt - strt) * 24
                 End If
              End If
              
              h = Cells(i, 4).Height / 2
   
         
              Set 予約テキスト = ActiveSheet.Shapes.AddShape( _
                                 msoTextOrientationHorizontal, _
                                 L, T, W, h)
              予約テキスト.TextFrame.Characters.Text = No & ":" & userNum & ":" & UsagItm
              予約テキスト.TextFrame.Characters.Font.Size = 11
              予約テキスト.TextFrame.Characters.Font.Bold = True
              予約テキスト.OnAction = "予約編集"
   
         Next j
      End If
    Next i
   Set 予約テキスト = Nothing
Next k

End Sub


Sub 新規予約()





    y1 = Year(Now) - 2021
    m1 = Month(Now) - 1
    d1 = Day(Now) - 1
    h1 = 0
    t1 = 0
    
    y2 = Year(Now) - 2021
    m2 = Month(Now) - 1
    d2 = Day(Now) - 1
    h2 = 23
    t2 = 3
    
    userNum = ""
    UsagItm = ""
    No = 0


    swich = 0

Do
    menu.Show
    If swich <> 0 Then Exit Do

Loop


 With Worksheets("予約表元データ")
 maxrow = .Cells(Rows.Count, 1).End(xlUp).Row

flag = 0

 For k = 3 To maxrow
       If CDate(.Cells(k, 3).Value + .Cells(k, 4).Value) < datSTRD And _
          CDate(.Cells(k, 5).Value + .Cells(k, 6).Value) > datSTRD Then
          flag = 1
       End If
       If CDate(.Cells(k, 3).Value + .Cells(k, 4).Value) < datENDD And _
          CDate(.Cells(k, 5).Value + .Cells(k, 6).Value) > datENDD Then
          flag = 1
       End If
 Next k

 
 If flag = 0 Then
    .Cells(maxrow + 1, 1).Value = .Cells(maxrow, 1).Value + 1
    .Cells(maxrow + 1, 2).Value = userNum
    .Cells(maxrow + 1, 3).Value = CDate(Year(datSTRD) & "/" & Month(datSTRD) & "/" & Day(datSTRD))
    .Cells(maxrow + 1, 4).Value = CDate(Hour(datSTRD) & ":" & Minute(datSTRD) & ":00")
    .Cells(maxrow + 1, 5).Value = CDate(Year(datENDD) & "/" & Month(datENDD) & "/" & Day(datENDD))
    .Cells(maxrow + 1, 6).Value = CDate(Hour(datENDD) & ":" & Minute(datENDD) & ":00")
    .Cells(maxrow + 1, 7).Value = UsagItm
 Else
    MsgBox ("予約が重複しています！！")
 End If
 End With

If swich = 1 Then Call 予約input
 
End Sub

Sub 予約編集()

 With Worksheets("予約表元データ")
 maxrow = .Cells(Rows.Count, 1).End(xlUp).Row
 
 No = Left(ActiveSheet.Shapes(Application.Caller).TextFrame.Characters.Text, InStr(ActiveSheet.Shapes(Application.Caller).TextFrame.Characters.Text, ":") - 1)
 y1 = Year(Now) - 2021
 m1 = Month(Now) - 1
 d1 = Day(Now) - 1
 h1 = 0
 t1 = 0
    
 y2 = Year(Now) - 2021
 m2 = Month(Now) - 1
 d2 = Day(Now) - 1
 h2 = 23
 t2 = 3
    
 userNum = ""
 UsagItm = ""
 
 For k = 3 To maxrow
    If .Cells(k, 1).Value = No Then
       y1 = Year(.Cells(k, 3).Value) - 2021
       m1 = Month(.Cells(k, 3).Value) - 1
       d1 = Day(.Cells(k, 3).Value) - 1
       h1 = Hour(.Cells(k, 4).Value)
       t1 = Minute(.Cells(k, 4).Value) / 15
    
       y2 = Year(.Cells(k, 5).Value) - 2021
       m2 = Month(.Cells(k, 5).Value) - 1
       d2 = Day(.Cells(k, 5).Value) - 1
       h2 = Hour(.Cells(k, 6).Value)
       t2 = Minute(.Cells(k, 6).Value) / 15
    
       userNum = .Cells(k, 2).Value
       UsagItm = .Cells(k, 7).Value
    End If
 Next k
    

'    BOTROW = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Row

    swich = 0

Do
    menu.Show
    If swich <> 0 Then Exit Do

Loop



 maxrow = .Cells(Rows.Count, 1).End(xlUp).Row
 flag = 0
 Nrow = 1
 For k = 3 To maxrow
    If .Cells(k, 1).Value = No Then
       Nrow = k
    Else
    
       If CDate(.Cells(k, 3).Value + .Cells(k, 4).Value) < datSTRD And _
          CDate(.Cells(k, 5).Value + .Cells(k, 6).Value) > datSTRD Then
          flag = 1
       End If
       If CDate(.Cells(k, 3).Value + .Cells(k, 4).Value) < datENDD And _
          CDate(.Cells(k, 5).Value + .Cells(k, 6).Value) > datENDD Then
          flag = 1
       End If
    End If
 Next k

 
 If flag = 0 Then
    .Cells(Nrow, 1).Value = No
    .Cells(Nrow, 2).Value = userNum
    .Cells(Nrow, 3).Value = CDate(Year(datSTRD) & "/" & Month(datSTRD) & "/" & Day(datSTRD))
    .Cells(Nrow, 4).Value = CDate(Hour(datSTRD) & ":" & Minute(datSTRD) & ":00")
    .Cells(Nrow, 5).Value = CDate(Year(datENDD) & "/" & Month(datENDD) & "/" & Day(datENDD))
    .Cells(Nrow, 6).Value = CDate(Hour(datENDD) & ":" & Minute(datENDD) & ":00")
    .Cells(maxrow, 7).Value = UsagItm
 Else
    MsgBox ("予約が重複しています！！")
 End If
 End With
 
If swich = 1 Then Call 予約input
 
End Sub


Sub 初期化()
    Dim sp As Shape
    Worksheets("予約表").Activate
    For Each sp In ActiveSheet.Shapes
        If Not sp.TextFrame.Characters.Text = "新規予約" Then
           If Not sp.TextFrame.Characters.Text = "表示更新" Then sp.Delete
        End If
    Next


End Sub
