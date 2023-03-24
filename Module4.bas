Attribute VB_Name = "Module1"
Sub Diag集計()
    ' ファイル操作
    Dim szPath, hzPath, Ptflag As String
    Dim objFileSystem As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim maxr As Long
    Dim bookN As String
    Dim buf As String
    Dim acv As String
    Dim tDays As Date
    Dim flag As Long
    flag = 0
    
    szPath = "C:\Program Files\STK Technology\Personal Tester\Diag\LogFolder"
    hzPath = "\\10.23.7.43\hin-kaiseki\10_解析結果格納用\07_LE\06_装置日常点検\PT_DIAG履歴\"
    
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    Set objFolder = objFileSystem.GetFolder(szPath)
    
    If err.Number <> 0 Then
      On Error GoTo 0
      szPath = "C:\Program Files (x86)\STK Technology\Personal Tester\Diag\LogFolder"
      Set objFolder = objFileSystem.GetFolder(szPath)
    End If
    On Error GoTo 0
    
    i = 0
    For Each objFile In objFolder.Files
        If Right(objFile.Name, 3) = "log" Then
'            Ptflag = "Other\"
            maxr = Cells(Rows.Count, 1).End(xlUp).Row
            
            Open objFile.Path For Input As #1
               Do Until EOF(1)
                  Line Input #1, buf
                  ' 日付の確認
                  If InStr(buf, "試験開始日時") > 0 Or InStr(buf, "DIAG Date") > 0 Then tDays = Mid(buf, InStr(buf, ": ") + 2, 10)

                      
                   ' 号機の確認
                       
                   If InStr(buf, "PT 装置名") > 0 Or InStr(buf, "PT Serial Number") > 0 Then
'                        acv = Mid(buf, InStr(buf, ": ") + 6, 7)
                        acv = Mid(buf, InStr(buf, "hin"), 7)
                        ' hin-001の場合の処理(シート関連だけ)
                        If acv = "hin-001" Then
                           bookN = "(車載)PT_Diag履歴.xlsm"
                           acv = "Diag履歴"
                           Ptflag = "hin-001\"
                            With Workbooks(bookN).Worksheets(acv)
                               If IsDate(.Cells(maxr, 2).Value) Then
                                 maxr = Cells(Rows.Count, 1).End(xlUp).Row
                                 If Month(tDays) <> Month(.Cells(maxr, 2).Value) Then
                                    .Copy after:=Worksheets(Worksheets.Count)
                                    ActiveSheet.Name = Year(.Cells(maxr, 2).Value) & "年" & Month(.Cells(maxr, 2).Value) & "月"
                                    .Activate
                                    .Range("A4:R" & maxr).ClearContents
                                  End If
                               End If
                            End With
                        ' hin-001以外の処理
                        Else
                           If acv = "PT-#033" Then acv = "hin-002"
                           If acv = "PT-#037" Then acv = "hin-003"
                           If acv = "PT-#050" Then acv = "hin-004"
                           If acv = "PT-#057" Then acv = "hin-005"
                           If acv = "AIM-#008" Then acv = "hin-006"
                           If acv = "AIM-#011" Then acv = "hin-007"
                           If acv = "AIM-#013" Then acv = "hin-008"
                           
                           If acv = "hin-002" Then Ptflag = "hin-002\"
                           If acv = "hin-003" Then Ptflag = "hin-003\"
                           If acv = "hin-004" Then Ptflag = "hin-004\"
                           If acv = "hin-005" Then Ptflag = "hin-005\"
                           If acv = "hin-006" Then Ptflag = "hin-006\"
                           If acv = "hin-007" Then Ptflag = "hin-007\"
                           If acv = "hin-008" Then Ptflag = "hin-008\"
                           If acv = "hin-009" Then Ptflag = "hin-009\"
                           If acv = "hin-010" Then Ptflag = "hin-010\"
                           If acv = "hin-011" Then Ptflag = "hin-011\"
                           If acv = "hin-012" Then Ptflag = "hin-012\"
                           If acv = "hin-013" Then Ptflag = "hin-013\"
                           If acv = "hin-014" Then Ptflag = "hin-014\"
                           If acv = "hin-015" Then Ptflag = "hin-015\"
                           If acv = "hin-016" Then Ptflag = "hin-016\"
                           If acv = "hin-017" Then Ptflag = "hin-017\"
                           If acv = "hin-018" Then Ptflag = "hin-018\"
                           If acv = "hin-019" Then Ptflag = "hin-019\"
                           If acv = "hin-020" Then Ptflag = "hin-020\"
                           
                           If InStr(acv, "hin-") = 0 Then
                              Ptflag = "Other\"
                              acv = "Other"
                           End If

                           bookN = "(車載以外)PT_Diag履歴.xlsm"
                            If InStr(acv, "hin-001") = 0 And flag = 0 Then flag = 1
                            If flag = 1 Then
                               Workbooks.Open Filename:="\\10.23.7.43\hin-kaiseki\10_解析結果格納用\07_LE\06_装置日常点検\PT_DIAG履歴\結果一覧表(Excel)\車載以外\(車載以外)PT_Diag履歴.xlsm"
                               flag = 2
                            End If

                               
                        End If
                        With Workbooks(bookN).Worksheets(acv)
                           maxr = .Cells(Rows.Count, 1).End(xlUp).Row
                           .Cells(maxr + 1, 2).Value = tDays
                           .Cells(maxr + 1, 1).Value = objFile.Name
                        End With
                  End If
                  
                  If acv <> "" Then
                     Call 結果表示(bookN, acv, objFile.Name, maxr, buf)
                  End If


               Loop
            Close #1
            Name objFile.Path As hzPath & Ptflag & objFile.Name
        End If
    Next

End Sub

Sub 結果表示(tBok As String, tShet As String, objFnum As String, maxrow As Long, buf As String)

                      With Workbooks(tBok).Worksheets(tShet)
'                      .Cells(maxrow + 1, 1).Value = objFnum
                      If InStr(buf, "PT 装置名") > 0 Or InStr(buf, "PT Serial Number") > 0 Then .Cells(maxrow + 1, 3).Value = Mid(buf, InStr(buf, ": ") + 2, Len(buf) - InStr(buf, ": ") + 2)
                      If InStr(buf, "DllOutCheck is PASS") > 0 Then .Cells(maxrow + 1, 4).Value = "PASS"
                      If InStr(buf, "ContactCheck is PASS") > 0 Then .Cells(maxrow + 1, 5).Value = "PASS"
                      If InStr(buf, "BAM_Logic_DIAG is PASS") > 0 Then .Cells(maxrow + 1, 6).Value = "PASS"
                      If InStr(buf, "Yamame_Logic_DIAG is PASS") > 0 Then .Cells(maxrow + 1, 7).Value = "PASS"
                      If InStr(buf, "ADCheck is PASS") > 0 Then .Cells(maxrow + 1, 8).Value = "PASS"
                      If InStr(buf, "CompDrvSkewCheck is PASS") > 0 Then .Cells(maxrow + 1, 9).Value = "PASS"
                      If InStr(buf, "ISOutAmpCheck is PASS") > 0 Then .Cells(maxrow + 1, 10).Value = "PASS"
                      If InStr(buf, "CHDCOutVoltCheck is PASS") > 0 Then .Cells(maxrow + 1, 11).Value = "PASS"
                      If InStr(buf, "REGOutVoltCheck is PASS") > 0 Then .Cells(maxrow + 1, 12).Value = "PASS"
                      If InStr(buf, "CHDCOutAmpCheck is PASS") > 0 Then .Cells(maxrow + 1, 13).Value = "PASS"
                      If InStr(buf, "REGOutAmpCheck is PASS") > 0 Then .Cells(maxrow + 1, 14).Value = "PASS"
                      If InStr(buf, "AnalogMonVoltCheck is PASS") > 0 Then .Cells(maxrow + 1, 15).Value = "PASS"
                      If InStr(buf, "DigitalMonCheck is PASS") > 0 Then .Cells(maxrow + 1, 16).Value = "PASS"
                      If InStr(buf, "RefFreqCheck is PASS") > 0 Then .Cells(maxrow + 1, 17).Value = "PASS"
                      If InStr(buf, "ILeakCheck is PASS") > 0 Then .Cells(maxrow + 1, 18).Value = "PASS"
                     End With

End Sub
