Attribute VB_Name = "ModuleMain"
Option Explicit

Private Const TOOL_NAME As String = "解析依頼システム_TOOL.xlsm"
Private Const MANUAL_NAME As String = "マニュアル_解析依頼システム.xlsm"

' TSB用メニュー
Public Sub TSB()
    ToolOpen 0
End Sub

' TNA用メニュー
Public Sub TNA()
    ToolOpen 1
End Sub

' 管理者用メニュー
Public Sub Admin()
    ToolOpen 2
End Sub

' テーマ長承認画面
Public Sub ThemeApprove()
    ToolOpen 3
End Sub

' 課長承認画面
Public Sub SectionApprove()
    ToolOpen 4
End Sub

' AIM承認画面
Public Sub AimApprove()
    ToolOpen 5
End Sub

' TNA上長認証画面
Public Sub TnaApprove()
    ToolOpen 6
End Sub

' 調達発注確認画面
Public Sub OrderConfirm()
    ToolOpen 7
End Sub

' ユーザ編集(TSB)画面
Public Sub UserEditTSB()
    ToolOpen 8
End Sub

Private Sub ToolOpen(userMode As Variant)
    Dim wb   As Workbook
    
    Dim filePath As String
    
    filePath = ThisWorkbook.Path & "\" & TOOL_NAME
    
    Set wb = Workbooks.Open(filePath)
    Application.Run Dir(filePath) & "!mdlAutoProcess.Auto_Open", userMode
End Sub

Public Sub ManualOpen()
    Dim wb   As Workbook
    
    Dim filePath As String
    
    filePath = ThisWorkbook.Path & "\" & MANUAL_NAME
    
    Set wb = Workbooks.Open(filePath)
End Sub
