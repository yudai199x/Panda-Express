Attribute VB_Name = "ModuleMain"
Option Explicit

Private Const TOOL_NAME As String = "��͈˗��V�X�e��_TOOL.xlsm"
Private Const MANUAL_NAME As String = "�}�j���A��_��͈˗��V�X�e��.xlsm"

' TSB�p���j���[
Public Sub TSB()
    ToolOpen 0
End Sub

' TNA�p���j���[
Public Sub TNA()
    ToolOpen 1
End Sub

' �Ǘ��җp���j���[
Public Sub Admin()
    ToolOpen 2
End Sub

' �e�[�}�����F���
Public Sub ThemeApprove()
    ToolOpen 3
End Sub

' �ے����F���
Public Sub SectionApprove()
    ToolOpen 4
End Sub

' AIM���F���
Public Sub AimApprove()
    ToolOpen 5
End Sub

' TNA�㒷�F�؉��
Public Sub TnaApprove()
    ToolOpen 6
End Sub

' ���B�����m�F���
Public Sub OrderConfirm()
    ToolOpen 7
End Sub

' ���[�U�ҏW(TSB)���
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
