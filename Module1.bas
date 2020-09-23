Attribute VB_Name = "Module1"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function ShellExec(FileName$, Parameters$, Directory$, Optional MsgBoxFlag As Boolean) As Boolean
    Dim l#
    l = ShellExecute(Form1.hwnd, "open", FileName, Parameters, Directory, vbNormalFocus)
    If l <= 32 Then
        If IsMissing(MsgBoxFlag) Then
            MsgBoxFlag = True
        End If
        If MsgBoxFlag Then
            MsgBox "Can't execute: " + FileName + " " + Parameters
        End If
        Exit Function
    End If
    ShellExec = True
End Function

