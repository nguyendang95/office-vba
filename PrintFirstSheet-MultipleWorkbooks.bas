Option Explicit

Public Sub PrintFirstWorksheet()
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim objFolderPicker As FileDialog
    Dim objWb As Excel.Workbook
    Dim objSh As Excel.Worksheet
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    With objFolderPicker
        .Title = "Chon thu muc chua cac tap tin Excel can in sheet dau tien"
        .Show
        If .SelectedItems.Count > 0 Then
            Set objFolder = objFSO.GetFolder(.SelectedItems(1))
        Else: Exit Sub
        End If
    End With
    For Each objFile In objFolder.Files
        Select Case GetFileExtension(objFile.Path)
            Case ".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltm", ".xltx"
            Set objWb = Application.Workbooks.Open(FileName:=objFile.Path, ReadOnly:=True)
            Set objSh = objWb.Worksheets(1)
            objSh.PrintOut , , 1
            objWb.Close SaveChanges:=False
        End Select
    Next
    Set objFSO = Nothing
    Set objFolder = Nothing
    Set objFile = Nothing
    Set objFolderPicker = Nothing
    Set objWb = Nothing
    Set objSh = Nothing
End Sub

Private Function GetFileExtension(FileName As String) As String
    On Error Resume Next
    GetFileExtension = Mid(FileName, InStrRev(FileName, "."))
    If Err.number = 5 Then
        GetFileExtension = vbNullString
    End If
End Function
