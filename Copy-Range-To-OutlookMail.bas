Option Explicit

Public Sub CopyRangeToOutlookMail()
    Dim objSel As Excel.Range
    Dim objOlApp As Object
    Dim objOlMail As Object
    Dim objWrdDoc As Object
    Dim objWrdRange As Object
    On Error Resume Next
    Set objSel = Application.InputBox("Chon vung du lieu can dua vao thu moi", "Chon vung du lieu", , , , , , 8)
    If Not objSel Is Nothing Then
        objSel.Copy
        On Error Resume Next
        Set objOlApp = GetObject(, "Outlook.Application")
        If Err.Number = 429 Then
            MsgBox "Ban can mo ung dung Outlook truoc khi chay macro nay. Vui long thu lai sau!", vbExclamation, "Loi: Chua khoi dong ung dung Outlook"
            Exit Sub
        End If
        Set objOlMail = objOlApp.CreateItem(0)
        With objOlMail
            .Display
            .Body = "Excel Range:"
            Set objWrdDoc = .GetInspector.WordEditor
            Set objWrdRange = objWrdDoc.Content
            With objWrdRange
                .Collapse (0)
                .Paragraphs.Add
                .InsertBreak
                .PasteSpecial DataType:=3
            End With
        End With
    Application.CutCopyMode = False
    End If
End Sub