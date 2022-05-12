Option Explicit

Private Sub PDF_EachPage()
    Dim objFolderPicker As Office.FileDialog
    Dim objDoc As Word.Document
    Dim lngNumPage As Long, lngPage As Long
    Dim strFolder As String
    Application.ScreenUpdating = False
    Set objDoc = ActiveDocument
    lngNumPage = objDoc.BuiltInDocumentProperties("Number of Pages")
    If objDoc.Saved Then
        Set objFolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
        With objFolderPicker
            .Title = "Select folder to export PDF files"
            .Show
            If .SelectedItems.Count = 1 Then
                strFolder = .SelectedItems(1)
            Else: Exit Sub
            End If
        End With
    End If
    Selection.HomeKey wdStory, wdMove
    For lngPage = 1 To lngNumPage
        If Selection.Information(wdActiveEndPageNumber) = 1 Then
            objDoc.Bookmarks("\Page").Range.ExportAsFixedFormat OutputFileName:=strFolder & "\" & objDoc.Name & "-Page" & lngPage & ".pdf", ExportFormat:=wdExportFormatPDF
            Selection.GoTo wdGoToPage, wdGoToNext, 1
        Else
            Selection.GoTo wdGoToPage, wdGoToNext, 1
            objDoc.Bookmarks("\Page").Range.ExportAsFixedFormat OutputFileName:=strFolder & "\" & objDoc.Name & "-Page" & lngPage & ".pdf", ExportFormat:=wdExportFormatPDF
        End If
    Next
    Application.ScreenUpdating = True
    Set objFolderPicker = Nothing
    Set objDoc = Nothing
End Sub