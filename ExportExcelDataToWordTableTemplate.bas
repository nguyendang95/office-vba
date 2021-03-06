Option Explicit

Public Sub ImportExcelDataToWordTable()
    Dim objWrdApp As Object
    Dim objWrdDoc1 As Object
    Dim objWrdDoc2 As Object
    Dim objWrdTable1 As Object
    Dim objWrdTable2 As Object
    Dim lngRow As Long, lngColumn As Long, lngLastRow As Long
    Dim arrRangeData()
    Dim objShell As Object
    Dim objFSO As Object
    Const wdAlertsNone As Byte = 0
    Const wdAlertsAll As Integer = -1
    Const wdFormatDocumentDefault As Integer = 16
    Const wdAlignParagraphLeft As Byte = 0
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Not objFSO.FolderExists(ActiveWorkbook.Path & "\Ketqua\") Then objFSO.CreateFolder ActiveWorkbook.Path & "\Ketqua"
    lngLastRow = Cells(Rows.Count, 1).End(xlUp).Row
    arrRangeData() = Cells(lngLastRow, 1).CurrentRegion.Value
    Set objWrdApp = CreateObject("Word.Application")
    objWrdApp.DisplayAlerts = wdAlertsNone
    Set objWrdDoc1 = objWrdApp.Documents.Open(Filename:=ActiveWorkbook.Path & "\TienAn.docx")
    If objWrdDoc1.Tables.Count > 0 Then
        Set objWrdTable1 = objWrdDoc1.Tables(1)
        For lngRow = 1 To lngLastRow - 3
            objWrdTable1.Cell(lngRow + 1, 1).Range.InsertAfter arrRangeData(lngRow + 1, 1)
            objWrdTable1.Cell(lngRow + 1, 2).Range.InsertAfter arrRangeData(lngRow + 1, 2)
            objWrdTable1.Cell(lngRow + 1, 3).Range.InsertAfter arrRangeData(lngRow + 1, 3)
            objWrdTable1.Cell(lngRow + 1, 4).Range.InsertAfter arrRangeData(lngRow + 1, 6)
            If Not lngRow = lngLastRow - 3 Then objWrdTable1.Rows.Add
        Next
    End If
    objWrdTable1.Columns.AutoFit
    objWrdTable1.Range.Paragraphs.Alignment = wdAlignParagraphLeft
    objWrdDoc1.SaveAs2 Filename:=ActiveWorkbook.Path & "\Ketqua\TienAn.docx", FileFormat:=wdFormatDocumentDefault
    objWrdDoc1.Close
    Set objWrdDoc2 = objWrdApp.Documents.Open(Filename:=ActiveWorkbook.Path & "\ThanhTien.docx")
    If objWrdDoc2.Tables.Count > 0 Then
        Set objWrdTable2 = objWrdDoc2.Tables(1)
        For lngRow = 1 To lngLastRow - 3
            objWrdTable2.Cell(lngRow + 1, 1).Range.InsertAfter arrRangeData(lngRow + 1, 1)
            objWrdTable2.Cell(lngRow + 1, 2).Range.InsertAfter arrRangeData(lngRow + 1, 2)
            objWrdTable2.Cell(lngRow + 1, 3).Range.InsertAfter arrRangeData(lngRow + 1, 3)
            objWrdTable2.Cell(lngRow + 1, 4).Range.InsertAfter arrRangeData(lngRow + 1, 11)
            If Not lngRow = lngLastRow - 3 Then objWrdTable2.Rows.Add
        Next
    End If
    objWrdTable2.Columns.AutoFit
    objWrdTable2.Range.Paragraphs.Alignment = wdAlignParagraphLeft
    objWrdDoc2.SaveAs2 Filename:=ActiveWorkbook.Path & "\Ketqua\ThanhTien.docx", FileFormat:=wdFormatDocumentDefault
    objWrdDoc2.Close
    objWrdApp.DisplayAlerts = wdAlertsAll
    objWrdApp.Quit
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "explorer """ & ActiveWorkbook.Path & "\Ketqua\" & "", vbNormalFocus
    Set objWrdApp = Nothing
    Set objWrdDoc1 = Nothing
    Set objWrdDoc2 = Nothing
    Set objWrdTable1 = Nothing
    Set objWrdTable2 = Nothing
    Set objShell = Nothing
    Set objFSO = Nothing
End Sub
