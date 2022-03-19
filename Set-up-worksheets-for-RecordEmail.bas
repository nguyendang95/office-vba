Option Explicit

Private Sub PrepareRecordEmails()
    Dim objSh As Excel.Worksheet
    Dim objPivotSh As Excel.Worksheet
    Dim objTable As Excel.ListObject
    Dim objPivotTable As Excel.PivotTable
    Dim objPivotCache As Excel.PivotCache
    Dim arrColumn()
    Set objSh = Application.ActiveSheet
    objSh.Name = "EmailDetails"
    arrColumn = Array("Subject", "From", "Email address", "Date and time received")
    objSh.Range("A1:D1").Value = arrColumn
    Set objTable = objSh.ListObjects.Add(xlSrcRange, objSh.Range("A1:D1"), , xlYes)
    Set objPivotSh = Application.ActiveWorkbook.Worksheets.Add(Before:=objSh)
    objPivotSh.Name = "EmailPivotTable"
    Set objPivotCache = Application.ActiveWorkbook.PivotCaches.Create(xlDatabase, objTable)
    Set objPivotTable = objPivotCache.CreatePivotTable(objPivotSh.Range("A1:D20"), objTable.Name)
    With objPivotTable.PivotFields("From")
        .Orientation = xlRowField
        .Position = 1
    End With
    With objPivotTable.PivotFields("Email address")
        .Orientation = xlRowField
        .Position = 2
    End With
    With objPivotTable.PivotFields("Date and time received")
        .Orientation = xlRowField
        .Position = 3
    End With
    With objPivotTable.PivotFields("Subject")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlFunction
    End With
    Set objSh = Nothing
    Set objPivotSh = Nothing
    Set objTable = Nothing
    Set objPivotTable = Nothing
    Set objPivotCache = Nothing
End Sub