Option Explicit

Private Sub ExportQueryResultToWordTable()
    Dim objTable As Word.Table
    Dim arrResult() As Variant
    Dim i As Long, j As Long
    On Error Resume Next
    arrResult = RunSQL("SELECT * FROM SanPham;")
    If Err.Number <> 0 Then Exit Sub
    If Not IsEmpty(arrResult) Then
        If Not Selection.Information(wdWithInTable) Then
            Set objTable = Application.ActiveDocument.Content.Tables.Add(Selection.Range, UBound(arrResult) + 1, UBound(arrResult, 2) + 1)
            With objTable
                .Borders.InsideLineStyle = wdLineStyleSingle
                .Borders.OutsideLineStyle = wdLineStyleSingle
                For i = 1 To UBound(arrResult, 2) + 1
                    .Cell(1, i).Range.Text = arrResult(0, i - 1)
                Next
                For i = 2 To UBound(arrResult) + 1
                    For j = 1 To UBound(arrResult, 2) + 1
                        If IsNull(arrResult(i - 1, j - 1)) Then
                            .Cell(i, j).Range.Text = vbNullString
                        Else: .Cell(i, j).Range.Text = arrResult(i - 1, j - 1)
                        End If
                    Next
                Next
            End With
        Else
            MsgBox "Please place your mouse cursor outside any tables before running this macro.", vbExclamation, "Error: Mouse Cursor Inside a Table"
            Exit Sub
        End If
    End If
    Set objTable = Nothing
End Sub

Private Function RunSQL(SQLStatement As String) As Variant
Dim objCnn As ADODB.Connection
    Dim objCmd As ADODB.Command
    Dim objRs As ADODB.Recordset
    Dim arrResult() As Variant, arrMergedArray() As Variant, arrColumnNames() As Variant
    Dim strCnn As String, strColumnName As String
    Dim i As Long
    strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\QuanLyBanHang.accdb;Persist Security Info=False;"
    Set objCnn = New ADODB.Connection
    With objCnn
        .ConnectionString = strCnn
        .Open
    End With
    Set objCmd = New ADODB.Command
    With objCmd
        .ActiveConnection = objCnn
        .CommandText = SQLStatement
        On Error Resume Next
        Set objRs = .Execute
        If Err.Number <> 0 Then Exit Function
    End With
    If Not objRs.BOF And Not objRs.EOF Then
        arrResult = Transpose2DArray(objRs.GetRows)
        ReDim arrColumnNames(0 To objRs.Fields.Count - 1)
        For i = 0 To objRs.Fields.Count - 1
            arrColumnNames(i) = objRs.Fields.Item(i).Name
        Next
        arrMergedArray = PrepareOutputData(arrColumnNames, arrResult)
        Erase arrColumnNames
        Erase arrResult
        objRs.Close
        objCnn.Close
        RunSQL = arrMergedArray
    Else
        Exit Function
    End If
    Set objCnn = Nothing
    Set objCmd = Nothing
    Set objRs = Nothing
End Function

Private Function PrepareOutputData(ColumnNames As Variant, TableData As Variant)
    Dim arrResult() As Variant
    Dim i As Long, j As Long
    If Is2DArray(TableData) Then
        ReDim arrResult(0 To UBound(TableData), 0 To UBound(ColumnNames))
        For i = 0 To UBound(ColumnNames)
            arrResult(0, i) = ColumnNames(i)
        Next
        For i = 1 To UBound(TableData)
            For j = 0 To UBound(ColumnNames)
                arrResult(i, j) = TableData(i, j + 1)
            Next
        Next
    Else
        ReDim arrResult(0 To 1, 0 To UBound(ColumnNames))
        For i = 0 To UBound(ColumnNames)
            arrResult(0, i) = ColumnNames(i)
        Next
        For j = 0 To UBound(ColumnNames)
            arrResult(1, j) = TableData(j + 1)
        Next
    End If
    PrepareOutputData = arrResult
End Function

Private Function Is2DArray(InputArray As Variant) As Boolean
    Dim lngMaxColumnIndex As Long
    On Error Resume Next
    lngMaxColumnIndex = UBound(InputArray, 2)
    If Err.Number = 9 Then
        Err.Clear
        Is2DArray = False
    Else: Is2DArray = True
    End If
End Function

Private Function Transpose2DArray(InputArray As Variant) As Variant
    Dim arrResult() As Variant
    Dim i As Integer, j As Integer
    Dim lngMaxColumnIndex As Long
    On Error Resume Next
    lngMaxColumnIndex = UBound(InputArray, 2)
    If Err.Number = 9 Then
        Err.Clear
        Transpose2DArray = InputArray
        Exit Function
    End If
    ReDim arrResult(1 To UBound(InputArray, 2) + 1, 1 To UBound(InputArray) + 1)
    For i = 1 To UBound(InputArray) + 1
        For j = 1 To UBound(InputArray, 2) + 1
            arrResult(j, i) = InputArray(i - 1, j - 1)
        Next
    Next
    Transpose2DArray = arrResult
End Function