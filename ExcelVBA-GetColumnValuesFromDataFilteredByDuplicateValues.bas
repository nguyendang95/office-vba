Option Explicit

Public Sub FilterAndCopyResult()
    Dim objRange As Excel.Range, objResultRange As Excel.Range
    Dim strSQL As String
    Dim arrValues() As Variant, arrColumnNames() As Variant, arrResult() As Variant
    Dim varColumnIndex As Variant
    Dim strSelectedField As String, strDataRange As String, strLastColumn As String
    On Error Resume Next
    Set objRange = Application.InputBox("Specify data range to process", "Specify Data Range", , , , , , 8)
    If Not objRange Is Nothing Then
        If IsArray(objRange.Value) Then
            arrValues = objRange.Value
        Else
            MsgBox "Please select data range with at least two columns and try again.", vbExclamation, "Error: Data Range Not Large Enough"
            Exit Sub
        End If
        varColumnIndex = InputBox("Specify column number to get data with an assumption that last column is always the column that contains duplicate values.", "Specify Column Number")
        If varColumnIndex <> vbNullString And IsNumeric(varColumnIndex) Then
            If varColumnIndex < UBound(arrValues, 2) Then
                strSelectedField = "[" & arrValues(1, varColumnIndex) & "]"
                strDataRange = "[" & ActiveSheet.Name & "$" & objRange.Address(False, False, xlA1) & "]"
                strLastColumn = "[" & arrValues(1, UBound(arrValues, 2)) & "]"
                strSQL = "SELECT " & strSelectedField & _
                        " FROM " & strDataRange & _
                        " WHERE " & strLastColumn & " NOT IN " & _
                        "( SELECT " & strLastColumn & _
                        " FROM " & strDataRange & _
                        " GROUP BY " & strLastColumn & _
                        " HAVING COUNT(" & strLastColumn & ") = 1 )"
                arrResult = RunSQL(strSQL)
                On Error Resume Next
                Set objResultRange = Application.InputBox("Select a cell to spill the filtered data", , , , , , , 8)
                If Not objResultRange Is Nothing Then
                    objResultRange.Offset(0, 0).Resize(UBound(arrResult), UBound(arrResult, 2)).Value = arrResult
                Else
                    MsgBox "You need to select a cell to spill the filtered data. Please try again!", vbExclamation, "Error"
                    Exit Sub
                End If
            Else
                MsgBox "The column index should not larger than last column index", vbExclamation, "Error"
                Exit Sub
            End If
        Else
            MsgBox "You must specify a number, not a text, for column index.", vbExclamation, "Error"
            Exit Sub
        End If
    End If
    Set objRange = Nothing
    Set objResultRange = Nothing
End Sub

Private Function RunSQL(SQLStatement As String) As Variant
Dim objCnn As ADODB.Connection
    Dim objCmd As ADODB.Command
    Dim objRs As ADODB.Recordset
    Dim arrResult() As Variant, arrMergedArray() As Variant, arrColumnNames() As Variant
    Dim strCnn As String, strColumnName As String
    Dim i As Long
    If ActiveWorkbook.Path <> vbNullString Then
        strCnn = GetConnectionString(ActiveWorkbook.FullName)
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
            If Err.Number <> 0 Then
                MsgBox Err.Description, vbExclamation, "Error"
                Exit Function
            End If
        End With
        If Not objRs Is Nothing Then
            On Error Resume Next
            arrResult = Transpose2DArray(objRs.GetRows)
            If Err.Number <> 0 Then
                RunSQL = vbNullString
                Exit Function
            End If
        End If
        objRs.Close
        objCnn.Close
        RunSQL = arrResult
    Else
        MsgBox "Please save this worbkook before performing this operation.", vbExclamation, "Error: Workbook Not Saved"
        Exit Function
    End If
    Set objCnn = Nothing
    Set objCmd = Nothing
    Set objRs = Nothing
End Function

Private Function GetConnectionString(FileName As String) As String
    Dim strConnectionString As String
    Select Case GetFileExtension(FileName)
        Case ".xlsx"
            strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FileName & ";Extended Properties='Excel 12.0 Xml;HDR=YES';"
        Case ".xlsm"
            strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FileName & ";Extended Properties='Excel 12.0 Macro;HDR=YES';"
        Case ".xlsb"
            strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FileName & ";Extended Properties='Excel 12.0;HDR=YES';"
        Case ".xls"
            strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FileName & ";Extended Properties='Excel 8.0;HDR=YES';"
    End Select
    GetConnectionString = strConnectionString
End Function

Private Function GetFileExtension(FileName As String) As String
    On Error Resume Next
    GetFileExtension = Mid(FileName, InStrRev(FileName, "."))
    If Err.Number = 5 Then
        GetFileExtension = vbNullString
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