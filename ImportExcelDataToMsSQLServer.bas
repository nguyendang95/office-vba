Option Explicit

Private Sub ImportExcelDataToMsSQLServer()
    Dim objCn As Object
    Dim objDataRange As Excel.Range
    Dim lngLastRow As Long, lngLastColumn, i As Long
    Dim strCnnStr As String
    Dim arrData()
    lngLastRow = Cells(Rows.Count, 2).End(xlUp).Row
    lngLastColumn = Cells(2, Columns.Count).End(xlToLeft).Column
    Set objDataRange = Range(Cells(3, 2), Cells(lngLastRow, lngLastColumn))
    arrData() = objDataRange.Value
    Set objCn = CreateObject("ADODB.Connection")
    strCnnStr = "Driver={SQL Server};Server=myserver\SQLEXPRESS;Database=QuanLyThuVien;User Id=admin;Password=admin;"
    With objCn
        .ConnectionString = strCnnStr
        .Open
        For i = 1 To UBound(arrData)
            .Execute "INSERT INTO T_BANDOC " & _
                    "VALUES (" & ConvertToUnicode(SingleQuote(arrData(i, 1))) & ", " & ConvertToUnicode(SingleQuote(arrData(i, 2))) & ", " & ConvertToUnicode(SingleQuote(arrData(i, 3))) & _
                    ", " & ConvertToUnicode(SingleQuote(arrData(i, 4))) & ", " & ConvertToUnicode(SingleQuote(arrData(i, 5))) & ", " & ConvertToUnicode(SingleQuote(arrData(i, 6))) & ", " & ConvertToUnicode(SingleQuote(arrData(i, 7))) & ")"
        Next
        .Close
    End With
    MsgBox "Thao tac nhap du lieu vao database da hoan tat.", vbInformation, "Thao tac hoan tat"
    Set objCn = Nothing
    Set objDataRange = Nothing
End Sub

Private Function SingleQuote(Text As Variant) As String
    SingleQuote = Chr(39) & Text & Chr(39)
End Function

Private Function ConvertToUnicode(Text As String) As String
    ConvertToUnicode = "N" & Text
End Function
