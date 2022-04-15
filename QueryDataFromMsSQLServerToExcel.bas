Option Explicit

Private Sub GetQueryDataFromMsSQLServer()
    Dim objCn As ADODB.Connection
    Dim objRs As ADODB.Recordset
    Dim objWs As Excel.Worksheet
    Dim objLO As Excel.ListObject
    Dim arrData()
    Dim arrTitle()
    Dim strSQLQuery As String
    Dim objDataRange As Excel.Range
    Dim objColRange As Excel.Range
    Dim i As Long, lngFieldsCount As Long
    Dim strCnnStr As String
    strCnnStr = "Driver={SQL Server};Server=myserver\SQLEXPRESS;Database=QuanLyThuVien;User Id=admin;Password=admin;"
    Set objCn = New ADODB.Connection
    If objCn.State = adStateOpen Then objCn.Close
    With objCn
        .ConnectionString = strCnnStr
        .Open
    End With
    Set objRs = New ADODB.Recordset
    With objRs
        If .State = adStateOpen Then .Close
        .ActiveConnection = objCn
        strSQLQuery = "SELECT T_BANDOC.TENBD, T_BANDOC.DIACHI, COUNT(T_MUONSACH.MABD) AS TSSACH FROM T_BANDOC INNER JOIN T_MUONSACH ON T_BANDOC.MABD = T_MUONSACH.MABD " & _
                        "GROUP BY T_BANDOC.TENBD, T_BANDOC.DIACHI;"
        .Open strSQLQuery, , adOpenKeyset, adLockOptimistic
        If .BOF And .BOF Then
            .Close
            objCn.Close
            Exit Sub
        Else
            lngFieldsCount = .Fields.Count
            ReDim Preserve arrTitle(0 To lngFieldsCount)
            For i = 0 To lngFieldsCount - 1
                arrTitle(i) = .Fields(i).Name
            Next
            arrData() = .GetRows
            .Close
        End If
    End With
    objCn.Close
    Set objWs = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    Set objColRange = objWs.Range("B1").Resize(1, lngFieldsCount)
    objColRange.Value = arrTitle()
    Set objDataRange = objWs.Range("B2").Resize(UBound(arrData, 2) + 1, UBound(arrData, 1) + 1)
    objDataRange.NumberFormat = "@"
    objDataRange.Value = Application.WorksheetFunction.Transpose(arrData())
    objColRange.EntireColumn.AutoFit
    Set objLO = objWs.ListObjects.Add(xlSrcRange, objWs.UsedRange, , xlYes)
    objLO.TableStyle = "TableStyleMedium3"
    Set objCn = Nothing
    Set objRs = Nothing
    Set objDataRange = Nothing
    Set objColRange = Nothing
    Set objLO = Nothing
End Sub
