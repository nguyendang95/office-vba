Option Explicit
#If VBA7 Then
    Declare PtrSafe Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
#Else
    Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
#End If
Public Sub AttachCurrentWBToEmailWithNoFormulas()
    Dim objNewWb As Excel.Workbook
    Dim objSheet As Excel.Worksheet
    Dim objRange As Excel.Range
    Dim varRangeValue As Variant
    Dim strTempPath As String
    Dim objOlApp As Object
    Dim objOlMail As Object
    Const olByValue As Byte = 1
    Const olMailItem As Byte = 0
    On Error Resume Next
    strTempPath = Environ$("TEMP") & ActiveWorkbook.Name
    If ActiveWorkbook.Path <> vbNullString Then
        ActiveWorkbook.SaveCopyAs strTempPath
        Set objNewWb = Application.Workbooks.Open(strTempPath)
        For Each objSheet In objNewWb.Sheets
            Set objRange = objSheet.UsedRange
            varRangeValue = objRange.value
            objRange.value = varRangeValue
        Next
        objNewWb.Save
        objNewWb.Close
        Set objOlApp = GetObject(, "Outlook.Application")
        If Err.number = 429 Then
            Set objOlApp = CreateObject("Outlook.Application")
            Err.Clear
        End If
        If Not objOlApp Is Nothing Then
            Set objOlMail = objOlApp.CreateItem(olMailItem)
            With objOlMail
                .Attachments.Add strTempPath, olByValue
                .Display
            End With
        End If
        DeleteFile strTempPath
    End If
    Set objNewWb = Nothing
    Set objSheet = Nothing
    Set objOlApp = Nothing
    Set objOlMail = Nothing
End Sub
