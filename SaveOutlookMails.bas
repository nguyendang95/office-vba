Option Explicit

Public Sub RestrictTimePeriod()
    Dim fld As Outlook.Folder
    Dim msg As Outlook.MailItem
    Dim filterCriteria As String
    Dim filterItems As Outlook.Items
    Dim i As Long
    Dim startdate As String, arrstartdate As Variant
    Dim enddate As String, arrenddate As Variant
    Set fld = Application.Session.GetDefaultFolder(olFolderSentMail)
    If Not fld Is Nothing Then
        startdate = CStr(ThisWorkbook.Worksheets("A6.Confirmation").Range("C9").Value)
        arrstartdate = Split(startdate, " ")
        arrstartdate(1) = Trim(Left(arrstartdate(1), InStrRev(arrstartdate(1), ":") - 1))
        startdate = Join(arrstartdate, " ")
        enddate = CStr(ThisWorkbook.Worksheets("A6.Confirmation").Range("C10").Value)
        arrenddate = Split(enddate, " ")
        arrenddate(1) = Trim(Left(arrenddate(1), InStrRev(arrenddate(1), ":") - 1))
        enddate = Join(arrenddate, " ")
        Const PR_SUBJECT_W As String = "http://schemas.microsoft.com/mapi/proptag/0x0037001F" 'Hoac, urn:schemas:httpmail:subject
        'ReceivedTime tuong duong voi urn:schemas:httpmail:datereceived
        If fld.Store.IsInstantSearchEnabled Then
            ' Su dung truy van DASL de tim kiem item
            filterCriteria = Quote("urn:schemas:httpmail:datereceived") & " > " & SingleQuote(startdate) & _
                     " And " & Quote("urn:schemas:httpmail:datereceived") & " < " & SingleQuote(enddate) & " AND " & Quote(PR_SUBJECT_W) & " ci_phrasematch " & SingleQuote("test")
        Else: filterCriteria = Quote("urn:schemas:httpmail:datereceived") & " > " & SingleQuote(startdate) & _
                     " And " & Quote("urn:schemas:httpmail:datereceived") & " < " & SingleQuote(enddate) & " AND " & Quote(PR_SUBJECT_W) & " LIKE " & SingleQuote("*test*")
        End If
        Set filterItems = fld.Items.Restrict(filterCriteria)
        If filterItems.Count > 0 Then
            For i = 1 To filterItems.Count
                If TypeOf filterItems.Item(i) Is Outlook.MailItem Then
                    Set msg = filterItems.Item(i)
                    msg.SaveAs ThisWorkbook.Path & "\" & i & ".msg", 3
                End If
            Next
        End If
    End If
End Sub

Private Function Quote(ByVal Text As String) As String
    Quote = Chr(34) & Text & Chr(34)
End Function

Private Function SingleQuote(ByVal Text As String) As String
    SingleQuote = Chr(39) & Text & Chr(39)
End Function
