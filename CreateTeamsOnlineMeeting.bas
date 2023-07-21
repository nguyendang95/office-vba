Option Explicit

Public Sub CreateNewEmailWithMeetingInvitation()
    Dim objSheet As Excel.Worksheet
    Set objSheet = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    Dim objDataRange As Excel.Range
    Set objDataRange = objSheet.Range("A2").CurrentRegion
    Dim arrDataRange As Variant
    arrDataRange = objDataRange.Value
    Dim objListOfMeetingSubjects As Scripting.Dictionary
    Set objListOfMeetingSubjects = New Scripting.Dictionary
    Dim i As Long
    For i = 2 To UBound(arrDataRange, 2)
        If objListOfMeetingSubjects.Exists(arrDataRange(i, 1)) = False Then objListOfMeetingSubjects.Add arrDataRange(i, 1), 1
    Next
    Dim arrListOfMeetingSubjects As Variant
    arrListOfMeetingSubjects = objListOfMeetingSubjects.Keys
    Dim colParticipants As Collection
    Dim j As Long
    Dim objMeeting As OnlineMeeting
    Dim objCnn As ADODB.Connection
    Set objCnn = New ADODB.Connection
    With objCnn
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & Application.ActiveWorkbook.FullName & ";Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
        .Open
    End With
    Dim objRs As ADODB.Recordset
    Dim objCmd As ADODB.Command
    Dim arrRows As Variant
    For i = LBound(arrListOfMeetingSubjects) To UBound(arrListOfMeetingSubjects)
        Set objCmd = New ADODB.Command
        With objCmd
            Set .ActiveConnection = objCnn
            .CommandText = "SELECT * FROM [Sheet1$" & objDataRange.Address(False, False, xlA1) & "] WHERE [" & arrDataRange(1, 1) & "] LIKE '" & arrListOfMeetingSubjects(i) & "';"
            Set objRs = .Execute
        End With
        If Not (objRs.BOF And objRs.EOF) Then
            arrRows = Application.WorksheetFunction.Transpose(objRs.GetRows)
            objRs.Close
            Set colParticipants = New Collection
            For j = 1 To UBound(arrRows)
                colParticipants.Add arrRows(j, 2)
            Next
            Set objMeeting = CreateNewTeamsMeeting(arrListOfMeetingSubjects(i), CDate(CStr(arrRows(1, 3)) & " " & CStr(arrRows(1, 4))), arrRows(1, 5))
            If Not objMeeting Is Nothing Then
                For j = 1 To colParticipants.Count
                    Call CreateNewEmail(colParticipants.Item(j), arrRows(1, 6), objMeeting.JoinWebUrl, objMeeting.Subject)
                Next
            End If
        End If
    Next
    objDataRange.AutoFilter
End Sub

Private Sub CreateNewEmail(ByVal ToEmailAddress As String, ByVal Description As String, ByVal InvitationLink As String, ByVal MeetingSubject As String)
    Dim objOlApp As Outlook.Application
    On Error Resume Next
    Set objOlApp = GetObject(, "Outlook.Application")
    If Err.Number = 429 Then
        MsgBox "You must open Outlook first before running this macro. Please try again!", vbExclamation, "Error"
        Exit Sub
    End If
    Dim objOlMail As Outlook.MailItem
    Set objOlMail = objOlApp.CreateItem(olMailItem)
    With objOlMail
        .To = ToEmailAddress
        .Subject = MeetingSubject
        .BodyFormat = olFormatHTML
        .HTMLBody = "<html><body><p>" & Description & "</p><a href=" & Chr(34) & InvitationLink & Chr(34) & ">Click here to join the meeting: " & MeetingSubject & "</a></body></html>"
        .Display
    End With
End Sub

Private Function CreateNewTeamsMeeting(ByVal Subject As String, ByVal StartDateTime As Date, ByVal Duration As Long) As OnlineMeeting
    Dim objMicrosoft As MicrosoftGraphOAuth2
    Set objMicrosoft = GetSession()
    If Not objMicrosoft Is Nothing Then
        Dim objMeeting As OnlineMeeting
        Set objMeeting = New OnlineMeeting
        With objMeeting
            .StartDateTime = Replace(JsonConverter.ConvertToIso(StartDateTime), "Z", "-07:00")
            .EndDateTime = Replace(JsonConverter.ConvertToIso(DateAdd("n", Duration, StartDateTime)), "Z", "-07:00")
            .Subject = Subject
        End With
        Dim objCreatedMeeting As OnlineMeeting
        On Error Resume Next
        Set objCreatedMeeting = objMicrosoft.CallsAndOnlineMeetings.OnlineMeetings.OnlineMeeting.Create(MeEnpoint, objMeeting)
        If Err.Number <> 0 Then
            MsgBox Err.Description, vbExclamation, "Error"
            Exit Function
        End If
        On Error GoTo 0
        If Not objCreatedMeeting Is Nothing Then Set CreateNewTeamsMeeting = objCreatedMeeting
    End If
End Function

Private Function GetSession() As MicrosoftGraphOAuth2
    Dim objMicrosoft As MicrosoftGraphOAuth2
    Set objMicrosoft = New MicrosoftGraphOAuth2
    With objMicrosoft
        .Tenant = Organizations
        .ApplicationName = "Teams"
        .ClientId = "510ce4dc-7eb1-48d9-a237-4efaa4c6852d"
        .Scope = Array("Files.ReadWrite.All", "Channel.ReadBasic.All", "OnlineMeetings.ReadWrite")
        On Error Resume Next
        .AuthorizeOAuth2
        If Err.Number = 0 Then Set GetSession = objMicrosoft
        On Error GoTo 0
    End With
End Function

Private Sub LogOut()
    Dim objMicrosoft As MicrosoftGraphOAuth2
    Set objMicrosoft = New MicrosoftGraphOAuth2
    With objMicrosoft
        .Tenant = Organizations
        .ApplicationName = "Teams"
        .ClientId = "510ce4dc-7eb1-48d9-a237-4efaa4c6852d"
        On Error Resume Next
        .LogOut
        If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, "Error"
        On Error GoTo 0
    End With
End Sub