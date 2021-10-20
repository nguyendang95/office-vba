Attribute VB_Name = "ReplyToEmailsofChoice"
Option Explicit

Sub ReplyToSelectedEmails()

Dim olApp As Object
Dim olExplorer As Object
Dim olSelection As Object
Dim olReplyMail As Object
Dim MailNum As Integer

On Error GoTo AppError
Set olApp = GetObject(, "Outlook.Application")
Set olExplorer = olApp.ActiveExplorer
Set olSelection = olExplorer.Selection
If olSelection.Count > 0 Then
    For MailNum = 1 To olSelection.Count
        On Error GoTo SelectionError
        Set olReplyMail = olSelection.Item(MailNum)
        With olReplyMail.Reply
            .Display
            .BodyFormat = olFormatHTML
            .HTMLBody = "<html><body style='font-size:12pt;font-family:Times New Roman;text-align:justify'><p>Thanks. I have received your email.</p></body></html>" & .HTMLBody
            On Error GoTo AddAttachmentError
            .Attachments.Add Source:=ActiveDocument.FullName
        End With
    Next
Else
    MsgBoxW "You need to select at least an email to reply.", vbExclamation, "No Email Selected"
End If
If Err.Number = 13 Then
Err.Clear
SelectionError: MsgBoxW "Your selection does not include email item(s). Please try again!", vbOKOnly + vbExclamation, "You need to select email item(s)"
Exit Sub
End If
If Err.Number = -2147024894 Then
Err.Clear
AddAttachmentError: MsgBoxW "You need to save your current document before attaching it to your reply mail.", vbExclamation + vbOKOnly, "Failed to Add Current Document As Attachment"
Exit Sub
End If
If Err.Number = 429 Then
Err.Clear
AppError: MsgBoxW "You need to launch Outlook in order to proceed next step. Please try again!", vbCritical, "Outlook Instance Not Found"
Exit Sub
End If
Set olApp = Nothing
Set olExplorer = Nothing
Set olReplyMail = Nothing
Set olSelection = Nothing

End Sub
