Option Explicit

Sub SaveAttachmentsFromSelectedMails()

Dim olExplorer As Explorer
Dim olSelection As Selection
Dim olMail As MailItem
Dim olAttch As Attachments
Dim i As Integer
Dim j As Integer
Dim AttchFldr As String

AttchFldr = Environ$("USERPROFILE") & "\OneDrive\Documents\Outlook Attachments\"
Set olExplorer = Application.ActiveExplorer
Set olSelection = olExplorer.Selection
If olSelection.Count > 0 Then
    For i = 1 To olSelection.Count
        If TypeOf olSelection.Item(i) Is MailItem Then
            If MsgBox("You have selected " & CStr(olSelection.Count) & "mails to save attachments. Do you want to proceed?", vbQuestion + vbOKOnly) = vbYes Then
                Set olMail = olSelection.Item(i)
                If olMail.Attachments.Count > 1 Then
                    For j = 1 To olMail.Attachments.Count
                        Set olAttch = olMail.Attachments.Item(i)
                        olAttch.SaveAsFile AttchFldr
                    Next
                Else: Exit For
                End If
            Else: Exit Sub
            End If
        Else
            MsgBox "Your selection does not include email item. Please try again!", vbExclamation, "No Mails Selected"
            Exit Sub
        End If
    Next
Else
    MsgBox "To run this macro, you need to select at least an email.", vbExclamation, "You did not select any email"
    Exit Sub
End If
If Err.Number = 0 Then MsgBox "All attachments in emails of your selection were saved to " & AttchFldr & ". Operation complete!", vbInformation, "Operation Complete"
Set olExplorer = Nothing
Set olSelection = Nothing
Set olMail = Nothing
Set olAttch = Nothing

End Sub
