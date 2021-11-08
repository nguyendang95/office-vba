Option Explicit

Sub MoveMailsToFolderOfItsName()

Dim MailFolderName As Folder
Dim olName As NameSpace
Dim SendersFld As Folder
Dim olInboxFld As Folder
Dim olMail As MailItem
Dim CountOfMail As Long

Set olName = Application.GetNamespace("MAPI")
If SendersFolderExists() Then
    Set SendersFld = olName.Folders("your_email_address@abcd.com").Folders("Senders")
Else
    Set SendersFld = olName.Folders("your_email_address@abcd.com").Folders.Add("Senders")
End If
Set olInboxFld = olName.Folders("your_email_address@abcd.com").Folders("Inbox")
For Each olMail In olInboxFld.Items
    If TypeOf olMail Is MailItem Then
        CountOfMail = CountOfMail + 1
        If MailFolderExists(olMail.SenderName) = False Then
            Set MailFolderName = SendersFld.Folders.Add(olMail.SenderName)
            olMail.Move MailFolderName
        Else
            Set MailFolderName = SendersFld.Folders(olMail.SenderName)
            olMail.Move MailFolderName
        End If
    End If
Next
olName.SendAndReceive True
If Err.Number = 0 Then MsgBox "Operation complete. " & CountOfMail & " mail items have been moved to folder of its name.", vbInformation + vbOKOnly, "Mail Items Successfully Moved To Folder Of Its Name"

End Sub

Function SendersFolderExists() As Boolean

Dim FldObj As Folder

For Each FldObj In Application.Session.Folders("your_email_address@abcd.com").Folders
    If FldObj.Name = "Senders" Then
        SendersFolderExists = True
        Exit For
    Else: SendersFolderExists = False
    End If
Next

End Function

Function MailFolderExists(FldName As String) As Boolean

Dim FldObj As Folder

For Each FldObj In Application.Session.Folders("your_email_address@abcd.com").Folders("Senders").Folders
    If FldObj.Name = FldName Then
        MailFolderExists = True
        Exit For
    Else: MailFolderExists = False
    End If
Next

End Function
