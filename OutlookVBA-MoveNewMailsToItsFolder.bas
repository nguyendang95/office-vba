'ThisOutlookSession
Private Sub Application_Startup()
Call Call_ItemAdd_Event
End Sub

'Call Class Module
Dim myClass As New MoveNewMailToFolder

Sub Call_ItemAdd_Event()
    Set myClass.olItems = Outlook.GetNamespace("MAPI").Folders("youremail").Folders("Inbox").Items
End Sub

'Class Module
Option Explicit
Public WithEvents olItems As Outlook.Items
 
Sub Application_Startup()

Dim olName As NameSpace
Dim olInboxFld As Folder

Set olName = Application.GetNamespace("MAPI")
Set olInboxFldr = olName.Folders("youremail").Folders("Inbox")
Set olItems = olInboxFldr.Items

End Sub

Sub olItems_ItemAdd(ByVal olItem As Object)

Dim NameOfSender As String
Dim olNewFldr As Folder
Dim olSendersFldr As Folder

If TypeOf olItem Is MailItem Then
    NameOfSender = olItem.SenderName
    If FolderExistsInSenders(NameOfSender) Then
        Set olNewFldr = Application.Session.Folders("youremail").Folders("Senders").Folders(NameOfSender)
        olItem.Move olNewFldr
    Else
        Set olNewFldr = Application.Session.Folders("youremail").Folders("Senders").Folders.Add(NameOfSender)
        olItem.Move olNewFldr
    End If
End If
Application.Session.SendAndReceive True

End Sub

Function FolderExistsInSenders(FolderName As String) As Boolean

Dim FolderObject As Folder
Dim SendersFldr As Folder

On Error GoTo FolderNotFoundErr
Set SendersFldr = Application.Session.Folders("youremail").Folders("Senders")
If SendersFldr.Folders.Count = 0 Then
    FolderExistsInSenders = False
    Exit Function
End If
For Each FolderObject In SendersFldr.Folders
    If FolderObject.Name = FolderName Then
        FolderExistsInSenders = True
        Exit For                               
    Else: FolderExistsInSenders = False
    End If
Next
If Err.Number = -2147221233 Then
FolderNotFoundErr: MsgBox "You need to create a folder named 'Senders' before running this macro. Please try again.", vbExclamation, "Senders Folder Not Found"
End If

End Function
