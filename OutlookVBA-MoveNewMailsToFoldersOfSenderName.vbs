'ThisOutlookSession
'To run this macro, create a folder named "Senders" first at the root folder, otherwise it will throw an error
Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)

Dim olName As NameSpace
Dim olItem As Object
Dim olSendersFldr As Folder
Dim olNewFldr As Folder
Dim i As Integer
Dim ArrID() As String

Set olName = Application.GetNamespace("MAPI")
ArrID() = Split(EntryIDCollection, ",")
For i = 0 To UBound(ArrID())
    Set olItem = olName.GetItemFromID(ArrID(i))
    If TypeOf olItem Is MailItem Then
        Set olSendersFldr = olItem.Parent.Parent.Folders("Senders")
        On Error GoTo Loi
        Set olNewFldr = olSendersFldr.Folders.Add(olItem.SenderName)
        olItem.Move olNewFldr
    End If
    If Err.Number = -2147221233 Then
Loi:
    Err.Clear
    Set olNewFldr = olSendersFldr.Folders(olItem.SenderName)
    olItem.Move olNewFldr
    End If
Next
MsgBox "New incoming email. Check it out!", vbInformation, "New Mail Received"

End Sub