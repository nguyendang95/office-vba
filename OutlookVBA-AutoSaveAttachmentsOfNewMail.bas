'ThisOutlookSession
Private Sub Application_Startup()
Call CallEventSaveAllAtts
End Sub

'Module CallEventSaveNewAttachments
Option Explicit

Dim myClass As New SaveAllAttachments_cls

Sub CallEventSaveAllAtts()
Set myClass.olItems = Outlook.Session.GetDefaultFolder(olFolderInbox).Items
End Sub

Option Explicit

Public WithEvents olItems As Items

Sub Application_Startup()

Dim olInboxFld As Folder
Dim olName As NameSpace

Set olName = Application.GetNamespace("MAPI")
Set olInboxFld = olName.GetDefaultFolder(olFolderInbox)
Set olItems = olInboxFld.Items

End Sub

Private Sub olItems_ItemAdd(ByVal olMail As Object)

Dim fs As Object
Dim fldrpath As String
Dim i As Integer
Dim olAtt As Attachment
Dim RCDatePath As String
Dim SenderPath As String

Set fs = CreateObject("Scripting.FileSystemObject")
fldrpath = Environ$("USERPROFILE") & "\OneDrive\Documents\BaiDich\"
If TypeOf olMail Is MailItem Then
    If olMail.Attachments.Count > 0 Then
        If Not fs.FolderExists(fldrpath & Trim(Replace(GetDateFromReceivedTime(olMail.ReceivedTime), "/", "-"))) Then
            RCDatePath = fldrpath & Trim(Replace(GetDateFromReceivedTime(olMail.ReceivedTime), "/", "-"))
            fs.CreateFolder RCDatePath
        Else: RCDatePath = fldrpath & Trim(Replace(GetDateFromReceivedTime(olMail.ReceivedTime), "/", "-"))
        End If
        If Not fs.FolderExists(fldrpath & Trim(Replace(GetDateFromReceivedTime(olMail.ReceivedTime), "/", "-")) & "\" & olMail.SenderName) Then
            SenderPath = fldrpath & Trim(Replace(GetDateFromReceivedTime(olMail.ReceivedTime), "/", "-")) & "\" & olMail.SenderName
            fs.CreateFolder SenderPath
        Else: SenderPath = fldrpath & Trim(Replace(GetDateFromReceivedTime(olMail.ReceivedTime), "/", "-")) & "\" & olMail.SenderName
        End If
        If olMail.Attachments.Count > 0 Then
            For i = 1 To olMail.Attachments.Count
                Set olAtt = olMail.Attachments.Item(i)
                olAtt.SaveAsFile SenderPath & "\" & olAtt.DisplayName
            Next i
        Else: Exit For
        End If
    End If
    If Err.Number = 0 Then
        MsgBox "A new email just received. Attachment(s) saved to " & SenderPath & ".", vbInformation, "New Attachments received"
        Shell "explorer """ & SenderPath & "", vbNormalFocus
    End If
End If

End Sub

Function GetDateFromReceivedTime(ReceivedDateString As String) As String

Dim RDString As String

RDString = Left(ReceivedDateString, InStr(ReceivedDateString, " "))
GetDateFromReceivedTime = RDString

End Function
