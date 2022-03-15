Option Explicit

Public Sub SaveAttachments()
    Dim objStore As Outlook.Store
    Dim objInboxFld As Outlook.Folder
    Dim objMail As Outlook.MailItem
    Dim objAtt As Outlook.Attachment
    Dim colItems As Outlook.Items
    Dim objItem As Object
    Dim objFSO As Object, objShell As Object
    Dim strFind As String
    Dim i As Long, j As Long, lngCountOfAttchs As Long
    Dim strFolderPath As String, strRCDatePath As String, strRCTimePath As String, strSenderPath As String, strFileName As String
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    strFolderPath = Environ$("USERPROFILE") & "\Documents\Outlook Attachments"
    If Not objFSO.FolderExists(strFolderPath) Then objFSO.CreateFolder strFolderPath
    Set objStore = Application.ActiveExplorer.CurrentFolder.Store
    Set objInboxFld = objStore.GetDefaultFolder(olFolderInbox)
    strFind = "@SQL=" & "%thismonth(" & Quote("urn:schemas:httpmail:datereceived") & ")%" & " And " & Quote("urn:schemas:httpmail:fromemail") & "='abc@gmail.com'" & " And " & Quote("urn:schemas:httpmail:hasattachment") & "=1"
    Set colItems = objInboxFld.Items.Restrict(strFind)
    If Not colItems Is Nothing Then
        For Each objItem In colItems
            If objItem.Class = olMail Then
                Set objMail = objItem
                If objItem.Attachments.Count > 0 Then
                    With objFSO
                        If Not .FolderExists(strFolderPath & "\" & objMail.SenderEmailAddress) Then
                            strSenderPath = strFolderPath & "\" & objMail.SenderEmailAddress
                            .CreateFolder strSenderPath
                        Else: strSenderPath = strFolderPath & "\" & objMail.SenderEmailAddress
                        End If
                        If Not .FolderExists(strSenderPath & "\" & Trim(Format(GetDateFromReceivedTime(objMail.ReceivedTime), "dd-mm-yyyy"))) Then
                            strRCDatePath = strSenderPath & "\" & Trim(Format(GetDateFromReceivedTime(objMail.ReceivedTime), "dd-mm-yyyy"))
                            .CreateFolder strRCDatePath
                        Else: strRCDatePath = strSenderPath & "\" & Trim(Format(GetDateFromReceivedTime(objMail.ReceivedTime), "dd-mm-yyyy"))
                        End If
                        If Not .FolderExists(strRCDatePath & "\" & Trim(Format(GetTimeFromReceivedTime(objMail.ReceivedTime), "hh.mm.ss am/pm"))) Then
                            strRCTimePath = strRCDatePath & "\" & Trim(Format(GetTimeFromReceivedTime(objMail.ReceivedTime), "hh.mm.ss am/pm"))
                            .CreateFolder strRCTimePath
                        Else: strRCTimePath = strRCDatePath & "\" & Trim(Format(GetTimeFromReceivedTime(objMail.ReceivedTime), "hh.mm.ss am/pm"))
                        End If
                    End With
                    For j = 1 To objMail.Attachments.Count
                        Set objAtt = objMail.Attachments.Item(j)
                        strFileName = strRCTimePath & "\" & objAtt.FileName
                        If Not objFSO.FileExists(strFileName) Then
                            objAtt.SaveAsFile strFileName
                            lngCountOfAttchs = lngCountOfAttchs + 1
                        Else
                            Kill strFileName
                            objAtt.SaveAsFile strFileName
                            lngCountOfAttchs = lngCountOfAttchs + 1
                        End If
                    Next
                End If
            End If
        Next
    Set objShell = CreateObject("WSCript.Shell")
    End If
    If Not lngCountOfAttchs = 0 Then objShell.Run "explorer """ & strFolderPath & "", vbNormalFocus
    Set objStore = Nothing
    Set objInboxFld = Nothing
    Set objMail = Nothing
    Set objAtt = Nothing
    Set colItems = Nothing
    Set objItem = Nothing
    Set objFSO = Nothing
    Set objShell = Nothing
End Sub

Private Function GetDateFromReceivedTime(strReceivedDate As String) As String
    GetDateFromReceivedTime = Left(strReceivedDate, InStr(strReceivedDate, " "))
End Function

Private Function GetTimeFromReceivedTime(strReceivedTime As String) As String
    GetTimeFromReceivedTime = Mid(strReceivedTime, InStr(strReceivedTime, " "))
End Function

Private Function Quote(Text As String)
    Quote = Chr(34) & Text & Chr(34)
End Function
