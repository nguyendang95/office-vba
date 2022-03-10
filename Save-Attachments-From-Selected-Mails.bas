Option Explicit

Sub SaveAttachmentsFromSelectedMails()
    Dim objSelection As Outlook.Selection
    Dim objMail As Outlook.MailItem
    Dim objAtt As Outlook.Attachment
    Dim objFSO, objShell As Object
    Dim i, j As Long
    Dim strFolderPath, strRCDatePath, strRCTimePath, strSenderPath, strFileName As String
    Dim lngCountOfAttchs As Long
    
    strFolderPath = Environ$("USERPROFILE") & "\Documents\Outlook Attachments"
    Set objSelection = Application.ActiveExplorer.Selection
    If objSelection.Count > 0 Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        If Not objFSO.FolderExists(strFolderPath) Then objFSO.CreateFolder strFolderPath
        Set objShell = CreateObject("WScript.Shell")
        For i = 1 To objSelection.Count
            If TypeOf objSelection.Item(i) Is Outlook.MailItem Then
                Set objMail = objSelection.Item(i)
                If Not objMail.Attachments.Count = 0 Then
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
    Else
        MsgBox "To run this macro, you need to select at least an email.", vbExclamation, "You did not select any email"
        Exit Sub
    End If
    If Not lngCountOfAttchs = 0 Then objShell.Run "explorer """ & strFolderPath & "", vbNormalFocus
    Set objSelection = Nothing
    Set objMail = Nothing
    Set objAttch = Nothing
    Set objFSO = Nothing
    Set objShell = Nothing
End Sub


Private Function GetDateFromReceivedTime(strReceivedDate As String) As String
    GetDateFromReceivedTime = Left(strReceivedDate, InStr(strReceivedDate, " "))
End Function

Private Function GetTimeFromReceivedTime(strReceivedTime As String) As String
    GetTimeFromReceivedTime = Mid(strReceivedTime, InStr(strReceivedTime, " "))
End Function
