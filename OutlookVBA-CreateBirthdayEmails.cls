Option Explicit

Private Sub Application_Reminder(ByVal Item As Object)
    Dim objAppt As Outlook.AppointmentItem
    If Item.Class = olAppointment Then
        Set objAppt = Item
        Call CreateBirthdayEmail(objAppt)
    End If
    Set objAppt = Nothing
End Sub

Private Sub CreateBirthdayEmail(Appointment As Outlook.AppointmentItem)
    Dim objMail As Outlook.MailItem
    Dim strRecipientName As String, strBirthDay As String
    Dim objAppt As Outlook.MailItem
    Dim objCont As Outlook.ContactItem
    Dim objwdDoc As Word.Document
    Dim objwdSel As Word.Selection
    Dim objSI As Outlook.StorageItem
    Dim objCalFld As Outlook.Folder
    Dim objUP As Outlook.UserProperty
    Dim strTemp As String
    Const strTemplate = "C:\Users\nguye\Email Templates\BirthdayMessage.oft"
    On Error Resume Next
    If InStr(1, Appointment.Subject, "'s Birthday") > 0 Then
        strTemp = Appointment.Subject
        strTemp = Replace(strTemp, "'s Birthday", vbNullString)
        Set objCont = GetContactLinkedToBirthday(strTemp)
        If Not objCont Is Nothing Then
            If objCont.Birthday <> #1/1/4501# Then
                Set objCalFld = Application.Session.GetDefaultFolder(olFolderCalendar)
                Set objSI = objCalFld.GetStorage("Birthday Emails Status", olIdentifyBySubject)
                If objSI.Size = 0 Then
                    Set objUP = objSI.UserProperties.Add(objCont.FullName, olText)
                    objUP.Value = CStr(Year(Date))
                    objSI.Save
                ElseIf objSI.UserProperties.Find(objCont.FullName, True) Is Nothing Then
                    Set objUP = objSI.UserProperties.Add(objCont.FullName, olText)
                    objUP.Value = CStr(Year(Date))
                    objSI.Save
                ElseIf objSI.UserProperties.Item(objCont.FullName).Value = CStr(Year(Date)) Then
                    Exit Sub
                End If
                Set objMail = Application.CreateItemFromTemplate(strTemplate)
                If Err.Number <> 0 Then
                    Set objMail = Application.CreateItem(olMailItem)
                    With objMail
                        .Subject = Appointment.Subject
                        .To = GetEmailAddressFromContact(objCont)
                        .Display
                    End With
                    Exit Sub
                End If
                With objMail
                    .Subject = .Subject & " " & objCont.FullName
                    .To = GetEmailAddressFromContact(objCont)
                    .Display
                End With
                strRecipientName = objCont.FullName
                strBirthDay = CStr(Format(objCont.Birthday, "dd.mm.yyyy", vbMonday))
                Set objwdDoc = objMail.GetInspector.WordEditor
                Set objwdSel = objwdDoc.Windows(1).Selection
                objwdSel.HomeKey wdStory
                With objwdSel.Find
                    .ClearFormatting
                    .Text = "Ten_Khach_Hang"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Replacement.Text = strRecipientName
                    .Execute Replace:=wdReplaceAll
                End With
                objwdSel.HomeKey wdStory
                With objwdSel.Find
                    .ClearFormatting
                    .Text = "Ngay_Sinh"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Replacement.Text = strBirthDay
                    .Execute Replace:=wdReplaceAll
                End With
            End If
        End If
    End If
    Set objMail = Nothing
    Set objAppt = Nothing
    Set objCont = Nothing
    Set objwdDoc = Nothing
    Set objwdSel = Nothing
    Set objSI = Nothing
    Set objUP = Nothing
    Set objCalFld = Nothing
End Sub

Private Function GetContactLinkedToBirthday(ContactFullName As String) As Outlook.ContactItem
    Dim objContact As Outlook.ContactItem
    Dim objFolder As Outlook.Folder
    Dim strFind As String
    Set objFolder = Application.Session.GetDefaultFolder(olFolderContacts)
    strFind = "[Full Name] = " & Chr(39) & ContactFullName & Chr(39)
    Set objContact = objFolder.Items.Find(strFind)
    Set GetContactLinkedToBirthday = objContact
    Set objContact = Nothing
    Set objFolder = Nothing
End Function

Private Function GetEmailAddressFromContact(Contact As Outlook.ContactItem) As String
    With Contact
        If .Email1Address <> vbNullString Then
            GetEmailAddressFromContact = .Email1Address
        ElseIf .Email2Address <> vbNullString Then
            GetEmailAddressFromContact = .Email2Address
        Else: GetEmailAddressFromContact = .Email3Address
        End If
    End With
End Function
