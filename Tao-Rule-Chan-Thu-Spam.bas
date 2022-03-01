Option Explicit

Public Sub CreateMoveMailsRule()
    Dim objStore As Outlook.Store
    Dim objRule As Outlook.Rule
    Dim colRules As Outlook.Rules
    Dim objMoveAction As Outlook.MoveOrCopyRuleAction
    Dim objFromCondition As Outlook.ToOrFromRuleCondition
    Dim objJunkFldr As Outlook.Folder
    Dim objSel As Outlook.Selection
    Dim objMail As Outlook.MailItem
    Dim i, j As Long
    Dim arrRecipients()
    Set objStore = Application.Session.DefaultStore
    Set objJunkFldr = objStore.GetDefaultFolder(olFolderJunk)
    Set colRules = objStore.GetRules
    Set objSel = Application.ActiveExplorer.Selection
    If objSel.Count > 0 Then
        If Not RuleExists("ProcessSpamMails") Then
            Set objRule = colRules.Create("ProcessSpamMails", olRuleReceive)
        Else: Set objRule = colRules.Item("ProcessSpamMails")
        End If
        For i = 1 To objSel.Count
            If TypeOf objSel.Item(i) Is Outlook.MailItem Then
                Set objMail = objSel.Item(i)
                With objRule
                    Set objFromCondition = .Conditions.From
                    With objFromCondition
                        .Enabled = True
                        If .Recipients.Item(objMail.SenderEmailAddress) Is Nothing Then
                            .Recipients.Add objMail.SenderEmailAddress
                            .Recipients.ResolveAll
                            j = j + 1
                            ReDim Preserve arrRecipients(1 To j)
                            arrRecipients(j) = objMail.SenderEmailAddress
                        End If
                    End With
                    Set objMoveAction = .Actions.MoveToFolder
                    With objMoveAction
                        .Enabled = True
                        .Folder = objJunkFldr
                    End With
                End With
                colRules.Save
            End If
        Next
        objRule.Execute
    Else
        MsgBox "Ban chua chon thu nao de them vao Rule loc thu spam.", vbExclamation, "Ban chua chon thu nao"
        Exit Sub
    End If
    MsgBox "Da them nhung nguoi gui co dia chi email: " & vbCrLf & Join(arrRecipients) & vbCrLf & "vao Rule: " & objRule.Name, vbInformation, "Thao tac hoan tat"
    Set objStore = Nothing
    Set objRule = Nothing
    Set colRules = Nothing
    Set objMoveAction = Nothing
    Set objFromCondition = Nothing
    Set objJunkFldr = Nothing
    Set objSel = Nothing
    Set objMail = Nothing
End Sub

Public Function RuleExists(RuleName As String) As Boolean
    Dim objRule As Outlook.Rule
    Dim colRules As Outlook.Rules
    Dim i As Long
    RuleExists = False
    Set colRules = Application.Session.DefaultStore.GetRules
    If colRules.Count > 0 Then
        For i = 1 To colRules.Count
            Set objRule = colRules.Item(i)
            If objRule.Name = RuleName Then
                RuleExists = True
                Exit For
            Else: RuleExists = False
            End If
        Next
    End If
    Set objRule = Nothing
    Set colRules = Nothing
End Function