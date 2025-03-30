Attribute VB_Name = "AddWL"
Public Sub AddSenderDomaintoWL()
    Dim objExplorer As Outlook.Explorer
    Dim objMail As Object
    Dim senderAddress As String
    Dim domain As String
    Dim filePath As String
    Dim fileNum As Integer

    Set objExplorer = Application.ActiveExplorer

    If objExplorer.Selection.Count <> 1 Then
        MsgBox "Please select exactly one email item.", vbExclamation
        Exit Sub
    End If

    Set objMail = objExplorer.Selection(1)

    ' Confirm selection is a MailItem
    If objMail.Class <> olMail Then
        MsgBox "Selected item is not an email.", vbExclamation
        Exit Sub
    End If

    ' Reliably retrieve sender SMTP address
    senderAddress = GetSenderSMTP(objMail)

    If senderAddress = "" Then
        MsgBox "Could not determine sender email address.", vbExclamation
        Exit Sub
    End If

    ' Extract domain from the sender email
    If InStr(senderAddress, "@") > 0 Then
        domain = Split(senderAddress, "@")(1)
    Else
        MsgBox "The sender address retrieved is invalid: " & senderAddress, vbExclamation
        Exit Sub
    End If

    ' Update path to your existing text file
    filePath = "D:\Users\kcorless\AppData\Local\Hammed\whitelist.txt"

    fileNum = FreeFile
    Open filePath For Append As #fileNum
    Print #fileNum, domain
    Close #fileNum

    MsgBox "Domain '" & domain & "' has been added to file.", vbInformation
End Sub

' Improved reliable helper function
Private Function GetSenderSMTP(mail As Outlook.mailItem) As String
    Dim sender As Outlook.AddressEntry
    Dim exchUser As Outlook.exchangeUser

    On Error GoTo fallbackMethod

    Set sender = mail.sender

    If sender Is Nothing Then GoTo fallbackMethod

    Select Case sender.AddressEntryUserType
        Case olExchangeUserAddressEntry, olExchangeRemoteUserAddressEntry
            Set exchUser = sender.GetExchangeUser
            If Not exchUser Is Nothing Then
                GetSenderSMTP = exchUser.PrimarySmtpAddress
            Else
                GetSenderSMTP = ""
            End If
        Case Else
            ' Standard SMTP address
            GetSenderSMTP = sender.Address
    End Select

    Exit Function

fallbackMethod:
    ' Fallback simple method
    GetSenderSMTP = mail.SenderEmailAddress
End Function

