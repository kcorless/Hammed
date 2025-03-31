VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private WithEvents InboxItems As Outlook.Items
Attribute InboxItems.VB_VarHelpID = -1

Private Sub Application_Startup()
    ' Set up event handling for inbox items as they arrive in the Inbox
    Set InboxItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Items

    Config.InitializeProject
    Logging.LogDebug "ThisOutlookSession has been initiated and is waiting for new arrivals to the Inbox"
End Sub

Private Sub InboxItems_ItemAdd(ByVal item As Object)
    On Error GoTo ErrorHandler

    Dim mail As Outlook.mailItem
    Dim senderEmail As String
    Dim senderDomain As String
    Dim replyMessage As String

    ' Only process mail items
    If Not TypeOf item Is Outlook.mailItem Then Exit Sub
    Set mail = item

    ' Skip items in Junk Email
    If mail.Parent.Name = "Junk E-mail" Then Exit Sub

    ' Reload whitelist if updated
    If DomainWhitelist.WhitelistChanged Then
        DomainWhitelist.LoadWhitelist
    End If

    ' Extract sender email and domain
    senderEmail = EmailUtils.GetSenderSmtpAddress(mail)
    senderDomain = DomainWhitelist.ExtractDomain(senderEmail)

    Logging.LogDebug "Checking message from ", senderEmail, " with senderDomain ", senderDomain

    ' Skip if whitelisted or contains secret word
    If DomainWhitelist.IsWhitelisted(senderDomain) Or EmailHandler.IsSecretWordPresent(mail) Then
        Logging.LogDebug "Email is untouched because domain ", senderDomain, " is whitelisted or secret word is present"
        Exit Sub
    End If

    ' Move message to Hammed folder
    Set mail = EmailHandler.MoveMail(mail, "Hammed")
    mail.UnRead = False
    mail.Save
    
    
    ' Skip reply if sender address begins with noreply or no-reply
    If LCase(Left(senderEmail, 7)) = "noreply" Or LCase(Left(senderEmail, 8)) = "no-reply" Then
        Logging.LogDebug "No reply drafted or sent because senderEmail begins with noreply/no-reply: ", senderEmail
        Exit Sub
    End If

    ' Create reply message
    Dim replyMail As Outlook.mailItem
    Set replyMail = mail.reply
    replyMessage = EmailHandler.GetReplyMessage()
    replyMail.Body = replyMessage & mail.Body
    replyMail.Subject = "Re: " & mail.Subject

    If Config.ShouldAutoSendReply Then
        replyMail.Send
        Logging.LogDebug "Auto-reply sent to: ", senderEmail
    Else
        replyMail.Save
        Logging.LogDebug "DONT_AUTO_SEND_REPLY_MODE is TRUE, reply saved to Drafts"
    End If

ExitPoint:
    Exit Sub

ErrorHandler:
    Logging.LogDebug "Error in InboxItems_ItemAdd", "Error", Err.Description
    Resume ExitPoint
End Sub

