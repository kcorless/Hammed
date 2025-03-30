Attribute VB_Name = "EmailUtils"
' EmailUtils.bas
' Robust logic to extract SMTP sender address from Outlook.MailItem.
' Fallback order:
' 1. mail.Sender.Address (most common for internal)
' 2. mail.Sender.GetExchangeUser().PrimarySmtpAddress
' 3. PR_SMTP_ADDRESS via PropertyAccessor
' 4. Return-Path header from PR_TRANSPORT_HEADERS
' 5. From: header (if Return-Path missing)
' 6. SenderEmailAddress as final fallback

Option Explicit

Private Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
Private Const PR_TRANSPORT_HEADERS As String = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"

Public Function GetSenderSmtpAddress(mail As Outlook.mailItem) As String
    Dim senderEmail As String
    Dim pa As Outlook.PropertyAccessor
    Dim exchangeUser As Outlook.exchangeUser

    On Error Resume Next

    ' 1. Try mail.Sender.Address
    If Not mail.sender Is Nothing Then
        senderEmail = CStr(mail.sender.Address)
        If InStr(senderEmail, "@") > 0 Then
            Logging.LogDebug "Resolved via Sender.Address: ", senderEmail
            GetSenderSmtpAddress = senderEmail
            Exit Function
        Else
            Logging.LogDebug "Sender.Address did not contain @: ", senderEmail
        End If
    Else
        Logging.LogDebug "mail.Sender was Nothing when trying Sender.Address"
    End If

    ' 2. Try GetExchangeUser
    Set exchangeUser = mail.sender.GetExchangeUser
    If Not exchangeUser Is Nothing Then
        senderEmail = exchangeUser.PrimarySmtpAddress
        If Len(senderEmail) > 0 Then
            Logging.LogDebug "Resolved via GetExchangeUser: ", senderEmail
            GetSenderSmtpAddress = senderEmail
            Exit Function
        Else
            Logging.LogDebug "ExchangeUser object found but PrimarySmtpAddress was blank"
        End If
    Else
        Logging.LogDebug "GetExchangeUser returned Nothing"
    End If

    ' 3. Try PR_SMTP_ADDRESS via PropertyAccessor
    Set pa = mail.PropertyAccessor
    senderEmail = pa.GetProperty(PR_SMTP_ADDRESS)
    If Len(senderEmail) > 0 And InStr(senderEmail, "@") > 0 Then
        Logging.LogDebug "Resolved via PropertyAccessor: ", senderEmail
        GetSenderSmtpAddress = senderEmail
        Exit Function
    Else
        Logging.LogDebug "PropertyAccessor returned blank or invalid SMTP address"
    End If

    ' 4. Try Return-Path from headers
    senderEmail = GetSmtpAddressFromHeaders(mail)
    If Len(senderEmail) > 0 Then
        Logging.LogDebug "Resolved via headers: ", senderEmail
        GetSenderSmtpAddress = senderEmail
        Exit Function
    Else
        Logging.LogDebug "No valid SMTP address found in headers"
    End If

    ' 5. Final fallback: SenderEmailAddress
    senderEmail = mail.SenderEmailAddress
    Logging.LogDebug "All methods failed. Using SenderEmailAddress: ", senderEmail
    GetSenderSmtpAddress = senderEmail
End Function

Private Function GetSmtpAddressFromHeaders(mail As Outlook.mailItem) As String
    Dim headers As String
    Dim addr As String

    On Error Resume Next
    headers = mail.PropertyAccessor.GetProperty(PR_TRANSPORT_HEADERS)
    On Error GoTo 0

    If Len(headers) = 0 Then
        Logging.LogDebug "PR_TRANSPORT_HEADERS returned empty"
        Exit Function
    End If

    ' Try Return-Path
    addr = ExtractAddressFromHeader(headers, "Return-Path:")
    If Len(addr) > 0 Then
        Logging.LogDebug "Extracted SMTP from Return-Path: ", addr
        GetSmtpAddressFromHeaders = addr
        Exit Function
    End If

    ' Try From
    addr = ExtractAddressFromHeader(headers, "From:")
    If Len(addr) > 0 Then
        Logging.LogDebug "Extracted SMTP from From: header: ", addr
        GetSmtpAddressFromHeaders = addr
        Exit Function
    End If
End Function

Private Function ExtractAddressFromHeader(headers As String, headerName As String) As String
    Dim startPos As Long, endPos As Long
    Dim lineEnd As Long, headerLine As String
    Dim addressCandidate As String

    startPos = InStr(1, headers, headerName, vbTextCompare)
    If startPos = 0 Then Exit Function

    lineEnd = InStr(startPos, headers, vbCrLf)
    If lineEnd = 0 Then lineEnd = Len(headers) + 1
    headerLine = Mid(headers, startPos, lineEnd - startPos)

    ' If <...> brackets exist, extract inside
    If InStr(headerLine, "<") > 0 And InStr(headerLine, ">") > InStr(headerLine, "<") Then
        startPos = InStr(headerLine, "<") + 1
        endPos = InStr(headerLine, ">")
        addressCandidate = Mid(headerLine, startPos, endPos - startPos)
    Else
        ' No angle brackets — get everything after colon
        startPos = InStr(headerLine, ":") + 1
        addressCandidate = Trim(Mid(headerLine, startPos))
    End If

    ' Clean it
    addressCandidate = Replace(addressCandidate, vbCr, "")
    addressCandidate = Replace(addressCandidate, vbLf, "")
    addressCandidate = Replace(addressCandidate, ",", "")

    ExtractAddressFromHeader = Trim(addressCandidate)
End Function

