Attribute VB_Name = "DomainWhitelist"
Option Explicit

Public Whitelist As Object
Private WhitelistFileSize As Long

' Load Whitelist from file
Public Sub LoadWhitelist()
    Dim line As String, fileNum As Integer
    Set Whitelist = CreateObject("Scripting.Dictionary")

    If Dir(Config.GetWhitelistFilePath()) = "" Then Exit Sub

    fileNum = FreeFile()
    Open Config.GetWhitelistFilePath() For Input As #fileNum
    Do Until EOF(fileNum)
        Line Input #fileNum, line
        line = LCase(Trim(line))
        If line <> "" Then Whitelist(line) = True
    Loop
    Close #fileNum

    WhitelistFileSize = FileLen(Config.GetWhitelistFilePath())
    Logging.LogDebug "Whitelist Loaded", "Entries", Whitelist.Count
End Sub

' Check if domain is whitelisted
Public Function IsWhitelisted(domain As String) As Boolean
    Dim item As Variant

    Logging.LogDebug "checking whitelist for domain " & domain

    ' Convert the domain to lowercase for case-insensitive comparison
    domain = LCase(domain)

    ' Iterate through each item in the whitelist collection
    For Each item In Whitelist
        ' Check if the domain is a substring of the whitelist item
        ' Logging.LogDebug "checking domain " & domain & " against whitelist item " & item
        If InStr(domain, LCase(item)) > 0 Then
            IsWhitelisted = True
            Logging.LogDebug "domain: ", domain, " found in whitelist.  Allowing email"
            Exit Function
        End If
    Next item

    ' If no match is found, return False
    IsWhitelisted = False
    Logging.LogDebug "domain: ", domain, " Not found in whitelist.  Beginning Hammed process"
End Function

' Add domain to whitelist
Public Sub AddDomain(domain As String)
    domain = LCase(domain)
    If Not Whitelist.Exists(domain) Then
        Whitelist.Add domain, True
        Dim fileNum As Integer: fileNum = FreeFile
        Open Config.GetWhitelistFilePath() For Append As #fileNum
        Print #fileNum, domain
        Close #fileNum
        Logging.LogDebug "Domain added", "Domain", domain
    End If
End Sub

' Check for external edits to whitelist file
Public Function WhitelistChanged() As Boolean
    WhitelistChanged = (FileLen(Config.GetWhitelistFilePath()) <> WhitelistFileSize)
End Function

' Extract domain from email
Public Function ExtractDomain(emailAddress As String) As String
    Dim atPos As Long
    atPos = InStr(emailAddress, "@")
    If atPos > 0 Then
        ExtractDomain = LCase(Mid(emailAddress, atPos + 1))
    Else
        ExtractDomain = ""
    End If
End Function

