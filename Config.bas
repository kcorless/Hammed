Attribute VB_Name = "Config"
Option Explicit

' Defaults
Public Const DEFAULT_REPLY_MESSAGE As String = "I am sorry but the recipient of your email does not reply to unsolicited commercial emails..."
Public Const DEFAULT_SECRET_WORD As String = "rosebud"

' Debug and AUTO_REPLY settings
Public DONT_AUTO_SEND_REPLY_MODE As Boolean
Public DEBUG_MODE As Boolean
Public MSGBOX_MODE As Boolean

' === Dynamic File Location Resolution ===
Public Function GetAllFilesPath() As String
    Dim username As String
    username = Environ$("USERNAME")
    GetAllFilesPath = "D:\Users\" & username & "\AppData\Local\Hammed\"
End Function

Public Function GetWhitelistFilePath() As String
    GetWhitelistFilePath = GetAllFilesPath() & "whitelist.txt"
End Function

Public Function GetLogFilePath() As String
    GetLogFilePath = GetAllFilesPath() & "HamDebugLog.txt"
End Function

Public Function GetBackupLogFilePath() As String
    GetBackupLogFilePath = GetAllFilesPath() & "HamDebugLog2.txt"
End Function

Public Function GetReplyFilePath() As String
    GetReplyFilePath = GetAllFilesPath() & "reply.txt"
End Function

Public Function GetSecretFilePath() As String
    GetSecretFilePath = GetAllFilesPath() & "secret.txt"
End Function

Public Function GetConfigFilePath() As String
    GetConfigFilePath = GetAllFilesPath() & "Hammed.ini"
End Function

' === Property Getters to Expose Config Values ===
Public Property Get ShouldAutoSendReply() As Boolean
    ShouldAutoSendReply = Not DONT_AUTO_SEND_REPLY_MODE
End Property

Public Property Get IsDebugMode() As Boolean
    IsDebugMode = DEBUG_MODE
End Property

Public Property Get IsMsgBoxMode() As Boolean
    IsMsgBoxMode = MSGBOX_MODE
End Property

Public Property Get DefaultReplyMessage() As String
    DefaultReplyMessage = DEFAULT_REPLY_MESSAGE
End Property

Public Property Get DefaultSecretWord() As String
    DefaultSecretWord = DEFAULT_SECRET_WORD
End Property

Public Property Get ReplyFilePath() As String
    ReplyFilePath = GetReplyFilePath()
End Property

Public Property Get SecretFilePath() As String
    SecretFilePath = GetSecretFilePath()
End Property

' === Initialization ===
Public Sub InitializeProject()
    Logging.LogDebug "Project Initialization begins"
    LoadRuntimeConfig
    Logging.LogDebug "DONT_AUTO_SEND_REPLY_MODE = ", CStr(DONT_AUTO_SEND_REPLY_MODE)
    Logging.LogDebug "DEBUG_MODE = ", CStr(DEBUG_MODE)
    Logging.LogDebug "MSGBOX_MODE = ", CStr(MSGBOX_MODE)

    If MSGBOX_MODE Then
        MsgBox "initializing project", vbInformation
        MsgBox "path: " & GetAllFilesPath() & " " & GetLogFilePath()
    End If

    Logging.LogDebug "About to load whitelist"
    DomainWhitelist.LoadWhitelist
    Logging.LogDebug "Project Initialization completed"
End Sub

Public Sub LoadRuntimeConfig()
    Dim fileNum As Integer
    Dim line As String
    Dim key As String, value As String
    Dim sepPos As Long

    ' Set defaults
    DONT_AUTO_SEND_REPLY_MODE = True
    DEBUG_MODE = True
    MSGBOX_MODE = False

    If Dir(GetConfigFilePath()) = "" Then Exit Sub

    fileNum = FreeFile
    Open GetConfigFilePath() For Input As #fileNum

    Do Until EOF(fileNum)
        Line Input #fileNum, line
        line = Trim(line)

        If line <> "" And InStr(line, "=") > 0 Then
            sepPos = InStr(line, "=")
            key = Trim(Left(line, sepPos - 1))
            value = Trim(Mid(line, sepPos + 1))

            Select Case UCase(key)
                Case "DONT_AUTO_SEND_REPLY_MODE": DONT_AUTO_SEND_REPLY_MODE = (LCase(value) = "true")
                Case "DEBUG_MODE": DEBUG_MODE = (LCase(value) = "true")
                Case "MSGBOX_MODE": MSGBOX_MODE = (LCase(value) = "true")
            End Select
        End If
    Loop

    Close #fileNum
End Sub

