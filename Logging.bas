Attribute VB_Name = "Logging"
' Logging.bas
Option Explicit

Public Sub LogDebug(ParamArray parts() As Variant)
    Dim message As String
    Dim i As Integer
    For i = LBound(parts) To UBound(parts)
        message = message & CStr(parts(i)) & " "
    Next i
    message = Trim(message)
    WriteLog GetLogFilePath(), message
End Sub

Private Sub WriteLog(ByVal filePath As String, ByVal message As String)
    Dim fileNum As Integer
    On Error GoTo BackupLog

    fileNum = FreeFile
    Open filePath For Append As #fileNum
    Print #fileNum, Now & " - " & message
    Close #fileNum
    Exit Sub

BackupLog:
    On Error GoTo FinalFail
    ' If writing to primary log fails, try writing to backup
    fileNum = FreeFile
    Open GetBackupLogFilePath() For Append As #fileNum
    Print #fileNum, Now & " - [BackupLog] " & message
    Close #fileNum
    Exit Sub

FinalFail:
    MsgBox "Critical Error: Unable to write to both primary and backup log files.", vbCritical, "Logging Failure"
    MsgBox "check to make sure this directory exists " & GetLogFilePath()
End Sub
