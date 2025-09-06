Attribute VB_Name = "Globals"
' === LaTeX Globals ===
Option Explicit

' Initialize the global variable
Public Sub InitializeLaTeX()
    On Error Resume Next
    Dim current As String
    current = GetSetting("LaTeXRenderer", "Settings", "Enabled", vbNullString)
    If current = vbNullString Then
        SetLaTeXEnabled True
    End If
End Sub
' LaTeX enablement
Public Function GetLaTeXEnabled() As Boolean
    On Error GoTo ErrorHandler
    Dim s As String
    s = GetSetting("LaTeXRenderer", "Settings", "Enabled", "True")
    GetLaTeXEnabled = (LCase(s) = "true")
    Exit Function
    
ErrorHandler:
    ' Default to True if there's any error
    GetLaTeXEnabled = True
    Err.Clear
End Function
Public Function SetLaTeXEnabled(value As Boolean)
    On Error Resume Next
    SaveSetting "LaTeXRenderer", "Settings", "Enabled", CStr(value)
End Function



' Runtime Timeout Management
Public Function GetTimeoutSeconds() As Double
    On Error GoTo ErrorHandler
    Dim s As String
    s = GetSetting("LaTeXRenderer", "Settings", "Timeout", "60") ' Default = 60s
    If IsNumeric(s) Then
        GetTimeoutSeconds = CDbl(s)
    Else
        GetTimeoutSeconds = 60
    End If
    Exit Function
    
ErrorHandler:
    GetTimeoutSeconds = 60 ' Safe default
    Err.Clear
End Function

Public Sub SetTimeoutSeconds(val As Double)
    On Error Resume Next
    SaveSetting "LaTeXRenderer", "Settings", "Timeout", CStr(val)
End Sub

Public Sub SetLaTeXTimeout()
    On Error Resume Next
    Dim userInput As String
    Dim newTimeout As Double
    
    userInput = InputBox("Enter timeout in seconds: ", "Set LaTeX Timeout", GetTimeoutSeconds())
    
    If Trim(userInput) = vbNullString Then Exit Sub ' User canceled
    
    If IsNumeric(userInput) Then
        newTimeout = CDbl(userInput)
        If newTimeout > 0 Then
            SetTimeoutSeconds newTimeout
            MsgBox "Timeout updated to " & newTimeout & " seconds.", vbInformation
        Else
            MsgBox "Please enter a positive number.", vbCritical
        End If
    Else
        MsgBox "Please enter a valid number.", vbCritical
    End If
End Sub
