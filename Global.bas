' === LaTeXGlobals Module ===
Option Explicit

' Global variable to track if add-in is enabled
Public Function GetLaTeXEnabled() As Boolean
    Dim s As String
    s = GetSetting("LaTeXRenderer", "Settings", "Enabled", "True")
    GetLaTeXEnabled = (LCase(s) = "true")
End Function
Public Function SetLaTeXEnabled(value As Boolean)
    SaveSetting "LaTeXRenderer", "Settings", "Enabled", CStr(value)
End Function

' Initialize the global variable
Public Sub InitializeLaTeX()
    Dim current As String
    current = GetSetting("LaTeXRenderer", "Settings", "Enabled", "")
    If current = "" Then
        SetLaTeXEnabled True
    End If
End Sub
' Toggle LaTeX rendering on/off
Public Sub ToggleLaTeXRenderer()
    Dim currentState As Boolean
    currentState = GetLaTeXEnabled()
    Call SetLaTeXEnabled(Not currentState)
    MsgBox "LaTeX Rendering is now " & IIf(GetLaTeXEnabled(), "ENABLED", "DISABLED"), vbInformation
End Sub


