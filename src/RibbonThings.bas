Attribute VB_Name = "RibbonThings"
' === Ribbon Globals ===
Dim g_ribbon As IRibbonUI

' Load ribbon
Public Sub OnRibbonLoad(Ribbon As IRibbonUI)
    On Error Resume Next
    Set g_ribbon = Ribbon
    Call InitializeLaTeX
End Sub


' Get renderer state
Public Function GetRendererState(control As IRibbonControl) As Boolean
    On Error GoTo ErrHandler
    GetRendererState = GetLaTeXEnabled()
    Exit Function
    
ErrHandler:
    ' Return safe default
    GetRendererState = True
    Err.Clear
End Function

' Toggle LaTeX rendering on/off
Public Sub ToggleLaTeXRenderer(control As IRibbonControl, pressed As Boolean)
    On Error Resume Next
    SetLaTeXEnabled pressed
    MsgBox "LaTeX Rendering is now " & IIf(pressed, "ENABLED", "DISABLED"), vbInformation
End Sub

' Timeout Handling
Public Function GetTimeoutText(control As IRibbonControl) As String
    On Error GoTo ErrHandler
    GetTimeoutText = CStr(GetTimeoutSeconds())
    Exit Function
    
ErrHandler:
    ' Return safe default
    GetTimeoutText = "60"
    Err.Clear
End Function

Public Sub OnTimeoutChange(control As IRibbonControl, text As String)
    On Error Resume Next
    If IsNumeric(text) Then
        Dim val As Double
        val = CDbl(text)
        If val > 0 Then
            SetTimeoutSeconds val
        Else
            MsgBox "Please enter a positive number.", vbCritical
        End If
    Else
        MsgBox "Please enter a numeric timeout value.", vbExclamation
    End If
End Sub
