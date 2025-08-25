' ===================================================================
' Worksheet Event Handler for LaTeX Renderer
' ===================================================================

Option Explicit
Private WithEvents App As Application
' Initialize add-in when workbook opens
Private Sub Workbook_Open()
    ' Initialize global variables
    Call InitializeLaTeX
    ' Set up application-level events
    Set App = Application
    Application.StatusBar = "LaTeX Renderer Add-in loaded successfully!"
    ' Clear status bar after 3 seconds
    Application.OnTime Now + TimeSerial(0, 0, 3), "ClearStatusBar"
End Sub

' Event handler for when user changed cell content
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If GetLaTeXEnabled() And Not Target Is Nothing Then
        Call LaTeX(Target)
    End If
End Sub




