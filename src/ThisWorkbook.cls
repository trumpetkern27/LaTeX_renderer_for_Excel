VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ===================================================================
' Worksheet Event Handler for LaTeX Renderer
' ===================================================================

Option Explicit
Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Private Sub Workbook_Open()
    ' Set up application-level events
    Set App = Application
    Application.StatusBar = "LaTeX Renderer Add-in loaded successfully!"
    ' Clear status bar after 3 seconds
    Application.OnTime Now + TimeSerial(0, 0, 3), "ClearStatusBar"
End Sub

' Event handler for when user changed cell content
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If GetLaTeXEnabled() And _
        Not Target Is Nothing And _
        ((Target.Columns.count = 1 And Target.Rows.count = 1) Or _
        Target.MergeCells) Then
        Call LaTeX(Target)
    End If
End Sub



