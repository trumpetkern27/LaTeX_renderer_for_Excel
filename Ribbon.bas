' Custom Ribbon/Menu Integration
Sub CreateLaTeXMenu()
    Dim menuBar As CommandBar
    Dim newMenu As CommandBarControl
    Dim menuItem As CommandBarControl
    
    ' Create custom menu
    Set menuBar = Application.CommandBars("Worksheet Menu Bar")
    Set newMenu = menuBar.Controls.Add(Type:=msoControlPopup, Before:=menuBar.Controls.count)
    newMenu.Caption = "LaTeX"
    
    ' Add toggle option
    Set menuItem = newMenu.Controls.Add(Type:=msoControlButton)
    menuItem.Caption = "Toggle LaTeX Renderer"
    menuItem.OnAction = "ToggleLaTeXRenderer"
End Sub

Sub RemoveLaTeXMenu()
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls("LaTeX").Delete
    On Error GoTo 0
End Sub


