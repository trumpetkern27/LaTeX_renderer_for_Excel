Attribute VB_Name = "renderLaTeX"
' ===================================================================
' Excel LaTeX Renderer Add-in
' Author: Zach Kern
' Description: Automatically renders LaTeX mathematical expressions in Excel cells
' Usage: Type LaTeX expressions between $ symbols (e.g., $\alpha + \beta^2$)
' ===================================================================

Option Explicit

' Clear status bar message
Public Sub ClearStatusBar()
    Application.StatusBar = False
End Sub


' Main LaTeX processing subroutine - called from worksheet change events
Public Sub LaTeX(ByVal Target As Range)
    
    Dim cell As Range
    Dim txt As String
    Dim startPos As Long, endPos As Long
    Dim mathBlock As String
    Dim rendered As String
    Dim iOffset As Long
    Dim t As Double: t = Timer
    Dim timeout As Double: timeout = GetTimeoutSeconds
    Dim mathChars As Object: Set mathChars = CreateObject("Scripting.Dictionary")
    mathChars.Add "^", Empty
    mathChars.Add "_", Empty
    mathChars.Add "\", Empty
    Dim lit As Variant
    lit = Array("å", "_", "æ", "$", "ç", "\", "ð", "^", "ò", "{", "õ", "}")
    Dim valsToReplace As Object: Set valsToReplace = CreateObject("Scripting.Dictionary")
    Dim replaceVals As Object: Set replaceVals = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 0 To 10 Step 2
        replaceVals.Add lit(i), lit(i + 1)
        valsToReplace.Add lit(i + 1), lit(i)
    Next i
    
    ' Error handling and disable events to prevent recursion
    On Error GoTo ErrHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' Process each cell in the target range
    For Each cell In Target.Cells
        If Not IsEmpty(cell.value) And IsText(cell.value) And Not Left(cell.value, 1) = "=" Then
            txt = CStr(cell.value)
            iOffset = 0

            ' Find and process all LaTeX expressions (delimited by $)
            Do While InStr(txt, "$") > 0
                If Timer - t >= timeout Then Err.Raise 1 ' Runtime error
                ' Find opening delimiter
                startPos = InStr(iOffset + 1, txt, "$")
                If startPos = 0 Then Exit Do
                
                ' Handle $ within \text
                Dim replaceEsc As Long
                If InStr(startPos + 1, txt, "\text{") > 0 Then
                    Dim startText As Long: startText = InStr(startPos, txt, "\text{") + 6
                    If InStr(startText, txt, "}") > 0 Then
                        Dim endText As Long: endText = InStr(startText, txt, "}")
                        replaceEsc = InStr(startText, txt, "$")
                        Do While replaceEsc > 0 And replaceEsc < endText
                            If Timer - t >= timeout Then Err.Raise 1
                            txt = Left(txt, replaceEsc - 1) & "æ" & Mid(txt, replaceEsc + 1)
                            cell.Characters(replaceEsc, 1).text = "æ"
                            replaceEsc = InStr(startText, txt, "$")
                        Loop
                    End If
                End If

                ' Find closing delimiter
                endPos = InStr(startPos + 1, txt, "$")
                If endPos < 2 Then Exit Do
                Do While Mid(txt, endPos - 1, 1) = "\" And Mid(txt, endPos - 2, 1) <> "\"
                    If Timer - t >= timeout Then Err.Raise 1
                    endPos = InStr(endPos + 1, txt, "$")
                Loop
                
                
                ' Check if this looks like actual LaTeX (contains LaTeX syntax)
                If InStr(Mid(txt, startPos, endPos - startPos + 1), "\") = 0 And _
                   InStr(Mid(txt, startPos, endPos - startPos + 1), "^") = 0 And _
                   InStr(Mid(txt, startPos, endPos - startPos + 1), "_") = 0 And _
                   InStr(Mid(txt, startPos, endPos - startPos + 1), "æ") = 0 And _
                   InStr(Mid(txt, startPos, endPos - startPos + 1), "ç") = 0 Then
                   iOffset = iOffset + 1
                   GoTo NextLoop
                Else
                    'Remove $ for valid block
                    cell.Characters(startPos, 1).text = vbNullString
                    endPos = endPos - 1
                    cell.Characters(endPos, 1).text = vbNullString
                    txt = cell.value
                End If
                
                ' Handle escaped characters (\$, \\, \{, \}, \^, \_)
                For i = startPos + 1 To endPos
                    If valsToReplace.exists(Mid(txt, i, 1)) And Mid(txt, i - 1, 1) = "\" Then
                        cell.Characters(i - 1, 2).text = valsToReplace(Mid(txt, i, 1))
                        txt = cell.value
                        endPos = endPos - 1
                    End If
                Next i
                
                Dim endBlock As Long
                ' Extract LaTeX code and render it
                Do
                    If Timer - t >= timeout Then Err.Raise 1
                    For i = startPos + 1 To endPos
                        endBlock = i
                        If mathChars.exists(CStr(Mid(txt, i, 1))) Then Exit For
                    Next i
                    If endBlock > endPos Or startPos >= endBlock Then Exit Do
                    mathBlock = Mid(txt, startPos, endBlock - startPos)
                    rendered = RenderLatexSimple(mathBlock)
                    Dim lastColor As Long: lastColor = cell.Characters(startPos, endBlock - startPos).Font.Color
    
                    ' Replace LaTeX with rendered Unicode, preserving formatting
                    With cell
                        .Characters(startPos, endBlock - startPos).text = rendered
                        .Characters(startPos, Len(rendered)).Font.Color = lastColor
                    End With
                    txt = cell.value
                    startPos = startPos + Len(rendered)
                    endPos = endPos - Len(mathBlock) + Len(rendered)
                    iOffset = startPos - 1
                Loop
                Dim endBrace As Long
                
                ' Process superscripts (^{...})
                Do
                    If Timer - t >= timeout Then Err.Raise 1
                    i = InStr(txt, "^{")
                    If i = 0 Then Exit Do
                    endBrace = ClosingBrace(txt, i + 2)
                    If endBrace = 0 Then Exit Do
                    ' Remove subscripts within superscripts
                    Do While InStr(1, Mid(txt, i + 2, endBrace - i - 2), "_") > 0
                        txt = Left(txt, i + InStr(1, Mid(txt, i + 2, endBrace - i - 2), "_")) & _
                            "^" & _
                            Mid(txt, i + 2 + InStr(1, Mid(txt, i + 2, endBrace - i - 2), "_"))
                    Loop
                    With cell
                        .Characters(i, endBrace - i + 1).text = Mid(txt, i + 2, endBrace - i - 2)
                        .Characters(i, endBrace - i - 2).Font.Superscript = True
                    End With
                    txt = cell.value
                Loop
                
                ' Process subscripts (_{...})
                Do
                    If Timer - t >= timeout Then Err.Raise 1
                    i = InStr(txt, "_{")
                    If i = 0 Then Exit Do
                    endBrace = ClosingBrace(txt, i + 2)
                    If endBrace = 0 Then Exit Do
                    With cell
                        .Characters(i, endBrace - i + 1).text = Mid(txt, i + 2, endBrace - i - 2)
                        .Characters(i, endBrace - i - 2).Font.Subscript = True
                    End With
                    txt = cell.value
                Loop

NextLoop:
            Loop
           
            ' Restore escaped characters
            txt = cell.value
            Dim escPos As Long
            For Each lit In replaceVals.keys
                Do While InStr(txt, lit) > 0
                    escPos = InStr(txt, lit)
                    cell.Characters(escPos, 1).text = replaceVals(lit)
                    txt = cell.value
                Loop
            Next lit
        End If
    Next cell
    GoTo ExitClean
ErrHandler:
    Select Case Err.Number
        Case 1004: MsgBox "Error processing LaTeX expression. Please check syntax.", vbCritical
        Case 94: MsgBox "Error: only one colour permitted per command.", vbCritical
        Case 9: MsgBox "Invalid character range detected.", vbCritical
        Case 1: MsgBox "Runtime error.", vbCritical
        Case Else: MsgBox "Unexpected error: " & Err.Description, vbCritical
    End Select
ExitClean:
    ' Always restore Excel's event handling and screen updating
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Core LaTeX rendering function - converts LaTeX symbols to Unicode
Private Function RenderLatexSimple(ByVal txt As String) As String
    Dim out As String
    out = txt
    
    ' Handle mathematical alphabets
    out = ProcessMathAlphabets(out)
    
    'Replace LaTeX symbols
    out = Replace(out, "\zeta", ChrW(&H3B6))
    out = Replace(out, "\Zeta", ChrW(&H396))
    
    out = Replace(out, "\yen", ChrW(&HA5))
    
    out = Replace(out, "\xi", ChrW(&H3BE))
    out = Replace(out, "\Xi", ChrW(&H39E))
    
    out = Replace(out, "\wp", ChrW(&H2118))
    out = Replace(out, "\wr", ChrW(&H2240))
    out = Replace(out, "\wedge", ChrW(&H2227))
    
    out = Replace(out, "\vee", ChrW(&H2228))
    out = Replace(out, "\vdots", ChrW(&H22EE))
    out = Replace(out, "\vdash", ChrW(&H22A2))
    out = Replace(out, "\vDash", ChrW(&H22A8))
    out = Replace(out, "\Vdash", ChrW(&H22A9))
    out = Replace(out, "\Vvdash", ChrW(&H22AA))
    out = Replace(out, "\vartriangleright", ChrW(&H22B3))
    out = Replace(out, "\vartriangleleft", ChrW(&H22B2))
    out = Replace(out, "\varsigma", ChrW(&H3C2))
    out = Replace(out, "\varrho", ChrW(&H3F1))
    out = Replace(out, "\varpi", ChrW(&H3D6))
    out = Replace(out, "\varkappa", ChrW(&H3F0))
    out = Replace(out, "\vartheta", ChrW(&H3D1))
    out = Replace(out, "\varphi", ChrW(&H3D5))
    out = Replace(out, "\varepsilon", ChrW(&H3F5))
    out = Replace(out, "\varnothing", ChrW(&H2205))
    
    out = Replace(out, "\upsilon", ChrW(&H3C5))
    out = Replace(out, "\Upsilon", ChrW(&H3A5))
    out = Replace(out, "\uparrow", ChrW(&H2191))
    out = Replace(out, "\Uparrow", ChrW(&H21D1))
    out = Replace(out, "\updownarrow", ChrW(&H2195))
    out = Replace(out, "\Updownarrow", ChrW(&H21D5))
    out = Replace(out, "\uplus", ChrW(&H228E))

    out = Replace(out, "\twoheadrightarrow", ChrW(&H21A0))
    out = Replace(out, "\triangleright", ChrW(&H25B7))
    out = Replace(out, "\triangleleft", ChrW(&H25C1))
    out = Replace(out, "\triangledown", ChrW(&H25BD))
    out = Replace(out, "\triangle", ChrW(&H25B3))
    out = Replace(out, "\top", ChrW(&H22A4))
    out = Replace(out, "\to", ChrW(&H2192))
    out = Replace(out, "\times", ChrW(&HD7))
    out = Replace(out, "\therefore", ChrW(&H2234))
    out = Replace(out, "\theta", ChrW(&H3B8))
    out = Replace(out, "\Theta", ChrW(&H398))
    out = Replace(out, "\thicksim", ChrW(&H223C))
    out = Replace(out, "\thickapprox", ChrW(&H2248))
    out = Replace(out, "\tau", ChrW(&H3C4))
    out = Replace(out, "\Tau", ChrW(&H3A4))
    
    out = Replace(out, "\swarrow", ChrW(&H2199))
    out = Replace(out, "\sum", ChrW(&H2211))
    out = Replace(out, "\supseteq", ChrW(&H2287))
    out = Replace(out, "\supset", ChrW(&H2283))
    out = Replace(out, "\surd", ChrW(&H221A))
    out = Replace(out, "\subseteq", ChrW(&H2286))
    out = Replace(out, "\subset", ChrW(&H2282))
    out = Replace(out, "\succsim", ChrW(&H227F))
    out = Replace(out, "\succeq", ChrW(&H227D))
    out = Replace(out, "\succ", ChrW(&H227B))
    out = Replace(out, "\star", ChrW(&H22C6))
    out = Replace(out, "\square", ChrW(&H25A1))
    out = Replace(out, "\sqrt", ChrW(&H221A))
    out = Replace(out, "\sqcup", ChrW(&H2294))
    out = Replace(out, "\sqcap", ChrW(&H2293))
    out = Replace(out, "\spadesuit", ChrW(&H2660))
    out = Replace(out, "\smile", ChrW(&H2323))
    out = Replace(out, "\smallsmile", ChrW(&H2323))
    out = Replace(out, "\smallsetminus", ChrW(&H2216))
    out = Replace(out, "\smallfrown", ChrW(&H2322))
    out = Replace(out, "\simeq", ChrW(&H2243))
    out = Replace(out, "\sim", ChrW(&H223C))
    out = Replace(out, "\sigma", ChrW(&H3C3))
    out = Replace(out, "\Sigma", ChrW(&H3A3))
    out = Replace(out, "\sharp", ChrW(&H266F))
    out = Replace(out, "\setminus", ChrW(&H2216))
    out = Replace(out, "\searrow", ChrW(&H2198))
    
    out = Replace(out, "\rtimes", ChrW(&H22CA))
    out = Replace(out, "\rightsquigarrow", ChrW(&H21DD))
    out = Replace(out, "\rightleftharpoons", ChrW(&H21CC))
    out = Replace(out, "\rightharpoonup", ChrW(&H21C0))
    out = Replace(out, "\rightharpoondown", ChrW(&H21C1))
    out = Replace(out, "\rightarrow", ChrW(&H2192))
    out = Replace(out, "\Rightarrow", ChrW(&H21D2))
    out = Replace(out, "\rho", ChrW(&H3C1))
    out = Replace(out, "\Rho", ChrW(&H3A1))
    out = Replace(out, "\rfloor", ChrW(&H230B))
    out = Replace(out, "\Re", ChrW(&H211C))
    out = Replace(out, "\rceil", ChrW(&H2309))
    out = Replace(out, "\rbrace", ChrW(&H7D))
    out = Replace(out, "\rangle", ChrW(&H27E9))

    out = Replace(out, "\psi", ChrW(&H3C8))
    out = Replace(out, "\Psi", ChrW(&H3A8))
    out = Replace(out, "\propto", ChrW(&H221D))
    out = Replace(out, "\prod", ChrW(&H220F))
    out = Replace(out, "\prime", ChrW(&H2032))
    out = Replace(out, "\precsim", ChrW(&H227E))
    out = Replace(out, "\preceq", ChrW(&H2AAF))
    out = Replace(out, "\prec", ChrW(&H227A))
    out = Replace(out, "\pm", ChrW(&HB1))
    out = Replace(out, "\pitchfork", ChrW(&H22D4))
    out = Replace(out, "\pi", ChrW(&H3C0))
    out = Replace(out, "\Pi", ChrW(&H3A0))
    out = Replace(out, "\phi", ChrW(&H3C6))
    out = Replace(out, "\Phi", ChrW(&H3A6))
    out = Replace(out, "\perp", ChrW(&H22A5))
    out = Replace(out, "\partial", ChrW(&H2202))
    
    out = Replace(out, "\owns", ChrW(&H220B))
    out = Replace(out, "\otimes", ChrW(&H2297))
    out = Replace(out, "\oslash", ChrW(&H2298))
    out = Replace(out, "\oplus", ChrW(&H2295))
    out = Replace(out, "\ominus", ChrW(&H2296))
    out = Replace(out, "\omicron", ChrW(&H3BF))
    out = Replace(out, "\Omicron", ChrW(&H39F))
    out = Replace(out, "\omega", ChrW(&H3C9))
    out = Replace(out, "\Omega", ChrW(&H3A9))
    out = Replace(out, "\odot", ChrW(&H2299))
    out = Replace(out, "\oint", ChrW(&H222E))

    out = Replace(out, "\nwarrow", ChrW(&H2196))
    out = Replace(out, "\nu", ChrW(&H3BD))
    out = Replace(out, "\Nu", ChrW(&H39D))
    out = Replace(out, "\nsupseteq", ChrW(&H2289))
    out = Replace(out, "\nsubseteq", ChrW(&H2288))
    out = Replace(out, "\nRightarrow", ChrW(&H21CF))
    out = Replace(out, "\nrightarrow", ChrW(&H219B))
    out = Replace(out, "\nLeftarrow", ChrW(&H21CD))
    out = Replace(out, "\nleftarrow", ChrW(&H219A))
    out = Replace(out, "\notin", ChrW(&H2209))
    out = Replace(out, "\nleq", ChrW(&H2270))
    out = Replace(out, "\ni", ChrW(&H220B))
    out = Replace(out, "\ngeq", ChrW(&H2271))
    out = Replace(out, "\nexists", ChrW(&H2204))
    out = Replace(out, "\neq", ChrW(&H2260))
    out = Replace(out, "\neg", ChrW(&HAC))
    out = Replace(out, "\nearrow", ChrW(&H2197))
    out = Replace(out, "\ne", ChrW(&H2260))
    out = Replace(out, "\nabla", ChrW(&H2207))
    
    
    out = Replace(out, "\models", ChrW(&H22A7))
    out = Replace(out, "\mu", ChrW(&H3BC))
    out = Replace(out, "\Mu", ChrW(&H39C))
    out = Replace(out, "\mp", ChrW(&H2213))
    out = Replace(out, "\mid", ChrW(&H2223))
    out = Replace(out, "\measuredangle", ChrW(&H2221))
    out = Replace(out, "\mapsto", ChrW(&H21A6))
    
    out = Replace(out, "\lt", ChrW(&H3C))
    out = Replace(out, "\lrcorner", ChrW(&H231F))
    out = Replace(out, "\lozenge", ChrW(&H25CA))
    out = Replace(out, "\longrightarrow", ChrW(&H27F6))
    out = Replace(out, "\longleftrightarrow", ChrW(&H27F7))
    out = Replace(out, "\longleftarrow", ChrW(&H27F5))
    out = Replace(out, "\Longrightarrow", ChrW(&H27F9))
    out = Replace(out, "\Longleftrightarrow", ChrW(&H27FA))
    out = Replace(out, "\Longleftarrow", ChrW(&H27F8))
    out = Replace(out, "\lor", ChrW(&H2228))
    out = Replace(out, "\ll", ChrW(&H226A))
    out = Replace(out, "\lhd", ChrW(&H25C1))
    out = Replace(out, "\lfloor", ChrW(&H230A))
    out = Replace(out, "\leq", ChrW(&H2264))
    out = Replace(out, "\leftrightarrow", ChrW(&H2194))
    out = Replace(out, "\Leftrightarrow", ChrW(&H21D4))
    out = Replace(out, "\Leftarrow", ChrW(&H21D0))
    out = Replace(out, "\leftarrow", ChrW(&H2190))
    out = Replace(out, "\leadsto", ChrW(&H21DD))
    out = Replace(out, "\ldots", ChrW(&H2026))
    out = Replace(out, "\lceil", ChrW(&H2308))
    out = Replace(out, "\langle", ChrW(&H27E8))
    out = Replace(out, "\land", ChrW(&H2227))
    out = Replace(out, "\lambda", ChrW(&H3BB))
    out = Replace(out, "\Lambda", ChrW(&H39B))
    
    out = Replace(out, "\kappa", ChrW(&H3BA))
    out = Replace(out, "\Kappa", ChrW(&H39A))
    
    out = Replace(out, "\jmath", ChrW(&H237))
    out = Replace(out, "\Join", ChrW(&H22C8))
    
    out = Replace(out, "\iota", ChrW(&H3B9))
    out = Replace(out, "\Iota", ChrW(&H399))
    out = Replace(out, "\int", ChrW(&H222B))
    out = Replace(out, "\infty", ChrW(&H221E))
    out = Replace(out, "\in", ChrW(&H2208))
    out = Replace(out, "\Im", ChrW(&H2111))
    out = Replace(out, "\imath", ChrW(&H131))
    
    out = Replace(out, "\hslash", ChrW(&H210F))
    out = Replace(out, "\hookrightarrow", ChrW(&H21AA))
    out = Replace(out, "\hookleftarrow", ChrW(&H21A9))
    out = Replace(out, "\hbar", ChrW(&H210F))
    out = Replace(out, "\heartsuit", ChrW(&H2665))
    
    out = Replace(out, "\gtrsim", ChrW(&H2273))
    out = Replace(out, "\gtrless", ChrW(&H2277))
    out = Replace(out, "\gtrapprox", ChrW(&H2A86))
    out = Replace(out, "\gg", ChrW(&H226B))
    out = Replace(out, "\geq", ChrW(&H2265))
    out = Replace(out, "\gamma", ChrW(&H3B3))
    out = Replace(out, "\Gamma", ChrW(&H393))
    
    out = Replace(out, "\forall", ChrW(&H2200))
    out = Replace(out, "\flat", ChrW(&H266D))
    out = Replace(out, "\frown", ChrW(&H2322))
    out = Replace(out, "\frac12", ChrW(&HBD))
    out = Replace(out, "\fallingdotseq", ChrW(&H2252))

    out = Replace(out, "\exists", ChrW(&H2203))
    out = Replace(out, "\epsilon", ChrW(&H3B5))
    out = Replace(out, "\equiv", ChrW(&H2261))
    out = Replace(out, "\eth", ChrW(&HF0))
    out = Replace(out, "\eta", ChrW(&H3B7))
    out = Replace(out, "\Eta", ChrW(&H397))
    out = Replace(out, "\Epsilon", ChrW(&H395))
    out = Replace(out, "\emptyset", ChrW(&H2205))
    out = Replace(out, "\ell", ChrW(&H2113))
    
    out = Replace(out, "\downarrow", ChrW(&H2193))
    out = Replace(out, "\Downarrow", ChrW(&H21D3))
    out = Replace(out, "\dots", ChrW(&H2026))
    out = Replace(out, "\div", ChrW(&HF7))
    out = Replace(out, "\diamondsuit", ChrW(&H2666))
    out = Replace(out, "\diamond", ChrW(&H22C4))
    out = Replace(out, "\delta", ChrW(&H3B4))
    out = Replace(out, "\Delta", ChrW(&H394))
    out = Replace(out, "\ddots", ChrW(&H22F1))
    out = Replace(out, "\ddagger", ChrW(&H2021))
    out = Replace(out, "\dagger", ChrW(&H2020))
    out = Replace(out, "\dashv", ChrW(&H22A3))
    
    out = Replace(out, "\curlywedge", ChrW(&H22CF))
    out = Replace(out, "\curlyvee", ChrW(&H22CE))
    out = Replace(out, "\cup", ChrW(&H222A))
    out = Replace(out, "\copyright", ChrW(&HA9))
    out = Replace(out, "\coprod", ChrW(&H2210))
    out = Replace(out, "\cong", ChrW(&H2245))
    out = Replace(out, "\complement", ChrW(&H2201))
    out = Replace(out, "\clubsuit", ChrW(&H2663))
    out = Replace(out, "\circ", ChrW(&H2218))
    out = Replace(out, "\chi", ChrW(&H3C7))
    out = Replace(out, "\Chi", ChrW(&H3A7))
    out = Replace(out, "\checkmark", ChrW(&H2713))
    out = Replace(out, "\cdots", ChrW(&H22EF))
    out = Replace(out, "\cdot", ChrW(&H22C5))
    out = Replace(out, "\cap", ChrW(&H2229))

    out = Replace(out, "\bullet", ChrW(&H2022))
    out = Replace(out, "\bowtie", ChrW(&H22C8))
    out = Replace(out, "\boxtimes", ChrW(&H22A0))
    out = Replace(out, "\Box", ChrW(&H25A1))
    out = Replace(out, "\bot", ChrW(&H22A5))
    out = Replace(out, "\blacktriangleright", ChrW(&H25B6))
    out = Replace(out, "\blacktriangleleft", ChrW(&H25C0))
    out = Replace(out, "\blacktriangledown", ChrW(&H25BC))
    out = Replace(out, "\blacktriangle", ChrW(&H25B2))
    out = Replace(out, "\blacksquare", ChrW(&H25A0))
    out = Replace(out, "\blacklozenge", ChrW(&H29EB))
    out = Replace(out, "\bigstar", ChrW(&H2605))
    out = Replace(out, "\beta", ChrW(&H3B2))
    out = Replace(out, "\Beta", ChrW(&H392))
    out = Replace(out, "\because", ChrW(&H2235))
    
    out = Replace(out, "\asymp", ChrW(&H224D))
    out = Replace(out, "\ast", ChrW(&H2A))
    out = Replace(out, "\approx", ChrW(&H2248))
    out = Replace(out, "\angle", ChrW(&H2220))
    out = Replace(out, "\Alpha", ChrW(&H391))
    out = Replace(out, "\alpha", ChrW(&H3B1))
    out = Replace(out, "\aleph", ChrW(&H2135))
    out = Replace(out, "\amalg", ChrW(&H2A3F))
    
    ' Handle single-character superscripts and subscripts
    out = ProcessSingleScripts(out, "^")
    out = ProcessSingleScripts(out, "_")
    
    
    
    ' Handle fractions
    out = ProcessFractions(out)
    
    RenderLatexSimple = out
End Function

' Helper function to process single-character superscripts and subscripts
Private Function ProcessSingleScripts(ByVal text As String, ByVal scriptChar As String) As String
    Dim out As String: out = text
    Dim skippedCount As Long: skippedCount = 0
    Dim pos As Long
    
    ' Process each script character
    Do While InStr(out, scriptChar) > 0
        pos = InStr(out, scriptChar)
        
        ' Skip if already has braces
        If Mid(out, pos + 1, 1) = "{" Then
            out = Left(out, pos - 1) & "Æ" & Mid(out, pos + 1, 1) & Mid(out, pos + 2)
            skippedCount = skippedCount + 1
            GoTo NextScript
        End If
        
        ' Add braces around single character
        out = Left(out, pos - 1) & "Æ{" & Mid(out, pos + 1, 1) & "}" & Mid(out, pos + 2)
        
NextScript:
    Loop
    
    ' Restore original script character
    out = Replace(out, "Æ", scriptChar)
    ProcessSingleScripts = out
End Function

' Helper function to process mathematical alphabets (\mathbf, \mathcal, etc.)
Private Function ProcessMathAlphabets(ByVal text As String) As String
    Dim out As String: out = text
    Dim alphabets As Variant
    Dim i As Long, sPos As Long, ePos As Long, middle As String
    
    alphabets = Array("mathbf", "mathcal", "mathbb", "mathfrak", "mathit", "mathsf", "mathtt", "text")
    
    Dim alpha As Variant
    For Each alpha In alphabets
        Do While InStr(out, "\" & alpha & "{") > 0
            sPos = InStr(out, "\" & alpha & "{") + Len(alpha) + 2
            ePos = InStr(sPos, out, "}")
            If ePos = 0 Then Exit Do
            
            middle = vbNullString
            On Error Resume Next
            For i = sPos To ePos - 1
                middle = middle & Application.Run(CStr(alpha), Mid(out, i, 1))
            Next i
            
            If Err.Number <> 0 Then
                middle = Mid(out, sPos, ePos - sPos)
                Err.Clear
            End If
            On Error GoTo 0
            
            out = Left(out, sPos - (Len(alpha) + 3)) & middle & Mid(out, ePos + 1)
        Loop
    Next alpha
    
    ProcessMathAlphabets = out
End Function

' Helper function to process fractions (\frac{numerator}{denominator})
Private Function ProcessFractions(ByVal text As String) As String
    Dim out As String: out = text
    Dim fracStart As Long, fracMid As Long, fracEnd As Long
    Dim numerator As String, denominator As String
    Do
        fracStart = InStr(out, "\frac{")
        If fracStart = 0 Then Exit Do
        
        fracMid = ClosingBrace(out, fracStart + 6)
        If fracMid = 0 Then Exit Do
        
        If Mid(out, fracMid + 1, 1) <> "{" Then Exit Do
        
        fracEnd = ClosingBrace(out, fracMid + 2)
        If fracEnd = 0 Then Exit Do
        
        
        numerator = RenderLatexSimple(Mid(out, fracStart + 6, fracMid - (fracStart + 6)))
        denominator = RenderLatexSimple(Mid(out, fracMid + 2, fracEnd - (fracMid + 2)))
        
        out = Left(out, fracStart - 1) & "^{" & numerator & "}/_{" & denominator & "}" & Mid(out, fracEnd + 1)
    Loop
    
    ProcessFractions = out
End Function

' Helper function to find closing brace
Private Function ClosingBrace(ByVal text As String, ByVal braceStart As Long) As Long
    Dim braceEnd As Long
    braceEnd = InStr(braceStart, text, "}")
    Do While braceEnd <> 0 And CountChars(Mid(text, braceStart, braceEnd - braceStart), "{") <> CountChars(Mid(text, braceStart, braceEnd - braceStart), "}")
        braceEnd = InStr(braceEnd + 1, text, "}")
    Loop
    ClosingBrace = braceEnd
End Function

' Helper function to count occurrences of a character
Private Function CountChars(ByVal text As String, ByVal char As String) As Long
    Dim i As Long, count As Long
    count = 0
    For i = 1 To Len(text)
        If Mid(text, i, 1) = char Then count = count + 1
    Next i
    CountChars = count
End Function

' Unicode character helper functions for mathematical alphabets
Private Function UnicodeChar(codePoint As Long) As String
    If codePoint <= 65535 Then
        UnicodeChar = ChrW(codePoint)
    Else
        Dim Uprime As Long: Uprime = codePoint - &H10000
        Dim highSurrogate As Long: highSurrogate = &HD800 + (Uprime \ &H400)
        Dim lowSurrogate As Long: lowSurrogate = &HDC00 + (Uprime Mod &H400)
        UnicodeChar = ChrW(highSurrogate) & ChrW(lowSurrogate)
    End If
End Function

' Mathematical alphabet functions
Private Function mathbf(c As String) As String
    Dim code As Long
    c = UCase(c)
    code = AscW(c)
    If code >= 65 And code <= 90 Then ' A-Z
        mathbf = UnicodeChar(&H1D400 + (code - 65))
    Else
        mathbf = c
    End If
End Function

Private Function mathcal(c As String) As String
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")
    
    ' Calligraphic alphabet mappings
    map.Add "A", &H1D49C: map.Add "B", &H212C: map.Add "C", &H1D49E
    map.Add "D", &H1D49F: map.Add "E", &H2130: map.Add "F", &H2131
    map.Add "G", &H1D4A2: map.Add "H", &H210B: map.Add "I", &H2110
    map.Add "J", &H1D4A5: map.Add "K", &H1D4A6: map.Add "L", &H2112
    map.Add "M", &H2133: map.Add "N", &H1D4A9: map.Add "O", &H1D4AA
    map.Add "P", &H1D4AB: map.Add "Q", &H1D4AC: map.Add "R", &H211B
    map.Add "S", &H1D4AE: map.Add "T", &H1D4AF: map.Add "U", &H1D4B0
    map.Add "V", &H1D4B1: map.Add "W", &H1D4B2: map.Add "X", &H1D4B3
    map.Add "Y", &H1D4B4: map.Add "Z", &H1D4B5
    
    c = UCase(c)
    If map.exists(c) Then
        mathcal = UnicodeChar(map(c))
    Else
        mathcal = c
    End If
End Function

Private Function mathbb(c As String) As String
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")
    
    ' Blackboard bold alphabet mappings
    map.Add "A", &H1D538: map.Add "B", &H1D539: map.Add "C", &H2102
    map.Add "D", &H1D53B: map.Add "E", &H1D53C: map.Add "F", &H1D53D
    map.Add "G", &H1D53E: map.Add "H", &H210D: map.Add "I", &H1D540
    map.Add "J", &H1D541: map.Add "K", &H1D542: map.Add "L", &H1D543
    map.Add "M", &H1D544: map.Add "N", &H2115: map.Add "O", &H1D546
    map.Add "P", &H2119: map.Add "Q", &H211A: map.Add "R", &H211D
    map.Add "S", &H1D54A: map.Add "T", &H1D54B: map.Add "U", &H1D54C
    map.Add "V", &H1D54D: map.Add "W", &H1D54E: map.Add "X", &H1D54F
    map.Add "Y", &H1D550: map.Add "Z", &H2124
    
    c = UCase(c)
    If map.exists(c) Then
        mathbb = UnicodeChar(map(c))
    Else
        mathbb = c
    End If
End Function

Private Function mathfrak(c As String) As String
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")
    
    ' Fraktur alphabet mappings
    map.Add "A", &H1D504: map.Add "B", &H1D505: map.Add "C", &H212D
    map.Add "D", &H1D507: map.Add "E", &H1D508: map.Add "F", &H1D509
    map.Add "G", &H1D50A: map.Add "H", &H210C: map.Add "I", &H2111
    map.Add "J", &H1D50D: map.Add "K", &H1D50E: map.Add "L", &H1D50F
    map.Add "M", &H1D510: map.Add "N", &H1D511: map.Add "O", &H1D512
    map.Add "P", &H1D513: map.Add "Q", &H1D514: map.Add "R", &H211C
    map.Add "S", &H1D516: map.Add "T", &H1D517: map.Add "U", &H1D518
    map.Add "V", &H1D519: map.Add "W", &H1D51A: map.Add "X", &H1D51B
    map.Add "Y", &H1D51C: map.Add "Z", &H2128
    
    c = UCase(c)
    If map.exists(c) Then
        mathfrak = UnicodeChar(map(c))
    Else
        mathfrak = c
    End If
End Function

Private Function mathit(c As String) As String
    Dim code As Long
    c = UCase(c)
    code = AscW(c)
    If code >= 65 And code <= 90 Then
        mathit = UnicodeChar(&H1D434 + (code - 65))
    Else
        mathit = c
    End If
End Function

Private Function mathsf(c As String) As String
    Dim code As Long
    c = UCase(c)
    code = AscW(c)
    If code >= 65 And code <= 90 Then
        mathsf = UnicodeChar(&H1D5A0 + (code - 65))
    Else
        mathsf = c
    End If
End Function

Private Function mathtt(c As String) As String
    Dim code As Long
    c = UCase(c)
    code = AscW(c)
    If code >= 65 And code <= 90 Then
        mathtt = UnicodeChar(&H1D670 + (code - 65))
    Else
        mathtt = c
    End If
End Function

Private Function text(c As String) As String
    c = Replace(c, "^", "ð")
    c = Replace(c, "_", "å")
    c = Replace(c, "\", "ç")
    text = c
End Function

' Additional helper function for checking if cell contains text
Private Function IsText(ByVal cellValue As Variant) As Boolean
    IsText = (VarType(cellValue) = vbString)
End Function


