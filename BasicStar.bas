' ============================================================
' QB64 Word Processor
' ============================================================
' Controls:
'   Arrows / Home / End / PgUp / PgDn      - navigation
'   Shift+Arrows                            - select text
'   Enter                                   - new line
'   Backspace                               - delete left / join lines
'   Delete                                  - delete right / join lines
'   Ctrl+C                                  - copy selection
'   Ctrl+X                                  - cut selection
'   Ctrl+V                                  - paste
'   Ctrl+S                                  - save
'   Ctrl+O                                  - open
'   Ctrl+N                                  - new document
'   Ctrl+Q                                  - quit
' ============================================================

Const MAXLINES = 5000
Const SCREENH = 23 ' text display rows (rows 1-23)
Const SCREENW = 80 ' display columns

Dim Shared lines$(1 To MAXLINES)
Dim Shared lineCount As Integer
Dim Shared curRow As Integer
Dim Shared curCol As Integer
Dim Shared topLine As Integer
Dim Shared docFile$
Dim Shared dirty As Integer ' -1 = unsaved changes
Dim Shared statusMsg$ ' message shown on row 25

' ---- Clipboard / selection ----
Dim Shared clipboard$ ' internal clipboard (newlines stored as CHR$(10))
Dim Shared selActive As Integer ' -1 = selection is active
Dim Shared selRow As Integer ' row where selection anchor was set
Dim Shared selCol As Integer ' col where selection anchor was set

' ---- Initialise ----
Screen 0
Width 80, 25
_Title "QB64 Word Processor"
_FullScreen
Color 7, 0
Cls

lineCount = 1
lines$(1) = ""
curRow = 1
curCol = 1
topLine = 1
dirty = 0
docFile$ = ""
statusMsg$ = ""
clipboard$ = ""
selActive = 0
selRow = 0
selCol = 0

DrawScreen
Locate curRow - topLine + 1, curCol

' ============================================================
' MAIN LOOP
' ============================================================
Dim running As Integer
Dim k As Long
Dim ctrlHeld As Integer
Dim shiftHeld As Integer

running = -1

Do While running

    k = _KeyHit

    If k <> 0 Then

        ctrlHeld = _KeyDown(100306) Or _KeyDown(100305) ' Left or Right Ctrl
        shiftHeld = _KeyDown(100304) Or _KeyDown(100303) ' Left or Right Shift

        Select Case k

            ' ---- Arrow: Up ----
            Case -18432
                If shiftHeld Then
                    If Not selActive Then SetAnchor
                Else
                    ClearSel
                End If
                If curRow > 1 Then
                    curRow = curRow - 1
                    ClampCol
                    If curRow < topLine Then topLine = curRow
                End If

                ' ---- Arrow: Down ----
            Case -20480
                If shiftHeld Then
                    If Not selActive Then SetAnchor
                Else
                    ClearSel
                End If
                If curRow < lineCount Then
                    curRow = curRow + 1
                    ClampCol
                    If curRow > topLine + SCREENH - 1 Then
                        topLine = curRow - SCREENH + 1
                    End If
                End If

                ' ---- Arrow: Left ----
            Case -19200
                If shiftHeld Then
                    If Not selActive Then SetAnchor
                Else
                    ClearSel
                End If
                If curCol > 1 Then
                    curCol = curCol - 1
                ElseIf curRow > 1 Then
                    curRow = curRow - 1
                    curCol = Len(lines$(curRow)) + 1
                    If curRow < topLine Then topLine = curRow
                End If

                ' ---- Arrow: Right ----
            Case -19712
                If shiftHeld Then
                    If Not selActive Then SetAnchor
                Else
                    ClearSel
                End If
                If curCol <= Len(lines$(curRow)) Then
                    curCol = curCol + 1
                ElseIf curRow < lineCount Then
                    curRow = curRow + 1
                    curCol = 1
                    If curRow > topLine + SCREENH - 1 Then
                        topLine = curRow - SCREENH + 1
                    End If
                End If

                ' ---- Home ----
            Case -18176
                ClearSel
                curCol = 1

                ' ---- End ----
            Case -20224
                ClearSel
                curCol = Len(lines$(curRow)) + 1

                ' ---- Page Up ----
            Case -18688
                ClearSel
                curRow = curRow - SCREENH
                If curRow < 1 Then curRow = 1
                topLine = curRow
                ClampCol

                ' ---- Page Down ----
            Case -20736
                ClearSel
                curRow = curRow + SCREENH
                If curRow > lineCount Then curRow = lineCount
                topLine = curRow - SCREENH + 1
                If topLine < 1 Then topLine = 1
                ClampCol

                ' ---- Enter ----
            Case 13
                If selActive Then DeleteSelection
                HandleEnter

                ' ---- Backspace ----
            Case 8
                If selActive Then
                    DeleteSelection
                Else
                    HandleBackspace
                End If

                ' ---- Delete ----
            Case 127, -21248
                If selActive Then
                    DeleteSelection
                Else
                    HandleDelete
                End If

                ' ---- Escape: clear selection and status message ----
            Case 27
                ClearSel
                statusMsg$ = ""

                ' ---- Printable ASCII (and Ctrl shortcuts) ----
            Case Else
                If k >= 32 And k <= 126 Then
                    If ctrlHeld Then
                        Select Case k
                            Case Asc("c"), Asc("C"): CopySelection
                            Case Asc("x"), Asc("X"): CutSelection
                            Case Asc("v"), Asc("V"): PasteClipboard
                            Case Asc("s"), Asc("S"): DoSave
                            Case Asc("o"), Asc("O"): DoOpen
                            Case Asc("n"), Asc("N"): DoNew
                            Case Asc("q"), Asc("Q"): running = 0
                        End Select
                    Else
                        If selActive Then DeleteSelection
                        InsertChar Chr$(k)
                    End If
                End If

        End Select

        DrawScreen
        Locate curRow - topLine + 1, curCol

    End If

    _Limit 60

Loop

System

' ============================================================
' SUB: SetAnchor
' Mark the current cursor position as the selection start
' ============================================================
Sub SetAnchor ()
    selActive = -1
    selRow = curRow
    selCol = curCol
End Sub

' ============================================================
' SUB: ClearSel
' Cancel any active selection
' ============================================================
Sub ClearSel ()
    selActive = 0
    selRow = 0
    selCol = 0
End Sub

' ============================================================
' FUNCTION: SelStartRow / SelStartCol / SelEndRow / SelEndCol
' Return the normalised start and end of the selection
' (anchor may be after the cursor, so we always sort them)
' ============================================================
Function SelStartRow ()
    If selRow < curRow Then SelStartRow = selRow
    If selRow > curRow Then SelStartRow = curRow
    If selRow = curRow Then
        If selCol <= curCol Then SelStartRow = selRow Else SelStartRow = curRow
    End If
End Function

Function SelStartCol ()
    If selRow < curRow Then SelStartCol = selCol
    If selRow > curRow Then SelStartCol = curCol
    If selRow = curRow Then
        If selCol <= curCol Then SelStartCol = selCol Else SelStartCol = curCol
    End If
End Function

Function SelEndRow ()
    If selRow < curRow Then SelEndRow = curRow
    If selRow > curRow Then SelEndRow = selRow
    If selRow = curRow Then
        If selCol <= curCol Then SelEndRow = curRow Else SelEndRow = selRow
    End If
End Function

Function SelEndCol ()
    If selRow < curRow Then SelEndCol = curCol
    If selRow > curRow Then SelEndCol = selCol
    If selRow = curRow Then
        If selCol <= curCol Then SelEndCol = curCol Else SelEndCol = selCol
    End If
End Function

' ============================================================
' SUB: CopySelection
' Copy selected text into clipboard$
' Lines are joined with CHR$(10) as a newline marker
' ============================================================
Sub CopySelection ()
    Dim sR As Integer, sC As Integer
    Dim eR As Integer, eC As Integer
    Dim i As Integer

    If Not selActive Then
        statusMsg$ = "No selection to copy."
        Exit Sub
    End If

    sR = SelStartRow: sC = SelStartCol
    eR = SelEndRow: eC = SelEndCol

    clipboard$ = ""

    If sR = eR Then
        ' Single line selection
        clipboard$ = Mid$(lines$(sR), sC, eC - sC)
    Else
        ' Multi-line: first partial line
        clipboard$ = Mid$(lines$(sR), sC)
        ' Middle lines (complete)
        For i = sR + 1 To eR - 1
            clipboard$ = clipboard$ + Chr$(10) + lines$(i)
        Next i
        ' Last partial line
        clipboard$ = clipboard$ + Chr$(10) + Left$(lines$(eR), eC - 1)
    End If

    statusMsg$ = "Copied."
End Sub

' ============================================================
' SUB: DeleteSelection
' Remove the selected text from the buffer.
' Leaves cursor at the start of the deleted range.
' ============================================================
Sub DeleteSelection ()
    Dim sR As Integer, sC As Integer
    Dim eR As Integer, eC As Integer
    Dim id As Integer

    If Not selActive Then Exit Sub

    sR = SelStartRow: sC = SelStartCol
    eR = SelEndRow: eC = SelEndCol

    If sR = eR Then
        ' Single line: just snip out the middle
        lines$(sR) = Left$(lines$(sR), sC - 1) + Mid$(lines$(sR), eC)
    Else
        ' Multi-line: join first line's left with last line's right
        lines$(sR) = Left$(lines$(sR), sC - 1) + Mid$(lines$(eR), eC)
        ' Remove lines sR+1 through eR by shifting up
        For id = sR + 1 To lineCount - (eR - sR)
            lines$(i) = lines$(i + (eR - sR))
        Next id
        ' Blank the vacated slots at the bottom
        For id = lineCount - (eR - sR) + 1 To lineCount
            lines$(id) = ""
        Next id
        lineCount = lineCount - (eR - sR)
    End If

    ' Place cursor at start of deleted range
    curRow = sR
    curCol = sC
    If curRow < topLine Then topLine = curRow

    ClearSel
    dirty = -1
End Sub

' ============================================================
' SUB: CutSelection
' Copy then delete the selection
' ============================================================
Sub CutSelection ()
    If Not selActive Then
        statusMsg$ = "No selection to cut."
        Exit Sub
    End If
    CopySelection
    DeleteSelection
    statusMsg$ = "Cut."
End Sub

' ============================================================
' SUB: PasteClipboard
' Insert clipboard$ at the cursor position.
' CHR$(10) in the clipboard becomes a new line.
' ============================================================
Sub PasteClipboard ()
    Dim ipaste As Integer
    Dim ch As String
    Dim position As Integer

    If clipboard$ = "" Then
        statusMsg$ = "Clipboard is empty."
        Exit Sub
    End If

    ' Insert one character/newline at a time
    For position = 1 To Len(clipboard$)
        ch = Mid$(clipboard$, position, 1)
        If ch = Chr$(10) Then
            HandleEnter
        Else
            InsertChar ch
        End If
    Next position

    dirty = -1
    statusMsg$ = "Pasted."
End Sub

' ============================================================
' SUB: InsertChar
' Insert a single character at the cursor position
' ============================================================
Sub InsertChar (ch$)
    lines$(curRow) = Left$(lines$(curRow), curCol - 1) + ch$ + Mid$(lines$(curRow), curCol)
    curCol = curCol + 1
    dirty = -1
End Sub

' ============================================================
' SUB: HandleEnter
' ============================================================
Sub HandleEnter ()
    Dim rightPart As String
    Dim i As Integer

    If lineCount >= MAXLINES Then
        statusMsg$ = "Document full - cannot add more lines."
        Exit Sub
    End If

    rightPart = Mid$(lines$(curRow), curCol)
    lines$(curRow) = Left$(lines$(curRow), curCol - 1)

    For i = lineCount To curRow + 1 Step -1
        lines$(i + 1) = lines$(i)
    Next i
    lineCount = lineCount + 1

    lines$(curRow + 1) = rightPart
    curRow = curRow + 1
    curCol = 1

    If curRow > topLine + SCREENH - 1 Then
        topLine = curRow - SCREENH + 1
    End If

    dirty = -1
End Sub

' ============================================================
' SUB: HandleBackspace
' ============================================================
Sub HandleBackspace ()
    Dim prevLen As Integer
    Dim i As Integer

    If curCol > 1 Then
        lines$(curRow) = Left$(lines$(curRow), curCol - 2) + Mid$(lines$(curRow), curCol)
        curCol = curCol - 1
    ElseIf curRow > 1 Then
        prevLen = Len(lines$(curRow - 1))
        lines$(curRow - 1) = lines$(curRow - 1) + lines$(curRow)
        For i = curRow To lineCount - 1
            lines$(i) = lines$(i + 1)
        Next i
        lines$(lineCount) = ""
        lineCount = lineCount - 1
        curRow = curRow - 1
        curCol = prevLen + 1
        If curRow < topLine Then topLine = curRow
    End If

    dirty = -1
End Sub

' ============================================================
' SUB: HandleDelete
' ============================================================
Sub HandleDelete ()
    Dim i As Integer

    If curCol <= Len(lines$(curRow)) Then
        lines$(curRow) = Left$(lines$(curRow), curCol - 1) + Mid$(lines$(curRow), curCol + 1)
    ElseIf curRow < lineCount Then
        lines$(curRow) = lines$(curRow) + lines$(curRow + 1)
        For i = curRow + 1 To lineCount - 1
            lines$(i) = lines$(i + 1)
        Next i
        lines$(lineCount) = ""
        lineCount = lineCount - 1
    End If

    dirty = -1
End Sub

' ============================================================
' SUB: ClampCol
' ============================================================
Sub ClampCol ()
    Dim maxCol As Integer
    maxCol = Len(lines$(curRow)) + 1
    If curCol > maxCol Then curCol = maxCol
    If curCol < 1 Then curCol = 1
End Sub

' ============================================================
' SUB: DrawScreen
' Redraw the visible text area and status bars.
' Selected text is shown in reverse colour.
' ============================================================
Sub DrawScreen ()
    Dim i As Integer
    Dim displayRow As Integer
    Dim ln As String
    Dim fileLabel As String
    Dim bar As String
    Dim sR As Integer
    Dim sC As Integer
    Dim eR As Integer
    Dim eC As Integer
    Dim c As Integer ' column counter for highlight

    If selActive Then
        sR = SelStartRow: sC = SelStartCol
        eR = SelEndRow: eC = SelEndCol
    End If

    ' ---- Text area ----
    For i = topLine To topLine + SCREENH - 1
        displayRow = i - topLine + 1
        Locate displayRow, 1

        If i <= lineCount Then
            ln = Left$(lines$(i), SCREENW)
        Else
            ln = ""
        End If

        If selActive And i >= sR And i <= eR Then
            ' Draw character by character to apply highlight
            For c = 1 To SCREENW
                Dim ch As String
                If c <= Len(ln) Then ch = Mid$(ln, c, 1) Else ch = " "

                Dim inSel As Integer
                inSel = 0
                If i = sR And i = eR Then
                    If c >= sC And c < eC Then inSel = -1
                ElseIf i = sR Then
                    If c >= sC Then inSel = -1
                ElseIf i = eR Then
                    If c < eC Then inSel = -1
                Else
                    inSel = -1
                End If

                If inSel Then
                    Color 0, 7 ' reverse: black on white
                Else
                    Color 7, 0 ' normal
                End If
                Print ch;
            Next c
        Else
            Color 7, 0
            Print ln;
            If Len(ln) < SCREENW Then Print Space$(SCREENW - Len(ln));
        End If
    Next i

    ' ---- Row 24: Status bar ----
    Color 0, 7
    Locate 24, 1
    If docFile$ = "" Then fileLabel = "[untitled]" Else fileLabel = docFile$
    If dirty Then fileLabel = fileLabel + " *"
bar = " " + fileLabel + "  Ln:" + LTRIM$(STR$(curRow)) + "  Co:" + LTRIM$(STR$(curCol)) _
    + "  | ^S=Save ^O=Open ^N=New ^C/X/V ^Q=Quit"
    Print Left$(bar + Space$(SCREENW), SCREENW);
    Color 7, 0

    ' ---- Row 25: Message line ----
    Locate 25, 1
    Color 14, 0
    Print Left$(statusMsg$ + Space$(SCREENW), SCREENW);
    Color 7, 0

End Sub

' ============================================================
' SUB: DoSave
' ============================================================
Sub DoSave ()
    Dim f As Integer
    Dim i As Integer

    If docFile$ = "" Then
        Locate 25, 1
        Color 14, 0
        Print "Save as (Enter to cancel): ";
        Color 7, 0
        Line Input docFile$
        docFile$ = RTrim$(LTrim$(docFile$))
    End If

    If docFile$ = "" Then
        statusMsg$ = "Save cancelled."
        Exit Sub
    End If

    f = FreeFile
    Open docFile$ For Output As #f
    For i = 1 To lineCount
        Print #f, lines$(i)
    Next i
    Close #f

    dirty = 0
    statusMsg$ = "Saved: " + docFile$
End Sub

' ============================================================
' SUB: DoOpen
' ============================================================
Sub DoOpen ()
    Dim fn As String
    Dim i As Integer
    Dim f As Integer

    Locate 25, 1
    Color 14, 0
    Print "Open file (Enter to cancel): ";
    Color 7, 0
    Line Input fn
    fn = RTrim$(LTrim$(fn))

    If fn = "" Then
        statusMsg$ = "Open cancelled."
        Exit Sub
    End If

    If Not _FileExists(fn) Then
        statusMsg$ = "File not found: " + fn
        Exit Sub
    End If

    For i = 1 To lineCount
        lines$(i) = ""
    Next i
    lineCount = 0

    f = FreeFile
    Open fn For Input As #f
    Do While Not EOF(f) And lineCount < MAXLINES
        lineCount = lineCount + 1
        Line Input #f, lines$(lineCount)
    Loop
    Close #f

    If lineCount = 0 Then
        lineCount = 1
        lines$(1) = ""
    End If

    docFile$ = fn
    curRow = 1
    curCol = 1
    topLine = 1
    dirty = 0
    statusMsg$ = "Opened: " + docFile$
End Sub

' ============================================================
' SUB: DoNew
' ============================================================
Sub DoNew ()
    Dim i As Integer

    For i = 1 To lineCount
        lines$(i) = ""
    Next i
    lineCount = 1
    lines$(1) = ""
    curRow = 1
    curCol = 1
    topLine = 1
    docFile$ = ""
    dirty = 0
    statusMsg$ = "New document."
    ClearSel
End Sub

