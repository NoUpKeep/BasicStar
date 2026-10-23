' ============================================================
' BasicStar Word Processor
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
'   Ctrl+Z                                  - undo
'   Ctrl+Y                                  - redo
'   Ctrl+S                                  - save
'   Ctrl+O                                  - open
'   Ctrl+N                                  - new document
'   Ctrl+Q                                  - quit
' ============================================================

$ExeIcon:'BasicStar.ico'

Const MAXLINES = 5000
Const SCREENH = 22 ' text display rows (rows 1-22)
Const SCREENW = 80 ' display columns
Const MAXUNDO = 30 ' undo/redo stack depth

Dim Shared lines$(1 To MAXLINES)
Dim Shared lineCount As Integer
Dim Shared curRow As Integer
Dim Shared curCol As Integer
Dim Shared topLine As Integer
Dim Shared docFile$
Dim Shared dirty As Integer ' -1 = unsaved changes
Dim Shared statusMsg$ ' message shown on row 25
Dim Shared discardOK As Integer ' set by ConfirmDiscard SUB

' ---- Clipboard / selection ----
Dim Shared clipboard$ ' internal clipboard (newlines stored as CHR$(10))
Dim Shared selActive As Integer ' -1 = selection is active
Dim Shared selRow As Integer ' row where selection anchor was set
Dim Shared selCol As Integer ' col where selection anchor was set

' ---- Undo stack ----
' Each record saves the state of a range of lines BEFORE an edit.
' undoToRow() is the last row in that saved range.
' The corresponding "after" last row = undoToRow + (currentLineCount - undoLineCount).
Dim Shared undoTop As Integer
Dim Shared undoSuspend As Integer ' when -1, PushUndoRange is a no-op
Dim Shared undoFromRow(1 To MAXUNDO) As Integer
Dim Shared undoToRow(1 To MAXUNDO) As Integer
Dim Shared undoContent$(1 To MAXUNDO)
Dim Shared undoLineCount(1 To MAXUNDO) As Integer
Dim Shared undoCurRow(1 To MAXUNDO) As Integer
Dim Shared undoCurCol(1 To MAXUNDO) As Integer

' ---- Redo stack (same structure as undo) ----
Dim Shared redoTop As Integer
Dim Shared redoFromRow(1 To MAXUNDO) As Integer
Dim Shared redoToRow(1 To MAXUNDO) As Integer
Dim Shared redoContent$(1 To MAXUNDO)
Dim Shared redoLineCount(1 To MAXUNDO) As Integer
Dim Shared redoCurRow(1 To MAXUNDO) As Integer
Dim Shared redoCurCol(1 To MAXUNDO) As Integer

' ---- Initialise ----
Screen 0
Width 80, 25
_Title "BasicStar Word Processor"
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
undoTop = 0
undoSuspend = 0
redoTop = 0

DrawScreen
Locate curRow - topLine + 1, curCol

' ============================================================
' MAIN LOOP
' ============================================================
Dim running As Integer
Dim k As Long
Dim ctrlHeld As Integer
Dim shiftHeld As Integer
Dim okAction As Integer

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
                            Case Asc("z"), Asc("Z"): DoUndo
                            Case Asc("y"), Asc("Y"): DoRedo
                            Case Asc("s"), Asc("S"): DoSave
                            Case Asc("o"), Asc("O"): DoOpen
                            Case Asc("n"), Asc("N")
                                ConfirmDiscard
                                If discardOK Then DoNew
                            Case Asc("q"), Asc("Q")
                                ConfirmDiscard
                                If discardOK Then running = 0
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
' SUB: PushUndoRange
' Save the content of lines(fromRow..toRow) onto the undo stack
' before an edit operation.  Clears the redo stack.
' No-op when undoSuspend is set (used during paste).
' ============================================================
Sub PushUndoRange (fromRow As Integer, toRow As Integer)
    Dim i As Integer
    Dim content As String

    If undoSuspend Then Exit Sub

    ' Shift stack down if full (oldest entry lost)
    If undoTop >= MAXUNDO Then
        For i = 1 To MAXUNDO - 1
            undoFromRow(i) = undoFromRow(i + 1)
            undoToRow(i) = undoToRow(i + 1)
            undoContent$(i) = undoContent$(i + 1)
            undoLineCount(i) = undoLineCount(i + 1)
            undoCurRow(i) = undoCurRow(i + 1)
            undoCurCol(i) = undoCurCol(i + 1)
        Next i
        undoTop = MAXUNDO - 1
    End If

    undoTop = undoTop + 1
    undoFromRow(undoTop) = fromRow
    undoToRow(undoTop) = toRow
    undoLineCount(undoTop) = lineCount
    undoCurRow(undoTop) = curRow
    undoCurCol(undoTop) = curCol

    ' Build content string: lines joined by CHR$(10)
    content = ""
    For i = fromRow To toRow
        If i > fromRow Then content = content + Chr$(10)
        If i >= 1 And i <= lineCount Then
            content = content + lines$(i)
        End If
    Next i
    undoContent$(undoTop) = content

    ' New edit invalidates redo history
    redoTop = 0
End Sub

' ============================================================
' SUB: PushRedoRange
' Save the content of lines(fromRow..toRow) onto the redo stack.
' Called by DoUndo before it restores the before-state.
' ============================================================
Sub PushRedoRange (fromRow As Integer, toRow As Integer, savedCR As Integer, savedCC As Integer)
    Dim i As Integer
    Dim content As String

    If redoTop >= MAXUNDO Then
        For i = 1 To MAXUNDO - 1
            redoFromRow(i) = redoFromRow(i + 1)
            redoToRow(i) = redoToRow(i + 1)
            redoContent$(i) = redoContent$(i + 1)
            redoLineCount(i) = redoLineCount(i + 1)
            redoCurRow(i) = redoCurRow(i + 1)
            redoCurCol(i) = redoCurCol(i + 1)
        Next i
        redoTop = MAXUNDO - 1
    End If

    redoTop = redoTop + 1
    redoFromRow(redoTop) = fromRow
    redoToRow(redoTop) = toRow
    redoLineCount(redoTop) = lineCount
    redoCurRow(redoTop) = savedCR
    redoCurCol(redoTop) = savedCC

    content = ""
    For i = fromRow To toRow
        If i > fromRow Then content = content + Chr$(10)
        If i >= 1 And i <= lineCount Then
            content = content + lines$(i)
        End If
    Next i
    redoContent$(redoTop) = content
End Sub

' ============================================================
' SUB: RestoreRange
' Replace the current lines [fromRow..oldToRow] with the
' content stored in newContent (covering fromRow..newToRow).
' Adjusts lineCount and repositions the cursor.
' ============================================================
Sub RestoreRange (fromRow As Integer, oldToRow As Integer, newToRow As Integer, newContent As String, newLC As Integer, newCR As Integer, newCC As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim p As Integer
    Dim nextP As Integer
    Dim oldRangeLen As Integer
    Dim newRangeLen As Integer
    Dim delta As Integer

    oldRangeLen = oldToRow - fromRow + 1
    newRangeLen = newToRow - fromRow + 1
    delta = newRangeLen - oldRangeLen

    If delta > 0 Then
        ' Expanding: shift rows after oldToRow down by delta
        For i = lineCount To oldToRow + 1 Step -1
            If i + delta <= MAXLINES Then
                lines$(i + delta) = lines$(i)
            End If
        Next i
    ElseIf delta < 0 Then
        ' Contracting: shift rows after oldToRow up by |delta|
        For i = oldToRow + 1 To lineCount
            If i + delta >= 1 Then
                lines$(i + delta) = lines$(i)
            End If
        Next i
        ' Clear vacated slots at the bottom
        For i = lineCount + delta + 1 To lineCount
            lines$(i) = ""
        Next i
    End If

    ' Write newContent into [fromRow..newToRow]
    j = fromRow
    If Len(newContent) = 0 Then
        lines$(fromRow) = ""
    Else
        p = 1
        Do While p <= Len(newContent) And j <= newToRow
            nextP = InStr(p, newContent, Chr$(10))
            If nextP = 0 Then
                lines$(j) = Mid$(newContent, p)
                j = j + 1
                p = Len(newContent) + 1
            Else
                lines$(j) = Mid$(newContent, p, nextP - p)
                j = j + 1
                p = nextP + 1
            End If
        Loop
    End If

    lineCount = newLC
    curRow = newCR
    curCol = newCC
    If curRow < 1 Then curRow = 1
    If curRow > lineCount Then curRow = lineCount
    If curCol < 1 Then curCol = 1
    If curRow < topLine Then topLine = curRow
    If curRow > topLine + SCREENH - 1 Then topLine = curRow - SCREENH + 1
    If topLine < 1 Then topLine = 1
End Sub

' ============================================================
' SUB: DoUndo  (Ctrl+Z)
' ============================================================
Sub DoUndo ()
    Dim fromRow As Integer
    Dim savedToRow As Integer
    Dim currToRow As Integer
    Dim savedLC As Integer

    If undoTop = 0 Then
        statusMsg$ = "Nothing to undo."
        Exit Sub
    End If

    fromRow = undoFromRow(undoTop)
    savedToRow = undoToRow(undoTop)
    savedLC = undoLineCount(undoTop)

    ' Current "after" last row (derived from the change in lineCount)
    currToRow = savedToRow + (lineCount - savedLC)

    ' Save the current (after) state onto the redo stack so Ctrl+Y works.
    ' The cursor position to restore on redo is the current cursor.
    PushRedoRange fromRow, currToRow, curRow, curCol

    ' Restore the before-state
    RestoreRange fromRow, currToRow, savedToRow, undoContent$(undoTop), savedLC, undoCurRow(undoTop), undoCurCol(undoTop)

    dirty = -1
    undoTop = undoTop - 1
    statusMsg$ = "Undo."
End Sub

' ============================================================
' SUB: DoRedo  (Ctrl+Y)
' ============================================================
Sub DoRedo ()
    Dim fromRow As Integer
    Dim savedToRow As Integer
    Dim currToRow As Integer
    Dim savedLC As Integer

    If redoTop = 0 Then
        statusMsg$ = "Nothing to redo."
        Exit Sub
    End If

    fromRow = redoFromRow(redoTop)
    savedToRow = redoToRow(redoTop)
    savedLC = redoLineCount(redoTop)

    ' Current "before" last row
    currToRow = savedToRow + (lineCount - savedLC)

    ' Push current state back onto undo stack (so Ctrl+Z still works after Ctrl+Y)
    ' We bypass PushUndoRange to avoid clearing the redo stack mid-replay.
    If undoTop < MAXUNDO Then
        undoTop = undoTop + 1
    Else
        Dim i As Integer
        For i = 1 To MAXUNDO - 1
            undoFromRow(i) = undoFromRow(i + 1)
            undoToRow(i) = undoToRow(i + 1)
            undoContent$(i) = undoContent$(i + 1)
            undoLineCount(i) = undoLineCount(i + 1)
            undoCurRow(i) = undoCurRow(i + 1)
            undoCurCol(i) = undoCurCol(i + 1)
        Next i
        undoTop = MAXUNDO
    End If

    Dim content As String
    Dim j As Integer
    content = ""
    For j = fromRow To currToRow
        If j > fromRow Then content = content + Chr$(10)
        If j >= 1 And j <= lineCount Then content = content + lines$(j)
    Next j

    undoFromRow(undoTop) = fromRow
    undoToRow(undoTop) = currToRow
    undoContent$(undoTop) = content
    undoLineCount(undoTop) = lineCount
    undoCurRow(undoTop) = curRow
    undoCurCol(undoTop) = curCol

    ' Restore the redo (after) state
    RestoreRange fromRow, currToRow, savedToRow, redoContent$(redoTop), savedLC, redoCurRow(redoTop), redoCurCol(redoTop)

    dirty = -1
    redoTop = redoTop - 1
    statusMsg$ = "Redo."
End Sub

' ============================================================
' SUB: SetAnchor
' ============================================================
Sub SetAnchor ()
    selActive = -1
    selRow = curRow
    selCol = curCol
End Sub

' ============================================================
' SUB: ClearSel
' ============================================================
Sub ClearSel ()
    selActive = 0
    selRow = 0
    selCol = 0
End Sub

' ============================================================
' FUNCTION: SelStartRow / SelStartCol / SelEndRow / SelEndCol
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
        clipboard$ = Mid$(lines$(sR), sC, eC - sC)
    Else
        clipboard$ = Mid$(lines$(sR), sC)
        For i = sR + 1 To eR - 1
            clipboard$ = clipboard$ + Chr$(10) + lines$(i)
        Next i
        clipboard$ = clipboard$ + Chr$(10) + Left$(lines$(eR), eC - 1)
    End If

    statusMsg$ = "Copied."
End Sub

' ============================================================
' SUB: DeleteSelection
' ============================================================
Sub DeleteSelection ()
    Dim sR As Integer, sC As Integer
    Dim eR As Integer, eC As Integer
    Dim i As Integer

    If Not selActive Then Exit Sub

    sR = SelStartRow: sC = SelStartCol
    eR = SelEndRow: eC = SelEndCol

    ' Save undo record before modifying
    PushUndoRange sR, eR

    If sR = eR Then
        lines$(sR) = Left$(lines$(sR), sC - 1) + Mid$(lines$(sR), eC)
    Else
        lines$(sR) = Left$(lines$(sR), sC - 1) + Mid$(lines$(eR), eC)
        For i = sR + 1 To lineCount - (eR - sR)
            lines$(i) = lines$(i + (eR - sR))
        Next i
        For i = lineCount - (eR - sR) + 1 To lineCount
            lines$(i) = ""
        Next i
        lineCount = lineCount - (eR - sR)
    End If

    curRow = sR
    curCol = sC
    If curRow < topLine Then topLine = curRow

    ' Record final lineCount into the undo entry
    If undoTop > 0 And Not undoSuspend Then
        undoLineCount(undoTop) = lineCount
    End If

    ClearSel
    dirty = -1
End Sub

' ============================================================
' SUB: CutSelection
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
' ============================================================
Sub PasteClipboard ()
    Dim idx As Integer
    Dim ch As String

    If clipboard$ = "" Then
        statusMsg$ = "Clipboard is empty."
        Exit Sub
    End If

    ' One undo record for the whole paste operation
    PushUndoRange curRow, curRow
    undoSuspend = -1

    For idx = 1 To Len(clipboard$)
        ch = Mid$(clipboard$, idx, 1)
        If ch = Chr$(10) Then
            HandleEnter
        Else
            InsertChar ch
        End If
    Next idx

    undoSuspend = 0
    ' Record final state into the undo entry
    If undoTop > 0 Then
        undoLineCount(undoTop) = lineCount
    End If

    dirty = -1
    statusMsg$ = "Pasted."
End Sub

' ============================================================
' SUB: InsertChar
' ============================================================
Sub InsertChar (ch$)
    PushUndoRange curRow, curRow
    lines$(curRow) = Left$(lines$(curRow), curCol - 1) + ch$ + Mid$(lines$(curRow), curCol)
    curCol = curCol + 1
    If undoTop > 0 And Not undoSuspend Then
        undoLineCount(undoTop) = lineCount
    End If
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

    PushUndoRange curRow, curRow

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

    If undoTop > 0 And Not undoSuspend Then
        undoLineCount(undoTop) = lineCount
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
        PushUndoRange curRow, curRow
        lines$(curRow) = Left$(lines$(curRow), curCol - 2) + Mid$(lines$(curRow), curCol)
        curCol = curCol - 1
    ElseIf curRow > 1 Then
        PushUndoRange curRow - 1, curRow
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

    If undoTop > 0 And Not undoSuspend Then
        undoLineCount(undoTop) = lineCount
    End If
    dirty = -1
End Sub

' ============================================================
' SUB: HandleDelete
' ============================================================
Sub HandleDelete ()
    Dim i As Integer

    If curCol <= Len(lines$(curRow)) Then
        PushUndoRange curRow, curRow
        lines$(curRow) = Left$(lines$(curRow), curCol - 1) + Mid$(lines$(curRow), curCol + 1)
    ElseIf curRow < lineCount Then
        PushUndoRange curRow, curRow + 1
        lines$(curRow) = lines$(curRow) + lines$(curRow + 1)
        For i = curRow + 1 To lineCount - 1
            lines$(i) = lines$(i + 1)
        Next i
        lines$(lineCount) = ""
        lineCount = lineCount - 1
    End If

    If undoTop > 0 And Not undoSuspend Then
        undoLineCount(undoTop) = lineCount
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
    Dim c As Integer

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
                    Color 0, 7
                Else
                    Color 7, 0
                End If
                Print ch;
            Next c
        Else
            Color 7, 0
            Print ln;
            If Len(ln) < SCREENW Then Print Space$(SCREENW - Len(ln));
        End If
    Next i

    ' ---- Row 23: File info bar ----
    Color 0, 7
    Locate 23, 1
    If docFile$ = "" Then fileLabel = "[untitled]" Else fileLabel = docFile$
    If dirty Then fileLabel = fileLabel + " *"
    bar = " " + fileLabel + "  Ln:" + LTrim$(Str$(curRow)) + "  Co:" + LTrim$(Str$(curCol))
    Print Left$(bar + Space$(SCREENW), SCREENW);

    ' ---- Row 24: Shortcut bar ----
    Locate 24, 1
    Print Left$("  ^S=Save  ^O=Open  ^N=New  ^Q=Quit  ^Z=Undo  ^Y=Redo  ^X=Cut  ^V=Paste" + Space$(SCREENW), SCREENW);
    Color 7, 0

    ' ---- Row 25: Message line ----
    Locate 25, 1
    Color 14, 0
    Print Left$(statusMsg$ + Space$(SCREENW), SCREENW);
    Color 7, 0

End Sub

' ============================================================
' SUB: ConfirmDiscard
' ============================================================
Sub ConfirmDiscard ()
    Dim k As Long

    If Not dirty Then
        discardOK = -1
        Exit Sub
    End If

    Locate 25, 1
    Color 0, 7
    Print Left$("  Unsaved changes. Save first? (Y)es  (N)o  (C)ancel  " + Space$(SCREENW), SCREENW);
    Color 7, 0

    Do
        k = _KeyHit
        If k = Asc("y") Or k = Asc("Y") Then
            DoSave
            discardOK = -1
            Exit Sub
        ElseIf k = Asc("n") Or k = Asc("N") Then
            discardOK = -1
            Exit Sub
        ElseIf k = Asc("c") Or k = Asc("C") Or k = 27 Then
            statusMsg$ = "Cancelled."
            discardOK = 0
            Exit Sub
        End If
        _Limit 60
    Loop
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
    ' Clear undo/redo history when a new file is loaded
    undoTop = 0
    redoTop = 0
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
    ' Clear undo/redo history for the new document
    undoTop = 0
    redoTop = 0
    statusMsg$ = "New document."
    ClearSel
End Sub

