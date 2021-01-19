Sub Process()
'
' Process Macro
' Processes data.
'
' Keyboard Shortcut: Ctrl+Shift+Q
'
' [Last edited on 02/25/19]
'

' +-------------------------------------+
' | PROCESSOR MACRO - EDIT WITH CAUTION |
' +-------------------------------------+

' Copies the team list to "DP19"

    Sheets("Teams").Select
    Range("B3:B203").Select
    Selection.Copy
    Sheets("DP19").Select
    Range("T2").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

' Ensures that a "stop value" exists in the copied list

    Dim rowA, rowB, colA, colB, blah, fact, ctr As Integer
    Let rowA = 2
    Let colA = 20
    Let blah = 0
    Let ctr = 0
    
    Do Until blah = 1
    Cells(rowA, colA).Select
    If Selection.Value = "" Then
    Let blah = 1
    Selection.Value = "[0x7effaf]"
    Exit Do
    Else
    rowA = rowA + 1
    End If
    Loop

' Prepares variables for data processing

    Let rowB = 2
    Let colB = 20
    Let blah = 0
    Let fact = Range("Indexer").Value
    Let ctr = 0
    
' Processes data
    
    Do
    Sheets("DP19").Select
    Cells(rowB, colB).Select
    ' Checks if the selected cell is the "stop value"
    If Selection.Value = "[0x7effaf]" Then
    ' If so, ends the Do loop
    Let blah = 1
    Selection.Clear
    Exit Do
    Else
    ' If not, copies the selected cell into "DP19"
    Selection.Copy
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    ' Copies data from "DP19" to "Storage"
    Range("B3:Q19").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Storage").Select
    Cells(3 + fact * ctr, 2).Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    ' Increments the variables
    Let rowB = rowB + 1
    Let ctr = ctr + 1
    End If
    Loop Until blah = 1

' Clears "DP19"

    Sheets("DP19").Select
    Range("A2").Select
    Selection.ClearContents
    Range("T2:T202").Select
    Selection.Clear
    Range("A1").Select

' Selects the "safe cells" in each sheet required for processing (cleanup)

    Sheets("Teams").Select
    Range("A1").Select
    Sheets("Storage").Select
    Range("A2").Select
    Range("A1").Select
End Sub
