Sub CleanRawTeams()
'
' CleanRawTeams Macro
' Removes 'junk' from the team list in "RAW INPUT". (For use with a custom scouting app)
'
' Keyboard Shortcut: Ctrl+Shift+P
' [Last edited on 01/18/21]
'

Dim ctr, col As Integer
Let ctr = 2
Let col = 1

Sheets("RAW INPUT").Select

Do Until ctr > 1001
' Selects a cell in the list
Cells(ctr, col).Select
Chopper:
If InStr(Selection.Value, "(") = 1 Then
' Cuts the bracket off of the cell
ActiveCell = Right(Selection.Value, Len(Selection.Value) - 1)
' Checks again
GoTo Chopper
End If
' Increments the counter
Let ctr = ctr + 1
Loop

Range("A1").Select

End Sub
