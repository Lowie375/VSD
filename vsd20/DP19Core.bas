Public Function arrLength(arr As Variant) As Integer
    arrLength = UBound(arr) - LBound(arr) + 1
End Function
Sub MDM()

' MDM - Match Data Merger
' Merges TBA match data into raw input data.

' Shortcut: Ctrl+Shift+K
' Last updated on 06/03/20

' +---------------------------------------+
' | MATCH DATA MERGER - EDIT WITH CAUTION |
' +---------------------------------------+
' +-------------------------------------------------+
' | CORE MACRO - SAVE + BACKUP BEFORE EVERY TEST!!! |
' +-------------------------------------------------+

' Declarations...
Dim wt As Worksheet, wm As Worksheet, ws As Worksheet, wj As Worksheet, wi As Worksheet, wg As Worksheet, active As Worksheet
Dim row As Integer, col As Integer, check As Integer, r2 As Integer, ctr As Integer, failCtr As Integer, arrCtr As Integer, mCols() As Integer, iCols() As Integer, bleakCtr As Integer
Dim tSave As Long, mSave As Long
Dim mLink As String, iLink As String, mColsT() As String, iColsT() As String

' ...and some more declarations...
Set wt = Worksheets("Teams")
Set wm = Worksheets("MDM")
Set ws = Worksheets("Storage")
Set wj = Worksheets("JSON")
Set wi = Worksheets("INPUT")
Set wg = Worksheets("GUIDE")
Set active = ActiveSheet
mLink = "3,4,7,8,9,10,11"
mColsT = Split(mLink, ",")
iLink = "5,16,19,27,24,23,25"
iColsT = Split(iLink, ",")

' ...and some re-declarations too!
ReDim mCols(0 To arrLength(mColsT) - 1) As Integer
ReDim iCols(0 To arrLength(iColsT) - 1) As Integer

' Awaits confirmation from the user (if not already given via Process)
wg.Activate
If wg.Range("MDMCheck").Value <> "[{{0x7effaf}}]" Then
    wg.Range("MDMCheck").Value = False
    MDMRunWarning.Show

    ' Checks if the user approved the action
    If wg.Range("MDMCheck").Value = False Then
        ' Not approved; cancel
        Exit Sub
    Else
        ' Approved; continue
        wg.Range("MDMCheck").Value = False
    End If
Else
    ' Approved via Process; continue
    wg.Range("MDMCheck").Value = False
End If

' Fixes the arrays
For m = 0 To arrLength(mColsT) - 1
    mCols(m) = CInt(mColsT(m))
Next m
For i = 0 To arrLength(iColsT) - 1
    iCols(i) = CInt(iColsT(i))
Next i

' Variable config
wj.Activate
row = wj.Range("MP.Rows")
col = wj.Range("MP.Cols")
check = 0
bleakCtr = 0

' Places a 'stop value' at the end of the TBA match list (optimized!)
Do Until check >= 1
    wj.Cells(row, col).Select
    If Selection.Value = "" Then
        check = 1
    Exit Do
    Else
        bleakCtr = bleakCtr + 1
        row = row + 100
    End If
Loop
' Checks if there is no data present
If bleakCtr = 0 Then
    ' Places the 'stop value' early and skips the rest of the algorithm
    wj.Cells(row, col).Value = "{[[0x7effaf]]}"
    check = 4
Else
    row = row - 10
End If
Do Until check >= 2
    wj.Cells(row, col).Select
    If Selection.Value <> "" Then
        check = 2
    Exit Do
    Else
        row = row - 10
    End If
Loop
row = row + 1
Do Until check >= 3
    wj.Cells(row, col).Select
    If Selection.Value = "" Then
        check = 3
        Selection.Value = "{[[0x7effaf]]}"
    Exit Do
    Else
        row = row + 1
    End If
Loop

' Clears "MDM"
wm.Activate
wm.Range("MDMData").ClearContents

' Splits + copies TBA data to "MDM"
row = wj.Range("MP.Rows")
r2 = 2
Do Until wj.Cells(row, col) = "{[[0x7effaf]]}"
    ' RED
    For i = 3 To 5
        ' Team number
        wj.Activate
        wj.Cells(row, col + i).Copy
        wm.Activate
        wm.Cells(r2, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        ' Match number
        wj.Activate
        wj.Cells(row, col).Copy
        wm.Activate
        wm.Cells(r2, 2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        ' Init line
        wj.Activate
        wj.Cells(row, col + i + 3).Copy
        wm.Activate
        wm.Cells(r2, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        ' Endgame
        wj.Activate
        wj.Cells(row, col + i + 6).Copy
        wm.Activate
        wm.Cells(r2, 4).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        ' General data
        wj.Activate
        wj.Range(wj.Cells(row, col + 12), wj.Cells(row, col + 17)).Copy
        wm.Activate
        wm.Cells(r2, 5).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        ' Result
        wj.Activate
        If wj.Cells(row, col + 2).Value = "R" Then
            wm.Activate
            wm.Cells(r2, 11).Value = "W"
        ElseIf wj.Cells(row, col + 2).Value = "B" Then
            wm.Activate
            wm.Cells(r2, 11).Value = "L"
        Else
            wm.Activate
            wm.Cells(r2, 11).Value = "T"
        End If
        
        ' Increments the MDM row counter
        r2 = r2 + 1
    Next i
    ' BLUE
    For i = 3 + wj.Range("MP.Shift") To 5 + wj.Range("MP.Shift")
        ' Team number
        wj.Activate
        wj.Cells(row, col + i).Copy
        wm.Activate
        wm.Cells(r2, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        ' Match number
        wj.Activate
        wj.Cells(row, col).Copy
        wm.Activate
        wm.Cells(r2, 2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        ' Init line
        wj.Activate
        wj.Cells(row, col + i + 3).Copy
        wm.Activate
        wm.Cells(r2, 3).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        ' Endgame
        wj.Activate
        wj.Cells(row, col + i + 6).Copy
        wm.Activate
        wm.Cells(r2, 4).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        ' General data
        wj.Activate
        wj.Range(wj.Cells(row, col + 12), wj.Cells(row, col + 17)).Copy
        wm.Activate
        wm.Cells(r2, 5).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        ' Result
        wj.Activate
        If wj.Cells(row, col + 2).Value = "B" Then
            wm.Activate
            wm.Cells(r2, 11).Value = "W"
        ElseIf wj.Cells(row, col + 2).Value = "R" Then
            wm.Activate
            wm.Cells(r2, 11).Value = "L"
        Else
            wm.Activate
            wm.Cells(r2, 11).Value = "T"
        End If
        
        ' Increments the MDM row counter
        r2 = r2 + 1
    Next i
    ' Increments the JSON row counter
    row = row + 1
Loop
wj.Activate
wj.Cells(row, col).ClearContents
wm.Activate
wm.Cells(r2, 1).Value = "[[[0x7effaf]]]"

' Data correction

' - Init line [3] (Exited --> Y, None --> N)
' - Rot/Pos [5/6] (True --> Partner, False --> N)
' - LvlRung [7] (IsLevel --> Y, notLevel --> N)
' - FMSCol [8] (Unknown --> N)

r2 = 2
Do Until wm.Cells(r2, 1) = "[[[0x7effaf]]]"
    ' Init line
    If wm.Cells(r2, 3) = "Exited" Then
        wm.Cells(r2, 3).Value = "Y"
    Else
        wm.Cells(r2, 3).Value = "N"
    End If
    ' Rot/Pos control
    For i = 5 To 6
        If wm.Cells(r2, i) = True Then
            wm.Cells(r2, i).Value = "Partner"
        Else
            wm.Cells(r2, i).Value = "N"
        End If
    Next i
    ' GenSwitch
    If wm.Cells(r2, 7) = "IsLevel" Then
        wm.Cells(r2, 7).Value = "Y"
    Else
        wm.Cells(r2, 7).Value = "N"
    End If
    ' FMS colour
    If wm.Cells(r2, 8) = "Unknown" Then
        wm.Cells(r2, 8).Value = "N"
    End If
    
    ' Increments the variable
    r2 = r2 + 1
Loop

' Merger (yay!)
ctr = 0
failCtr = 0
For r3 = 3 To 10370
    ' If applicable, checks if the hard-coded limit has been reached
    If wm.Range("HardLimitCheck").Value = True And ctr > wm.Range("HardLimit").Value Then
        ' Hard limit reached; escapes the merger loop
        Exit For
    End If
    
    ' Checks if there is team and match number
    wi.Activate
    If wi.Cells(r3, 1) <> "" And wi.Cells(r3, 2) <> "" Then
        ' Saves the team and match number
        wi.Activate
        tSave = wi.Cells(r3, 1).Value
        mSave = wi.Cells(r3, 2).Value
        
        ' Resets the other variables
        r2 = 2
        ctr = 0
        check = 0
        
        ' Loops through all the TBA match data, looking for a match
        wm.Activate
        Do Until wm.Cells(r2, 2) = "[[[0x7effaf]]]"
            If wm.Cells(r2, 2).Value = mSave Then
                check = 1
                Exit Do
            Else
                r2 = r2 + 6
            End If
        Loop
        If check = 1 Then
            ' If exact match is found, searches for a matching team
            For j = 0 To 5
                If wm.Cells(r2 + j, 1).Value = tSave Then
                    ' Exact match found! Begins merging TBA match data with "INPUT" data
                    check = 2
                    
                    ' Result
                    wm.Cells(r2 + j, 11).Copy
                    wi.Activate
                    wi.Cells(r3, 25).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                    ' Alliance colour
                    If j >= 0 And j <= 2 Then
                        wi.Cells(r3, 3).Value = "R"
                    ElseIf j >= 3 And j <= 5 Then
                        wi.Cells(r3, 3).Value = "B"
                    End If
                    ' RotControl/PosControl (5, 13)/(6, 14)
                    For k = 0 To 1
                        If wi.Cells(r3, 13 + k).Value <> "Yes" And wi.Cells(r3, 13 + k).Value <> "Y" And wi.Cells(r3, 13 + k).Value <> "Bot" _
                            And wi.Cells(r3, 13 + k).Value <> "B" And wi.Cells(r3, 13 + k).Value <> 1 And wi.Cells(r3, 13 + k).Value <> True Then
                            
                            If wm.Cells(r2 + j, 5 + k) = "Y" Then
                                wi.Activate
                                wi.Cells(r3, 13 + k).Value = "Partner"
                            Else
                                wi.Activate
                                wi.Cells(r3, 13 + k).Value = "N"
                            End If
                        End If
                    Next k
                    
                    ' All others
                    arrCtr = 0
                    For Each xCol In mCols
                        wm.Activate
                        wm.Cells(r2 + j, xCol).Copy
                        wi.Activate
                        wi.Cells(r3, iCols(arrCtr)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                        
                        ' Increments the variable
                        arrCtr = arrCtr + 1
                    Next
                End If
            Next j
            If check <> 2 Then
                ' If no matching team is found, increments the fail counter and jumps to the next row
                failCtr = failCtr + 1
            End If
        Else
            ' If no matching match is found, increments the fail counter and jumps to the next row
            failCtr = failCtr + 1
        End If
    Else
        ' If either T# or M# are missing, skips the row (exact match can't be confirmed) and increments the counter
        ctr = ctr + 1
    End If
Next r3

' If failCtr > 0 Then
    'MDMAltDone.Show ' New UserForm
' Else
    'MDMDone.Show ' New UserForm
' End If

' View cleanup
wt.Activate
wt.Range("C3").Select
wt.Range("A1").Select
wm.Activate
wm.Range("C3").Select
wm.Range("A1").Select
ws.Activate
ws.Range("C3").Select
ws.Range("A1").Select
wj.Activate
wj.Range("C3").Select
wj.Range("A1").Select
wi.Activate
wi.Range("C3").Select
wi.Range("A1").Select
wg.Activate
wg.Range("C3").Select
wg.Range("A1").Select
active.Activate

End Sub
Sub Process()

' Process - Data Processor
' Processes raw data and copies it to storage for analysis and calculations.

' Shortcut: Ctrl+Shift+Q
' Last updated on 07/03/2020

' +-------------------------------------+
' | PROCESSOR MACRO - EDIT WITH CAUTION |
' +-------------------------------------+
' +-------------------------------------------------+
' | CORE MACRO - SAVE + BACKUP BEFORE EVERY TEST!!! |
' +-------------------------------------------------+

' Declarations...
Dim row As Integer, col As Integer, fact As Integer, ctr As Integer, check As Integer

' ...and some more declarations too!
Dim wt As Worksheet, wp As Worksheet, ws As Worksheet, wg As Worksheet, active As Worksheet
Set wt = Worksheets("Teams")
Set wp = Worksheets("DP19")
Set ws = Worksheets("Storage")
Set wg = Worksheets("GUIDE")
Set active = ActiveSheet

ws.Activate
If ws.Range("AutoPullCheck") = True Then
    ' Awaits confirmation from the user
    wg.Activate
    wg.Range("FullProcessCheck").Value = 0
    FullProcessWarning.Show

    ' Checks if the user approved the action
    If wg.Range("FullProcessCheck").Value = 2 Then
        ' Approved; continue
        wg.Range("FullProcessCheck").Value = 0
        wg.Range("MDMCheck").Value = "[{{0x7effaf}}]"
        Call GetMatches
        Call MDM
    ElseIf wg.Range("FullProcessCheck").Value = 1 Then
        ' Partially approved; skip TBA extras
        wg.Range("FullProcessCheck").Value = 0
    Else
        ' Not approved; cancel
        Exit Sub
    End If
End If

' Copies team numbers from team list
wt.Activate
wt.Range("B3:B502").Select
Selection.Copy
wp.Activate
wp.Range("Y2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

' Variable config
row = wp.Range("X3").Value
col = wp.Range("X4").Value
fact = wt.Range("INDEX").Value
check = 0

' Places a 'stop value' at the end of the copied team list (19gen)
Do Until check = 1
    wp.Cells(row, col).Select
    If Selection.Value = "" Then
        Selection.Value = "{[{0x7effaf}]}"
        check = 1
    Exit Do
    Else
        row = row + 1
    End If
Loop

' Variable reconfig
row = wp.Range("X3").Value
ctr = 0
check = 0

' For each team... process data:
Do Until check = 1
    wp.Activate
    wp.Cells(row, col).Select
    
    ' Checks if the selected cell is the 'stop value'
    If Selection.Value = "{[{0x7effaf}]}" Then
        ' If so, ends the Do loop
        Let check = 1
        Selection.Clear
    Exit Do
    Else
        ' If not, copies the selected cell into "DP19" to begin data processing
        Selection.Copy
        wp.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        
        ' Copies data from "DP19" to "Storage"
        wp.Range("ProcessorCore").Copy
        ws.Activate
        ws.Cells(wt.Range("InitialIndex").Value + 1 + fact * ctr, 2).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        
        ' Increments the variables
        Let row = row + 1
        Let ctr = ctr + 1
    End If
Loop
    
' Spring cleaning
wp.Activate
wp.Range("A2").Value = "X"
wp.Range("Y2:Y501").Clear

' View cleanup
wp.Activate
wp.Range("C3").Select
wp.Range("A1").Select
ws.Activate
ws.Range("C3").Select
ws.Range("A1").Select
wt.Activate
wt.Range("C3").Select
wt.Range("A1").Select
wg.Activate
wg.Range("C3").Select
wg.Range("A1").Select
active.Activate

End Sub
