Sub GetTeams()

' GetTeams - TBA Team Puller [xTBA]
' Gets the team list for a specified event from thebluealliance.com

' Shortcut: Ctrl+Shift+T
' Last updated on 01/17/21

' +--------------------------------------------+
' | TEAM PULL MACRO [xTBA] - EDIT WITH CAUTION |
' +--------------------------------------------+

' Declarations...
Dim aReq  As Object, xJson As Object
Dim r As Long, c As Long, rSave As Long
Dim key As String, xURL As String, response As String, token As String, lmSHT As String, xHead As String, t As String

' ...and more declarations...
Dim wt As Worksheet, wj As Worksheet, active As Worksheet
Set wt = Worksheets("Teams")
Set wj = Worksheets("JSON")
Set active = ActiveSheet

' ...and even more declarations...
key = LCase(wt.Range("ECode").Value)
token = wt.Range("TOKEN").Value

xURL = "https://www.thebluealliance.com/api/v3/event/" & key & "/teams"
wj.Activate
If wj.Range("TP.Origin") <> "" And xURL <> wj.Range("TP.Origin") Then
    wj.Range("TP.LM").ClearContents
End If
lmSHT = wj.Range("TP.LM").Value
' Okay, that's all of them now.

' Sends the HTTP request
Set aReq = CreateObject("MSXML2.XMLHTTP")
    With aReq
        .Open "GET", xURL, False
        .SetRequestHeader "X-TBA-Auth-Key", token
        If lmSHT <> "" Then
            .SetRequestHeader "If-Modified-Since", lmSHT
        End If
        .Send
    End With

xHead = aReq.GetAllResponseHeaders()

wj.Range("TP.Origin") = xURL
wj.Range("TP.Status") = aReq.Status & ": " & aReq.StatusText
wj.Range("TP.Stamp") = aReq.GetResponseHeader("Date")

' Checks for unwanted statuses
If aReq.Status = 304 Then ' "Not Modified"
    ' Loads dialog box
    Stat304Msg.Show
    
    ' View cleanup & macro escape
    wj.Activate
    wj.Range("C3").Select
    wj.Range("A1").Select
    wt.Activate
    wt.Range("C3").Select
    wt.Range("A1").Select
    active.Activate
    Application.CutCopyMode = False
    Exit Sub
ElseIf aReq.Status = 401 Then ' "Unauthorized"
    ' Loads dialog box
    Stat401Msg.Show

    ' View cleanup & macro escape
    wj.Activate
    wj.Range("C3").Select
    wj.Range("A1").Select
    wt.Activate
    wt.Range("C3").Select
    wt.Range("A1").Select
    active.Activate
    Application.CutCopyMode = False
    Exit Sub
ElseIf aReq.Status = 404 Then ' "Not Found"
    ' Loads dialog box
    Stat404xMsg.Show

    ' View cleanup & macro escape
    wj.Activate
    wj.Range("C3").Select
    wj.Range("A1").Select
    wt.Activate
    wt.Range("C3").Select
    wt.Range("A1").Select
    active.Activate
    Application.CutCopyMode = False
    Exit Sub
ElseIf aReq.Status <> 200 Then ' Other errors
    ' View cleanup & macro escape
    wj.Activate
    wj.Range("C3").Select
    wj.Range("A1").Select
    wt.Activate
    wt.Range("C3").Select
    wt.Range("A1").Select
    active.Activate
    Application.CutCopyMode = False
    Exit Sub
End If

' Parses data
response = "{""data"":" & aReq.ResponseText & "}"
Set xJson = JsonConverter.ParseJson(response)

wj.Range("TP.LM") = aReq.GetResponseHeader("Last-Modified")
wj.Range("TP.CC") = aReq.GetResponseHeader("Cache-Control")

' Spring cleaning
wj.Range("TP.Output").ClearContents

' Data puller
r = Range("TP.Rows")
For Each Item In xJson("data")
    c = wj.Range("TP.Cols")
    wj.Cells(r, c) = Item("team_number")
    wj.Cells(r, c + 1) = Item("nickname")
    
    ' Location builder
    Dim loc As String
    loc = ""
    If Item("city") <> "" Then
        loc = Item("city")
    End If
    If Item("state_prov") <> "" And loc <> "" Then
        loc = loc & ", " & Item("state_prov")
    ElseIf Item("state_prov") <> "" And loc = "" Then
        loc = Item("state_prov")
    End If
    If Item("country") <> "" And loc <> "" Then
        loc = loc & ", " & Item("country")
    ElseIf Item("country") <> "" And loc = "" Then
        loc = Item("country")
    End If
    wj.Cells(r, c + 2) = loc
    
    wj.Cells(r, c + 3) = Item("website")
    wj.Cells(r, c + 4) = Item("rookie_year")
    r = r + 1
Next
' Sorts team numbers in numerical order
wj.Sort.SortFields.Clear
wj.Sort.SortFields.Add key:=Range("A9"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With wj.Sort
    .SetRange Range("TP.Output")
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

' Adds "stop values" to the end of the list
wj.Cells(r, c) = "[{[0x7effaf]}]"
wj.Cells(r, c + 3) = "[[{0x7effaf}]]"
rSave = r

' Copy/paste + spring cleaning
wt.Activate
wt.Range("B3:F502").Select
Selection.ClearContents
wj.Activate
wj.Range("TP.Output").Select
Selection.Copy
wt.Activate
wt.Range("B3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

' TBA team link inserter
r = 3
c = 2
t = wt.Cells(r, c).Value
Do While t <> "[{[0x7effaf]}]"
    ' Select and link the team's TBA page
    wt.Cells(r, c).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="https://www.thebluealliance.com/team/" & t & "/2020", TextToDisplay:=t
    
    ' Reset formatting
    Selection.Font.ColorIndex = xlColorIndexAutomatic
    Selection.Font.Underline = False
    Selection.Font.Bold = True
    
    ' Increment
    r = r + 1
    t = wt.Cells(r, c).Value
Loop
wt.Cells(r, c).ClearContents

' Team website link inserter
r = 3
c = 5
t = wt.Cells(r, c).Value
Do While t <> "[[{0x7effaf}]]"
    If t <> "" Then
        ' Select and link the team's TBA page
        wt.Cells(r, c).Select
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=t, TextToDisplay:=t
    
        ' Reset formatting
        Selection.Font.ColorIndex = xlColorIndexAutomatic
        Selection.Font.Underline = False
        Selection.Font.Italic = True
    End If
    
    ' Increment
    r = r + 1
    t = wt.Cells(r, c).Value
Loop
wt.Cells(r, c).ClearContents

' Removes the stop values from "JSON"
wj.Activate
wj.Cells(rSave, wj.Range("TP.Cols").Value).ClearContents
wj.Cells(rSave, wj.Range("TP.Cols").Value + 3).ClearContents

If wt.Range("AutoInitCheck") = True Then
    Call SetWorkbook
Else
    ' Add re-init warning message
    wt.Activate
    wt.Range("C1").Select
    Selection.Font.Color = RGB(227, 45, 145)
    Selection.Value = wt.Range("A23").Value

    ' View cleanup
    wj.Activate
    wj.Range("C3").Select
    wj.Range("A1").Select
    wt.Activate
    wt.Range("C3").Select
    wt.Range("A1").Select
    Application.CutCopyMode = False
    active.Activate
End If

End Sub
Sub GetMatches()

' GetMatches - TBA Match Puller [xTBA] [WIP]
' Gets match data for a specified event from thebluealliance.com

' Shortcut: Ctrl+Shift+M
' Last updated on 01/17/21

' +---------------------------------------------+
' | MATCH PULL MACRO [xTBA] - EDIT WITH CAUTION |
' +---------------------------------------------+

' Declarations...
Dim aReq As Object, xJson As Object
Dim reds() As Integer, blues() As Integer
Dim r As Long, c As Long, i As Long
Dim key As String, xURL As String, response As String, token As String, lmSHT As String, xHead As String, attribs() As String, link As String

' ...and some more declarations!
Dim wt As Worksheet, wj As Worksheet, wi As Worksheet, active As Worksheet
Set wt = Worksheets("Teams")
Set wj = Worksheets("JSON")
Set wi = Worksheets("INPUT")
Set active = ActiveSheet
link = "initLineRobot1,initLineRobot2,initLineRobot3,endgameRobot1,endgameRobot2,endgameRobot3,stage2Activated,stage3Activated,endgameRungIsLevel,stage3TargetColor,rp"
attribs = Split(link, ",")

wt.Activate
wj.Activate

key = LCase(wt.Range("ECode").Value)
token = wt.Range("TOKEN").Value

xURL = "https://www.thebluealliance.com/api/v3/event/" & key & "/matches"
If wj.Range("MP.Origin") <> "" And xURL <> wj.Range("MP.Origin") Then
    wj.Range("MP.LM").ClearContents
End If
lmSHT = wj.Range("MP.LM").Value

Set aReq = CreateObject("MSXML2.XMLHTTP")
    With aReq
        .Open "GET", xURL, False
        .SetRequestHeader "X-TBA-Auth-Key", token
        If lmSHT <> "" Then
            .SetRequestHeader "If-Modified-Since", lmSHT
        End If
        .Send
    End With

xHead = aReq.GetAllResponseHeaders()

wj.Range("MP.Origin") = xURL
wj.Range("MP.Status") = aReq.Status & ": " & aReq.StatusText
wj.Range("MP.Stamp") = aReq.GetResponseHeader("Date")

' Checks for unwanted statuses
If aReq.Status = 304 Then ' "Not Modified"
    ' Loads dialog box
    Stat304Msg.Show
    
    ' View cleanup & macro escape
    wj.Activate
    wj.Range("C3").Select
    wj.Range("A1").Select
    wt.Activate
    wt.Range("C3").Select
    wt.Range("A1").Select
    wi.Activate
    wi.Range("C3").Select
    wi.Range("A1").Select
    active.Activate
    Application.CutCopyMode = False
    Exit Sub
ElseIf aReq.Status = 401 Then ' "Unauthorized"
    ' Loads dialog box
    Stat401Msg.Show
    
    ' View cleanup & macro escape
    wj.Activate
    wj.Range("C3").Select
    wj.Range("A1").Select
    wt.Activate
    wt.Range("C3").Select
    wt.Range("A1").Select
    wi.Activate
    wi.Range("C3").Select
    wi.Range("A1").Select
    active.Activate
    Application.CutCopyMode = False
    Exit Sub
ElseIf aReq.Status = 404 Then ' "Not Found"
    ' Loads dialog box
    Stat404xMsg.Show

    ' View cleanup & macro escape
    wj.Activate
    wj.Range("C3").Select
    wj.Range("A1").Select
    wt.Activate
    wt.Range("C3").Select
    wt.Range("A1").Select
    wi.Activate
    wi.Range("C3").Select
    wi.Range("A1").Select
    active.Activate
    Application.CutCopyMode = False
    Exit Sub
ElseIf aReq.Status <> 200 Then ' Other errors
    ' Loads dialog box
    GenericErrorMsg.Show

    ' View cleanup & macro escape
    wj.Activate
    wj.Range("C3").Select
    wj.Range("A1").Select
    wt.Activate
    wt.Range("C3").Select
    wt.Range("A1").Select
    wi.Activate
    wi.Range("C3").Select
    wi.Range("A1").Select
    active.Activate
    Application.CutCopyMode = False
    Exit Sub
End If

response = "{""data"":" & aReq.ResponseText & "}"
Set xJson = JsonConverter.ParseJson(response)

wj.Range("MP.LM") = aReq.GetResponseHeader("Last-Modified")
wj.Range("MP.CC") = aReq.GetResponseHeader("Cache-Control")

' Spring cleaning
wj.Range("MP.Output").ClearContents

' Data puller
r = Range("MP.Rows")
For Each ItemX In xJson("data")
    c = wj.Range("MP.Cols")
    If ItemX("comp_level") = "qm" And ItemX("actual_time") <> "Null" Then
        wj.Cells(r, c) = ItemX("match_number")
        wj.Cells(r, c + 1) = ItemX("comp_level")
        If ItemX("winning_alliance") = "red" Then
            wj.Cells(r, c + 2) = "R"
        ElseIf ItemX("winning_alliance") = "blue" Then
            wj.Cells(r, c + 2) = "B"
        Else
            wj.Cells(r, c + 2) = "T"
        End If
        i = 3
        
            ' RED
            For Each ItemR In ItemX("alliances")("red")("team_keys")
                wj.Cells(r, c + i) = Mid(ItemR, 4)
                i = i + 1
            Next
            For Each a In attribs
                wj.Cells(r, c + i) = ItemX("score_breakdown")("red")(a)
                i = i + 1
            Next
            wj.Cells(r, c + i) = ItemX("alliances")("red")("score")
            i = i + 1
    
            ' BLUE
            For Each ItemB In ItemX("alliances")("blue")("team_keys")
                wj.Cells(r, c + i) = Mid(ItemB, 4)
                i = i + 1
            Next
            For Each a In attribs
                wj.Cells(r, c + i) = ItemX("score_breakdown")("blue")(a)
                i = i + 1
            Next
            wj.Cells(r, c + i) = ItemX("alliances")("blue")("score")
    
            ' ItemX("alliances")("red"/"blue")("score")
            ' ("data")(i)("score_breakdown")("red"/"blue")
            ' ("data")(i)("alliances")("red"/"blue")("team_keys")
            r = r + 1

    End If
Next

' View cleanup
wj.Activate
wj.Range("C3").Select
wj.Range("A1").Select
wt.Activate
wt.Range("C3").Select
wt.Range("A1").Select
wi.Activate
wi.Range("C3").Select
wi.Range("A1").Select
Application.CutCopyMode = False
active.Activate

End Sub
Sub UselessButton()

' UselessButton - Button Pusher Macro
' Useless macro for button pushing purposes. Because pushing buttons is fun.

' Shortcut: N/A
' Last updated on 01/17/21

' This is an easter egg that lets useless buttons be pushed. Yay!
' I've always wanted to add one of these to a project of mine.

' Thanks for being a curious being, or for trying to fix my broken code and stumbling upon this in the process, haha.
' Keep doing what you're doing. Rock on. :)

' - nxo / N
'   Lead Scout, an undisclosed FRC team

End Sub
