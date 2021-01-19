Sub ClearData()

' ClearData - Data Deletion Macro
' Clears all raw data in the VSD.

' Shortcut: N/A
' Last updated on 03/03/2020

' +-----------------------------------------+
' | DATA DELETION MACRO - EDIT WITH CAUTION |
' +-----------------------------------------+

' Declarations!
Dim wi As Worksheet, wm As Worksheet, ws As Worksheet, wt As Worksheet, wg As Worksheet, active As Worksheet
Set wi = Worksheets("INPUT")
Set wm = Worksheets("MDM")
Set ws = Worksheets("Storage")
Set wt = Worksheets("Teams")
Set wg = Worksheets("GUIDE")
Set active = ActiveSheet

' Awaits confirmation from the user
wg.Activate
wg.Range("DataClearCheck").Value = False
DataClearWarning.Show

' Checks if the user approved the action
If wg.Range("DataClearCheck").Value = False Then
    ' Not approved; cancel
    Exit Sub
Else
    ' Approved; continue
    wg.Range("DataClearCheck").Value = False
End If

' Clears "INPUT"
wi.Activate
wi.Range("RawData").ClearContents

' Clears "Storage"
ws.Activate
ws.Range(ws.Cells(3, 1), ws.Cells(1 + wt.Range("INDEX").Value * 500, 22)).Clear

' Clears "MDM"
wm.Activate
wm.Range("MDMData").ClearContents

' View cleanup
wi.Activate
wi.Range("C3").Select
wi.Range("A1").Select
ws.Activate
ws.Range("C3").Select
ws.Range("A1").Select
wm.Activate
wm.Range("C3").Select
wm.Range("A1").Select
wg.Activate
wg.Range("C3").Select
wg.Range("A1").Select
active.Activate

End Sub
Sub ClearWorkbookCall()

' ClearWorkbookCall - Call Macro for ClearWorkbook
' Calls ClearWorkbook. That's it.

' Shortcut: N/A
' Last updated on 31/01/20

' Awaits confirmation from the user
VSDClearWarning.Show

End Sub
Sub ClearWorkbook()

' ClearWorkbook - Soft Reset Macro
' Clears all VSD data and standard configurations, performing a "soft reset".

' Shortcut: N/A
' Last updated on 03/03/2020

' +--------------------------------------+
' | SOFT RESET MACRO - EDIT WITH CAUTION |
' +--------------------------------------+

' Declarations!
Dim wi As Worksheet, wm As Worksheet, ws As Worksheet, wt As Worksheet, wa As Worksheet, wl As Worksheet, wj As Worksheet, wg As Worksheet, active As Worksheet
Set wi = Worksheets("INPUT")
Set wm = Worksheets("MDM")
Set ws = Worksheets("Storage")
Set wt = Worksheets("Teams")
Set wa = Worksheets("Averages")
Set wl = Worksheets("Picklist")
Set wj = Worksheets("JSON")
Set wg = Worksheets("GUIDE")
Set active = ActiveSheet

' Awaits confirmation from the user
wg.Activate
wg.Range("VSDClearCheck").Value = False
VSDClearWarning.Show

' Checks if the user approved the action
If wg.Range("VSDClearCheck").Value = False Then
    ' Not approved; cancel
    Exit Sub
Else
    ' Approved; continue
    wg.Range("VSDClearCheck").Value = False
End If

' Clears "INPUT"
wi.Activate
wi.Range("RawData").ClearContents

' Clears "Storage"
ws.Activate
ws.Range(ws.Cells(3, 1), ws.Cells(1 + wt.Range("INDEX").Value * 500, 22)).Clear

' Clears "MDM"
wm.Activate
wm.Range("MDMData").ClearContents

' Clears "Teams"
wt.Activate
wt.Range("B3:F502").ClearContents
wt.Range("I3:Q502").ClearContents
wt.Range("ECode").ClearContents
wt.Range("TOKEN").ClearContents

' Clears "Averages"
wa.Activate
wa.Range("B3:B502").ClearContents

' Clears "Picklist"
wl.Activate
wl.Range("C5:C28").ClearContents
wl.Range("F5:F504").ClearContents
wl.Range("MyTeam").ClearContents

' Clears "JSON"
wj.Activate
wj.Range("TP.Output").ClearContents
wj.Range("MP.Output").ClearContents
wj.Range("C2:C6").ClearContents
wj.Range("L2:L6").ClearContents

' View cleanup
wi.Activate
wi.Range("C3").Select
wi.Range("A1").Select
ws.Activate
ws.Range("C3").Select
ws.Range("A1").Select
wm.Activate
wm.Range("C3").Select
wm.Range("A1").Select
wt.Activate
wt.Range("C3").Select
wt.Range("A1").Select
wa.Activate
wa.Range("C3").Select
wa.Range("A1").Select
wl.Activate
wl.Range("C3").Select
wl.Range("A1").Select
wj.Activate
wj.Range("C3").Select
wj.Range("A1").Select
wg.Activate
wg.Range("C3").Select
wg.Range("A1").Select
active.Activate
End Sub
