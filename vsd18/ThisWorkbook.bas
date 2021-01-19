Private Sub Workbook_Open()
' Stability macro: prevents calculation lag
    Sheets("Data Processor").EnableCalculation = False
End Sub
