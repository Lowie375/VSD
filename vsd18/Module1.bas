Sub Process()
' Stability macro: prevents calculation lag
' Added 01/18/21 for VSD archive

' Shortcut: Ctrl+Shift+Q
    
    Sheets("Data Processor").EnableCalculation = True
    ActiveSheet.Calculate
    Sheets("Data Processor").EnableCalculation = False
End Sub
