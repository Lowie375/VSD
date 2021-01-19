' Highlights "Cancel" button
btnSubmitOff.Visible = True
btnCancelOff.Visible = False

End Sub
Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

' Resets both buttons
btnSubmitOff.Visible = True
btnCancelOff.Visible = True

End Sub
'
' Clicky stuffs:
'
Private Sub btnCancelOn_Click()

' Unloads the UserForm
Unload Me

End Sub
Private Sub btnSubmitOn_Click()

' Checks if CONFIRM was entered
If Me.tbxConfirm.Value = "CONFIRM" Then
    ' Approves the operation
    Worksheets("GUIDE").Range("VSDClearCheck") = True
End If

' Unloads the UserForm
Unload Me

End Sub
