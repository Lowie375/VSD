'
' Mousey move stuffs:
'
Sub btnRunOff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

' Highlights "Run" button
btnCancelOff.Visible = True
btnRunOff.Visible = False

End Sub
Sub btnCancelOff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

' Highlights "Cancel" button
btnRunOff.Visible = True
btnCancelOff.Visible = False

End Sub
Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

' Resets both buttons
btnRunOff.Visible = True
btnCancelOff.Visible = True

End Sub
'
' Clicky stuffs:
'
Private Sub btnCancelOn_Click()

' Unloads the UserForm
Unload Me

End Sub
Private Sub btnRunOn_Click()

' Approves the operation
Worksheets("GUIDE").Range("MDMCheck").Value = True

' Unloads the UserForm
Unload Me

End Sub
