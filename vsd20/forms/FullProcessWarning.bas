'
' Mousey move stuffs:
'
Sub btnRunOff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

' Highlights "Run All" button
btnCancelOff.Visible = True
btnOnlyOff.Visible = True
btnRunOff.Visible = False

End Sub
Sub btnCancelOff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

' Highlights "Cancel" button
btnRunOff.Visible = True
btnOnlyOff.Visible = True
btnCancelOff.Visible = False

End Sub
Sub btnOnlyOff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

' Highlights "Process Only" button
btnCancelOff.Visible = True
btnRunOff.Visible = True
btnOnlyOff.Visible = False

End Sub
Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

' Resets all buttons
btnRunOff.Visible = True
btnOnlyOff.Visible = True
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
Worksheets("GUIDE").Range("FullProcessCheck").Value = 2

' Unloads the UserForm
Unload Me

End Sub
Private Sub btnOnlyOn_Click()

' Approves only a regular process
Worksheets("GUIDE").Range("FullProcessCheck").Value = 1

' Unloads the UserForm
Unload Me

End Sub
