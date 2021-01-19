Sub btnOKOff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

' Highlights "OK" button
btnOKOff.Visible = False

End Sub

Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

' Resets "OK" button
btnOKOff.Visible = True

End Sub
'
' Clicky stuffs:
'
Private Sub btnOKOn_Click()

' Cancels the operation
Unload Me

End Sub
