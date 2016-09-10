Private Sub Command1_Click()
  If Text1.Text = "123" Then
  Form1.Show
  Unload Me
Else
  MsgBox "密码错误！", , "提示"
End If

End Sub
