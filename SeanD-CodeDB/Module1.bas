Attribute VB_Name = "Module1"
Global user As String
Global password As String
Global group As Integer
Global score As Integer
Global test As Date


Sub Logout()

frmLogin.Show

Module1.user = ""
Module1.password = ""
Module1.group = 12
Module1.score = 12

End Sub

