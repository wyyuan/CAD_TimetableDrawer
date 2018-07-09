VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "请输入站间距"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Cancel_Click()
Me.Hide
UserForm1.Show
End Sub

Private Sub Cmd_OK_Click()
Dim i As Integer
For i = 0 To 2
station(i) = TextBox1.Text
Next i
Me.Hide
UserForm1.Show
End Sub
