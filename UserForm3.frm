VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim date1, date2 As Date
Dim temp As String
temp = TextBox1.Text & ":" & TextBox2.Text & ":" & TextBox3.Text
date1 = CDate(temp)
temp = TextBox4.Text & ":" & TextBox5.Text & ":" & TextBox6.Text
date2 = CDate(temp)
'Call drawtime(date1, date2, Int(TextBox7.Text), Int(TextBox8.Text))
ZoomExtents
End Sub

Private Sub CommandButton2_Click()
Me.Hide
UserForm1.Show
End Sub
