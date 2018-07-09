VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3060
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

End Sub

Private Sub CmdDraw_Click()
'Dim beginTime, endTime As Double
Dim StartPoint(0 To 2) As Double     ' 定义起点坐标
Dim EndPoint(0 To 2) As Double
StartPoint(0) = 0               ' 提取起点x坐标
StartPoint(1) = 0
StartPoint(2) = 0
' 提取起点y坐标
EndPoint(0) = 100              ' 提取终点x坐标
EndPoint(1) = 100
EndPoint(2) = 0
'绘制直线
Call draw(StartPoint(), EndPoint())
End Sub
'将时间转化为长度
Sub time2lenth()

End Sub
Sub draw(StartPoint() As Double, EndPoint() As Double)
Dim LineObj As AcadLine          ' 定义Line对象
' 设置Line对象
Set LineObj = ThisDrawing.ModelSpace.AddLine(StartPoint, EndPoint)
End Sub

Private Sub Cmdstart_Click()
Dim i, j As Integer
Dim StartPoint(0 To 2) As Double     ' 定义起点坐标
Static station(3) As Double
station(0) = 20
station(1) = 50
station(2) = 70
station(3) = 110
Dim EndPoint(0 To 2) As Double
'+++++++++画2分钟的竖线++++++++++++++先画20条
For i = 0 To 10
StartPoint(0) = i * 20
StartPoint(1) = 0
StartPoint(2) = 0
EndPoint(0) = i * 20
EndPoint(1) = 100
EndPoint(2) = 0
Call draw(StartPoint(), EndPoint())
Next i
'+++++++++画2分钟的横线++++++++++++++
For j = 0 To 3
StartPoint(0) = 0
StartPoint(1) = station(j)
StartPoint(2) = 0
EndPoint(0) = i * 20
EndPoint(1) = station(j)
EndPoint(2) = 0
Call draw(StartPoint(), EndPoint())
Next j
End Sub
