VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Picture slide Effects"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command13 
      Caption         =   "TNC13"
      Height          =   255
      Left            =   6075
      TabIndex        =   15
      Top             =   3315
      Width           =   945
   End
   Begin VB.CommandButton Command12 
      Caption         =   "TNC12"
      Height          =   255
      Left            =   6060
      TabIndex        =   14
      Top             =   2985
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Height          =   3615
      Left            =   240
      ScaleHeight     =   3555
      ScaleWidth      =   4635
      TabIndex        =   13
      Top             =   240
      Width           =   4695
   End
   Begin VB.CommandButton Command11 
      Caption         =   "9"
      Height          =   255
      Left            =   6600
      TabIndex        =   12
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "11"
      Height          =   255
      Left            =   6210
      TabIndex        =   11
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "8"
      Height          =   255
      Left            =   6600
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "10"
      Height          =   255
      Left            =   6600
      TabIndex        =   9
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   255
      Left            =   6600
      TabIndex        =   8
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   255
      Left            =   5760
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   255
      Left            =   5760
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.PictureBox Picture33 
      Height          =   3615
      Left            =   11175
      Picture         =   "Form1.frx":0ECA
      ScaleHeight     =   3555
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   1875
      Width           =   4695
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   240
      Picture         =   "Form1.frx":9714
      ScaleHeight     =   3555
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim barnumber As Integer, barwidth As Integer
Dim i, j As Integer
barwidth = 300
Picture2.Picture = Picture33.Picture
barnumber = Picture1.ScaleWidth / barwidth
On Error Resume Next
For i = 1 To barwidth
For j = 0 To barnumber
Picture2.PaintPicture Picture1.Picture, j * barwidth, 0, i, Picture1.ScaleHeight, j * barwidth, 0, i, Picture1.ScaleHeight, &HCC0029
DoEvents
Next
Next
End Sub

Private Sub Command10_Click()
Dim barwidth, barheight As Integer
Dim i As Single
If Picture1.ScaleWidth > Picture1.ScaleHeight Then
barwidth = Picture1.ScaleWidth - Picture1.ScaleHeight
barheight = 1
ElseIf Picture1.ScaleWidth < Picture1.ScaleHeight Then
barwidth = 1
barheight = Picture1.ScaleHeight - Picture1.ScaleWidth
Else
barwidth = 1: barheight = 1
End If
On Error Resume Next
For i = 1 To Picture1.ScaleWidth - barwidth
Picture2.PaintPicture Picture1.Picture, Int((Picture1.ScaleWidth - barwidth) / 2), Int((Picture1.ScaleHeight - barheight) / 2), barwidth, barheight, Int((Picture1.ScaleWidth - barwidth) / 2), barwidth, barheight, &HCC0020
barwidth = barwidth + 1: barheight = barheight + 1
DoEvents
Next
End Sub

Private Sub Command11_Click()
Dim barwidth As Integer, barheight As Integer
Dim i As Integer
Picture2.Picture = Picture33.Picture
barwidth = 1: barheight = Picture1.ScaleHeight
On Error Resume Next
For i = 1 To Picture1.ScaleWidth / 2
Picture2.PaintPicture Picture1.Picture, (Picture1.ScaleWidth - barwidth) / 2, 0, barwidth, barheight, (Picture1.ScaleWidth - barwidth) / 2, 0, barwidth, barheight, &HCC0020
barwidth = barwidth + 4
DoEvents
Next
End Sub


Private Sub Command12_Click()
Dim barnumber As Integer, barwidth, barheight As Integer
Dim i, j As Integer
barwidth = 300
barheigth = 300
Picture2.Picture = Picture33.Picture
barnumber = Picture1.ScaleWidth / barwidth
On Error Resume Next
For i = 1 To barwidth / 2 Step 3
For j = 0 To barnumber
Picture2.PaintPicture Picture1.Picture, j * barwidth, 0, i, Picture1.ScaleHeight, j * barwidth, 0, i, Picture1.ScaleHeight, &HCC0029
DoEvents
Next
Next
For i = 1 To barheigth Step 11
For j = 0 To barnumber
Picture2.PaintPicture Picture1.Picture, 0, j * barheigth, Picture1.ScaleWidth, i, 0, j * barheigth, Picture1.ScaleWidth, i, &HCC0020
DoEvents
Next
Next

End Sub

Private Sub Command13_Click()
Dim barnumber As Integer, barwidth, barheight As Integer
Dim i, j As Integer
barwidth = 300
barheigth = 300
On Error Resume Next
Picture2.Picture = Picture33.Picture
For i = 1 To Picture1.ScaleWidth Step 12
Picture2.PaintPicture Picture1.Picture, 0, 0, i, Picture1.ScaleHeight, 0, 0, i, Picture1.ScaleHeight, &HCC0020
Picture2.PaintPicture Picture1.Picture, 0, j * barheigth, Picture1.ScaleWidth, i, 0, j * barheigth, Picture1.ScaleWidth, i, &HCC0020
DoEvents
Next

End Sub

Private Sub Command14_Click()

End Sub

Private Sub Command2_Click()
Dim barnumber As Integer, barheigth As Integer
Dim i, j As Integer
barheigth = 300
Picture2.Picture = Picture33.Picture
barnumber = Picture1.ScaleHeight / barheigth
On Error Resume Next
For i = 1 To barheigth
For j = 0 To barnumber
Picture2.PaintPicture Picture1.Picture, 0, j * barheigth, Picture1.ScaleWidth, i, 0, j * barheigth, Picture1.ScaleWidth, i, &HCC0020
DoEvents
Next
Next
End Sub

Private Sub Command3_Click()
Dim i As Single
On Error Resume Next
For i = 1 To Picture1.ScaleWidth Step 12
Picture2.PaintPicture Picture1.Picture, 0, 0, i, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, &HCC0020
Next
End Sub

Private Sub Command4_Click()
Dim i As Single
On Error Resume Next
Picture2.Picture = Picture33.Picture
For i = 1 To Picture1.ScaleWidth Step 12
Picture2.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 0, i, Picture1.ScaleHeight, &HCC0020
DoEvents
Next
End Sub

Private Sub Command5_Click()
Dim i As Single
On Error Resume Next
Picture2.Picture = Picture33.Picture
For i = 1 To Picture1.ScaleWidth Step 12
Picture2.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth, i, &HCC0020
DoEvents
Next
End Sub

Private Sub Command6_Click()
Dim i As Single
On Error Resume Next
Picture2.Picture = Picture33.Picture
For i = 1 To Picture1.ScaleWidth Step 12
Picture2.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, i, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, &HCC0020
DoEvents
Next
End Sub

Private Sub Command7_Click()
Dim i As Single
On Error Resume Next
Picture2.Picture = Picture33.Picture
For i = 1 To Picture1.ScaleWidth Step 12
Picture2.PaintPicture Picture1.Picture, 0, 0, i, Picture1.ScaleHeight, 0, 0, i, Picture1.ScaleHeight, &HCC0020
DoEvents
Next
End Sub

Private Sub Command8_Click()
Dim barwidth, barheight As Integer
Dim i As Integer

barwidth = Picture1.ScaleWidth
barheight = 1
On Error Resume Next
For i = 1 To Picture1.ScaleHeight / 2
Picture2.PaintPicture pictrue1.Picture, 0, (Picture1.ScaleHeight - barheight) / 2, barwidth, barheight, 0, (Picture1.ScaleHeight - barheight) / 2, barwidth, barheight, &HCC0020
barheight = barheight + 4
DoEvents
Next
End Sub

Private Sub Command9_Click()
Dim i As Single
On Error Resume Next
Picture2.Picture = Picture33.Picture
For i = 1 To Picture1.ScaleWidth Step 12
Picture2.PaintPicture Picture1.Picture, Picture1.ScaleWidth - i, 0, i, Picture1.ScaleHeight, Picture1.ScaleWidth - i, 0, i, Picture1.ScaleHeight, &HCC0020
DoEvents
Next
End Sub












Private Sub Form_Load()

End Sub
