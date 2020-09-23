VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "exit"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do It"
      Height          =   855
      Left            =   720
      TabIndex        =   2
      Top             =   3840
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   5655
      Left            =   3840
      ScaleHeight     =   373
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   373
      TabIndex        =   1
      Top             =   0
      Width           =   5655
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   0
      Top             =   0
      Width           =   3750
   End
   Begin VB.Label Label1 
      Caption         =   "Destination"
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   5
      Top             =   5640
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Source"
      Height          =   615
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   2760
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Sub Command1_Click()
MsgBox "very slov, but interesting, if anybody does thios much faster, do please send me a copy: kozlicki@yahoo.com"
diameter = Picture1.ScaleHeight

Picture2.ScaleWidth = diameter * 2
Picture2.ScaleHeight = diameter * 2

Cls
Do
'nariše krog - make egg
c1 = 6.27 / 0.001
c = 0
For kot = 0 To 6.27 Step 0.001
c = c + 1
x = (Picture1.ScaleWidth) - (Picture1.ScaleWidth / c1) * (c + 10) '15 = neneèrtovan popravek slike
mycolor = GetPixel(Picture1.hdc, x, y)
a = Sin(kot) * diameter
b = Cos(kot) * diameter
SetPixel Picture2.hdc, a + (Picture2.ScaleWidth / 2), b + (Picture2.ScaleHeight / 2), mycolor
'SetPixel Picture1.hdc, X, Y, QBColor(12)
'PSet (a + 100, b + 100), RGB(0, 0, MyColor)
'For i = 0 To 10000
'Next i

Next kot
Picture2.Refresh
diameter = diameter - 1

y = y + 1
x = 0
c = 0

'mycolor = mycolor - mycolor / diameter
Loop Until y = Picture1.ScaleHeight
Beep

End Sub

Private Sub Command2_Click()
End
End Sub
