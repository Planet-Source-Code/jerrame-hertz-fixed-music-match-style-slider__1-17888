VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "Picclp32.ocx"
Begin VB.Form Form1 
   Caption         =   "Slider"
   ClientHeight    =   2610
   ClientLeft      =   10725
   ClientTop       =   3660
   ClientWidth     =   1680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   1680
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   1080
      ScaleHeight     =   1695
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   3000
      Top             =   240
      _ExtentX        =   714
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   2
      Picture         =   "Form1.frx":044A
   End
   Begin VB.Line Line2 
      X1              =   960
      X2              =   1200
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   960
      X2              =   1200
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   960
      X2              =   1200
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With Picture1   ' Set the thumb picture and position
        .Picture = PictureClip1.GraphicCell(0)
        .Left = Text1.Left + Text1.Width + 20
    End With
    With Picture2   ' Set the slider position and width
        .Left = Picture1.Left + 50
        .Width = Picture1.Width - 100
    End With
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' swap the image for down position and show the slider value
    Picture1.Picture = PictureClip1.GraphicCell(1)
    Text1.Visible = True
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TheTop As Integer ' dimesion for the thumb position
    If Button = 1 Then
        If Y < 0 Then ' Set the lowest value for the thumb
            Picture1.Top = Picture2.Top - (Picture1.Height / 2)
            Text1.Top = Picture1.Top
        ElseIf Y > 1695 Then    ' Set the highest value for the thumb
            Picture1.Top = Picture2.Top + Picture2.Height - (Picture1.Height / 2)
            Text1.Top = Picture1.Top
        Else ' allow the thumb to slide within lowest and highest values
            TheTop = Y
            Picture1.Top = TheTop + Picture2.Top - (Picture1.Height / 2)
            Text1.Top = Picture1.Top
            Picture2.BackColor = TheTop * 0.15
            Text1.Text = Fix(TheTop / 169.5)
        End If
    End If
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' swap the image for up position and hide the slider value
    Picture1.Picture = PictureClip1.GraphicCell(0)
    Text1.Visible = False
End Sub
