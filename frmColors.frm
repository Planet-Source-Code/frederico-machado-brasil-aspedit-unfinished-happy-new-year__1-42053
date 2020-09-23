VERSION 5.00
Begin VB.Form frmColors 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Colors"
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3165
   Icon            =   "frmColors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   153
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   211
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   120
      ScaleHeight     =   300
      ScaleWidth      =   705
      TabIndex        =   1
      Top             =   90
      Width           =   735
   End
   Begin VB.PictureBox picColors 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      Picture         =   "frmColors.frx":000C
      ScaleHeight     =   121
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   211
      TabIndex        =   0
      Top             =   480
      Width           =   3165
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "#FFFFFF"
      Height          =   195
      Left            =   1320
      TabIndex        =   2
      Top             =   150
      Width           =   645
   End
   Begin VB.Line Line3 
      X1              =   210
      X2              =   210
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   210
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   32
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mColor As Long

Private Sub picColors_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    mColor = -1
    Unload Me
  End If
End Sub

Private Sub picColors_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mColor = picColors.point(X, Y)
  picPreview.BackColor = mColor
  lblColor = FormatRGBString(mColor)
End Sub

Private Sub picColors_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mColor = picColors.point(X, Y)
  Unload Me
End Sub

Private Sub picPreview_Click()
  mColor = -1
  Unload Me
End Sub
