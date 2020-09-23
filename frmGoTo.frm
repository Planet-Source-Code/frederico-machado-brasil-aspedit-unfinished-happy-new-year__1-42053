VERSION 5.00
Begin VB.Form frmGoTo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Goto Line"
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2925
   Icon            =   "frmGoTo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   53
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "GO!"
      Default         =   -1  'True
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txtLine 
      Height          =   285
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "0"
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Line number:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   270
      Width           =   915
   End
End
Attribute VB_Name = "frmGoTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  With frmMain.ASPEdit
    lngStart = SendMessage(.hWnd, EM_LINEINDEX, CLng(txtLine) - 1, 0&)
    If lngStart = -1 Then 'Invalid line number
      MsgBox "Can't go. The Line number is invalid.", vbCritical, "Invalid Line"
      Exit Sub
    End If
    .SelStart = lngStart 'Go To line
  End With
End Sub

Private Sub txtLine_GotFocus()
  txtLine.SelStart = 0
  txtLine.SelLength = Len(txtLine)
End Sub
