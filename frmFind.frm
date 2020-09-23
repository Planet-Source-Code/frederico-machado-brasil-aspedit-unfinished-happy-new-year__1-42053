VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find and Replace"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton FindButton 
      Caption         =   "Find"
      Default         =   -1  'True
      Height          =   325
      Left            =   3705
      TabIndex        =   2
      Top             =   150
      Width           =   1095
   End
   Begin VB.CommandButton FindNextButton 
      Caption         =   "Find Next"
      Height          =   325
      Left            =   4920
      TabIndex        =   3
      Top             =   150
      Width           =   1095
   End
   Begin VB.CommandButton ReplaceButton 
      Caption         =   "Replace"
      Enabled         =   0   'False
      Height          =   325
      Left            =   3705
      TabIndex        =   4
      Top             =   630
      Width           =   1095
   End
   Begin VB.CommandButton ReplaceAllButton 
      Caption         =   "Replace All"
      Enabled         =   0   'False
      Height          =   325
      Left            =   4920
      TabIndex        =   5
      Top             =   630
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   325
      Left            =   4920
      TabIndex        =   8
      Top             =   1090
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Case sensitive"
      Height          =   240
      Left            =   360
      TabIndex        =   6
      Top             =   1110
      Width           =   1440
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Whole word only"
      Height          =   240
      Left            =   2205
      TabIndex        =   7
      Top             =   1110
      Width           =   1560
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   150
      Width           =   2115
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   630
      Width           =   2100
   End
   Begin VB.Label Label1 
      Caption         =   "Find what"
      Height          =   240
      Left            =   240
      TabIndex        =   10
      Top             =   150
      Width           =   810
   End
   Begin VB.Label Label2 
      Caption         =   "Replace with"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   630
      Width           =   1485
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Position As Integer

Private Sub FindButton_Click()

   Dim FindFlags As Integer

   On Error GoTo BSS_ErrorHandler

   Position = 0
   FindFlags = Check1.Value * 4 + Check2.Value * 2
   Position = frmMain.ASPEdit.Find(Text1.Text, Position + 1, , FindFlags)
   frmMain.ASPEdit.SelLength = Len(Trim$(Text1.Text))
   If Position >= 0 Then
      ReplaceButton.Enabled = True
      ReplaceAllButton.Enabled = True
   Else
      frmMain.ASPEdit.SelStart = 0
      MsgBox "Done. Not found in current document.", vbInformation
      ReplaceButton.Enabled = False
      ReplaceAllButton.Enabled = False
   End If

   Exit Sub

BSS_ErrorHandler:

   If Err.Number > 0 Then
      Resume Next
   End If

End Sub

Private Sub FindNextButton_Click()

   Dim FindFlags As Variant

   On Error GoTo BSS_ErrorHandler

   FindFlags = Check1.Value * 4 + Check2.Value * 2
   Position = frmMain.ASPEdit.Find(Text1.Text, Position + 1, , FindFlags)
   frmMain.ASPEdit.SelLength = Len(Trim$(Text1.Text))
   If Position > 0 Then

   Else
      MsgBox "Done. It doesn't exist more strings in current document.", vbInformation
      ReplaceButton.Enabled = False
      ReplaceAllButton.Enabled = False
   End If

   Exit Sub

BSS_ErrorHandler:

   If Err.Number > 0 Then
      Resume Next
   End If

End Sub

Private Sub Command5_Click()

   On Error GoTo BSS_ErrorHandler

   Unload Me

   Exit Sub

BSS_ErrorHandler:

   If Err.Number > 0 Then
      Resume Next
   End If

End Sub

Private Sub Form_Activate()

   On Error GoTo BSS_ErrorHandler

   Text1.SetFocus

   Exit Sub

BSS_ErrorHandler:

   If Err.Number > 0 Then
      Resume Next
   End If

End Sub

Private Sub Form_GotFocus()

   On Error GoTo BSS_ErrorHandler

   Text1.SetFocus

   Exit Sub

BSS_ErrorHandler:

   If Err.Number > 0 Then
      Resume Next
   End If

End Sub

Private Sub ReplaceButton_Click()

   Dim FindFlags As Integer

   On Error GoTo BSS_ErrorHandler

   frmMain.ASPEdit.SelText = Text2.Text
   FindFlags = Check1.Value * 4 + Check2.Value * 2
   Position = frmMain.ASPEdit.Find(Text1.Text, Position + 1, , FindFlags)
   If Position > 0 Then

   Else
      ReplaceButton.Enabled = False
      ReplaceAllButton.Enabled = False
   End If

   Exit Sub

BSS_ErrorHandler:

   If Err.Number > 0 Then
      Resume Next
   End If

End Sub

Private Sub ReplaceAllButton_Click()

   Dim FindFlags As Integer, I As Integer

   On Error GoTo BSS_ErrorHandler

   I = 0
   FindFlags = Check1.Value * 4 + Check2.Value * 2
   frmMain.ASPEdit.SelText = Text2.Text
   Position = frmMain.ASPEdit.Find(Text1.Text, Position + 1, , FindFlags)
   While Position > 0
      I = I + 1
      frmMain.ASPEdit.SelText = Text2.Text
      Position = frmMain.ASPEdit.Find(Text1.Text, Position + 1, , FindFlags)
   Wend
   If I = 0 Then
      I = 1
   End If
   ReplaceButton.Enabled = False
   ReplaceAllButton.Enabled = False
   MsgBox "Done. " & I & " item(s) replaced.", vbInformation

   Exit Sub

BSS_ErrorHandler:

   If Err.Number > 0 Then
      Resume Next
   End If

End Sub
