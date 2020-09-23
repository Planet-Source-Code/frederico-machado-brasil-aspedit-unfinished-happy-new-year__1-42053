VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTags 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Tags/Codes"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   174
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   204
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabTags 
      Height          =   2640
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   4657
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   4
      TabHeight       =   476
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "ASP Tags"
      TabPicture(0)   =   "frmTags.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstASPTags"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ASP Codes"
      TabPicture(1)   =   "frmTags.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tvASPCodes"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "HTML Tags"
      TabPicture(2)   =   "frmTags.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstHTMLTags"
      Tab(2).ControlCount=   1
      Begin VB.ListBox lstHTMLTags 
         Height          =   2295
         IntegralHeight  =   0   'False
         ItemData        =   "frmTags.frx":0054
         Left            =   -75000
         List            =   "frmTags.frx":0056
         TabIndex        =   3
         Top             =   300
         Width           =   3075
      End
      Begin VB.ListBox lstASPTags 
         Height          =   2295
         IntegralHeight  =   0   'False
         ItemData        =   "frmTags.frx":0058
         Left            =   0
         List            =   "frmTags.frx":005A
         TabIndex        =   2
         Top             =   300
         Width           =   3075
      End
      Begin MSComctlLib.TreeView tvASPCodes 
         Height          =   2295
         Left            =   -75000
         TabIndex        =   1
         Top             =   300
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   4048
         _Version        =   393217
         Indentation     =   0
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "frmTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  
  LoadTags Path & "Data\asp.tags", lstASPTags
  LoadTags Path & "Data\html.tags", lstHTMLTags
  
End Sub

Private Sub Form_Resize()
  tabTags.Width = ScaleWidth + 3
  tabTags.Height = ScaleHeight + 2
  lstASPTags.Width = Width - 105
  lstASPTags.Height = Height - 690
  tvASPCodes.Width = lstASPTags.Width
  tvASPCodes.Height = lstASPTags.Height
  lstHTMLTags.Width = lstASPTags.Width
  lstHTMLTags.Height = lstASPTags.Height
End Sub

Private Sub tabTags_Click(PreviousTab As Integer)
  Caption = tabTags.Caption
End Sub

Sub LoadTags(sFileName As String, lstListBox As ListBox)
  
  Dim sFileData As String
  
  If Not FileExists(sFileName) Then Exit Sub
  
  Open sFileName For Binary As #1
    sFileData = String(LOF(1), 0)
    Get #1, 1, sFileData
  Close #1
  
  Dim sTmpArray() As String, i As Integer
  
  sTmpArray = Split(sFileData, vbCrLf)
  
  For i = 0 To UBound(sTmpArray)
    If sTmpArray(i) = "" Then GoTo Jump
    lstListBox.AddItem sTmpArray(i)
Jump:
  Next
  
End Sub
