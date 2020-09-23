VERSION 5.00
Begin VB.Form frmPProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Page Properties"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "frmPProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   328
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   473
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBaseTarget 
      Height          =   285
      Left            =   1680
      TabIndex        =   17
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CheckBox chkLoop 
      Caption         =   "Sound Loop"
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowseSound 
      Caption         =   "Browse..."
      Height          =   285
      Left            =   3000
      TabIndex        =   15
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtBGSound 
      Height          =   285
      Left            =   1680
      TabIndex        =   14
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtMarginHeight 
      Height          =   285
      Left            =   4200
      MaxLength       =   7
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtMarginWidth 
      Height          =   285
      Left            =   4200
      MaxLength       =   7
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtTopMargin 
      Height          =   285
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   11
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtLeftMargin 
      Height          =   285
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.PictureBox picActiveColor 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4200
      Picture         =   "frmPProperties.frx":000C
      ScaleHeight     =   270
      ScaleWidth      =   270
      TabIndex        =   32
      Top             =   2160
      Width           =   270
   End
   Begin VB.PictureBox picVisitedColor 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4200
      Picture         =   "frmPProperties.frx":0094
      ScaleHeight     =   270
      ScaleWidth      =   270
      TabIndex        =   31
      Top             =   1800
      Width           =   270
   End
   Begin VB.PictureBox picLinkColor 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1680
      Picture         =   "frmPProperties.frx":011C
      ScaleHeight     =   270
      ScaleWidth      =   270
      TabIndex        =   30
      Top             =   2160
      Width           =   270
   End
   Begin VB.PictureBox picTextColor 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1680
      Picture         =   "frmPProperties.frx":01A4
      ScaleHeight     =   270
      ScaleWidth      =   270
      TabIndex        =   29
      Top             =   1800
      Width           =   270
   End
   Begin VB.PictureBox picBGColor 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1680
      Picture         =   "frmPProperties.frx":022C
      ScaleHeight     =   270
      ScaleWidth      =   270
      TabIndex        =   28
      Top             =   1440
      Width           =   270
   End
   Begin VB.TextBox txtActiveColor 
      Height          =   285
      Left            =   4560
      MaxLength       =   7
      TabIndex        =   9
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtVisitedColor 
      Height          =   285
      Left            =   4560
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtLinkColor 
      Height          =   285
      Left            =   2040
      MaxLength       =   7
      TabIndex        =   7
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtTextColor 
      Height          =   285
      Left            =   2040
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtBGColor 
      Height          =   285
      Left            =   2040
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.OptionButton optBGFixed 
      Caption         =   "Fixed"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.OptionButton optBGTiled 
      Caption         =   "Tiled (Default)"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton cmdBGBrowse 
      Caption         =   "Browse..."
      Height          =   285
      Left            =   4440
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtBgImage 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   19
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5880
      TabIndex        =   18
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Base Target:"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   3765
      Width           =   1455
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Background Sound:"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   3405
      Width           =   1455
   End
   Begin VB.Label lblFolder 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   720
      TabIndex        =   38
      Top             =   4440
      Width           =   4695
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Document Folder:"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Margin Height:"
      Height          =   255
      Left            =   3000
      TabIndex        =   36
      Top             =   2925
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Margin Width:"
      Height          =   255
      Left            =   3000
      TabIndex        =   35
      Top             =   2565
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Top Margin:"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   2925
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Left Margin:"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   2565
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Active Links:"
      Height          =   255
      Left            =   3000
      TabIndex        =   27
      Top             =   2205
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Visited Links:"
      Height          =   255
      Left            =   3000
      TabIndex        =   26
      Top             =   1845
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Links:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2205
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Text:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1845
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Background:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1485
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BG Properties:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Background Image:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   765
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   285
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   375
      X2              =   375
      Y1              =   15
      Y2              =   311
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   376
      X2              =   376
      Y1              =   16
      Y2              =   310
   End
End
Attribute VB_Name = "frmPProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Color As Long

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  If GetTitle() <> txtTitle Then
    ChangeTitle txtTitle
    frmMain.txtTitle = txtTitle
  End If
  SetDocInfo
  frmMain.ASPEdit.SelStart = 0
  Unload Me
End Sub

Private Sub Form_Load()

  On Local Error Resume Next
  Dim whole As String, attrib As String
  Dim pos As Long, pos2 As Long
    
  PProp = GetPageProperties
    
  txtTitle = PProp.Title
  txtBgImage = PProp.Background
  If LCase(PProp.BGProperties) = "fixed" Then
    optBGFixed.Value = True
  Else
    optBGTiled.Value = True
  End If
  txtBGColor = PProp.BGColor
  picBGColor.BackColor = HexCol2Long(txtBGColor)
  txtTextColor = PProp.Text
  picTextColor.BackColor = HexCol2Long(txtTextColor)
  txtLinkColor = PProp.Link
  picLinkColor.BackColor = HexCol2Long(txtLinkColor)
  txtVisitedColor = PProp.Visited
  picVisitedColor.BackColor = HexCol2Long(txtVisitedColor)
  txtActiveColor = PProp.Active
  picActiveColor.BackColor = HexCol2Long(txtActiveColor)
  txtLeftMargin = PProp.LeftMargin
  txtTopMargin = PProp.TopMargin
  txtMarginWidth = PProp.MarginWidth
  txtMarginHeight = PProp.MarginHeight
  If strDocFolder = "" Then strDocFolder = App.Path
  lblFolder = strDocFolder
  
  'bgsound
  pos = InStr(1, frmMain.ASPEdit.Text, "<BGSOUND", vbTextCompare)
  pos2 = InStr(pos + 1, frmMain.ASPEdit.Text, ">", vbTextCompare)
  whole = Mid$(frmMain.ASPEdit.Text, pos, pos2 - pos)
  attrib = ReadAttrib("src", whole)
  If attrib <> whole Then txtBGSound.Text = attrib
  attrib = ReadAttrib("loop", whole)
  If attrib = "-1" Then chkLoop.Value = 1

  'base target
  pos = InStr(1, frmMain.ASPEdit.Text, "<BASE", vbTextCompare)
  pos2 = InStr(pos + 1, frmMain.ASPEdit.Text, ">", vbTextCompare)
  whole = Mid$(frmMain.ASPEdit.Text, pos, pos2 - pos)
  attrib = ReadAttrib("target", whole)
  If attrib <> whole Then txtBaseTarget.Text = attrib

End Sub

Private Sub picActiveColor_Click()
  Color = GetColor()
  If Color = -1 Then Exit Sub
  picActiveColor.BackColor = Color
  txtActiveColor = FormatRGBString(picActiveColor.BackColor)
End Sub

Private Sub picBGColor_Click()
  Color = GetColor()
  If Color = -1 Then Exit Sub
  picBGColor.BackColor = Color
  txtBGColor = FormatRGBString(picBGColor.BackColor)
End Sub

Private Sub picLinkColor_Click()
  Color = GetColor()
  If Color = -1 Then Exit Sub
  picLinkColor.BackColor = Color
  txtLinkColor = FormatRGBString(picLinkColor.BackColor)
End Sub

Private Sub picTextColor_Click()
  Color = GetColor()
  If Color = -1 Then Exit Sub
  picTextColor.BackColor = Color
  txtTextColor = FormatRGBString(picTextColor.BackColor)
End Sub

Private Sub picVisitedColor_Click()
  Color = GetColor()
  If Color = -1 Then Exit Sub
  picVisitedColor.BackColor = Color
  txtVisitedColor = FormatRGBString(picVisitedColor.BackColor)
End Sub

Private Sub txtActiveColor_LostFocus()
  picActiveColor.BackColor = HexCol2Long(txtActiveColor)
End Sub

Private Sub txtBGColor_LostFocus()
  picBGColor.BackColor = HexCol2Long(txtBGColor)
End Sub

Private Sub txtLinkColor_LostFocus()
  picLinkColor.BackColor = HexCol2Long(txtLinkColor)
End Sub

Private Sub txtTextColor_LostFocus()
  picTextColor.BackColor = HexCol2Long(txtTextColor)
End Sub

Private Sub txtVisitedColor_LostFocus()
  picVisitedColor.BackColor = HexCol2Long(txtVisitedColor)
End Sub

Sub SetDocInfo()
  On Error Resume Next
  Dim pos As Long, pos2 As Long, whole As String
  'set BGSOUND
  pos = InStr(1, frmMain.ASPEdit.Text, "<bgsound", vbTextCompare)
  pos2 = InStr(pos + 1, frmMain.ASPEdit.Text, ">", vbTextCompare)
  If (pos = 0 Or pos2 = 0) And txtBGSound.Text <> "" Then
    pos = InStr(1, frmMain.ASPEdit.Text, "<HEAD>" & vbCrLf, vbTextCompare) + Len("<HEAD>" & vbCrLf)
    If pos = Len("<HEAD>" & vbCrLf) Then pos = InStr(1, frmMain.ASPEdit.Text, "<HEAD>", vbTextCompare) + Len("<HEAD>")
    If pos = Len("<HEAD>") Then pos = 1
    frmMain.ASPEdit.SelStart = pos - 1
    frmMain.ASPEdit.SelText = "<BGSOUND src=(Wait...)>" & vbCrLf
    pos = InStr(1, frmMain.ASPEdit.Text, "<bgsound", vbTextCompare)
    pos2 = InStr(pos + 1, frmMain.ASPEdit.Text, ">", vbTextCompare)
  End If
  Dim strLoop As String
  strLoop = IIf(chkLoop.Value = 1, " loop=" & Chr(34) & "-1" & Chr(34) & " ", "")
  frmMain.ASPEdit.SelStart = pos - 1
  frmMain.ASPEdit.SelLength = pos2 - pos + 1
  If txtBGSound = "" Then
    If InStr(frmMain.ASPEdit.SelText, "<bgsound") > 0 Then frmMain.ASPEdit.SelText = ""
  Else
    frmMain.ASPEdit.SelText = "<BGSOUND src=" & Chr(34) & txtBGSound.Text & Chr(34) & strLoop & ">"
  End If
nxt2:
  'set BASE target
  pos = InStr(1, frmMain.ASPEdit.Text, "<base", vbTextCompare)
  pos2 = InStr(pos + 1, frmMain.ASPEdit.Text, ">", vbTextCompare)
  If (pos = 0 Or pos2 = 0) And txtBaseTarget.Text <> "" Then
    pos = InStr(1, frmMain.ASPEdit.Text, "<HEAD>" & vbCrLf, vbTextCompare) + Len("<HEAD>" & vbCrLf)
    If pos = Len("<HEAD>" & vbCrLf) Then pos = InStr(1, frmMain.ASPEdit.Text, "<HEAD>", vbTextCompare) + Len("<HEAD>")
    If pos = Len("<HEAD>") Then pos = 1
    frmMain.ASPEdit.SelStart = pos - 1
    frmMain.ASPEdit.SelText = "<BASE target=(Wait...)>" & vbCrLf
    pos = InStr(1, frmMain.ASPEdit.Text, "<base", vbTextCompare)
    pos2 = InStr(pos + 1, frmMain.ASPEdit.Text, ">", vbTextCompare)
  End If
  pos2 = pos2 + 1 '1=len(">")
  whole = Mid$(frmMain.ASPEdit.Text, pos, pos2 - pos)
  If txtBaseTarget.Text = "" Then DelAttrib whole, "target", pos Else SaveAttrib "target", txtBaseTarget.Text, "<BASE", ">"
  'set body things. Complicated...
  pos = InStr(1, frmMain.ASPEdit.Text, "<body", vbTextCompare)
  pos2 = InStr(pos + 1, frmMain.ASPEdit.Text, ">", vbTextCompare)
  If pos = 0 Or pos2 = 0 Then
    pos = InStr(1, frmMain.ASPEdit.Text, "</HTML>", vbTextCompare)
    If pos = 0 Then pos = Len(frmMain.ASPEdit.Text)
    frmMain.ASPEdit.SelStart = pos - 1
    frmMain.ASPEdit.SelText = "<BODY>" & vbCrLf & vbCrLf & "</BODY>" & vbCrLf
    pos = InStr(1, frmMain.ASPEdit.Text, "<body", vbTextCompare)
    pos2 = InStr(pos + 1, frmMain.ASPEdit.Text, ">", vbTextCompare)
  End If
  pos2 = pos2 + 1
  whole = Mid$(frmMain.ASPEdit.Text, pos, pos2 - pos)
  'whole is now the body tag (w/ attributes etc.)
  'If any textbox is empty, delete the attribute in the tag if it already
  'exists. Else just don't add.
  If txtTextColor.Text = "" Then DelAttrib whole, "text", pos Else SaveAttrib "text", txtTextColor.Text
  If txtBGColor.Text = "" Then DelAttrib whole, "bgcolor", pos Else SaveAttrib "bgcolor", txtBGColor.Text
  If txtBgImage.Text = "" Then DelAttrib whole, "background", pos Else SaveAttrib "background", txtBgImage.Text
  If optBGFixed.Value = 0 Then DelAttrib whole, "bgproperties", pos Else SaveAttrib "bgproperties", "fixed"
  If txtLinkColor.Text = "" Then DelAttrib whole, "link", pos Else SaveAttrib "link", txtLinkColor.Text
  If txtVisitedColor.Text = "" Then DelAttrib whole, "vlink", pos Else SaveAttrib "vlink", txtVisitedColor.Text
  If txtActiveColor.Text = "" Then DelAttrib whole, "alink", pos Else SaveAttrib "alink", txtActiveColor.Text
  If txtLeftMargin.Text = "" Then DelAttrib whole, "leftmargin", pos Else SaveAttrib "leftmargin", txtLeftMargin.Text
  If txtTopMargin.Text = "" Then DelAttrib whole, "topmargin", pos Else SaveAttrib "topmargin", txtTopMargin.Text
  If txtMarginHeight.Text = "" Then DelAttrib whole, "marginheight", pos Else SaveAttrib "marginheight", txtMarginHeight.Text
  If txtMarginWidth.Text = "" Then DelAttrib whole, "marginwidth", pos Else SaveAttrib "marginwidth", txtMarginWidth.Text
End Sub

Function DelAttrib(Where As String, ID As String, StartPos As Long) As String
  On Error Resume Next
  Dim pos As Long, pos2 As Long
  pos = InStr(1, Where, " " & ID & "=", vbTextCompare)
  If pos = 0 Then Exit Function
  pos2 = InStr(pos + 1, Where, " ")
  If Mid$(Where, pos + Len(ID) + 3, 1) = Chr(34) Then pos2 = InStr(pos + 1, Where, Chr(34))
  If Mid$(Where, pos + Len(ID) + 3, 1) = "'" Then pos2 = InStr(pos + 1, Where, "'")
  If pos2 = 0 Then pos2 = Len(Where)
  frmMain.ASPEdit.SelStart = StartPos + pos - 2
  frmMain.ASPEdit.SelLength = pos2 - pos
  frmMain.ASPEdit.SelText = ""
End Function

Function SaveAttrib(ID As String, Value As String, Optional tagStart As String = "<BODY", Optional tagEnd As String = ">") As String
  On Error Resume Next
  Dim pos As Long, pos2 As Long
  Dim lStart As Long
  Dim Where As String
  pos = InStr(1, frmMain.ASPEdit.Text, tagStart, vbTextCompare)
  pos2 = InStr(pos + 1, frmMain.ASPEdit.Text, tagEnd, vbTextCompare)
  If pos = 0 Or pos2 = 0 Then Exit Function
  lStart = pos 'save the position of tag in lstart
  Where = Mid$(frmMain.ASPEdit.Text, pos, pos2 + 1 - pos)
  'Where contains the BODY tag now
  pos = 0: pos2 = 0
  'every time this function is called, a new BODY tag has to be
  'calculated as a previous call might have altered it.
  'where contains the entire body tag e.g. <BODY a="b" c="d">
  pos = InStr(1, Where, " " & ID & "=", vbTextCompare)
  'find attrib, e.g. attrib 'name' is searched for as " name="
  If pos = 0 Then GoTo not_existing Else pos = pos + Len(ID) + 2 'skip over the id and get the value part
  pos2 = InStr(pos + 1, Where, " ") - 1 'it'll be -1 if no space exists
  'by default, value ends at the next space. If it is enclosed in " or ', calculate
  'the position of the closing " or '.
  If Mid$(Where, pos, 1) = Chr(34) Then pos2 = InStr(pos + 1, Where, Chr(34))
  If Mid$(Where, pos, 1) = "'" Then pos2 = InStr(pos + 1, Where, "'")
  If pos2 = -1 Then pos2 = Len(Where) Else pos2 = pos2 + 1
  frmMain.ASPEdit.SelStart = lStart + pos - 2
  frmMain.ASPEdit.SelLength = pos2 - pos
  frmMain.ASPEdit.SelText = Chr(34) & Value & Chr(34)
  
  Exit Function
not_existing:
  frmMain.ASPEdit.SelStart = lStart + Len(Where) - 2 'to end of string
  frmMain.ASPEdit.SelText = " " & ID & "=" & Chr(34) & Value & Chr(34)
End Function
