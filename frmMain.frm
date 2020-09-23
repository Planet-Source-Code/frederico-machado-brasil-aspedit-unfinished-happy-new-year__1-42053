VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Fredisoft ASPEdit"
   ClientHeight    =   4890
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7530
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   326
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   502
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSave 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   840
      Picture         =   "frmMain.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   120
      Width           =   240
   End
   Begin VB.PictureBox picOpen 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   480
      Picture         =   "frmMain.frx":0C4C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   120
      Width           =   240
   End
   Begin VB.PictureBox picNew 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   120
      Picture         =   "frmMain.frx":130E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   120
      Width           =   240
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1965
      TabIndex        =   5
      Top             =   105
      Width           =   2175
   End
   Begin VB.PictureBox picAEHlp 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   4800
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Label lblTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   45
         Width           =   60
      End
      Begin VB.Line lBottom 
         X1              =   0
         X2              =   144
         Y1              =   20
         Y2              =   20
      End
      Begin VB.Line lRight 
         X1              =   144
         X2              =   144
         Y1              =   0
         Y2              =   21
      End
      Begin VB.Line lLeft 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   21
      End
      Begin VB.Line lTop 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   144
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Timer tmrLines 
      Interval        =   25
      Left            =   5760
      Top             =   840
   End
   Begin VB.PictureBox picLines 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   -75
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   1
      Top             =   480
      Width           =   720
   End
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   6240
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Supported Documents|*.asp;*.htm*|"
   End
   Begin RichTextLib.RichTextBox ASPEdit 
      Height          =   4095
      Left            =   660
      TabIndex        =   0
      Top             =   480
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   7223
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      RightMargin     =   64000
      TextRTF         =   $"frmMain.frx":19D0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   496
      Y1              =   31
      Y2              =   31
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   0
      X2              =   496
      Y1              =   31
      Y2              =   31
   End
   Begin VB.Label Label1 
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   150
      Width           =   375
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   88
      X2              =   88
      Y1              =   5
      Y2              =   27
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   89
      X2              =   89
      Y1              =   5
      Y2              =   27
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Fredisoft ASPEdit version 1.00 Alpha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   4650
      Width           =   2640
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   496
      Y1              =   305
      Y2              =   305
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   497
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   497
      Y1              =   0.667
      Y2              =   0.667
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^W
      End
      Begin VB.Menu sep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu sep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "&Preview in Browser"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFilesep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecent 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu sep03 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu sep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Cle&ar"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu sep05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoTo 
         Caption         =   "&Goto Line"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find and Replace..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuInsImage 
         Caption         =   "&Image"
      End
      Begin VB.Menu mnuInsRollImage 
         Caption         =   "&Rollover Image"
      End
      Begin VB.Menu mnuInsMedia 
         Caption         =   "&Media"
         Begin VB.Menu mnuInsMediaFlash 
            Caption         =   "&Flash"
         End
         Begin VB.Menu mnuInsMediaShock 
            Caption         =   "&Shockwave"
         End
         Begin VB.Menu mnuInsMediaGen 
            Caption         =   "&Generator"
         End
         Begin VB.Menu mnuInsMediaApplet 
            Caption         =   "&Applet"
         End
         Begin VB.Menu mnuInsMediaPlugin 
            Caption         =   "&Plugin"
         End
         Begin VB.Menu mnuInsMediaAX 
            Caption         =   "A&ctiveX"
         End
      End
      Begin VB.Menu inssep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsTable 
         Caption         =   "&Table"
      End
      Begin VB.Menu mnuInsLayer 
         Caption         =   "La&yer"
      End
      Begin VB.Menu mnuInsFrameS 
         Caption         =   "Frame&s"
         Begin VB.Menu mnuInsFrameLeft 
            Caption         =   "&Left"
         End
         Begin VB.Menu mnuInsFrameRight 
            Caption         =   "&Right"
         End
         Begin VB.Menu mnuInsFrameTop 
            Caption         =   "&Top"
         End
         Begin VB.Menu mnuInsFrameBottom 
            Caption         =   "&Bottom"
         End
         Begin VB.Menu mnuInsFrameLnT 
            Caption         =   "L&eft and Top"
         End
         Begin VB.Menu mnuInsFrameLT 
            Caption         =   "Le&ft Top"
         End
         Begin VB.Menu mnuInsFrameTL 
            Caption         =   "T&op Left"
         End
         Begin VB.Menu mnuInsFrameSplit 
            Caption         =   "&Split"
         End
      End
      Begin VB.Menu inssep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsForm 
         Caption         =   "&Form"
      End
      Begin VB.Menu mnuInsFormObj 
         Caption         =   "Form O&bjects"
         Begin VB.Menu mnuInsFormTextF 
            Caption         =   "&Text Field"
         End
         Begin VB.Menu mnuInsFormButton 
            Caption         =   "&Button"
         End
         Begin VB.Menu mnuInsFormCheckB 
            Caption         =   "&Check Box"
         End
         Begin VB.Menu mnuInsFormRadioB 
            Caption         =   "&Radio Button"
         End
         Begin VB.Menu mnuInsFormLMenu 
            Caption         =   "&List/Menu"
         End
         Begin VB.Menu mnuInsFormFileF 
            Caption         =   "&File Field"
         End
         Begin VB.Menu mnuInsFormImageF 
            Caption         =   "&Image Field"
         End
         Begin VB.Menu mnuInsFormHiddenF 
            Caption         =   "&Hidden Field"
         End
         Begin VB.Menu insformsep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuInsFormJumpM 
            Caption         =   "&Jump Menu"
         End
      End
      Begin VB.Menu inssep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsEmailLink 
         Caption         =   "Email &Link"
      End
      Begin VB.Menu mnuInsDate 
         Caption         =   "&Date"
      End
      Begin VB.Menu mnuInsTabularData 
         Caption         =   "T&abular Data"
      End
      Begin VB.Menu mnuInsHorizontalRule 
         Caption         =   "Hori&zontal Rule"
      End
      Begin VB.Menu inssep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsInvTags 
         Caption         =   "In&visible Tags"
         Begin VB.Menu mnuInsInvTNamedA 
            Caption         =   "&Named Anchor"
         End
         Begin VB.Menu mnuInsInvTScript 
            Caption         =   "&Script"
         End
         Begin VB.Menu mnuInsInvTComment 
            Caption         =   "&Comment"
         End
      End
      Begin VB.Menu mnuInsHeadTags 
         Caption         =   "&Head Tags"
         Begin VB.Menu mnuInsHeadTMeta 
            Caption         =   "&Meta"
         End
         Begin VB.Menu mnuInsHeadTKeyw 
            Caption         =   "&Keywords"
         End
         Begin VB.Menu mnuInsHeadTDescrip 
            Caption         =   "&Description"
         End
         Begin VB.Menu mnuInsHeadTRefresh 
            Caption         =   "&Refresh"
         End
         Begin VB.Menu mnuInsHeadTBase 
            Caption         =   "&Base"
         End
         Begin VB.Menu mnuInsHeadTLink 
            Caption         =   "&Link"
         End
      End
      Begin VB.Menu mnuInsSpecialChars 
         Caption         =   "Special &Characters"
         Begin VB.Menu mnuInsSpecialCLineB 
            Caption         =   "Lin&e Break"
         End
         Begin VB.Menu mnuInsSpecialCNonBS 
            Caption         =   "Non-Brea&king Space"
         End
         Begin VB.Menu scsep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuInsSpecialCCopyr 
            Caption         =   "&Copyright"
         End
         Begin VB.Menu mnuInsSpecialCReg 
            Caption         =   "&Registered"
         End
         Begin VB.Menu mnuInsSpecialCTradeM 
            Caption         =   "&Trademark"
         End
         Begin VB.Menu mnuInsSpecialCPound 
            Caption         =   "&Pound"
         End
         Begin VB.Menu mnuInsSpecialCLeftQ 
            Caption         =   "&Left Quote"
         End
         Begin VB.Menu mnuInsSpecialCRightQ 
            Caption         =   "R&ight Quote"
         End
         Begin VB.Menu mnuInsSpecialCOther 
            Caption         =   "&Other..."
         End
      End
      Begin VB.Menu inssep05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsASP 
         Caption         =   "ASP"
         Index           =   0
         Begin VB.Menu mnuInsASPDatabase 
            Caption         =   "&Database"
            Begin VB.Menu mnuInsASPDataBaseDB 
               Caption         =   "&Database"
            End
            Begin VB.Menu mnuInsASPDataBaseADO 
               Caption         =   "&ADO Database"
            End
            Begin VB.Menu mnuInsASPDatabaseSQL 
               Caption         =   "&SQL Database"
            End
         End
         Begin VB.Menu mnuInsASPTagS 
            Caption         =   "&ASP Tags"
            Begin VB.Menu mnuInsASPTagSelectC 
               Caption         =   "Select &Case"
            End
            Begin VB.Menu mnuInsASPTagsIf 
               Caption         =   "&If Then Else"
            End
            Begin VB.Menu mnuInsASPTagDoWhile 
               Caption         =   "&Do While"
            End
            Begin VB.Menu mnuInsASPTagForN 
               Caption         =   "F&or Next"
            End
            Begin VB.Menu mnuInsASPTagOA 
               Caption         =   "&<%"
            End
            Begin VB.Menu mnuInsASPTagCA 
               Caption         =   "&%>"
            End
            Begin VB.Menu mnuInsASPTagSub 
               Caption         =   "&Sub"
            End
            Begin VB.Menu mnuInsASPTagFunc 
               Caption         =   "&Function"
            End
         End
         Begin VB.Menu mnuInsASPCreateObj 
            Caption         =   "Create &Objects"
            Begin VB.Menu mnuInsASPCreateObj1 
               Caption         =   "ADODB.Recordset"
            End
            Begin VB.Menu mnuInsASPCreateObj2 
               Caption         =   "ADODB.Connection"
            End
            Begin VB.Menu mnuInsASPCreateObj3 
               Caption         =   "ASPEmail Object"
            End
            Begin VB.Menu mnuInsASPCreateObj4 
               Caption         =   "ASPMail Object"
            End
            Begin VB.Menu mnuInsASPCreateObj5 
               Caption         =   "ASPUpload Object"
            End
            Begin VB.Menu mnuInsASPCreateObj6 
               Caption         =   "CDONTS.NewMail"
            End
            Begin VB.Menu mnuInsASPCreateObj7 
               Caption         =   "JMail Object"
            End
            Begin VB.Menu mnuInsASPCreateObj8 
               Caption         =   "MSWC.AdRotator"
            End
            Begin VB.Menu mnuInsASPCreateObj9 
               Caption         =   "MSWC.BrowserType"
            End
            Begin VB.Menu mnuInsASPCreateObj10 
               Caption         =   "Scripting.FileSystemObject"
            End
         End
         Begin VB.Menu mnuInsASPRequestVar 
            Caption         =   "&Request Variables"
         End
         Begin VB.Menu mnuInsASPResponseRed 
            Caption         =   "Response.Re&direct"
         End
         Begin VB.Menu mnuInsASPResponseWri 
            Caption         =   "Response.&Write"
         End
         Begin VB.Menu mnuInsASPCookies 
            Caption         =   "&Cookies"
         End
         Begin VB.Menu mnuInsASPDate 
            Caption         =   "ASP Da&te"
         End
         Begin VB.Menu mnuInsASPInclude 
            Caption         =   "&Include"
         End
      End
   End
   Begin VB.Menu mnuModify 
      Caption         =   "&Modify"
      Begin VB.Menu mnuPageProperties 
         Caption         =   "&Page Properties..."
      End
      Begin VB.Menu mnuModifysep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuText 
      Caption         =   "&Text"
      Begin VB.Menu mnuTextParFormat 
         Caption         =   "&Paragraph Format"
         Begin VB.Menu mnuParFormatPar 
            Caption         =   "&Paragraph"
         End
         Begin VB.Menu mnuParFormatHead1 
            Caption         =   "Heading &1"
         End
         Begin VB.Menu mnuParFormatHead2 
            Caption         =   "Heading &2"
         End
         Begin VB.Menu mnuParFormatHead3 
            Caption         =   "Heading &3"
         End
         Begin VB.Menu mnuParFormatHead4 
            Caption         =   "Heading &4"
         End
         Begin VB.Menu mnuParFormatHead5 
            Caption         =   "Heading &5"
         End
         Begin VB.Menu mnuParFormatHead6 
            Caption         =   "Heading &6"
         End
         Begin VB.Menu mnuParFormatPreFT 
            Caption         =   "P&reformatted Text"
         End
      End
      Begin VB.Menu mnuTextAlign 
         Caption         =   "&Align"
         Begin VB.Menu mnuTextAlignLeft 
            Caption         =   "&Left"
         End
         Begin VB.Menu mnuTextAlignCenter 
            Caption         =   "&Center"
         End
         Begin VB.Menu mnuTextAlignRight 
            Caption         =   "&Right"
         End
      End
      Begin VB.Menu textsep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTextFont 
         Caption         =   "&Font"
         Begin VB.Menu mnuTextFontList 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu textfontsep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTextFontEditFL 
            Caption         =   "&Edit Font List..."
         End
      End
      Begin VB.Menu mnuTextStyle 
         Caption         =   "&Style"
         Begin VB.Menu mnuTextStyleBold 
            Caption         =   "&Bold"
         End
         Begin VB.Menu mnuTextStyleItalic 
            Caption         =   "&Italic"
         End
         Begin VB.Menu mnuTextStyleUnderl 
            Caption         =   "&Underline"
         End
         Begin VB.Menu stylesep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTextStyleStrikeT 
            Caption         =   "&Strikethrough"
         End
         Begin VB.Menu mnuTextStypeTeleT 
            Caption         =   "&Teletype"
         End
         Begin VB.Menu mnuTextStyleEmphasis 
            Caption         =   "&Emphasis"
         End
         Begin VB.Menu mnuTextStyleStrong 
            Caption         =   "St&ong"
         End
         Begin VB.Menu stylesep02 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTextStyleCode 
            Caption         =   "&Code"
         End
         Begin VB.Menu mnuTextStyleVariable 
            Caption         =   "&Variable"
         End
         Begin VB.Menu mnuTextStyleSample 
            Caption         =   "S&ample"
         End
         Begin VB.Menu mnuTextStyleKeyboard 
            Caption         =   "&Keyboard"
         End
         Begin VB.Menu stylesep03 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTextStyleCitation 
            Caption         =   "Citati&on"
         End
         Begin VB.Menu mnuTextStyleDefinition 
            Caption         =   "&Definition"
         End
      End
      Begin VB.Menu mnuTextSize 
         Caption         =   "Si&ze"
         Begin VB.Menu mnuTextSize1 
            Caption         =   "&1"
         End
         Begin VB.Menu mnuTextSize2 
            Caption         =   "&2"
         End
         Begin VB.Menu mnuTextSize3 
            Caption         =   "&3"
         End
         Begin VB.Menu mnuTextSize4 
            Caption         =   "&4"
         End
         Begin VB.Menu mnuTextSize5 
            Caption         =   "&5"
         End
         Begin VB.Menu mnuTextSize6 
            Caption         =   "&6"
         End
         Begin VB.Menu mnuTextSize7 
            Caption         =   "&7"
         End
      End
      Begin VB.Menu textsep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTextColor 
         Caption         =   "&Color..."
      End
   End
   Begin VB.Menu mnuPlugins 
      Caption         =   "&Plugins"
      Begin VB.Menu mnuPluginManager 
         Caption         =   "&Plugin Manager..."
      End
      Begin VB.Menu pluginsep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPluginList 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuWindowObjects 
         Caption         =   "&Objects"
      End
      Begin VB.Menu windowsep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowASPTags 
         Caption         =   "&ASP Tags"
      End
      Begin VB.Menu mnuWindowASPCodes 
         Caption         =   "ASP &Codes"
      End
      Begin VB.Menu mnuWindowHTMLTags 
         Caption         =   "&HTML Tags"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuUsing 
         Caption         =   "&Contents..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu sepHelp01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRegister 
         Caption         =   "&Register ASPEdit (FREE)"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About ASPEdit..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' // ASPEdit
' // Developed by Frederico Machado (indiofu@bol.com.br)
' /////////////////////////////////////////////////////////

Option Explicit

Private nIndent As Integer
Private bInVar As Boolean
Private isTag As Boolean

Private Type sAEHelpLang
  sLangName As String
  sAESubs() As String
End Type

Private sAEHelp() As sAEHelpLang
Private iAEHelpCount As Integer

Private Sub ASPEdit_Click()
  Dim lTagSP As Long, lTagLenght, sTmpTag As String
  On Error Resume Next
  lTagSP = InStrRev(ASPEdit.Text, "<", ASPEdit.SelStart)
  If InStrRev(ASPEdit.Text, ">", ASPEdit.SelStart) > lTagSP Then
    ' no tag
    Exit Sub
  End If
  lTagLenght = InStr(ASPEdit.SelStart, ASPEdit.Text, ">") - lTagSP + 1
  sTmpTag = Mid$(ASPEdit.Text, lTagSP, lTagLenght)
  If Left(sTmpTag, 3) = "<!-" Then
    ' comment
    Exit Sub
  End If
  TagProp = GetTagProperties(sTmpTag)
  Dim i As Integer
  For i = 1 To UBound(TagProp.Properties)
    If TagProp.Properties(i, 0) = "" Then Exit For
    ' tags
  Next
End Sub

Private Sub ASPEdit_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Tab key
  If KeyCode = 9 Then
    Dim iStart As Integer, iLength As Integer
    iStart = ASPEdit.SelStart
    iLength = ASPEdit.SelLength
    ASPEdit.SelText = "  " & ASPEdit.SelText
    ASPEdit.SelStart = iStart + 2
    ASPEdit.SelLength = iLength
    KeyCode = 0
  End If
  If KeyCode = 37 Or KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Then
    If picAEHlp.Visible Then picAEHlp.Visible = False
  End If
End Sub

Private Sub ASPEdit_KeyPress(KeyAscii As Integer)
  Dim PT As POINTAPI
  
  If KeyAscii = Asc(")") Or KeyAscii = Asc(" ") Or KeyAscii = 8 Then
    If picAEHlp.Visible Then picAEHlp.Visible = False
    Exit Sub
  End If
  If KeyAscii = Asc("(") Then
    Dim i As Integer, j As Integer, sCommand As String
    For i = 0 To UBound(sAEHelp)
      For j = 0 To UBound(sAEHelp(i).sAESubs)
        sCommand = Left$(sAEHelp(i).sAESubs(j), InStr(sAEHelp(i).sAESubs(j), "=") - 1)
        If LCase(Mid$(ASPEdit.Text, ASPEdit.SelStart - Len(sCommand) + 1, Len(sCommand))) = LCase(sCommand) Then
          lblTip = Mid$(sAEHelp(i).sAESubs(j), InStr(sAEHelp(i).sAESubs(j), "=") + 1)
          picAEHlp.Width = lblTip.Width + 15
          lTop.X2 = picAEHlp.ScaleWidth - 1: lBottom.X2 = lTop.X2
          lRight.X1 = lTop.X2: lRight.X2 = lRight.X1
          GetCaretPos PT
          If PT.X > (ScaleWidth - picAEHlp.Width) Then
            picAEHlp.Left = ScaleWidth - picAEHlp.Width
          Else
            picAEHlp.Left = PT.X
          End If
          picAEHlp.Top = PT.Y + 45
          picAEHlp.Visible = True
        End If
      Next
    Next
    Exit Sub
  End If
End Sub

Private Sub ASPEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If picAEHlp.Visible Then picAEHlp.Visible = False
End Sub

Private Sub Form_Load()
  Path = App.Path
  If Right$(Path, 1) <> "\" Then Path = Path & "\"
  
  If bURegistered Then mnuRegister.Visible = False
  LoadRecentFiles
  LoadFonts
  LoadAEHelpFile Path & "Data\asp.aeh", "ASP"
  NewDocument
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  Line1.X2 = ScaleWidth
  Line2.X2 = ScaleWidth
  Line5.X2 = ScaleWidth
  Line6.X2 = ScaleWidth
  Line7.X2 = ScaleWidth
  picLines.Height = ScaleHeight - 55
  ASPEdit.Height = ScaleHeight - 55
  ASPEdit.Width = ScaleWidth - 44
  Line7.Y1 = ASPEdit.Top + ASPEdit.Height
  Line7.Y2 = Line7.Y1
  lblStatus.Top = Line7.Y1 + 5
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If frmTags.Visible Then Unload frmTags
  CleanTempFolder
  DoEvents
  End
End Sub

Private Sub lblTip_Click()
  picAEHlp.Visible = False
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show 1
End Sub

Private Sub mnuClear_Click()
  ASPEdit.SelText = ""
End Sub

Private Sub mnuClose_Click()
  Unload Me
End Sub

Private Sub mnuCopy_Click()
  If ASPEdit.SelText = "" Then Exit Sub
  Clipboard.Clear
  Clipboard.SetText ASPEdit.SelText
End Sub

Private Sub mnuCut_Click()
  If ASPEdit.SelText = "" Then Exit Sub
  Clipboard.Clear
  Clipboard.SetText ASPEdit.SelText
  ASPEdit.SelText = ""
End Sub

Private Sub mnuFind_Click()
  frmFind.Text1 = ASPEdit.SelText
  frmFind.Show , Me
End Sub

Private Sub mnuGoTo_Click()
  frmGoTo.Show , Me
End Sub

Private Sub mnuNew_Click()
  NewDocument
End Sub

Private Sub mnuOpen_Click()
    On Error GoTo noopen
    cmDialog.FileName = ""
    cmDialog.DialogTitle = "Open Page"
    cmDialog.Filter = "All Supported Documents|*.asp;*.htm*;*.js;*.jsp;*.css|"
    cmDialog.FLAGS = &H4 Or &H1000
    cmDialog.ShowOpen
    strFileName = cmDialog.FileName
    If strFileName = "" Then Exit Sub
    strDocFolder = Left$(strFileName, InStrRev(strFileName, "\") - 1)
    Me.MousePointer = 11
    
    OpenFile strFileName
    If InStr(frmMain.ASPEdit, "<title>") Then txtTitle = GetTitle()
    
    Me.MousePointer = 0
    
    Add2RecentList strFileName
    
    Exit Sub
noopen:
    If Err = 32755 Then Exit Sub
    MsgBox "An error has occured while trying to access the file [" & Err & "]:" & vbCrLf & vbCrLf & Error, vbExclamation
    strDocFolder = "": txtTitle = ""
End Sub

Private Sub mnuOptions_Click()
  frmOptions.Show 1
End Sub

Private Sub mnuPageProperties_Click()
  frmPProperties.Show 1
End Sub

Private Sub mnuPaste_Click()
  ASPEdit.SelText = Clipboard.GetText
End Sub

Private Sub mnuPreview_Click()
  Dim File As String, iNumber As Integer
  If Not DirExists(Path & "temp") Then MkDir Path & "temp"
  Randomize
  iNumber = Rnd * 10240
  File = Path & "temp\tmppage" & Format$(iNumber, "00000") & ".htm"
  If FileExists(File) Then Kill File
  DoEvents
  SaveFile File
  ShellExecute hWnd, "Open", File, 0&, File, 1
End Sub

Private Sub mnuRecent_Click(Index As Integer)
    On Error GoTo noopen
    strFileName = mnuRecent(Index).Tag
    strDocFolder = Left$(strFileName, InStrRev(strFileName, "\") - 1)
    Me.MousePointer = 11
    OpenFile strFileName
    If InStr(frmMain.ASPEdit, "<title>") Then txtTitle = GetTitle()
    Me.MousePointer = 0
    Add2RecentList strFileName
    Exit Sub
noopen:
    If Err = 32755 Then Exit Sub
    MsgBox "An error has occured while trying to access the file [" & Err & "]:" & vbCrLf & vbCrLf & Error, vbExclamation
    strDocFolder = "": txtTitle = ""
End Sub

Private Sub mnuRegister_Click()
  frmRegister.Show 1
End Sub

Private Sub mnuSave_Click()
  Me.MousePointer = 11
  If strFileName = "" Then
    Call mnuSaveAs_Click
  Else
    SaveFile strFileName
  End If
  Me.MousePointer = 0
End Sub

Private Sub mnuSaveAs_Click()
  On Error GoTo noSaveAs
  cmDialog.DialogTitle = "Save Page As"
  cmDialog.FLAGS = &H4 Or &H1000
  cmDialog.FileName = strFileName
  cmDialog.ShowSave
  strFileName = cmDialog.FileName
  If strFileName = "" Then Exit Sub
  strDocFolder = Left$(strFileName, InStrRev(strFileName, "\") - 1)
  SaveFile strFileName
    
  Exit Sub
  
noSaveAs:
    If Err = 32755 Then Exit Sub
    MsgBox "An error has occured while trying to access the file [" & Err & "]:" & vbCrLf & vbCrLf & Error, vbExclamation
End Sub

Private Sub mnuSelectAll_Click()
  ASPEdit.SelStart = 0
  ASPEdit.SelLength = Len(ASPEdit.Text)
End Sub

Private Sub mnuWindowASPCodes_Click()
  frmTags.Caption = "ASP Codes"
  frmTags.tabTags.Tab = 1
  frmTags.Show , Me
End Sub

Private Sub mnuWindowASPTags_Click()
  frmTags.Caption = "ASP Tags"
  frmTags.tabTags.Tab = 0
  frmTags.Show , Me
End Sub

Private Sub mnuWindowHTMLTags_Click()
  frmTags.Caption = "HTML Tags"
  frmTags.tabTags.Tab = 2
  frmTags.Show , Me
End Sub

Private Sub picAEHlp_Click()
  picAEHlp.Visible = False
End Sub

Private Sub tmrLines_Timer()
  DrawLines picLines, ASPEdit
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    ChangeTitle txtTitle
  End If
End Sub

Private Sub txtTitle_LostFocus()
  ChangeTitle txtTitle
End Sub

Function GetLine(strText As String, curPos As Integer) As String
    Dim strRet As String
    If curPos = 0 Then Exit Function
    
    Do
        If Mid$(strText, curPos, 1) = Chr(10) Then Exit Do
        strRet = Mid$(strText, curPos, 1) & strRet
        curPos = curPos - 1
    Loop Until curPos = 0
    GetLine = strRet
End Function

Sub CleanTempFolder()
  
  Dim strTemp As String
  
  strTemp = Dir$(Path & "temp\*.*")
  If strTemp <> "" Then
    Kill Path & "temp\" & strTemp
    strTemp = Dir$
    While Len(strTemp) > 0
      Kill Path & "temp\" & strTemp
      DoEvents
      strTemp = Dir$
    Wend
  End If
  
End Sub

Sub LoadFonts()
  
  If Not FileExists(Path & "Data\font.lst") Then Exit Sub
  
  Dim strFileData As String
  Open Path & "Data\font.lst" For Binary As #1
    strFileData = String(LOF(1), 0)
    Get #1, 1, strFileData
  Close #1
  
  Dim sTmpArray() As String, i As Integer
  
  sTmpArray = Split(strFileData, vbCrLf)
  For i = 0 To UBound(sTmpArray)
    If i > 0 Then Load mnuTextFontList(i)
    mnuTextFontList(i).Caption = sTmpArray(i)
  Next
  
End Sub

Sub LoadRecentFiles()
  
  Dim sFileName As String, sFileData As String
  
  sFileName = Path & "Data\recent.ini"
  
  If Not FileExists(sFileName) Then Exit Sub
  
  Open sFileName For Binary As #1
    sFileData = String(LOF(1), 0)
    Get #1, 1, sFileData
  Close #1
  
  Dim sTmpArray() As String, i As Integer
  
  sTmpArray = Split(sFileData, vbCrLf)
  For i = 0 To UBound(sTmpArray)
    If sTmpArray(i) = "" Then GoTo Jump
    If i = 5 Then Exit Sub
    If i > 0 Then Load mnuRecent(i)
    mnuRecent(i).Caption = "&" & (i + 1) & " " & Mid$(sTmpArray(i), InStrRev(sTmpArray(i), "\") + 1)
    mnuRecent(i).Tag = sTmpArray(i)
    mnuRecent(i).Visible = True
    sep03.Visible = True
Jump:
  Next
  
End Sub

Sub ClearRecentList()
  
  Dim i As Integer
  
  For i = 0 To mnuRecent.UBound
    mnuRecent(i).Visible = False
    If i > 0 Then Unload mnuRecent(i)
  Next
  sep03.Visible = False
  
End Sub

Sub Add2RecentList(sFileName As String)
  
  If Not mnuRecent(0).Visible Then
    Open Path & "Data\recent.ini" For Output As #1
      Print #1, sFileName
    Close #1
    GoTo Done
  End If
  
  Dim sData As String, i As Integer
  sData = sFileName
  
  For i = 0 To mnuRecent.UBound
    If i = 4 Then Exit For
    If sFileName <> mnuRecent(i).Tag Then
      sData = sData & vbCrLf & mnuRecent(i).Tag
    End If
  Next
  
  Open Path & "Data\recent.ini" For Output As #1
    Print #1, sData
  Close #1
  
Done:
  
  ClearRecentList
  LoadRecentFiles
  
End Sub

Sub LoadAEHelpFile(sFileName As String, sLang As String)
  
  iAEHelpCount = iAEHelpCount + 1
  ReDim Preserve sAEHelp(iAEHelpCount - 1)
  sAEHelp(iAEHelpCount - 1).sLangName = sLang
  
  Dim sFileData As String
  
  Open sFileName For Binary As #1
    sFileData = String(LOF(1), 0)
    Get #1, 1, sFileData
  Close #1
  
  sAEHelp(iAEHelpCount - 1).sAESubs = Split(sFileData, vbCrLf)
  
End Sub
