VERSION 5.00
Begin VB.Form frmRegister 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Registration (FREE)"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   285
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox DataArrival 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   360
      TabIndex        =   18
      Top             =   3720
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register!"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtHomepage 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Text            =   "http://"
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox txtMail 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   1680
      Width           =   3255
   End
   Begin VB.ComboBox cboResidence 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmRegister.frx":038A
      Left            =   1200
      List            =   "frmRegister.frx":0451
      Sorted          =   -1  'True
      TabIndex        =   4
      Text            =   "United States of America"
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox txtSex 
      Height          =   285
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtAge 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtOccupation 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Homepage:"
      Height          =   195
      Left            =   285
      TabIndex        =   15
      Top             =   2040
      Width           =   825
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail Addr.:"
      Height          =   195
      Left            =   210
      TabIndex        =   14
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Residence:"
      Height          =   195
      Left            =   300
      TabIndex        =   13
      Top             =   1320
      Width           =   810
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex:"
      Height          =   195
      Left            =   2760
      TabIndex        =   12
      Top             =   960
      Width           =   315
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age:"
      Height          =   195
      Left            =   780
      TabIndex        =   11
      Top             =   960
      Width           =   330
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fredisoft Corp."
      Height          =   195
      Left            =   3480
      TabIndex        =   10
      Top             =   3360
      Width           =   1020
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmRegister.frx":06F6
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation:"
      Height          =   195
      Left            =   210
      TabIndex        =   8
      Top             =   600
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   960
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bTrans As Boolean
Private m_iStage As Integer
Private Sock As Integer
Private RC As Integer
Private Bytes As Integer
Private ResponseCode As Integer

Private Const mailserver As String = "smtp.nho.terra.com.br"
Private Const Tobox As String = "indiofu@terra.com.br"
Private Const Subject As String = "ASPEdit User Registration!"

Private Frombox As String

'This is for the WaitforResponse Routine
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub cmdCancel_Click()
  On Error Resume Next
    
  closesocket Sock
  RC = WSACleanup()
  Unload Me
End Sub

Private Sub cmdRegister_Click()

  If txtName = "" Then
    MsgBox "You need to type your name to register ASPEdit.", vbCritical, "Error"
    txtName.SetFocus
    Exit Sub
  End If
  If txtOccupation = "" Then
    MsgBox "You need to type your occupation.", vbCritical, "Error"
    txtOccupation.SetFocus
    Exit Sub
  End If
  If txtAge = "" Then
    MsgBox "You need to type your age.", vbCritical, "Error"
    txtAge.SetFocus
    Exit Sub
  End If
  If cboResidence = "" Then
    MsgBox "You need to choose where your residence is.", vbCritical, "Error"
    cboResidence.SetFocus
    Exit Sub
  End If
  If txtMail = "" Or InStr(txtMail, "@") = 0 Then
    MsgBox "You forgot to type your e-mail address or it is an invalid format.", vbCritical, "Error"
    txtMail.SetFocus
    Exit Sub
  End If

  Frombox = txtMail

  Dim StartupData As WSADataType
  Dim SocketBuffer As sockaddr
  Dim IpAddr As Long
    
  'Ini the Winsocket
  RC = WSAStartup(&H101, StartupData)
  RC = WSAStartup(&H101, StartupData)
    
  'Open a free Socket (with this source code you can also
  'open several connections! Very useful for E-Mail Applications...)
  Sock = Socket(AF_INET, SOCK_STREAM, 0)
  If Sock = SOCKET_ERROR Then
    MsgBox "Cannot Create Socket.", vbCritical
    Exit Sub
  End If

  'Checks if the Hostname exists
  If RC = SOCKET_ERROR Then Exit Sub
  IpAddr = GetHostByNameAlias(mailserver)
  If IpAddr = -1 Then
    MsgBox "Mail host not found or busy, try again later, please.", vbCritical
    Exit Sub
  End If

  'This part is responsible for the connection
  SocketBuffer.sin_family = AF_INET
  SocketBuffer.sin_port = htons(25)
  SocketBuffer.sin_addr = IpAddr
  SocketBuffer.sin_zero = String$(8, 0)
    
  RC = Connect(Sock, SocketBuffer, Len(SocketBuffer))

  'If an error occured close the connection and
  'send an error message to the text window
  If RC = SOCKET_ERROR Then
        MsgBox "Cannot Connect to the mail server, try again later, please." + _
                            Chr$(13) + Chr$(10) + _
                            GetWSAErrorString(WSAGetLastError()), vbCritical
        closesocket Sock
        RC = WSACleanup()
        Exit Sub
  End If

  'Select Receive Window
  RC = WSAAsyncSelect(Sock, DataArrival.hWnd, ByVal &H202, ByVal FD_READ Or FD_CLOSE)
  If RC = SOCKET_ERROR Then
    MsgBox "Cannot Process Asynchronously.", vbCritical
    closesocket Sock
    RC = WSACleanup()
    Exit Sub
  End If

  bTrans = True
  m_iStage = 0
  DataArrival = ""

  ResponseCode = 220
  Call WaitForResponse
  
End Sub

Private Sub Transmit(iStage As Integer)
Dim Helo As String, temp As String
Dim pos As Integer

Select Case m_iStage

Case 1:
    Helo = Frombox
    pos = Len(Helo) - InStr(Helo, "@")
    Helo = Right$(Helo, pos)
    
    ResponseCode = 250
    WinsockSendData ("HELO " & Helo & vbCrLf)
    Call WaitForResponse

Case 2:
    ResponseCode = 250
    WinsockSendData ("MAIL FROM: <" & Trim(Frombox) & ">" & vbCrLf)
    Call WaitForResponse

Case 3:
    ResponseCode = 250
    WinsockSendData ("RCPT TO: <" & Trim(Tobox) & ">" & vbCrLf)
    Call WaitForResponse

Case 4:
    ResponseCode = 354
    WinsockSendData ("DATA" & vbCrLf)
    Call WaitForResponse

Case 5:

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If you want additional Headers like Date,Message-Id,...etc. !
'simply add them below                                      !
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    temp = temp & "From: " & Frombox & vbNewLine
    temp = temp & "To: " & Tobox & vbNewLine
    temp = temp & "Subject: " & Subject & vbNewLine

    'Header + Message
    temp = temp & vbCrLf & _
    "Name: " & txtName & vbCrLf & _
    "Occupation: " & txtOccupation & vbCrLf & _
    "Age: " & txtAge & vbCrLf & "Sex: " & txtSex & vbCrLf & _
    "Residence: " & cboResidence & vbCrLf & _
    "E-Mail Address: " & txtMail & vbCrLf & _
    "Homepage: " & txtHomepage

    'Send the Message & close connection
    WinsockSendData (temp)
    WinsockSendData (vbCrLf & "." & vbCrLf)
    ResponseCode = 250
    Call WaitForResponse

Case 6:
    MsgBox "You are successfuly registered!" & vbCrLf & "Thank you for registering ASPEdit.", vbInformation
    bURegistered = True
    WinsockSendData ("QUIT" & vbCrLf)
    ResponseCode = 221
    Call WaitForResponse
    m_iStage = 0
    bTrans = False
End Select
End Sub

'***************************************************************
'Routine for arraving Data
'***************************************************************

Private Sub DataArrival_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MsgBuffer As String * 2048


    
On Error Resume Next

 

    If Sock > 0 Then
        'Receive up to 2048 chars
        Bytes = recv(Sock, ByVal MsgBuffer, 2048, 0)
        
        If Bytes > 0 Then
            
                
        If bTrans Then
            If ResponseCode = Left(MsgBuffer, 3) Then
            MsgBuffer = vbNullString
            m_iStage = m_iStage + 1
            Transmit m_iStage
            Else
                closesocket (Sock)
                RC = WSACleanup()
                Sock = 0
                MsgBox "The Server responds with an unexpected Response Code!", vbOKOnly, "Error!"
                Exit Sub
            End If
        End If

        ElseIf WSAGetLastError() <> WSAEWOULDBLOCK Then
            closesocket (Sock)
            RC = WSACleanup()
            Sock = 0
        End If
    End If

Refresh


End Sub

'**************************************************************
' Waits until time out, while waiting for response
'**************************************************************

Private Sub WaitForResponse()
Dim Start As Long
Dim Tmr As Long

'Works with an Api Declaration because it's more precious

Start = timeGetTime
While Bytes > 0
    Tmr = timeGetTime - Start
    DoEvents ' Let System keep checking for incoming response
        
    'Wait 50 seconds for response
    If Tmr > 50000 Then
        MsgBox "SMTP service error, timed out while waiting for response", 64, "Error!"
        End
    End If
Wend
End Sub

Private Sub WinsockSendData(DatatoSend As String)
Dim RC As Integer
Dim MsgBuffer As String * 2048

MsgBuffer = DatatoSend

RC = Send(Sock, ByVal MsgBuffer, Len(DatatoSend), 0)
    
'If an error occurs send an error message and
'reset the winsock
If RC = SOCKET_ERROR Then
    MsgBox "Cannot Send Request." + _
                            Chr$(13) + Chr$(10) + _
                            Str$(WSAGetLastError()) + _
                            GetWSAErrorString(WSAGetLastError())
    closesocket Sock
    RC = WSACleanup()
    Exit Sub
End If


End Sub

Private Sub txtHomepage_GotFocus()
  txtHomepage.SelStart = Len(txtHomepage)
End Sub
