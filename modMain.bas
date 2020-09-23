Attribute VB_Name = "modMain"
' ASPEdit
' Developed by Frederico Machado

Option Explicit

Global Path As String
Global strFileName As String
Global strDocFolder As String

Global bURegistered As Boolean

Public Type PageProperties
  Title As String
  Background As String
  BGProperties As String
  BGColor As String
  Text As String
  Link As String
  Visited As String
  Active As String
  LeftMargin As String
  TopMargin As String
  MarginWidth As String
  MarginHeight As String
End Type

Global PProp As PageProperties

Public Type TagProperties
  Name As String
  Properties(1 To 20, 0 To 1) As String
End Type

Global TagProp As TagProperties

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type POINTAPI
  X As Long
  Y As Long
End Type

Public POINTAPI As POINTAPI

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Enum ModifyTypes
    AddText = 0
    DeleteText = 1
    ReplaceText = 2
    CutText = 3
    PasteText = 4
End Enum

Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_GETFIRSTVISIBLELINE = &HCE

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub DrawLines(picTo As PictureBox, rtf As RichTextBox)
  Dim iLine As Long, cLine As Long, vLine As Long
  Dim sWidth As Single, tmp As Single
  'count the lines
  iLine = SendMessage(rtf.hWnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
  'current line
  cLine = 1 + rtf.GetLineFromChar(rtf.SelStart)
  'first visible line
  vLine = SendMessage(rtf.hWnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
  picTo.Cls

  picTo.Font = rtf.Font
  picTo.ForeColor = &H8000000C
  Dim i As Integer
  For i = vLine + 1 To iLine
    sWidth = picTo.TextWidth(CStr(i))
    picTo.CurrentX = (picTo.Width - 10) - sWidth
    If i <> cLine Then
      picTo.ForeColor = &H8000000C
      picTo.Print i
    Else
      If i = cLine Then
        picTo.ForeColor = &H80000018
        picTo.Print i
      End If
    End If
  Next i
End Sub

Sub NewDocument()
  Dim strNew As String
  
  strNew = "<html>" & vbCrLf & "<head>" & vbCrLf
  strNew = strNew & "<title>Untitled Page</title>" & vbCrLf
  strNew = strNew & "<meta name=" & Chr$(34) & "PROGRAM GENERATOR" & Chr$(34)
  strNew = strNew & " content=" & Chr$(34) & "Fredisoft ASPEdit" & Chr$(34) & ">" & vbCrLf
  strNew = strNew & "</head>" & vbCrLf & "<body bgcolor=" & Chr$(34)
  strNew = strNew & "#FFFFFF" & Chr$(34) & " text=" & Chr$(34) & "#000000" & Chr$(34) & ">" & vbCrLf
  strNew = strNew & vbCrLf & "</body>" & vbCrLf & "</html>"
  
  PProp = ChangePProp("Untitled Page", "", "", "#FFFFFF", "#000000", "", "", "", "", "", "", "")
  
  frmMain.ASPEdit.Text = strNew
  
  frmMain.txtTitle = "Untitled Page"
End Sub

Sub OpenFile(strFileName As String)
  On Error GoTo noLoad
  Dim strFileData As String
  
  Open strFileName For Binary As #1
    strFileData = String(LOF(1), 0)
    Get #1, 1, strFileData
  Close #1
    
  frmMain.ASPEdit.Text = strFileData
    
  strFileData = ""
    
  Exit Sub
  
noLoad:
    MsgBox "An error has occured while trying to load the file: [" & Err & "]" & vbCrLf & vbCrLf, vbCritical
End Sub

Sub SaveFile(strFileName As String)
  On Error GoTo noSaveAs
  frmMain.MousePointer = 11
  Open strFileName For Output As #1
    Print #1, frmMain.ASPEdit.Text
  Close #1
    
  frmMain.MousePointer = 0
  frmMain.Add2RecentList strFileName
  Exit Sub
    
noSaveAs:
    frmMain.MousePointer = 0
    MsgBox "An error has occured while trying to save the file, [" & Err & "]:" & vbCrLf & vbCrLf & Error, vbExclamation
End Sub

Function GetPageProperties() As PageProperties

  Dim sBody As String, iB As Integer, sChar As String

  GetPageProperties.Title = GetTitle
  iB = InStr(frmMain.ASPEdit.Text, "<body")
  If iB > 0 Then
    Do While sChar <> ">"
      sChar = Mid$(frmMain.ASPEdit.Text, iB, 1)
      sBody = sBody & sChar
      iB = iB + 1
    Loop
  End If
  
  Dim sTemp As String, iTemp As Integer, sCharTmp
  
  iB = InStr(sBody, "background")
  If iB > 0 Then
    If Mid$(sBody, iB + 11, 1) = Chr$(34) Then
      iTemp = iB + 11
      Do While sCharTmp <> Chr$(34)
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    Else
      iTemp = iB + 10
      Do While sCharTmp <> " "
        If sCharTmp = ">" Then Exit Do
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    End If
    GetPageProperties.Background = sTemp
  Else
    GetPageProperties.Background = ""
  End If
  
  sTemp = "": sCharTmp = ""
  iB = InStr(sBody, "bgproperties")
  If iB > 0 Then
    If Mid$(sBody, iB + 13, 1) = Chr$(34) Then
      iTemp = iB + 13
      Do While sCharTmp <> Chr$(34)
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    Else
      iTemp = iB + 12
      Do While sCharTmp <> " "
        If sCharTmp = ">" Then Exit Do
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    End If
    GetPageProperties.BGProperties = sTemp
  Else
    GetPageProperties.BGProperties = ""
  End If
  
  sTemp = "": sCharTmp = ""
  iB = InStr(sBody, "bgcolor")
  If iB > 0 Then
    If Mid$(sBody, iB + 8, 1) = Chr$(34) Then
      iTemp = iB + 8
      Do While sCharTmp <> Chr$(34)
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    Else
      iTemp = iB + 7
      Do While sCharTmp <> " "
        If sCharTmp = ">" Then Exit Do
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    End If
    GetPageProperties.BGColor = sTemp
  Else
    GetPageProperties.BGColor = ""
  End If

  sTemp = "": sCharTmp = ""
  iB = InStr(sBody, "text")
  If iB > 0 Then
    If Mid$(sBody, iB + 5, 1) = Chr$(34) Then
      iTemp = iB + 5
      Do While sCharTmp <> Chr$(34)
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    Else
      iTemp = iB + 4
      Do While sCharTmp <> " "
        If sCharTmp = ">" Then Exit Do
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    End If
    GetPageProperties.Text = sTemp
  Else
    GetPageProperties.Text = ""
  End If
  
  sTemp = "": sCharTmp = ""
  iB = InStr(sBody, "link")
  If iB > 0 Then
    If Mid$(sBody, iB + 5, 1) = Chr$(34) Then
      iTemp = iB + 5
      Do While sCharTmp <> Chr$(34)
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    Else
      iTemp = iB + 4
      Do While sCharTmp <> " "
        If sCharTmp = ">" Then Exit Do
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    End If
    GetPageProperties.Link = sTemp
  Else
    GetPageProperties.Link = ""
  End If

  sTemp = "": sCharTmp = ""
  iB = InStr(sBody, "vlink")
  If iB > 0 Then
    If Mid$(sBody, iB + 6, 1) = Chr$(34) Then
      iTemp = iB + 6
      Do While sCharTmp <> Chr$(34)
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    Else
      iTemp = iB + 5
      Do While sCharTmp <> " "
        If sCharTmp = ">" Then Exit Do
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    End If
    GetPageProperties.Visited = sTemp
  Else
    GetPageProperties.Visited = ""
  End If
  
  sTemp = "": sCharTmp = ""
  iB = InStr(sBody, "alink")
  If iB > 0 Then
    If Mid$(sBody, iB + 6, 1) = Chr$(34) Then
      iTemp = iB + 6
      Do While sCharTmp <> Chr$(34)
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    Else
      iTemp = iB + 5
      Do While sCharTmp <> " "
        If sCharTmp = ">" Then Exit Do
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    End If
    GetPageProperties.Active = sTemp
  Else
    GetPageProperties.Active = ""
  End If
  
  sTemp = "": sCharTmp = ""
  iB = InStr(sBody, "leftmargin")
  If iB > 0 Then
    If Mid$(sBody, iB + 11, 1) = Chr$(34) Then
      iTemp = iB + 11
      Do While sCharTmp <> Chr$(34)
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    Else
      iTemp = iB + 10
      Do While sCharTmp <> " "
        If sCharTmp = ">" Then Exit Do
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    End If
    GetPageProperties.LeftMargin = sTemp
  Else
    GetPageProperties.LeftMargin = ""
  End If
  
  sTemp = "": sCharTmp = ""
  iB = InStr(sBody, "topmargin")
  If iB > 0 Then
    If Mid$(sBody, iB + 10, 1) = Chr$(34) Then
      iTemp = iB + 10
      Do While sCharTmp <> Chr$(34)
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    Else
      iTemp = iB + 9
      Do While sCharTmp <> " "
        If sCharTmp = ">" Then Exit Do
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    End If
    GetPageProperties.TopMargin = sTemp
  Else
    GetPageProperties.TopMargin = ""
  End If
  
  sTemp = "": sCharTmp = ""
  iB = InStr(sBody, "marginwidth")
  If iB > 0 Then
    If Mid$(sBody, iB + 12, 1) = Chr$(34) Then
      iTemp = iB + 12
      Do While sCharTmp <> Chr$(34)
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    Else
      iTemp = iB + 11
      Do While sCharTmp <> " "
        If sCharTmp = ">" Then Exit Do
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    End If
    GetPageProperties.MarginWidth = sTemp
  Else
    GetPageProperties.MarginWidth = ""
  End If
  
  sTemp = "": sCharTmp = ""
  iB = InStr(sBody, "marginheight")
  If iB > 0 Then
    If Mid$(sBody, iB + 13, 1) = Chr$(34) Then
      iTemp = iB + 13
      Do While sCharTmp <> Chr$(34)
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    Else
      iTemp = iB + 12
      Do While sCharTmp <> " "
        If sCharTmp = ">" Then Exit Do
        iTemp = iTemp + 1
        sTemp = sTemp & sCharTmp
        sCharTmp = Mid$(sBody, iTemp, 1)
      Loop
    End If
    GetPageProperties.MarginHeight = sTemp
  Else
    GetPageProperties.MarginHeight = ""
  End If

End Function

Sub ChangeTitle(NewTitle As String)
  Dim iStart As Integer, iEnd As Integer
  
  If InStr(frmMain.ASPEdit, "<title>") = 0 Then Exit Sub
  
  iStart = InStr(frmMain.ASPEdit.Text, "<title>") + 6
  iEnd = InStr(frmMain.ASPEdit.Text, "</title>")
  
  frmMain.ASPEdit.SelStart = iStart
  frmMain.ASPEdit.SelLength = iEnd - iStart - 1
  frmMain.ASPEdit.SelText = NewTitle
  frmMain.ASPEdit.SelStart = 0
End Sub

Function GetTitle() As String
  Dim iStart As Integer, iEnd As Integer
  
  If InStr(frmMain.ASPEdit, "<title>") = 0 Then Exit Function
  
  iStart = InStr(frmMain.ASPEdit.Text, "<title>") + 6
  iEnd = InStr(frmMain.ASPEdit.Text, "</title>")
  
  frmMain.ASPEdit.SelStart = iStart
  frmMain.ASPEdit.SelLength = iEnd - iStart - 1
  GetTitle = frmMain.ASPEdit.SelText
  frmMain.ASPEdit.SelStart = 0
End Function

Function ChangePProp(Title As String, Background As String, BGProperties As String, BGColor As String, Text As String, Link As String, _
    Visited As String, Active As String, LeftMargin As String, TopMargin As String, MarginWidth As String, MarginHeight As String) As PageProperties
    
  ChangePProp.Title = Title
  ChangePProp.Background = Background
  ChangePProp.BGProperties = BGProperties
  ChangePProp.BGColor = BGColor
  ChangePProp.Text = Text
  ChangePProp.Link = Link
  ChangePProp.Visited = Visited
  ChangePProp.Active = Active
  ChangePProp.LeftMargin = LeftMargin
  ChangePProp.TopMargin = TopMargin
  ChangePProp.MarginWidth = MarginWidth
  ChangePProp.MarginHeight = MarginHeight
    
End Function

Function GetTagProperties(sTagLine As String) As TagProperties

  sTagLine = Replace(sTagLine, "<", "")
  sTagLine = Replace(sTagLine, ">", "")
  
  If InStr(sTagLine, " ") = 0 Then Exit Function
  
  GetTagProperties.Name = Left$(sTagLine, InStr(sTagLine, " ") - 1)
  
  sTagLine = Replace(sTagLine, GetTagProperties.Name & " ", "")

  Dim iPropCount As Integer
  Dim i As Integer, sTmpChar As String, sNewLine
  For i = 1 To Len(sTagLine)
    sTmpChar = Mid$(sTagLine, i, 1)
    sNewLine = sNewLine & sTmpChar
    If sTmpChar = "=" Then
      iPropCount = iPropCount + 1
      sNewLine = Replace(sNewLine, " ", "")
      sNewLine = Replace(sNewLine, "=", "")
      GetTagProperties.Properties(iPropCount, 0) = sNewLine
      sNewLine = ""
      i = i + 1
      sTmpChar = Mid$(sTagLine, i, 1)
      If sTmpChar = Chr$(34) Then
        i = i + 1
        sTmpChar = Mid$(sTagLine, i, 1)
        Do While sTmpChar <> Chr$(34)
          sNewLine = sNewLine & sTmpChar
          i = i + 1
          sTmpChar = Mid$(sTagLine, i, 1)
        Loop
        GetTagProperties.Properties(iPropCount, 1) = sNewLine
        sNewLine = ""
      Else
        If InStr(i, sTagLine, " ") = 0 Then
          sNewLine = Mid$(sTagLine, i, Len(sTagLine) - i + 1)
        Else
          sNewLine = Mid$(sTagLine, i, InStr(i, sTagLine, " ") - i)
          i = i + Len(sNewLine)
        End If
        GetTagProperties.Properties(iPropCount, 1) = sNewLine
        sNewLine = ""
      End If
    End If
  Next
  
End Function

Function GetColor() As String
  Dim point As POINTAPI
  GetCursorPos point
  frmColors.Move point.X * 15, point.Y * 15
  frmColors.Show 1
  GetColor = frmColors.mColor
End Function

Public Function FormatRGBString(val As Long) As String

  Dim Color As String
  Dim pad As Long
  Dim r As String
  Dim g As String
  Dim b As String

  Color = Hex$(val)
  pad = 6 - Len(Color)
    
  If pad Then
    Color = String$(pad, "0") & Color
  End If
  r = Right$(Color, 2)
  g = Mid$(Color, 3, 2)
  b = Left$(Color, 2)

  Color = "#" & r & g & b
  
  FormatRGBString = Color

End Function

Public Function HexCol2Long(val As String) As Long

  Dim HexCol As String
  
  If val = "" Then HexCol2Long = &HE0E0E0: Exit Function
  
  Select Case LCase(val)
    Case "white"
      HexCol2Long = vbWhite: Exit Function
    Case "black"
      HexCol2Long = vbBlack: Exit Function
    Case "red"
      HexCol2Long = vbRed: Exit Function
    Case "green"
      HexCol2Long = vbGreen: Exit Function
    Case "blue"
      HexCol2Long = vbBlue: Exit Function
    Case "yellow"
      HexCol2Long = vbYellow: Exit Function
    Case "magenta"
      HexCol2Long = vbMagenta: Exit Function
    Case "cyan"
      HexCol2Long = vbCyan: Exit Function
  End Select
  
  HexCol = Replace(val, "#", "")
  
  HexCol2Long = RGB(CByte("&H" & Left(HexCol, 2)), CByte("&H" & Mid(HexCol, 3, 2)), CByte("&H" & Right(HexCol, 2)))

End Function

Function FileExists(FullFileName As String) As Boolean
    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        Open FullFileName For Input As #1
        Close #1
        'no error, file exists
        FileExists = True
    Exit Function
MakeF:
        'error, file does Not exist
        FileExists = False
    Exit Function
End Function

Public Function DirExists(ByVal strDirName As String) As Integer
    Const strWILDCARD$ = "*.*"

    Dim strDummy As String

    On Error Resume Next

    If Right$(strDirName, 1) <> "\" Then strDirName = strDirName & "\"
    strDummy = Dir$(strDirName & strWILDCARD, vbDirectory)
    DirExists = Not (strDummy = vbNullString)

    Err = 0
End Function

Function ReadAttrib(ByVal lpAttribName As String, ByVal lpString As String) As String
'readattrib("name","<BODY name=foo>") returns foo
On Error Resume Next
Dim lnPos1 As Long, lnPos2 As Long, tmp As String

  lnPos1 = InStr(1, lpString, " " & lpAttribName & "=", vbTextCompare)
  If lnPos1 > 0 Then lnPos1 = lnPos1 + Len(lpAttribName) + 2 Else ReadAttrib = lpString: Exit Function
  
  lnPos2 = InStr(lnPos1 + 1, lpString, " ")
  If Mid$(lpString, lnPos1, 1) = "'" Then lnPos2 = InStr(lnPos1 + 1, lpString, "'")
  If Mid$(lpString, lnPos1, 1) = Chr(34) Then lnPos2 = InStr(lnPos1 + 1, lpString, Chr(34))
  If lnPos2 = 0 Then lnPos2 = Len(lpString) + 1
  tmp = Mid$(lpString, lnPos1, lnPos2 - lnPos1)
  
  If Left(tmp, 1) = Chr(34) Or Left(tmp, 1) = "'" Then tmp = Right(tmp, Len(tmp) - 1)
  If Right(tmp, 1) = Chr(34) Or Right(tmp, 1) = "'" Then tmp = Left(tmp, Len(tmp) - 1)
  
  tmp = Replace(tmp, "\'", "'")
  tmp = Replace(tmp, "\" & Chr(34), Chr(34))
  
  ReadAttrib = tmp
End Function
