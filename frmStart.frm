VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data"
   ClientHeight    =   3975
   ClientLeft      =   8955
   ClientTop       =   4755
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5370
   Begin VB.CheckBox Check1 
      Caption         =   "Reuse test-data"
      Height          =   225
      Left            =   3600
      TabIndex        =   24
      Top             =   3690
      Value           =   1  'Checked
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   21
      Top             =   3105
      Width           =   1695
   End
   Begin VB.TextBox txt7 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   18
      Text            =   "1000"
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox txt6 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   15
      Text            =   "{STDOUT}"
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox txt5 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   12
      Text            =   "{STDIN}"
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox txt4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Text            =   "{FileName}\*[0].out"
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox txt3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Text            =   "{FileName}\*[0].in"
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txt2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Text            =   "{FileName}.exe"
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lbln2 
      Caption         =   "0"
      Height          =   255
      Left            =   4320
      TabIndex        =   23
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label lbln1 
      Caption         =   "0"
      Height          =   255
      Left            =   4320
      TabIndex        =   22
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lbl7 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   4800
      TabIndex        =   20
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Timeout :"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lbl6 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "file-out :"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lbl5 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "file-in :"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lbl4 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "STD-out :"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lbl3 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "STD-in :"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lbl2 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Exefile :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Filename :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lbl1 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UpdateLblnl()
  If lbln2.Caption <> lbln1.Caption Or lbln1.Caption = "0" Or lbln2.Caption = "0" Then
    funstate lbl3, False
    funstate lbl4, False
  Else
    funstate lbl3, True
    funstate lbl4, True
  End If
End Sub

Private Sub UpdateIOFile(s As String, ByRef lbl As Label, ByRef lbln As Label, IO As Boolean)
  Dim i As Integer, j As Integer
  If Len(s) > 11 And Right(s, 11) = "\config.ini" Then
    Dim sbin() As Byte
    Open App.Path & "\" & s For Binary As #233
      ReDim sbin(LOF(233) - 1)
      Get #233, , sbin
    Close #233
    Dim tst As String
    tst = StrConv(sbin, vbUnicode)
    Dim spstr() As String
    spstr = Split(tst, vbLf)
    If IO = True Then
      stdoutnum = Val(spstr(0))
      For i = 1 To stdoutnum
        stdout(i) = App.Path & "\" & Split(s, "\")(0) & "\" & Split(spstr(i), "|")(1)
      Next i
    Else
      stdinnum = Val(spstr(0))
      For i = 1 To stdinnum
        stdin(i) = App.Path & "\" & Split(s, "\")(0) & "\" & Split(spstr(i), "|")(0)
      Next i
    End If
    funstate lbl, True
    lbln.Caption = Trim(str(Val(spstr(0))))
    UpdateLblnl
    Exit Sub
  End If
  i = InStr(1, s, "[")
  If InStr(1, s, "*") > 0 And i = 0 Then
    Dim num%, file$
    Dim dp$
    Dim snp%: snp = InStr(1, s, "\")
    If snp <= 0 Then: snp = Len(s)
    dp = App.Path & "\" & Left(s, snp)
    file = Dir(App.Path & "\" & s)
    Do Until file = ""
      num = num + 1
      If IO = True Then
        stdout(num) = dp & file
      Else
        stdin(num) = dp & file
      End If
      file = Dir()
    Loop
    If IO = True Then
      stdoutnum = num
    Else
      stdinnum = num
    End If
    If num > 0 Then
      funstate lbl, True
      lbln.Caption = Trim(str(num))
    Else
      funstate lbl, False
      lbln.Caption = "0"
    End If
    UpdateLblnl
    Exit Sub
  End If
  If i = 0 Then
    'MsgBox "No support!", vbInformation + vbOKOnly
    'funstate lbl, False
    'Exit Sub
    Dim fss As Object
    Set fss = CreateObject("Scripting.FileSystemObject")
    If fss.FileExists(App.Path & "\" & s) Then
      funstate lbl, True
      lbln.Caption = "1"
      If IO = True Then
        stdoutnum = 1
        stdout(1) = App.Path & "\" & s
      Else
        stdinnum = 1
        stdin(1) = App.Path & "\" & s
      End If
    Else
      funstate lbl, False
      lbln.Caption = "0"
    End If
    UpdateLblnl
    Set fss = Nothing
    Exit Sub
  End If
  j = i
  Dim s1 As String
  s1 = ""
  Do While j < Len(s)
    j = j + 1
    If Mid(s, j, 1) <> "]" Then
      s1 = s1 + Mid(s, j, 1)
    Else
      Exit Do
    End If
  Loop
  If s1 = "" Then
    MsgBox "Unknow ""]""!", vbExclamation + vbOKOnly
    funstate lbl, False
    Exit Sub
  End If
  Dim k As Integer, p As Integer, sf As Integer
  Dim lv%
  sf = Val(s1)
  k = 0
  For p = 200 To sf Step -1
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    'If fs.FileExists(App.Path & "\" & Mid(s, 1, i - 1) & Trim(Str(p)) & Mid(s, j + 1, Len(s) - j)) Then
    snp = InStr(1, s, "\")
    If snp <= 0 Then: snp = Len(s)
    dp = App.Path & "\" & Left(s, snp)
    file = Trim(Dir(App.Path & "\" & Mid(s, 1, i - 1) & Trim(str(p)) & Mid(s, j + 1, Len(s) - j)))
recomp:
    If file <> "" Then
      k = k + 1
      If IO = True Then
        stdout(k) = dp & file
        For lv = 1 To k - 1
          If stdout(k) = stdout(lv) Then
            k = k - 1
            file = Dir()
            GoTo recomp
          End If
        Next lv
      Else
        stdin(k) = dp & file
        For lv = 1 To k - 1
          If stdin(k) = stdin(lv) Then
            k = k - 1
            file = Dir()
            GoTo recomp
          End If
        Next lv
      End If
    Else
      If k > 0 Then: Exit For
    End If
  Next p
  Dim tss$
  If IO = True Then
    For lv = 1 To stdoutnum \ 2
      tss = stdout(lv)
      stdout(lv) = stdout(stdoutnum - lv + 1)
      stdout(stdoutnum - lv + 1) = tss
    Next lv
    stdoutnum = k
  Else
    For lv = 1 To stdinnum \ 2
      tss = stdin(lv)
      stdin(lv) = stdin(stdinnum - lv + 1)
      stdin(stdinnum - lv + 1) = tss
    Next lv
    stdinnum = k
  End If
  If k = 0 Then
    funstate lbl, False
  Else
    funstate lbl, True
  End If
  lbln.Caption = Trim(str(k))
  UpdateLblnl
End Sub

Private Function allg()
  If lbl1.Caption <> "Yes" Then: GoTo be
  If lbl2.Caption <> "Yes" Then: GoTo be
  If lbl3.Caption <> "Yes" Then: GoTo be
  If lbl4.Caption <> "Yes" Then: GoTo be
  If lbl5.Caption <> "Yes" Then: GoTo be
  If lbl6.Caption <> "Yes" Then: GoTo be
  If lbl7.Caption <> "Yes" Then: GoTo be
  allg = True
  Exit Function
be:
  allg = False
End Function
Private Sub funstate(lbl As Label, state As Boolean)
  If Not state Then
    lbl.ForeColor = vbRed
    lbl.Caption = "No"
  Else
    lbl.ForeColor = vbGreen
    lbl.Caption = "Yes"
  End If
End Sub
Private Sub Form_Load()
  Show
  funstate lbl1, False
  funstate lbl2, False
  funstate lbl3, False
  funstate lbl4, False
  funstate lbl5, False
  funstate lbl6, False
  funstate lbl7, True
  timeout = 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub txt1_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    Call Command1_Click
  End If
End Sub

Private Sub txt1_LostFocus()
  If txt1.Text = "" Then
    funstate lbl1, False
    Exit Sub
  End If
  funstate lbl1, True
  exefile = App.Path & "\" & txt1.Text
  
  Dim fs As Object
  Set fs = CreateObject("Scripting.FileSystemObject")
  If fs.FileExists(exefile & "\config.ini") Then
    txt3.Text = "{FileName}" & "\config.ini"
    txt4.Text = "{FileName}" & "\config.ini"
  End If
  
  txt2_LostFocus
  txt3_LostFocus
  txt4_LostFocus
  txt5_LostFocus
  txt6_LostFocus
End Sub

Private Sub txt2_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    Call Command1_Click
  End If
End Sub

Private Sub txt2_LostFocus()
  Dim s As String
  s = LCase(txt2.Text)
  s = Replace(s, "{filename}", "{fn}")
  s = Replace(s, "{fn}", txt1.Text)
  Dim fs As Object
  Set fs = CreateObject("Scripting.FileSystemObject")
  If Not fs.FileExists(App.Path & "\" & s) Then
    funstate lbl2, False
  Else
    funstate lbl2, True
    exefile = App.Path & "\" & s
  End If
End Sub

Private Sub txt3_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    Call Command1_Click
  End If
End Sub

Private Sub txt3_LostFocus()
  lbln1.Caption = "0"
  If txt3.Text = "" Or txt1.Text = "" Then
    funstate lbl3, False
    Exit Sub
  End If
  Dim s As String
  s = LCase(txt3.Text)
  s = Replace(s, "{filename}", "{fn}")
  s = Replace(s, "{fn}", txt1.Text)
  UpdateIOFile s, lbl3, lbln1, False
  UpdateLblnl
End Sub

Private Sub txt4_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    Call Command1_Click
  End If
End Sub

Private Sub txt4_LostFocus()
  lbln2.Caption = "0"
  If txt4.Text = "" Or txt1.Text = "" Then
    funstate lbl4, False
    Exit Sub
  End If
  Dim s As String
  s = LCase(txt4.Text)
  s = Replace(s, "{filename}", "{fn}")
  s = Replace(s, "{fn}", txt1.Text)
  UpdateIOFile s, lbl4, lbln2, True
  UpdateLblnl
End Sub

Private Sub txt5_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    Call Command1_Click
  End If
End Sub

Private Sub txt5_LostFocus()
 If txt5.Text = "" Or txt1.Text = "" Then
    funstate lbl5, False
    Exit Sub
  End If
  Dim s As String
  s = LCase(txt5.Text)
  s = Replace(s, "{std.in}", "{fn}.in")
  s = Replace(s, "{filename}", "{fn}")
  s = Replace(s, "{fn}", txt1.Text)
  funstate lbl5, True
  infile = s
  If Trim(UCase(s)) <> "{STDIN}" Then
    infile = App.Path & "\" & infile
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(infile) Then: Kill infile
  End If
End Sub

Private Sub txt6_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    Call Command1_Click
  End If
End Sub

Private Sub txt6_LostFocus()
  If txt6.Text = "" Or txt1.Text = "" Then
    funstate lbl6, False
    Exit Sub
  End If
  Dim s As String
  s = LCase(txt6.Text)
  s = Replace(s, "{std.out}", "{fn}.out")
  s = Replace(s, "{filename}", "{fn}")
  s = Replace(s, "{fn}", txt1.Text)
  funstate lbl6, True
  outfile = s
  If Trim(UCase(s)) <> "{STDOUT}" Then
    outfile = App.Path & "\" & outfile
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(outfile) Then: Kill outfile
  End If
End Sub

Private Sub txt7_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    Call Command1_Click
  End If
End Sub

Private Sub txt7_LostFocus()
  Dim i As Long
  i = Val(txt7.Text)
  If i < 100 Or i > 60000 Then
    funstate lbl7, False
  Else
    funstate lbl7, True
    timeout = i
  End If
End Sub
Private Sub txt1_KeyPress(key As Integer)
  'If key = vbKeyBack Or key = vbKeyDelete Or key = vbKeyInsert Or key = vbKeyClear Then: Exit Sub
  'If Not Chr(key) Like "[a-zA-Z0-9./\]" Then: key = 0
End Sub
Private Sub txt2_KeyPress(key As Integer)
  'If key = vbKeyBack Or key = vbKeyDelete Or key = vbKeyInsert Or key = vbKeyClear Then: Exit Sub
  'If Not Chr(key) Like "[a-zA-Z0-9{}./\]" Then: key = 0
End Sub
Private Sub txt3_KeyPress(key As Integer)
  'If key = vbKeyBack Or key = vbKeyDelete Or key = vbKeyInsert Or key = vbKeyClear Or key = Asc("[") Or key = Asc("]") Then: Exit Sub
  'If Not Chr(key) Like "[a-zA-Z0-9{}./\]" Then: key = 0
End Sub
Private Sub txt4_KeyPress(key As Integer)
  'If key = vbKeyBack Or key = vbKeyDelete Or key = vbKeyInsert Or key = vbKeyClear Or key = Asc("[") Or key = Asc("]") Then: Exit Sub
  'If Not Chr(key) Like "[a-zA-Z0-9{}./\]" Then: key = 0
End Sub
Private Sub txt5_KeyPress(key As Integer)
  'If key = vbKeyBack Or key = vbKeyDelete Or key = vbKeyInsert Or key = vbKeyClear Then: Exit Sub
  'If Not Chr(key) Like "[a-zA-Z0-9{}./\]" Then: key = 0
End Sub
Private Sub txt6_KeyPress(key As Integer)
  'If key = vbKeyBack Or key = vbKeyDelete Or key = vbKeyInsert Or key = vbKeyClear Then: Exit Sub
  'If Not Chr(key) Like "[a-zA-Z0-9{}./\]" Then: key = 0
End Sub
Private Sub txt7_KeyPress(key As Integer)
  If key = Asc("-") Then: key = 0
  If key = vbKeyBack Or key = vbKeyDelete Or key = vbKeyInsert Or key = vbKeyClear Then: Exit Sub
  If (Not Chr(key) Like "[0-9]") Then: key = 0
End Sub
Private Sub Command1_Click()
  Unload frmTest
  txt1_LostFocus
  txt2_LostFocus
  txt3_LostFocus
  txt4_LostFocus
  txt5_LostFocus
  txt6_LostFocus
  Dim okstd As Boolean
  okstd = True
  If Trim(UCase(txt5.Text)) = "{STDIN}" Then: okstd = IIf(Trim(UCase(txt6.Text)) = "{STDOUT}", True, False)
  If Trim(UCase(txt6.Text)) = "{STDOUT}" Then: okstd = IIf(Trim(UCase(txt5.Text)) = "{STDIN}", True, False)
  If Not allg() Then
    MsgBox "Can't start!", vbExclamation + vbOKOnly
    Exit Sub
  End If
  If Not okstd Then
    MsgBox "Use file or std???" & vbCrLf & "Can't start!", vbExclamation + vbOKOnly
    Exit Sub
  End If
  Hide
  frmTest.Show
End Sub

