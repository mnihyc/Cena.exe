VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmView 
   AutoRedraw      =   -1  'True
   Caption         =   "View"
   ClientHeight    =   6795
   ClientLeft      =   7695
   ClientTop       =   3870
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   5550
   Begin VB.CommandButton CommandDiff 
      Caption         =   "Diff"
      Height          =   285
      Left            =   1170
      TabIndex        =   15
      Top             =   2700
      Width           =   735
   End
   Begin VB.CommandButton CommandJump 
      Caption         =   "Jump"
      Height          =   315
      Index           =   1
      Left            =   2505
      TabIndex        =   14
      Top             =   4485
      Width           =   750
   End
   Begin VB.CommandButton CommandJump 
      Caption         =   "Jump"
      Height          =   315
      Index           =   0
      Left            =   2460
      TabIndex        =   11
      Top             =   2685
      Width           =   750
   End
   Begin RichTextLib.RichTextBox Text3 
      Height          =   1335
      Left            =   0
      TabIndex        =   10
      Top             =   4800
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2355
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmView.frx":0000
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   1455
      Left            =   0
      TabIndex        =   9
      Top             =   3000
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2566
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmView.frx":009D
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1455
      Left            =   0
      TabIndex        =   8
      Top             =   1200
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2566
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmView.frx":013A
   End
   Begin VB.Label LabelLine2 
      Alignment       =   2  'Center
      Caption         =   "Line: "
      Height          =   195
      Left            =   3255
      TabIndex        =   13
      Top             =   4530
      Width           =   2205
   End
   Begin VB.Label LabelLine1 
      Alignment       =   2  'Center
      Caption         =   "Line: "
      Height          =   195
      Left            =   3195
      TabIndex        =   12
      Top             =   2745
      Width           =   2310
   End
   Begin VB.Label Label9 
      Caption         =   "1000ms"
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Running time: "
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "standard111.out"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   4560
      Width           =   2385
   End
   Begin VB.Label Label6 
      Caption         =   ".in"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Extra information"
      BeginProperty Font 
         Name            =   "ËÎÌו"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   6240
      Width           =   5175
   End
   Begin VB.Label Label3 
      Caption         =   "Output:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Input:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Accept"
      BeginProperty Font 
         Name            =   "ËÎÌו"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FormOldWidth As Long
Public FormOldHeight As Long
Private res As type_res
Private Const SHOW_LENGTH As Long = 20480
Private Const LINE_DIFF As Long = 200
Dim s1() As String, s2() As String
Dim curLine1&, curLine2&

Friend Sub SetRes(tres As type_res)
  res = tres
  s1 = Split(res.out, vbCrLf)
  If res.stdreaded = False Then
    res.sincontent = ReadFromFile(res.sin, SHOW_LENGTH, True)
    res.soutcontent = ReadFromFile(res.sout)
  End If
  s2 = Split(res.soutcontent, vbCrLf)
End Sub

Public Sub ReSet()
With Me
  Dim strLine As String, sstr As String, str1 As String
  sstr = Left(res.soutcontent, SHOW_LENGTH)
  str1 = Left(res.sincontent, SHOW_LENGTH)
  If .Visible = False And .FormOldWidth > 0 And .FormOldHeight > 0 Then
    .Width = .FormOldWidth
    .Height = .FormOldHeight
  End If
  .Show
  funstate .Label1, res.rw, True
  Dim leng&
  leng = IIf(res.stdreaded = True, Len(res.sincontent), FileLen(res.sin))
  If leng > SHOW_LENGTH Then
    .Text1.Text = Left(str1, SHOW_LENGTH)
    .Text1.Text = .Text1.Text & "<skip around " & Trim(Format(leng - SHOW_LENGTH, "###,###")) & "bytes>"
  Else
    .Text1.Text = str1
  End If
  
  leng = Len(res.out)
  If leng > SHOW_LENGTH Then
    .Text2.Text = Left(res.out, SHOW_LENGTH)
    .Text2.Text = .Text2.Text & "<skip around " & Trim(Format(leng - SHOW_LENGTH, "###,###")) & "bytes>"
  Else
    .Text2.Text = res.out
  End If
  
  leng = Len(res.soutcontent)
  If leng > SHOW_LENGTH Then
    .Text3.Text = Left(sstr, SHOW_LENGTH)
    .Text3.Text = .Text3.Text & "<skip around " & Trim(Format(leng - SHOW_LENGTH, "###,###")) & "bytes>"
  Else
    .Text3.Text = sstr
  End If
  .Label5.Caption = res.err
  .Label6.Caption = Right(res.sin, Len(res.sin) - InStrRev(res.sin, "\"))
  .Label7.Caption = Right(res.sout, Len(res.sout) - InStrRev(res.sout, "\"))
  .Label9.Caption = Format(res.runningtime, "###,###") & "ms"
  curLine1 = 1: curLine2 = 1
  .Text1.SelStart = 0: .Text1.SelLength = 0: .Text1.Refresh
  .Text2.SelStart = 0: .Text2.SelLength = 0: .Text2.Refresh
  .Text3.SelStart = 0: .Text3.SelLength = 0: .Text3.Refresh
  Call Text2_SelChange: Call Text3_SelChange
  .Refresh
  .SetFocus
End With
  If res.stdreaded = False Then
    res.sincontent = ""
    res.soutcontent = ""
  End If
End Sub

' Use ByRef to prevent unnecessary waste of time
Private Sub DoTextJump(Text As RichTextBox, ByRef sarr() As String, ByRef curLine&, Optional line& = 1)
  Dim i&, leng&: leng = UBound(sarr) + 1
  If line <= 0 Then: line = 0
  If line > leng Then: line = leng
  Dim str$: str = ""
  Dim sline&, eline&
  sline = IIf(line - LINE_DIFF >= 1, line - LINE_DIFF, 1)
  curLine = sline
  eline = IIf(line + LINE_DIFF <= leng, line + LINE_DIFF, leng)
  Dim sbyte&
  Text.Text = ""
  If sline > 1 Then
    sbyte = 0
    For i = 1 To sline - 1
      sbyte = sbyte + Len(sarr(i - 1)) + 2
    Next i
    sbyte = sbyte - 2
    Text.Text = Text.Text & "<skip around " & Trim(Format(sbyte, "###,###")) & "bytes>"
  End If
  
  Dim ss&
  For i = sline To eline
    If i = line Then: ss = Len(Text.Text)
    Text.Text = Text.Text & sarr(i - 1) & vbCrLf
  Next i
  
  If eline < leng Then
    sbyte = 0
    For i = eline + 1 To leng
      sbyte = sbyte + Len(sarr(i - 1)) + 2
    Next i
    sbyte = sbyte - 2
    Text.Text = Text.Text & "<skip around " & Trim(Format(sbyte, "###,###")) & "bytes>"
  End If
  Text.SelStart = Len(Text.Text): Text.SelLength = 0
  
  'Dim LineIndex As Long
  'LineIndex = SendMessage(Text.hwnd, EM_LINEINDEX, line - curLine + 1, ByVal 0&)
  'SendMessage Text.hwnd, EM_SETSEL, LineIndex, ByVal LineIndex + 1
  Text.SelStart = ss: Text.SelLength = 0
End Sub

Private Sub CommandDiff_Click()
  Dim i&, ss1$, ss2$
  i = SendMessage(Text2.hwnd, EM_LINEFROMCHAR, Text2.SelStart, 0&) + curLine1 - 1
  Do While i <= UBound(s1) Or i <= UBound(s2)
    If i > UBound(s1) Then
      ss1 = ""
    Else
      ss1 = s1(i)
    End If
    If i > UBound(s2) Then
      ss2 = ""
    Else
      ss2 = s2(i)
    End If
    If Trim(ss1) <> Trim(ss2) Then
      Call DoTextJump(Text2, s1, curLine1, i + 1)
      Call DoTextJump(Text3, s2, curLine2, i + 1)
      Exit Do
    End If
    i = i + 1
  Loop
  Text2.SetFocus
End Sub

Private Sub CommandJump_Click(Index As Integer)
  Dim Text As RichTextBox
  If Index = 0 Then
    Set Text = Text2
  Else
    Set Text = Text3
  End If
  
  Dim line&
  line = Val(ShowDialog("Input which line to jump", Me))
  If Index = 0 Then
    Call DoTextJump(Text, s1, curLine1, line)
  Else
    Call DoTextJump(Text, s2, curLine2, line)
  End If
  
  Text.SetFocus
End Sub

Private Sub Form_Load()
  'Text1.Locked = True
  'Text2.Locked = True
  'Text3.Locked = True
  Dim Obj As Control
  FormOldWidth = Me.ScaleWidth
  FormOldHeight = Me.ScaleHeight
  For Each Obj In Me
    Obj.Tag = Obj.Left & " " & Obj.Top & " " & Obj.Width & " " & Obj.Height & " "
  Next Obj
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = -1
  Hide
  frmTest.Show
End Sub

Private Sub Form_Resize()
  Dim PosS() As String
  Dim Obj As Control
  Dim Pos(4) As Double, ScaleX As Double, ScaleY As Double
  Dim i As Long
  ScaleX = Me.ScaleWidth / FormOldWidth
  ScaleY = Me.ScaleHeight / FormOldHeight
  For Each Obj In Me
    PosS = Split(Obj.Tag, " ")
    For i = 0 To 3
      Pos(i) = CDbl(PosS(i))
    Next i
    Obj.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
  Next Obj
End Sub

Private Sub Text2_SelChange()
  LabelLine1.Caption = "Line: " & Trim(str(SendMessage(Text2.hwnd, EM_LINEFROMCHAR, Text2.SelStart, 0&) + curLine1)) & "/" & Trim(str(UBound(s1) + 1))
End Sub

Private Sub Text3_SelChange()
  LabelLine2.Caption = "Line: " & Trim(str(SendMessage(Text3.hwnd, EM_LINEFROMCHAR, Text3.SelStart, 0&) + curLine2)) & "/" & Trim(str(UBound(s2) + 1))
End Sub
