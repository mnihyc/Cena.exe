VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Result"
   ClientHeight    =   6000
   ClientLeft      =   8565
   ClientTop       =   3030
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8055
   Begin VB.CommandButton cmd_rta 
      Caption         =   "Re-test ALL"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   36
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   6840
      Top             =   5520
   End
   Begin VB.CommandButton cmd_rt 
      Caption         =   "Re-test"
      Height          =   375
      Index           =   6
      Left            =   6600
      TabIndex        =   35
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton cmd_rt 
      Caption         =   "Re-test"
      Height          =   375
      Index           =   5
      Left            =   6600
      TabIndex        =   34
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmd_rt 
      Caption         =   "Re-test"
      Height          =   375
      Index           =   4
      Left            =   6600
      TabIndex        =   33
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmd_rt 
      Caption         =   "Re-test"
      Height          =   375
      Index           =   3
      Left            =   6600
      TabIndex        =   32
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmd_rt 
      Caption         =   "Re-test"
      Height          =   375
      Index           =   2
      Left            =   6600
      TabIndex        =   31
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmd_rt 
      Caption         =   "Re-test"
      Height          =   375
      Index           =   1
      Left            =   6600
      TabIndex        =   30
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmd_rt 
      Caption         =   "Re-test"
      Height          =   375
      Index           =   0
      Left            =   6600
      TabIndex        =   29
      Top             =   120
      Width           =   855
   End
   Begin VB.VScrollBar vs1 
      Height          =   6015
      Left            =   7680
      Max             =   100
      Min             =   1
      TabIndex        =   0
      Top             =   0
      Value           =   1
      Width           =   375
   End
   Begin VB.Label lbl_rs 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5280
      TabIndex        =   28
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lbl_rs 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   5280
      TabIndex        =   27
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lbl_rs 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   5280
      TabIndex        =   26
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lbl_rs 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5280
      TabIndex        =   25
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lbl_rs 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   24
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lbl_rs 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5280
      TabIndex        =   23
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lbl_rs 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5280
      TabIndex        =   22
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbl_txt 
      Caption         =   "Waiting ..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2880
      TabIndex        =   21
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label lbl_txt 
      Caption         =   "Waiting ..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2880
      TabIndex        =   20
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label lbl_txt 
      Caption         =   "Waiting ..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2880
      TabIndex        =   19
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label lbl_txt 
      Caption         =   "Waiting ..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2880
      TabIndex        =   18
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label lbl_txt 
      Caption         =   "Waiting ..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   17
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lbl_txt 
      Caption         =   "Waiting ..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   16
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lbl_txt 
      Caption         =   "Waiting ..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   15
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lbl_num 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2280
      TabIndex        =   14
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label lbl_num 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2280
      TabIndex        =   13
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lbl_num 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2280
      TabIndex        =   12
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label lbl_num 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   11
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lbl_num 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   10
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lbl_num 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   9
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lbl_num 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.Label slbl 
      Caption         =   "Test Point "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label slbl 
      Caption         =   "Test Point "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label slbl 
      Caption         =   "Test Point "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label slbl 
      Caption         =   "Test Point "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label slbl 
      Caption         =   "Test Point "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label slbl 
      Caption         =   "Test Point "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label slbl 
      Caption         =   "Test Point "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim res() As type_res
Dim ylbl_num(6) As String
Dim stoptest As Boolean
Private Const BUFFER_LENGTH As Long = 102400
Dim resenabled(1050) As Boolean

Private Sub cmd_rt_Click(Index As Integer)
  If Val(lbl_num(Index)) > stdinnum Or Val(lbl_num(Index)) <= 0 Then: Exit Sub
  res_id Val(lbl_num(Index)), False
  res(Val(lbl_num(Index))).state = False
  ylbl_num(Index) = ""
  Call doexec(res(Val(lbl_num(Index))))
  res_id Val(lbl_num(Index)), True
End Sub

Public Sub cmd_rta_Click()
  If vs1.Enabled = True Then: vs1.SetFocus
  If cmd_rta.Caption = "Stop testing" Then
    cmd_rta.Caption = "Re-test ALL"
    stoptest = True
    cmd_rta.Enabled = False
    'cmd_rta.Enabled = True
    Exit Sub
  End If
  cmd_rta.Caption = "Stop testing"
  Dim i%
  For i = 1 To stdinnum
    res(i).state = False
    res_id i, False
  Next i
  For i = 0 To 6
    cmd_rt(i).Enabled = False
    ylbl_num(i) = ""
    Call print_res(lbl_num(i), lbl_txt(i), lbl_rs(i), cmd_rt(i))
  Next i
  For i = 1 To stdinnum
    Refresh
    If stoptest Then: Exit For
    doexec res(i)
  Next i
  cmd_rta.Caption = "Re-test ALL"
  stoptest = False
  cmd_rta.Enabled = True
  For i = 1 To stdinnum
    res_id i, True
  Next i
End Sub

Private Sub Form_Click()
  bFrmTestFocus = True
End Sub

Private Sub Form_GotFocus()
  bFrmTestFocus = True
End Sub

Private Sub Form_Load()
  stoptest = False
  Timer2.Enabled = True
  bFrmTestFocus = True
  vs1.Max = stdinnum \ 7 + IIf(stdinnum Mod 7 = 0, 0, 1)
  If stdinnum <= 7 Then
    vs1.Enabled = False
  Else
    HookMouse (hwnd)
  End If
  Dim i
  ReDim res(1 To stdinnum) As type_res
  For i = 1 To stdinnum
    With res(i)
      .id = i
      .state = False
      .sin = stdin(i)
      .sout = stdout(i)
      .out = ""
      .err = ""
      .runningtime = 0
      .stdreaded = False
      .sincontent = ""
      .soutcontent = ""
    End With
  Next i
  Show
  Refresh
  cmd_rta_Click
End Sub

Private Sub Form_LostFocus()
  bFrmTestFocus = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If bMouseFlag Then: UnHookMouse (hwnd)
  Unload frmView
  Timer2.Enabled = False
  stoptest = True
  cmd_rta.Enabled = False
  frmStart.Show
  Unload Me
End Sub

Private Function ReadOutput(hRead As Long) As String
  Dim fs As Object
  Set fs = CreateObject("Scripting.FileSystemObject")
  Dim strLine$, sstr$, lgSize&, ssstr$, ltst&
  Dim sBuffer(0 To (BUFFER_LENGTH)) As Byte
  If Trim(UCase(outfile)) <> "{STDOUT}" Then
    If fs.FileExists(outfile) Then
     sstr = ReadFromFile(outfile)
    End If
  Else
    Do While ReadFile(hRead, sBuffer(0), BUFFER_LENGTH, lgSize, ByVal 0&)
      ssstr = StrConv(sBuffer(), vbUnicode)
      'trim() could lost SPACES !!!!!!!!!!!!!
      Erase sBuffer()
      ltst = InStr(1, ssstr, Chr(0)) - 1
      If ltst > 0 Then: ssstr = Left(ssstr, ltst)
      sstr = sstr & ssstr
      If Len(ssstr) <> BUFFER_LENGTH Then: Exit Do
      DoEvents
    Loop
    sstr = Replace(sstr, vbCr, "")
    sstr = Replace(sstr, vbLf, vbCrLf)
  End If
  ReadOutput = Trim(sstr)
End Function

Private Function StrCompare(s1 As String, s2 As String) As Boolean
  Dim as1() As String, as2() As String
  as1 = Split(s1, vbCrLf)
  as2 = Split(s2, vbCrLf)
  Dim i&, ass1$, ass2$
  i = 0
  Do While i <= UBound(as1) Or i <= UBound(as2)
    If i > UBound(as1) Then
      ass1 = ""
    Else
      ass1 = as1(i)
    End If
    If i > UBound(as2) Then
      ass2 = ""
    Else
      ass2 = as2(i)
    End If
    If Trim(ass1) <> Trim(ass2) Then
      StrCompare = False
      Exit Function
    End If
    i = i + 1
  Loop
  StrCompare = True
End Function

Private Sub doexec(res As type_res)
  res_id res.id, False
  Dim fs As Object
  Set fs = CreateObject("Scripting.FileSystemObject")
  'On Error Resume Next
  If fs.FileExists(infile) Then: Kill infile
  If fs.FileExists(outfile) Then: Kill outfile
  'On Error GoTo 0
  With res
      .state = False
      .out = ""
      .err = ""
      .runningtime = 0
      .judging = -1
    End With

  If Not res.stdreaded Then
    If Not fs.FileExists(res.sin) Or Not fs.FileExists(res.sout) Or Not fs.FileExists(exefile) Then
      MsgBox "Can't find stdin/stdout/exe files!" & vbCrLf & res.sin & vbCrLf & res.sout & vbCrLf & exefile, vbCritical, "Error"
      res_id res.id, True
      GoTo endsub
    End If
    res.sincontent = ReadFromFile(res.sin, , True)
    res.soutcontent = ReadFromFile(res.sout)
    res.stdreaded = True
  End If
  
  Dim pstdin&, pstdinread&, pstdinwrite&
  Dim si As STARTUPINFO
  Dim sa As SECURITY_ATTRIBUTES
  With sa
    .nLength = Len(sa)
    .bInheritHandle = 1&
    .lpSecurityDescriptor = 0&
  End With
  If Trim(UCase(infile)) <> "{STDIN}" Then
    FileCopy res.sin, infile
    If Not fs.FileExists(infile) Then
      MsgBox "Error when copying the input file!", vbCritical, "Error"
      res_id res.id, True
      GoTo endsub
    End If
  Else
    'pstdin = CreateFile(res.sin, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
    'If pstdin = INVALID_HANDLE_VALUE Or pstdin = 0 Then
    '  MsgBox "Error: FAILED to OPEN STD-IN ! (CreateFile with " & Str(GetLastError()) & ")", vbCritical, "FAULT"
    '  res_id res.id, True
    '  GoTo endsub
    'End If
    'Dim nSize&: nSize = GetFileSize(pstdin, 0)
    'SetFilePointer pstdin, 0, 0, FILE_BEGIN
    Dim bBytes() As Byte, nSize&
    bBytes = StrConv(res.sincontent, vbFromUnicode)
    nSize = UBound(bBytes) + 1
    
    Dim dwRet&, dwRet1&
    dwRet = CreatePipe(pstdinread, pstdinwrite, sa, (nSize \ BUFFER_LENGTH + 1) * BUFFER_LENGTH)
    If dwRet = 0 Then
      MsgBox "Error: Failed to create a pipe! (CreatePipe() with " & str(GetLastError()) & ")", vbCritical, "FAULT"
      res_id res.id, True
      GoTo endsub
    End If
    'Do
      'ReadFile pstdin, bBytes(0), BUFFER_LENGTH, dwRet, ByVal 0&
      'WriteFile pstdinwrite, bBytes(0), BUFFER_LENGTH, dwRet1, ByVal 0&
      'DoEvents
    'Loop Until (dwRet <> BUFFER_LENGTH)
    'CloseHandle pstdin
    WriteFile pstdinwrite, bBytes(0), nSize, dwRet1, ByVal 0&
    CloseHandle pstdinwrite
    With si
      .dwFlags = .dwFlags Or STARTF_USESTDHANDLES
      .hStdInput = pstdinread
    End With
  End If
  'Dim pid As Long: pid = Shell(exefile, vbHide)
  'Dim phandle As Long: phandle = OpenProcess(&H1F0FFF, True, pid)
  Dim pi As PROCESS_INFORMATION
  Dim phandle&, retval&
  Dim hRead&, hWrite&
  
  retval = CreatePipe(hRead, hWrite, sa, (FileLen(res.sout) * 2 \ BUFFER_LENGTH + 1) * BUFFER_LENGTH)
  If retval = 0 Then
    MsgBox "Error: Error: Failed to create a pipe! (CreatePipe() with " & str(GetLastError()) & ")", vbCritical, "FAULT"
    res_id res.id, True
    GoTo endsub
  End If
  
  With si
    .cb = Len(si)
    .dwFlags = .dwFlags Or STARTF_USESHOWWINDOW
    .wShowWindow = SW_HIDE
    .hStdOutput = hWrite
  End With
  
  res.judging = 1
  Sleep 100
  retval = CreateProcess(vbNullString, exefile & vbNullChar, sa, sa, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, App.Path & vbNullString, si, pi)
  phandle = pi.hProcess
  If retval = 0 Or phandle = 0 Or phandle = INVALID_HANDLE_VALUE Then
    MsgBox "Error: Failed to create a process! (CreateProcess() with " & str(GetLastError()) & ")", vbCritical, "FAULT"
    res_id res.id, True
    GoTo endsub
  End If
  
  Dim Savetime As Double: Savetime = timeGetTime
  Dim ec As Long
  Dim strLine As String, sstr As String, str1 As String
  While timeGetTime < Savetime + timeout
    DoEvents
    Sleep 10
    retval = GetExitCodeProcess(phandle, ec)
    If retval = 0 Then
      MsgBox "Error: Failed to call GetExitCodeProcess() ! (GetExitCodeProcess() with " & str(GetLastError()) & ")", vbCritical, "FAULT"
      res_id res.id, True
      GoTo endsub
    End If
    If ec <> &H103 Then
      CloseHandle phandle
      'res.state = True
      res.runningtime = timeGetTime - Savetime
      If hWrite Then: CloseHandle hWrite
      res.judging = 2
      DoEvents
      sstr = ReadOutput(hRead)
      If ec <> 0 Then
        res.rw = 3
        res.out = sstr
        res.err = "The program returned  " & Trim(str(ec))
        res_id res.id, True
        res.state = True
        GoTo endsub
      End If
      If (Trim(UCase(outfile)) <> "{STDOUT}" And Not fs.FileExists(outfile)) Then
        res.rw = 0
        res.err = "No Output"
        res.out = ""
        res_id res.id, True
        res.state = True
        GoTo endsub
      End If
      If sstr = "" Then
        res.rw = 0
        res.err = "Empty Output"
        res_id res.id, True
        res.state = True
        GoTo endsub
      End If
      str1 = res.soutcontent
      res.judging = 3
      DoEvents
      If Not StrCompare(sstr, str1) Then
        res.rw = 0
        res.out = sstr
        res.err = "Wrong Answer"
        res_id res.id, True
        res.state = True
        GoTo endsub
      End If
      res.rw = 1
      res.out = sstr
      res_id res.id, True
      res.state = True
      Exit Sub
    End If
  Wend
  Call TerminateProcess(phandle, 0)
  'Dim ret&
  'ret = WaitForSingleObject(phandle, 1000)
  'If ret = STATUS_TIMEOUT Then
  '  Call TerminateProcess(phandle, 1)
  '  CloseHandle phandle
  '  MsgBox "Error at terminating the process !", vbCritical, "Error"
  'End If
  res.runningtime = timeGetTime - Savetime
  res.rw = 2
  res.err = "Time Limit Excceed"
  If hWrite Then: CloseHandle hWrite
  res.judging = 2
  DoEvents
  res.out = ReadOutput(hRead)
  res_id res.id, True
  res.state = True
  
endsub:
  If frmStart.Check1.Value = 0 Then
    res.sincontent = ""
    res.soutcontent = ""
    res.stdreaded = False
  End If
  res.judging = 0
  CloseHandle phandle
  If pstdin Then: CloseHandle pstdin
  If hRead Then: CloseHandle hRead
  'If hWrite Then: CloseHandle hWrite
  If pstdinread Then: CloseHandle pstdinread
End Sub

Private Sub print_res(lbl As Label, lbl_txt As Label, lbl_rs As Label, cmd_rt As CommandButton)
  Dim p As Integer: p = Val(Trim(lbl.Caption))
  If p = 0 Or p > stdinnum Then
    lbl_txt.Caption = ""
    funstate lbl_rs, -1
    cmd_rt.Enabled = False
    Exit Sub
  End If
  If res(p).state = False Then
    If res(p).judging = -1 Then
      lbl_txt.Caption = "Preparing.."
    ElseIf res(p).judging = 1 Then
      lbl_txt.Caption = "Running ..."
     ElseIf res(p).judging = 2 Then
      lbl_txt.Caption = "Reading ..."
    ElseIf res(p).judging = 3 Then
      lbl_txt.Caption = "Comparing.."
    Else
      lbl_txt.Caption = "Waiting ..."
    End If
    funstate lbl_rs, -1
    'cmd_rt.Enabled = False
    Exit Sub
  End If
  On Error GoTo 0
  lbl_txt.Caption = "..........."
  funstate lbl_rs, res(p).rw
  cmd_rt.Enabled = True
  Refresh
End Sub


Private Sub lbl_rs_Click(Index As Integer)
With frmView
  If res(Val(lbl_num(Index).Caption)).state = False Then: Exit Sub
  Call .SetRes(res(Val(lbl_num(Index))))
  Call .ReSet
End With
End Sub

Private Sub Timer2_Timer()
  Dim i%
  For i = 0 To 6
stf: If i > 6 Then: Exit For
    If Val(lbl_num(i).Caption) > stdinnum Or Val(lbl_num(i).Caption) <= 0 Then
      Call print_res(lbl_num(i), lbl_txt(i), lbl_rs(i), cmd_rt(i))
      i = i + 1
      GoTo stf
    End If
    If resenabled(Val(lbl_num(i).Caption)) = False Then
      cmd_rt(i).Enabled = False
    Else
      cmd_rt(i).Enabled = True
    End If
    If ylbl_num(i) <> lbl_num(i).Caption Then
      If res(Val(lbl_num(i).Caption)).state = True Then ylbl_num(i) = lbl_num(i).Caption
      Call print_res(lbl_num(i), lbl_txt(i), lbl_rs(i), cmd_rt(i))
    End If
  Next i
  DoEvents
End Sub

Private Sub res_id(id As Integer, state As Boolean)
  resenabled(id) = state
  Dim i%
  For i = 0 To 6
    If Val(lbl_num(i)) = id Then
      cmd_rt(i).Enabled = state
      Exit Sub
    End If
  Next i
End Sub

Private Sub vs1_Change()
  Dim p As Integer: p = vs1.Value
  Dim i%
    For i = 0 To 6
      lbl_num(i) = IIf(p * 7 - 7 + i + 1 <= stdinnum, p * 7 - 7 + i + 1, 0)
      ylbl_num(i) = ""
    Next i
End Sub
