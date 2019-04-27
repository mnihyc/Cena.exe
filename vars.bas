Attribute VB_Name = "vars"
Option Explicit
Public exefile As String, infile As String, outfile As String, timeout As Double
Public stdin(1 To 1050) As String, stdinnum As Integer
Public stdout(1 To 1050) As String, stdoutnum As Integer
Public Type type_res
  id As Integer
  state As Boolean
  judging As Integer
  rw As Integer
  sin As String
  sout As String
  out As String
  err As String
  runningtime As Integer
  stdreaded As Boolean
  sincontent As String
  soutcontent As String
End Type

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMs As Long) As Long
Public Const STATUS_TIMEOUT = &H102

Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Public Type STARTUPINFO
    cb                              As Long
    lpReserved                      As String
    lpDesktop                       As String
    lpTitle                         As String
    dwX                             As Long
    dwY                             As Long
    dwXSize                         As Long
    dwYSize                         As Long
    dwXCountChars                   As Long
    dwYCountChars                   As Long
    dwFillAttribute                 As Long
    dwFlags                         As Long
    wShowWindow                     As Integer
    cbReserved2                     As Integer
    lpReserved2                     As Long
    hStdInput                       As Long
    hStdOutput                      As Long
    hStdError                       As Long
End Type
Public Type PROCESS_INFORMATION
    hProcess                        As Long
    hThread                         As Long
    dwProcessId                     As Long
    dwThreadId                      As Long
End Type
Public Type SECURITY_ATTRIBUTES
    nLength                         As Long
    lpSecurityDescriptor            As Long
    bInheritHandle                  As Long
End Type
Public Const NORMAL_PRIORITY_CLASS  As Long = &H20&
Public Const STARTF_USESTDHANDLES   As Long = &H100&
Public Const STARTF_USESHOWWINDOW   As Long = &H1&
Public Const SW_HIDE                As Long = 0&
Public Const INFINITE               As Long = &HFFFF&
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSize As Currency) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Const FILE_BEGIN = 0


Enum DesiredAccess
    GENERIC_READ = &H80000000
    GENERIC_WRITE = &H40000000
    GENERIC_EXECUTE = &H20000000
    GENERIC_ALL = &H10000000
End Enum
  
Enum ShareMode
    FILE_SHARE_READ = &H1
    FILE_SHARE_WRITE = &H2
    FILE_SHARE_DELETE = &H4
End Enum
  
'This parameter must be one of the following values, which cannot be combined:
Enum CreationDisposition
    TRUNCATE_EXISTING = 5
    OPEN_ALWAYS = 4
    OPEN_EXISTING = 3
    CREATE_ALWAYS = 2
    CREATE_NEW = 1
End Enum
  
Enum FlagsAndAttributes
    'attributes
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_COMPRESSED = &H800
    FILE_ATTRIBUTE_DIRECTORY = &H10
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_NORMAL = &H80 'The file does not have other attributes set. This attribute is valid only if used alone.
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_TEMPORARY = &H100
    'flags
    FILE_FLAG_BACKUP_SEMANTICS = &H2000000
    FILE_FLAG_DELETE_ON_CLOSE = &H4000000
    FILE_FLAG_NO_BUFFERING = &H20000000
    FILE_FLAG_OVERLAPPED = &H40000000
    FILE_FLAG_POSIX_SEMANTICS = &H1000000
    FILE_FLAG_RANDOM_ACCESS = &H10000000
    FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
    FILE_FLAG_WRITE_THROUGH = &H80000000
End Enum
  
Public Const INVALID_HANDLE_VALUE = -1

Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function SendMessageByRef Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Long, lParam As Any) As Long
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_SETSEL = &HB1
Public Const SB_LINEDOWN = 1
Public Const SB_LINEUP = 0
Public Const SB_VERT = 1
Public Const WM_VSCROLL = &H115
Public Declare Function GetScrollPos Lib "User32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Public Declare Function SetScrollPos Lib "User32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long


Public Sub funstate(lbl As Label, state As Integer, Optional full As Boolean = False)
  If state = 0 Then
    lbl.ForeColor = vbRed
    lbl.Caption = IIf(full, "Wrong Answer", "WA")
  ElseIf state = 1 Then
    lbl.ForeColor = vbGreen
    lbl.Caption = IIf(full, "Accept", "AC")
  ElseIf state = 2 Then
    lbl.ForeColor = &HFF00FF
    lbl.Caption = IIf(full, "Time Limit Excceed", "TLE")
  ElseIf state = 3 Then
    lbl.ForeColor = vbRed
    lbl.Caption = IIf(full, "Runtime Error", "RE")
  Else
    lbl.Caption = ""
  End If
End Sub

Public Function ReadFromFile(fn As String, Optional size = 0, Optional org As Boolean = False) As String
  Dim fln&: fln = size
  If size = 0 Then: fln = FileLen(fn)
  Dim hFile&: hFile = FreeFile()
  Open fn For Binary Access Read As hFile
  ReadFromFile = Space(fln)
  DoEvents
  Get hFile, , ReadFromFile
  Close hFile
  If org Then: Exit Function
  ReadFromFile = Replace(ReadFromFile, vbCr, "")
  ReadFromFile = Replace(ReadFromFile, vbLf, vbCrLf)
  ReadFromFile = Trim(ReadFromFile)
End Function

Public Function ShowDialog(str As String, tf As Form, Optional txstr$ = "") As String
  ReSetDialog str, tf.Left + (tf.Width - frmDialog.Width) \ 2, tf.Top + (tf.Height - frmDialog.Height) \ 2, txstr
  frmDialog.Show vbModal
  ShowDialog = frmDialog.data
End Function

Public Sub ReSetDialog(str As String, Optional tL& = 0, Optional tT& = 0, Optional tS$ = "")
  ' The form has been initialized before when accessing frmDialog.Width/Height
  frmDialog.Label1.Caption = str
  frmDialog.Move tL, tT
  frmDialog.Text1.Text = tS
End Sub





