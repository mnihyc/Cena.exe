Attribute VB_Name = "mouse_check"
Option Explicit
Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A
Global lpPrevWndProcA As Long
Public bMouseFlag As Boolean
Public bFrmTestFocus As Boolean
Public Sub HookMouse(ByVal hwnd As Long)
  lpPrevWndProcA = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
  bMouseFlag = True
End Sub

Public Sub UnHookMouse(ByVal hwnd As Long)
  SetWindowLong hwnd, GWL_WNDPROC, lpPrevWndProcA
  bMouseFlag = False
End Sub

Private Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Select Case uMsg
  Case WM_MOUSEWHEEL
    Dim wzDelta, wKeys As Integer
    wzDelta = HIWORD(wParam)
    wKeys = LOWORD(wParam)
    If wzDelta < 0 Then
      If bFrmTestFocus Then: frmTest.vs1.Value = frmTest.vs1.Value + IIf(frmTest.vs1.Value < frmTest.vs1.Max, 1, 0)
    Else
      If bFrmTestFocus Then: frmTest.vs1.Value = frmTest.vs1.Value - IIf(frmTest.vs1.Value > 1, 1, 0)
    End If
  Case Else
    WindowProc = CallWindowProc(lpPrevWndProcA, hw, uMsg, wParam, lParam)
  End Select
End Function

Private Function HIWORD(LongIn As Long) As Integer
  HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function
Private Function LOWORD(LongIn As Long) As Integer
  LOWORD = LongIn And &HFFFF&
End Function

