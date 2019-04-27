VERSION 5.00
Begin VB.Form frmDialog 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Query"
   ClientHeight    =   1755
   ClientLeft      =   8910
   ClientTop       =   7125
   ClientWidth     =   3660
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2340
      TabIndex        =   3
      Top             =   1245
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Okay"
      Height          =   360
      Left            =   240
      TabIndex        =   2
      Top             =   1260
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   360
      MaxLength       =   30
      TabIndex        =   1
      Top             =   690
      Width           =   3000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Input which username to use "
      BeginProperty Font 
         Name            =   "ËÎÌו"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   3345
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public data As String

Private Sub Command1_Click()
  data = Text1.Text
  Unload Me
End Sub

Private Sub Command2_Click()
  data = ""
  Unload Me
End Sub

Private Sub Form_Load()
  data = ""
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    Call Command1_Click
  End If
End Sub

