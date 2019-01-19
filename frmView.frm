VERSION 5.00
Begin VB.Form frmView 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View"
   ClientHeight    =   5355
   ClientLeft      =   7620
   ClientTop       =   3795
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text3 
      Height          =   975
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3720
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2400
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label9 
      Caption         =   "1000ms"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Running time: "
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   ".out"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   3480
      Width           =   3735
   End
   Begin VB.Label Label6 
      Caption         =   ".in"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
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
      TabIndex        =   7
      Top             =   4800
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "StdOutput:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Output:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Input:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Width           =   4215
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'Text1.Locked = True
'Text2.Locked = True
'Text3.Locked = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = -1
Hide
frmTest.Show
End Sub

