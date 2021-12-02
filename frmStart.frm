VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reversi"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   1680
      TabIndex        =   4
      Top             =   1560
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1245
   End
   Begin VB.OptionButton optColor 
      Caption         =   "&Player 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   345
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   1305
   End
   Begin VB.OptionButton optColor 
      Caption         =   "&Player 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Value           =   -1  'True
      Width           =   1425
   End
   Begin VB.Label lbl 
      Caption         =   "Which colour do you want to choose?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   525
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2475
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ChosenColor As Integer
Public OK As Boolean

Private Sub cmdCancel_Click()
   OK = False: Hide
End Sub

Private Sub cmdOK_Click()
   OK = True: Hide
End Sub

Private Sub Form_Load()
Dim Zwart As Long
Dim Wit As Long
Zwart = frmOth.Text7.Text
Wit = frmOth.Text8.Text
      ChosenColor = 0
      optColor(0).ForeColor = Zwart
      optColor(1).ForeColor = Wit
End Sub

Private Sub optColor_Click(Index As Integer)
   ChosenColor = Index
End Sub

