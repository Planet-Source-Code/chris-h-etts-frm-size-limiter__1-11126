VERSION 5.00
Object = "*\A..\EttsLimiter.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   ScaleHeight     =   1080
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin EttsLimiter.EL EL1 
      Left            =   3600
      Top             =   120
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2 More Forms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
    Form2.Show
    Form3.Show
End Sub

'   Be sure to set your reference to YOUR SSUBTMR6.DLL

Private Sub Form_Load()
    With EL1
        .FormInQuestion = Me
        .EnableLimiter = True
        .MinHeight = 200
        .MinWidth = 275
        Me.Height = .MinHeight
        Me.Width = .MinWidth
        .About
        .CenterOnLoad = True
    End With
End Sub
