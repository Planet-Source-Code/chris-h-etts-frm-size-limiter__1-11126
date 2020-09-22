VERSION 5.00
Object = "*\A..\EttsLimiter.vbp"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1725
   LinkTopic       =   "Form3"
   ScaleHeight     =   1680
   ScaleWidth      =   1725
   StartUpPosition =   3  'Windows Default
   Begin EttsLimiter.EL EL1 
      Left            =   480
      Top             =   240
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With EL1
        .FormInQuestion = Me
        .EnableLimiter = True
        .MinHeight = 200
        .MinWidth = 275
        Me.Height = .MinHeight
        Me.Width = .MinWidth
        .CenterOnLoad = True
    End With
End Sub
