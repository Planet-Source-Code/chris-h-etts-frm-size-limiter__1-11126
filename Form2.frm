VERSION 5.00
Object = "*\A..\EttsLimiter.vbp"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2175
   LinkTopic       =   "Form2"
   ScaleHeight     =   1995
   ScaleWidth      =   2175
   StartUpPosition =   3  'Windows Default
   Begin EttsLimiter.EL EL1 
      Left            =   600
      Top             =   240
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
End
Attribute VB_Name = "Form2"
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
