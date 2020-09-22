VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1150
      Left            =   50
      TabIndex        =   0
      Top             =   -40
      Width           =   5055
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Multi"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0048FC00&
         Height          =   405
         Index           =   0
         Left            =   1025
         TabIndex        =   6
         Top             =   570
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Form Size Limiter"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1890
         TabIndex        =   2
         Top             =   670
         Width           =   1710
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   435
         Left            =   3750
         TabIndex        =   3
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Effective Transcription Solutions"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   1
         Top             =   200
         Width           =   3705
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   780
         Left            =   120
         Picture         =   "frmAbout.frx":0000
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   435
         Left            =   3780
         TabIndex        =   4
         Top             =   630
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Multi"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   405
         Index           =   1
         Left            =   1055
         TabIndex        =   7
         Top             =   600
         Width           =   690
      End
   End
   Begin VB.PictureBox EL1 
      Height          =   480
      Left            =   4320
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   5
      Top             =   3000
      Width           =   1200
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Load()
    Label5.ForeColor = 4783104
    
    Dim str As String
    If App.Minor = 0 Then
        str = "o"
        Else
        str = App.Minor
    End If
    
    Label3.Caption = "v." & App.Major & "." & str
    Label5.Caption = "v." & App.Major & "." & str
    CenterForm Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub Label1_Click()
    Unload Me
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub Label5_Click()
    Unload Me
End Sub
