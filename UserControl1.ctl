VERSION 5.00
Begin VB.UserControl EL 
   BackColor       =   &H00FFC0C0&
   CanGetFocus     =   0   'False
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   735
   InvisibleAtRuntime=   -1  'True
   Picture         =   "UserControl1.ctx":0000
   ScaleHeight     =   735
   ScaleWidth      =   735
   ToolboxBitmap   =   "UserControl1.ctx":1B44
End
Attribute VB_Name = "EL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'Property Variables:
'   This is what I called
'   My control, so be sure
'   to change this to yours
Private objeL As EL

'   this will Initialize the
'   Class
Private cM As New cWinMinMax

'   Var's
Dim m_MinHeight As Integer
Dim m_MinWidth As Integer
Dim m_MaxHeight As Integer
Dim m_MaxWidth As Integer
Dim m_EnableLimiter As Boolean
Dim m_frmCenter As Boolean
Dim m_FormInQuestion As Object

'Default Property Values:
Const m_def_MinHeight = 0
Const m_def_MinWidth = 0
Const m_def_MaxHeight = 0
Const m_def_MaxWidth = 0
Const m_def_EnableLimiter = 0
Const m_def_frmCenter = 0

'Win32 api
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Win32 API Function declarations
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

'Win32 API Constant declarations
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BDR_RAISED = &H5





'==================================================================
'                  00:16      Date: 08/31/00
'NOTE:  Using Api make it look like a button
'==================================================================
Private Sub UserControl_Paint()
    On Error Resume Next
    Dim rct As RECT
    GetClientRect UserControl.hwnd, rct
    DrawEdge UserControl.hdc, rct, BDR_RAISED, BF_RECT
End Sub





'==================================================================
'                  00:17      Date: 08/31/00
'NOTE:  This Kinda makes it look like a button
'==================================================================
Private Sub UserControl_Resize()
    On Error Resume Next
    UserControl.Size 48 * Screen.TwipsPerPixelX, 48 * _
    Screen.TwipsPerPixelY
End Sub





'==================================================================
'                  00:17      Date: 08/31/00
'NOTE:  Load property values from storage
'==================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Set m_FormInQuestion = PropBag.ReadProperty("FormInQuestion", Nothing)
    m_EnableLimiter = PropBag.ReadProperty("EnableLimiter", m_def_EnableLimiter)
    m_frmCenter = PropBag.ReadProperty("CenterOnLoad", m_def_frmCenter)
    m_MinHeight = PropBag.ReadProperty("MinHeight", m_def_MinHeight)
    m_MinWidth = PropBag.ReadProperty("MinWidth", m_def_MinWidth)
    m_MaxHeight = PropBag.ReadProperty("MaxHeight", m_def_MaxHeight)
    m_MaxWidth = PropBag.ReadProperty("MaxWidth", m_def_MaxWidth)
End Sub





'==================================================================
'                  00:18      Date: 08/31/00
'NOTE:  CleanUp  (This is IMPORTANT)
'==================================================================
Private Sub UserControl_Terminate()
    On Error Resume Next
    cM.Detach
End Sub





'==================================================================
'                  00:18      Date: 08/31/00
'NOTE:  Write property values to storage
'==================================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    Call PropBag.WriteProperty("FormInQuestion", m_FormInQuestion, Nothing)
    Call PropBag.WriteProperty("CenterOnLoad", m_frmCenter, m_def_frmCenter)
    Call PropBag.WriteProperty("EnableLimiter", m_EnableLimiter, m_def_EnableLimiter)
    Call PropBag.WriteProperty("MinHeight", m_MinHeight, m_def_MinHeight)
    Call PropBag.WriteProperty("MinWidth", m_MinWidth, m_def_MinWidth)
    Call PropBag.WriteProperty("MaxHeight", m_MaxHeight, m_def_MaxHeight)
    Call PropBag.WriteProperty("MaxWidth", m_MaxWidth, m_def_MaxWidth)
End Sub





'==================================================================
'                  00:18      Date: 08/31/00
'NOTE:  Initialize Properties for User Control
'   ZERO's will prevent Errors
'==================================================================
Private Sub UserControl_InitProperties()
    On Error Resume Next
    m_EnableLimiter = m_def_EnableLimiter
    m_MinHeight = m_def_MinHeight
    m_MinWidth = m_def_MinWidth
    m_MaxHeight = m_def_MaxHeight
    m_MaxWidth = m_def_MaxWidth
End Sub





'==================================================================
'                  00:19      Date: 08/31/00
'NOTE:  Min Height
'==================================================================
Public Property Get MinHeight() As Integer
    On Error Resume Next
    MinHeight = m_MinHeight
    cM.MinTrackHeight = m_MinHeight
End Property

Public Property Let MinHeight(ByVal New_MinHeight As Integer)
    On Error Resume Next
    m_MinHeight = New_MinHeight
    PropertyChanged "MinHeight"
    cM.MinTrackHeight = m_MinHeight
End Property





'==================================================================
'                  00:19      Date: 08/31/00
'NOTE:  Min Width
'==================================================================
Public Property Get MinWidth() As Integer
    On Error Resume Next
    MinWidth = m_MinWidth
End Property

Public Property Let MinWidth(ByVal New_MinWidth As Integer)
    On Error Resume Next
    m_MinWidth = New_MinWidth
    PropertyChanged "MinWidth"
    cM.MinTrackWidth = m_MinWidth
End Property





'==================================================================
'                  00:19      Date: 08/31/00
'NOTE:  Max Height
'==================================================================
Public Property Get MaxHeight() As Integer
    On Error Resume Next
    MaxHeight = m_MaxHeight
End Property

Public Property Let MaxHeight(ByVal New_MaxHeight As Integer)
    On Error Resume Next
    m_MaxHeight = New_MaxHeight
    PropertyChanged "MaxHeight"
    cM.MaxTrackHeight = m_MaxHeight
End Property





'==================================================================
'                  00:19      Date: 08/31/00
'NOTE:  Max Width
'==================================================================
Public Property Get MaxWidth() As Integer
    On Error Resume Next
    MaxWidth = m_MaxWidth
End Property

Public Property Let MaxWidth(ByVal New_MaxWidth As Integer)
    On Error Resume Next
    m_MaxWidth = New_MaxWidth
    PropertyChanged "MaxWidth"
    cM.MaxTrackWidth = m_MaxWidth
End Property





'==================================================================
'                  00:20      Date: 08/31/00
'NOTE:  Ive got this here cause i use it alot
'==================================================================
Public Property Let CenterOnLoad(ByVal New_CenterOnLoad As Boolean)
    On Error Resume Next
    m_frmCenter = New_CenterOnLoad
    PropertyChanged "frmCenter"
    If m_frmCenter = True Then
        UserControl.Extender.Parent.Refresh
        CenterForm m_FormInQuestion
        Else
        DoEvents
    End If

End Property





'==================================================================
'                  00:21      Date: 08/31/00
'NOTE:  Need to get Form Name   ( Initially I used the
'     Usercontrol property but I got alot of STRANGE error)
'==================================================================
Public Property Let FormInQuestion(ByVal New_FormInQuestion As Object)
    On Error Resume Next
    Set m_FormInQuestion = New_FormInQuestion
    PropertyChanged "FormInQuestion"
End Property





'==================================================================
'                  00:21      Date: 08/31/00
'NOTE:  Turn on subclassing
'==================================================================
Public Property Let EnableLimiter(ByVal New_EnableLimiter As Boolean)
    On Error Resume Next
    m_EnableLimiter = New_EnableLimiter
    PropertyChanged "EnableLimiter"
    If m_FormInQuestion Is Nothing Then
        Exit Property
        Else
        cM.Attach m_FormInQuestion.hwnd
        End If
End Property





'==================================================================
'                  00:22      Date: 08/31/00
'NOTE:   Load your own but hey I wouldnt mind
'      a little recognition?!?!
'==================================================================
Public Sub About()
Attribute About.VB_UserMemId = -552
    On Error Resume Next
    frmAbout.Show 1
End Sub
