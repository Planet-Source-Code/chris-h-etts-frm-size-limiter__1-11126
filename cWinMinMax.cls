VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWinMinMax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Type PointAPI
        X As Long
        Y As Long
End Type
Private Type MINMAXINFO
        ptReserved As PointAPI
        ptMaxSize As PointAPI
        ptMaxPosition As PointAPI
        ptMinTrackSize As PointAPI
        ptMaxTrackSize As PointAPI
End Type
Private Const WM_GETMINMAXINFO = &H24
Private CurrentHwnd As Long
Implements ISubclass
Private mvarMaxWidth As Integer
Private mvarMaxHeight As Integer
Private mvarMaxLeft As Integer
Private mvarMaxTop As Integer
Private mvarMinTrackWidth As Integer
Private mvarMinTrackHeight As Integer
Private mvarMaxTrackWidth As Integer
Private mvarMaxTrackHeight As Integer
Public Property Let MaxTrackHeight(ByVal vData As Integer)
Attribute MaxTrackHeight.VB_Description = "Specifies the maximum tracking height."
    mvarMaxTrackHeight = vData
End Property
Public Property Get MaxTrackHeight() As Integer
    MaxTrackHeight = mvarMaxTrackHeight
End Property
Public Property Let MaxTrackWidth(ByVal vData As Integer)
Attribute MaxTrackWidth.VB_Description = "Specifies the maximum tracking width."
    mvarMaxTrackWidth = vData
End Property
Public Property Get MaxTrackWidth() As Integer
    MaxTrackWidth = mvarMaxTrackWidth
End Property
Public Property Let MinTrackHeight(ByVal vData As Integer)
Attribute MinTrackHeight.VB_Description = "Specifies the minimum tracking height."
    mvarMinTrackHeight = vData
End Property
Public Property Get MinTrackHeight() As Integer
    MinTrackHeight = mvarMinTrackHeight
End Property
Public Property Let MinTrackWidth(ByVal vData As Integer)
Attribute MinTrackWidth.VB_Description = "Specifies the minimum tracking width."
    mvarMinTrackWidth = vData
End Property
Public Property Get MinTrackWidth() As Integer
    MinTrackWidth = mvarMinTrackWidth
End Property
Public Property Let MaxTop(ByVal vData As Integer)
Attribute MaxTop.VB_Description = "Specifies the position of the top side of the maximized window."
    mvarMaxTop = vData
End Property
Public Property Get MaxTop() As Integer
    MaxTop = mvarMaxTop
End Property
Public Property Let MaxLeft(ByVal vData As Integer)
Attribute MaxLeft.VB_Description = "Specifies the position of the left side of the maximized window."
    mvarMaxLeft = vData
End Property
Public Property Get MaxLeft() As Integer
    MaxLeft = mvarMaxLeft
End Property
Public Property Let MaxHeight(ByVal vData As Integer)
Attribute MaxHeight.VB_Description = "Specifies the maximized height of the window."
    mvarMaxHeight = vData
End Property
Public Property Get MaxHeight() As Integer
    MaxHeight = mvarMaxHeight
End Property
Public Property Let MaxWidth(ByVal vData As Integer)
Attribute MaxWidth.VB_Description = "Specifies the maximized width of the window."
    mvarMaxWidth = vData
End Property
Public Property Get MaxWidth() As Integer
    MaxWidth = mvarMaxWidth
End Property
Public Sub Attach(ByVal hwnd As Long)
    Detach
    CurrentHwnd = hwnd
    AttachMessage Me, hwnd, WM_GETMINMAXINFO
End Sub
Public Sub Detach()
If Not CurrentHwnd = 0 Then
    DetachMessage Me, CurrentHwnd, WM_GETMINMAXINFO
End If
End Sub
Private Sub Class_Terminate()
Detach
End Sub









Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
'
End Property


Private Property Get ISubclass_MsgResponse() As EMsgResponse
ISubclass_MsgResponse = emrConsume
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim MinMax As MINMAXINFO

CopyMemory MinMax, ByVal lParam, LenB(MinMax) 'Getting value

If mvarMaxLeft > 0 Then MinMax.ptMaxPosition.X = mvarMaxLeft
If mvarMaxTop > 0 Then MinMax.ptMaxPosition.Y = mvarMaxTop
If mvarMaxWidth > 0 Then MinMax.ptMaxSize.X = mvarMaxWidth
If mvarMaxHeight > 0 Then MinMax.ptMaxSize.Y = mvarMaxHeight
If mvarMaxTrackWidth > 0 Then MinMax.ptMaxTrackSize.X = mvarMaxTrackWidth
If mvarMaxTrackHeight > 0 Then MinMax.ptMaxTrackSize.Y = mvarMaxTrackHeight
If mvarMinTrackWidth > 0 Then MinMax.ptMinTrackSize.X = mvarMinTrackWidth
If mvarMinTrackHeight > 0 Then MinMax.ptMinTrackSize.Y = mvarMinTrackHeight

CopyMemory ByVal lParam, MinMax, LenB(MinMax) 'Returning Value

End Function


'Private Property Let ISubclass_MsgResponse(ByVal RHS As EttsSubTmr6.EMsgResponse)
''
'End Property
'
'Private Property Get ISubclass_MsgResponse() As EttsSubTmr6.EMsgResponse
'ISubclass_MsgResponse = emrConsume
'End Property
'
'Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Dim MinMax As MINMAXINFO
'
'CopyMemory MinMax, ByVal lParam, LenB(MinMax) 'Getting value
'
'If mvarMaxLeft > 0 Then MinMax.ptMaxPosition.X = mvarMaxLeft
'If mvarMaxTop > 0 Then MinMax.ptMaxPosition.Y = mvarMaxTop
'If mvarMaxWidth > 0 Then MinMax.ptMaxSize.X = mvarMaxWidth
'If mvarMaxHeight > 0 Then MinMax.ptMaxSize.Y = mvarMaxHeight
'If mvarMaxTrackWidth > 0 Then MinMax.ptMaxTrackSize.X = mvarMaxTrackWidth
'If mvarMaxTrackHeight > 0 Then MinMax.ptMaxTrackSize.Y = mvarMaxTrackHeight
'If mvarMinTrackWidth > 0 Then MinMax.ptMinTrackSize.X = mvarMinTrackWidth
'If mvarMinTrackHeight > 0 Then MinMax.ptMinTrackSize.Y = mvarMinTrackHeight
'
'CopyMemory ByVal lParam, MinMax, LenB(MinMax) 'Returning Value
'
'End Function

