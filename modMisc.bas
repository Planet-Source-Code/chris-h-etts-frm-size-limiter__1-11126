Attribute VB_Name = "modMisc"
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" _
                                              (ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "user32" _
                                           Alias "GetWindowLongA" _
                                           (ByVal hwnd As Long, _
                                           ByVal nIndex As Long) As Long
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17




Public Sub CenterForm(frm As Form)
      On Error Resume Next
            Dim Left As Long, Top As Long
            Left = (Screen.TwipsPerPixelX _
                * (GetSystemMetrics(SM_CXFULLSCREEN) / 2)) - _
                (frm.Width / 2)
            Top = (Screen.TwipsPerPixelY * _
                (GetSystemMetrics(SM_CYFULLSCREEN) / 2)) - _
                (frm.Height / 2)
            frm.Move Left, Top
            
            'Just to let me know that we just got done "Centering the Form"
            Debug.Print """CenterForm"" was Just used"
End Sub


