
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal lngx As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const GWL_STYLE As Long = (-16)           'The offset of a window's style
Private Const WS_CAPTION As Long = &HC00000       'Style to add a titlebar

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim hwndform As Long
    Dim IStyle As Long
    Dim lngx As Long
    
    '- --------------------------------------- -
    '- remove the window title bar
    '- --------------------------------------- -
    hwndform = FindWindow("ThunderDframe", Me.Caption)
    
    IStyle = GetWindowLong(hwndform, GWL_STYLE)
    IStyle = IStyle And Not WS_CAPTION
    lngx = SetWindowLong(hwndform, GWL_STYLE, IStyle)
    
    'repaint the window
    SetWindowPos hwndform, 0, 0, 0, 0, 0, &H20 Or &H2 Or &H4 Or &H1
    
    On Error Resume Next
    Label5.SetFocus
    On Error GoTo 0
End Sub

Public Sub setViewMode()
    With Me
        .Height = 233
    End With
    
End Sub