Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#Else
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal lngx As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#End If

Private Const GWL_STYLE As Long = (-16)           'The offset of a window's style
Private Const WS_CAPTION As Long = &HC00000       'Style to add a titlebar

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
#If VBA7 Then
    Dim hwndform As LongPtr
    Dim IStyle As LongPtr
    Dim lngx As LongPtr
#Else
    Dim hwndform As Long
    Dim IStyle As Long
    Dim lngx As Long
#End If
    
    '- --------------------------------------- -
    '- remove the window title bar
    '- --------------------------------------- -
    hwndform = FindWindow("ThunderDframe", Me.Caption)
    
    IStyle = GetWindowLong(hwndform, GWL_STYLE)
    IStyle = IStyle And Not WS_CAPTION
    lngx = SetWindowLong(hwndform, GWL_STYLE, IStyle)
    
    'repaint the window
    SetWindowPos hwndform, 0, 0, 0, 0, 0, &H20 Or &H2 Or &H4 Or &H1
End Sub
