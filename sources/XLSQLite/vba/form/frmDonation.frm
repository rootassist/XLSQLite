
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal lngx As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const GWL_STYLE As Long = (-16)           'The offset of a window's style
Private Const WS_CAPTION As Long = &HC00000       'Style to add a titlebar

Private Sub lblCancel_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim hwndform As Long
    Dim IStyle As Long
    Dim lngx As Long
    Dim URL_ As String
    Dim headers As String
    Dim Data As Variant
    Dim bytes() As Byte
    
    '- --------------------------------------- -
    '- remove the window title bar
    '- --------------------------------------- -
    hwndform = FindWindow("ThunderDframe", Me.Caption)
    
    IStyle = GetWindowLong(hwndform, GWL_STYLE)
    IStyle = IStyle And Not WS_CAPTION
    lngx = SetWindowLong(hwndform, GWL_STYLE, IStyle)
    
    'repaint the window
    SetWindowPos hwndform, 0, 0, 0, 0, 0, &H20 Or &H2 Or &H4 Or &H1
    
    '- --------------------------------------- -
    '- Navigate to Paypal Donation page
    '- --------------------------------------- -
    URL_ = "https://www.paypal.com/cgi-bin/webscr"

    headers = "Content-Type: application/x-www-form-urlencoded"

    Data = "cmd=_s-xclick&hosted_button_id=9DRLD3WCF5886"
    
    Data = StrConv(Data, vbFromUnicode) 'convert from unicode
    bytes = Data    'transfer to a byte array

    Application.Cursor = xlWait

    On Error Resume Next
    bwrDonation.Navigate URL_, 64, , bytes, headers
    
    'wait till page is loaded
    While bwrDonation.Busy
        DoEvents
    Wend
    
    Application.Cursor = xlDefault
    
    txtURL.Value = bwrDonation.LocationURL
    
    If Err.Number <> 0 Then
        Unload Me
        
        MsgBox Err.description, vbCritical, "XLSQLite [Paypal donation]"
    End If
    On Error GoTo 0
End Sub