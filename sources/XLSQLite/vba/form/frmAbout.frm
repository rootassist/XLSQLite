Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal lngx As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const GWL_STYLE As Long = (-16)           'The offset of a window's style
Private Const WS_CAPTION As Long = &HC00000       'Style to add a titlebar

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub imgClose_Click()
    Unload Me
End Sub

Private Sub imgDonate_Click()
    Unload Me
    
    frmDonation.Show
End Sub

Private Sub imgGMail_Click()
    Dim gmail As String
    
    gmail = "mailto:gatekeeper.excel@gmail.com?subject=XLSQLite: "
    
    On Error Resume Next
    
    ActiveWorkbook.FollowHyperlink Address:=gmail, NewWindow:=True
    
    Unload Me
    
    If Err.Number <> 0 Then
        MsgBox "Unable to open email application.", vbCritical, "XLSQLite [About {contactus}]"
    End If
End Sub

Private Sub imgSQLite_Click()
    Dim sqlite As String
    
    sqlite = "http://www.sqlite.org"
    
    On Error Resume Next
    
    ActiveWorkbook.FollowHyperlink Address:=sqlite, NewWindow:=True
    
    If Err.Number <> 0 Then
        MsgBox "Unable to open web browser.", vbCritical, "XLSQLite [About {SQLite}]"
    End If
End Sub

Private Sub Label1_Click()
    Dim frmLicence As frmLicense1101
    
    Me.Hide
    
    Set frmLicence = New frmLicense1101
    frmLicence.setViewMode
    
    frmLicence.Show
    
    Unload frmLicence
    
    Me.Show
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub Label3_Click()
    Unload Me
End Sub

Private Sub Label4_Click()
    Unload Me
End Sub

Private Sub Label6_Click()
    Dim frmLicence As frmLicense1101
    
    Me.Hide
    
    Set frmLicence = New frmLicense1101
    frmLicence.setViewMode
    
    frmLicence.Show
    
    Unload frmLicence
    
    Me.Show
End Sub

Private Sub Label7_Click()
    Dim xlsqlite As String
    
    xlsqlite = "http://www.gatekeeperforexcel.com/other-freebies.html"
    
    On Error Resume Next
    
    ActiveWorkbook.FollowHyperlink Address:=xlsqlite, NewWindow:=True
    
    If Err.Number <> 0 Then
        MsgBox "Unable to open web browser.", vbCritical, "XLSQLite [About]"
    End If
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub lblLicense_Click()
    Me.Hide
    
    frmLicenseSQLite.Show
    
    Me.Show
End Sub

Private Sub lblLicenseStatus_Click()
    Unload Me
End Sub

Private Sub lblsqlite4xl_Click()
    Dim sqlite4xl As String
    
    sqlite4xl = "http://sqliteforexcel.codeplex.com/"
    
    On Error Resume Next
    
    ActiveWorkbook.FollowHyperlink Address:=sqlite4xl, NewWindow:=True
    
    If Err.Number <> 0 Then
        MsgBox "Unable to open web browser.", vbCritical, "XLSQLite [About {SQLite for Excel}]"
    End If
End Sub

Private Sub UserForm_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim hwndform As Long
    Dim IStyle As Long
    Dim lngx As Long
    
    hwndform = FindWindow("ThunderDframe", Me.Caption)
    
    IStyle = GetWindowLong(hwndform, GWL_STYLE)
    IStyle = IStyle And Not WS_CAPTION
    lngx = SetWindowLong(hwndform, GWL_STYLE, IStyle)
    
    'repaint the window
    SetWindowPos hwndform, 0, 0, 0, 0, 0, &H20 Or &H2 Or &H4 Or &H1
End Sub