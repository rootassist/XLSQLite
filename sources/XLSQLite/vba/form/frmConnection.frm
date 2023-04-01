'XLSQlite
'
'LICENSE: The MIT License (MIT)
'
'Copyright ｩ 2013 Mark Camilleri
'
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and  associated _
 documentation files (the "Software"), to deal in the Software without restriction, including without limitation _
 the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and _
 to permit persons to whom the Software is furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in all copies or substantial portions of _
 the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO _
 THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE _
 AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, _
 TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE _
 SOFTWARE.
 
 
Option Explicit

Private Sub btnBrowse_Create_Click()
    Dim txtFile As String
    
    txtFile = saveFile("SQLite Editor - [Select Database]", ThisWorkbook.path & "\sqlite", "SQLite DB files (*.sqlite), *.sqlite")

    If txtFile <> "False" Then
        txtDB_Create.Value = txtFile
        
        btnConnect.enabled = True
        
        On Error Resume Next
        btnConnect.SetFocus
        On Error GoTo 0
    Else
        txtDB_Create.Value = ""
        
        btnConnect.enabled = False
        
        On Error Resume Next
        btnBrowse_Create.SetFocus
        On Error GoTo 0
    End If

End Sub

Private Function saveFile(winTitle As String, Optional startPath As String = "C:\", Optional filter As String = "Text files (*.txt), *.txt") As String
    Dim inti As Integer
    Dim intLastOccur As Integer
    
    saveFile = Application.GetSaveAsFilename(fileFilter:=filter, _
                        InitialFileName:=startPath, _
                        title:=winTitle)
                        
    Application.StatusBar = False
End Function

Private Sub btnConnect_Click()
    Dim rtrn As Long
    
    If sqlManager Is Nothing Then
        Set sqlManager = New clsSQLiteManager
    Else
        sqlManager.closeDB
    End If
    
    rtrn = sqlManager.openDB(txtDB_Create.Value)
    
    If rtrn = 0 Then
        MsgBox "Connection to database successful.", vbOKOnly + vbInformation, "SQLite - [Database Connection]"
        
        Unload Me
    Else
        MsgBox "Error:" & vbCr & vbCr & sqlManager.getError, vbCritical + vbOKOnly, "SQLite - [Database Connection]"
    End If
End Sub

Private Sub UserForm_Activate()
    If Not sqlManager Is Nothing Then
        txtDB_Create.Value = sqlManager.getCurrentDBPath
        
        If sqlManager.isDBOpen Then
            btnConnect.enabled = True
        Else
            btnConnect.enabled = False
        End If
    Else
        btnConnect.enabled = False
    End If
End Sub

Private Sub UserForm_Initialize()
    If Not sqlManager Is Nothing Then
        txtDB_Create.Value = sqlManager.getCurrentDBPath
        
        If sqlManager.isDBOpen Then
            btnConnect.enabled = True
        Else
            btnConnect.enabled = False
        End If
    Else
        btnConnect.enabled = False
    End If
End Sub