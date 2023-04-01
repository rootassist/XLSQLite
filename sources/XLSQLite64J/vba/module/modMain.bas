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

Private frmEditor As frmSQLEditor
Private frmDDL As frmSQLite_DDL

Public resultWB As Excel.Workbook
Public sqlEditorOrigValue As String
Public sqlManager As clsSQLiteManager


Public Sub SQLiteEditor(Optional dummyParam As String)
    If frmEditor Is Nothing Then
        Set frmEditor = New frmSQLEditor
    End If
    
    On Error Resume Next
    If Workbooks(frmEditor.cmbWB_SQL.Value) Is Nothing Then
    
        If Err.Number <> 0 Then
            frmEditor.cmbWB_SQL.Value = ActiveWorkbook.Name
            
            frmEditor.cmbWS_SQL.Value = ActiveWorkbook.ActiveSheet.Name
        End If
        
    End If
    
    On Error GoTo 0
    frmEditor.Show 0
End Sub

Public Sub SQLiteDDL(Optional dummyParam As String)
    If frmDDL Is Nothing Then
        Set frmDDL = New frmSQLite_DDL
    End If
        
    frmDDL.Show 1
End Sub

Public Sub XLSQLiteAbout(Optional dummyParam As String)
    frmAbout.Show
End Sub

Public Function fileExists(fullPath As String) As Boolean
    Dim fso As Object
    Dim fsoFile As Object

    fileExists = False
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    Set fsoFile = fso.getFile(fullPath)

    If Err.Number = 0 Then
    'the file exists
        fileExists = True
    End If
    On Error GoTo 0
End Function

Public Function folderExists(path As String) As Boolean
    Dim fso As Object
    Dim fsoFolder As Object
    Dim varItems As Variant
    Dim inti As Integer

    folderExists = False
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Right(path, 1) <> Application.PathSeparator Then
        varItems = Split(path, Application.PathSeparator)

        path = ""
        For inti = 0 To UBound(varItems) - 1
            path = path & varItems(inti) & Application.PathSeparator
        Next inti
    End If

    On Error Resume Next
    Set fsoFolder = fso.getFolder(path)

    If Err.Number = 0 Then
    'the folder exists
        folderExists = True
    End If
    On Error GoTo 0
End Function

Public Sub initialise()
    Set sqlManager = New clsSQLiteManager
End Sub