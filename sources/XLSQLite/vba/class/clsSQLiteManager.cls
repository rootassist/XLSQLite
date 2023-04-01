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

#If Win64 Then
Private dbHandle As LongPtr
#Else
Private dbHandle As Long
#End If

Private dbFullPath As String
Private initReturn As Long      'the return value upon initialisation of sqlite; 0=ok, 1=fail

Public Event beforeCreateDB(Cancel As Boolean)
Public Event afterCreateDB(ByVal fullPath As String, ByVal RetVal As Long)

Public Function version() As String
    version = SQLite3LibVersion()
End Function

Public Function isDBOpen() As Boolean
'use to check whether a database is open or not
    isDBOpen = IIf(dbHandle > 0, True, False)
End Function

Public Function isInitialised() As Boolean
    isInitialised = IIf(initReturn = 0, True, False)
End Function

Public Function openDB(fullPath As String, Optional createDB As Boolean = True) As Long
    Dim fso As Object
    Dim fsoFile As Object
    Dim fsoFolder As Object
    Dim blnCancel As Boolean
    Dim inti As Integer
    Dim folder As String
    
    If initReturn <> SQLITE_INIT_OK Then
    'DLL missing
        openDB = -1
        Exit Function
    End If
    
    If dbHandle <> 0 Then
    'another db is already open
        openDB = -2
        Exit Function
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    Set fsoFile = fso.getFile(fullPath)
    
    If Err.Number = 0 Then
    'the file exists
        RaiseEvent beforeCreateDB(blnCancel)
        
        If Not blnCancel Then
            openDB = SQLite3Open(fullPath, dbHandle)
            dbFullPath = fullPath
            RaiseEvent afterCreateDB(fullPath, openDB)
        End If
    Else
    'the file does not exist, check whether the database needs to be created or the path is wrong
        Set fso = Nothing
        Set fso = CreateObject("Scripting.FileSystemObject")
        Err.Clear
        
        folder = ""
        For inti = Len(fullPath) To 1 Step -1
            If Mid(fullPath, inti, 1) = Application.PathSeparator Then
                folder = Left(fullPath, inti)
                Exit For
            End If
        Next inti
        
        If folder = "" Then
        'fullpath is invalid
            openDB = SQLITE_IOERR
        Else
            Err.Clear   'just in case
            Set fsoFolder = fso.getFolder(folder)
            
            If Err.Number = 0 Then
                If createDB Then
                    'the database needs to be created
                    RaiseEvent beforeCreateDB(blnCancel)
                    If Not blnCancel Then
                        openDB = SQLite3Open(fullPath, dbHandle)
                        dbFullPath = fullPath
                        RaiseEvent afterCreateDB(fullPath, openDB)
                    End If
                Else
                    'the function was called with the createDB flag set to false
                    openDB = SQLITE_IOERR
                End If
            Else
                'the folder does not exist
                openDB = SQLITE_IOERR
            End If
            
            Set fsoFolder = Nothing
        End If
    End If
    
    On Error GoTo 0
    
    Set fso = Nothing
    Set fsoFile = Nothing
End Function

Public Function closeDB() As Long
'Closes the currently open DB; a return value of -1 means that no db was opened yet

    closeDB = -1    'signifies that no DB was open yet
    If dbHandle <> 0 Then
        closeDB = SQLite3Close(dbHandle)
        dbHandle = 0
    End If
End Function

#If Win64 Then
Function ColumnValue(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
#Else
Function ColumnValue(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
#End If
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            ColumnValue = SQLite3ColumnInt32(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_FLOAT:
            ColumnValue = SQLite3ColumnDouble(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_TEXT:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_BLOB:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_NULL:
            ColumnValue = Null
    End Select
End Function

Public Function executeQuery(sql As String, ByRef resultset As Collection) As Long
    #If Win64 Then
    Dim stmtHandle As LongPtr
    #Else
    Dim stmtHandle As Long
    #End If
    
    Dim varRow() As Variant
    Dim colCount As Long
    Dim colType As Long
    Dim lngi As Long
    Dim lngj As Long
    Dim lngRowCount As Long
    
    Set resultset = New Collection    'will contain the result set, each row is placed in an array

    ' Create the sql statement - getting a StmtHandle back
    executeQuery = SQLite3PrepareV2(dbHandle, sql, stmtHandle)
    If executeQuery <> 0 Then Exit Function
    
    'start running the statment
    executeQuery = SQLite3Step(stmtHandle)
    
    If executeQuery = SQLITE_ROW Then
        colCount = SQLite3ColumnCount(stmtHandle)
        ReDim varRow(0 To colCount - 1) As Variant
        
        'get the column names and write to the record set
        For lngi = 0 To colCount - 1
            varRow(lngi) = SQLite3ColumnName(stmtHandle, lngi)
        Next lngi
        
        resultset.add varRow, "ColumnNames"
    End If
    
    lngRowCount = 1
    While executeQuery = SQLITE_ROW
        For lngi = 0 To colCount - 1
            colType = SQLite3ColumnType(stmtHandle, lngi)
            'colTypeName = TypeName(colType)
            varRow(lngi) = IIf(IsNull(ColumnValue(stmtHandle, lngi, colType)), "", ColumnValue(stmtHandle, lngi, colType))
        Next
        
        'add the row to the result set
        resultset.add varRow, CStr(lngRowCount)
        
        executeQuery = SQLite3Step(stmtHandle) 'read the next row
        lngRowCount = lngRowCount + 1
    Wend
    
    If executeQuery <> SQLITE_DONE Then
    'something went wrong
        Set resultset = Nothing
    Else
        executeQuery = SQLite3Finalize(stmtHandle)
    End If
End Function

Public Function executeNonQuery(sql As String, Optional ByRef linesAffected As Long) As Long
'will run any sql against the currently open database, e.g. UPDATE, INSERT, etc.  excluding SELECT as it will not return any result set
    #If Win64 Then
    Dim stmtHandle As LongPtr
    #Else
    Dim stmtHandle As Long
    #End If
    
    ' Create the sql statement - getting a StmtHandle back
    executeNonQuery = SQLite3PrepareV2(dbHandle, sql, stmtHandle)
    If executeNonQuery <> 0 Then Exit Function
    
    ' Start running the statement
    executeNonQuery = SQLite3Step(stmtHandle)
    If executeNonQuery <> SQLITE_DONE Then Exit Function
    
    ' Finalize (delete) the statement
    executeNonQuery = SQLite3Finalize(stmtHandle)
    
    ' Return the number of affected lines in linesAffected
    linesAffected = SQLite3Changes(dbHandle)
End Function

Private Sub Class_Initialize()
    initReturn = SQLite3Initialize ' Default path is ThisWorkbook.Path but can specify other path where the .dlls reside.
End Sub

Private Sub Class_Terminate()
    closeDB
End Sub

Public Function getError() As String
    getError = SQLite3ErrMsg(dbHandle)
End Function

Public Function getLinesChanged() As Long
    getLinesChanged = SQLite3Changes(dbHandle)
End Function

Public Function getCurrentDBPath() As String
    getCurrentDBPath = dbFullPath
End Function
