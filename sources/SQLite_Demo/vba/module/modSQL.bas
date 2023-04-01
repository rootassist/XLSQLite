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

Private Function putResultInArray(resultset As Collection, Optional rangeRows As Long = 0, Optional rangeColumns As Long = 0) As Variant()
'This function puts the values from a result set collection into an array.  It also indicates whether more rows or columns _
 are required to show all the values in the result set.  'rangeRows', and 'rangeColumns' are used to indicate the size of _
 the array formula range.  The 'resultset' collection is a collection of arrays, each item is a tuple from the query.
 
    Dim lngi As Long
    Dim lngj As Long
    Dim rows As Long
    Dim columns As Long
    Dim varElements As Variant
    Dim tempArray() As Variant
    Dim blnNotEnoughRows As Boolean
    Dim blnNotEnoughColumns As Boolean
    Dim lngMinRows As Long
    
    If resultset Is Nothing Then
        Exit Function
    End If
    
    If resultset.count = 0 Then
        Exit Function
    End If
    
    If resultset.count > 0 Then
        rows = resultset.count  'the number of rows returned by the query
        
        varElements = resultset.item(1)
        
        If Not IsArray(varElements) Then
            Exit Function
        End If
        
        columns = UBound(resultset.item(1)) + 1     'the number of columns returned by the query
        
        If rangeColumns > 0 And columns > rangeColumns Then
            blnNotEnoughColumns = True
        End If
        
        'get which has got the minimum number of rows, the range or the result set
        lngMinRows = IIf(rows < rangeRows, rows, rangeRows)
        
        'redim the array to the min number of rows
        ReDim tempArray(1 To lngMinRows) As Variant
        
        If rangeRows > 0 And rows > rangeRows Then
            blnNotEnoughRows = True
        End If
        
        Application.StatusBar = "XLSQLite: Placing values in array..."
        
        'transfer the values from the resultset to the array
        For lngi = 1 To lngMinRows
            tempArray(lngi) = resultset.item(lngi)
            
            If lngi = rangeRows And blnNotEnoughRows Then
                'if there aren't enough rows put an ellipses in the last row
                For lngj = 0 To rangeColumns - 1
                    tempArray(lngi)(lngj) = "...+" & rows - rangeRows & "r"
                Next lngj
                
                Exit For
            ElseIf blnNotEnoughColumns Then
                'if there aren't enough columns put an ellipses in the last column
                tempArray(lngi)(rangeColumns - 1) = "...+" & columns - rangeColumns & "c"
            End If
            
            DoEvents
        Next lngi
        
        Application.StatusBar = False
        
        putResultInArray = tempArray
        
        Application.StatusBar = False
    End If
End Function

Public Function SQLite_Query(dbPath As String, sql As String) As Variant()
'This function will return an array containing the result from a Select query.  The parameters required are the database _
 to which to connect (full path) and the query to process.  This function MUST be called from within an array formula _
 for example 'sqlite_query("c:\database\testdatabase.sqlite", "Select name, surname from client")'. _
 _
 If the array formula is pasted in a range of cells that is smaller than the result set the formula will indicate how many _
 more rows or columns had to be excluded, for example if array formula range is 4 columns smaller than the result set _
 the function will show '...+4c" in the last (range) column.  If it 4 rows short it will place '...+4r' in the last (range) _
 row. _
 _
 If the array formula is larger than the result set, the extra rows/columns will be set to "#N/A".
 
    Dim r As Long
    Dim clxRtrn As Collection
    Dim Query As String
    Dim sqlManager As clsSQLiteManager
    Dim lngColumns As Long
    Dim lngRows As Long
    Dim rngRange As Range
    Dim errMsg(0 To 0) As String
    
    Query = sql
    
    Set sqlManager = New clsSQLiteManager
    
    Select Case TypeName(Application.Caller)
    Case Is = "Range"
        Set rngRange = Range(Application.Caller.Address)
        
        'get the number of rows in the range where the values will be pasted
        lngRows = rngRange.rows.count
        lngColumns = rngRange.columns.count
        
        Set rngRange = Nothing
    End Select
    
    'Check if the database is already open and if not try to open it
    If Not sqlManager.isDBOpen Then
        r = sqlManager.openDB(dbPath, False)
    End If
    
    If r <> 0 Then
        errMsg(0) = "no such database: " & dbPath
        Set clxRtrn = New Collection
        clxRtrn.add errMsg
        
        SQLite_Query = putResultInArray(clxRtrn, lngRows, lngColumns)
        
        Set clxRtrn = Nothing
        Exit Function
    End If
    
    If Right(Query, 1) = ";" Then
        Query = Left(Query, Len(Query) - 1)
    End If
    
    If Trim(CStr(Query)) <> "" Then
        Application.StatusBar = "XLSQLite: Executing query..."
        
        r = sqlManager.executeQuery(Trim(CStr(Query)), clxRtrn)
        
        If r <> 0 Then
            errMsg(0) = sqlManager.getError
            clxRtrn.add errMsg
            
            SQLite_Query = putResultInArray(clxRtrn, lngRows, lngColumns)
            Set clxRtrn = Nothing
            Exit Function
        End If
        
        If clxRtrn.count = 0 Then
            errMsg(0) = "no tuple meets the given criteria"
            clxRtrn.add errMsg
            
            SQLite_Query = putResultInArray(clxRtrn, lngRows, lngColumns)
            Set clxRtrn = Nothing
            Exit Function
        Else
            Application.StatusBar = "XLSQLite: Placing values in worksheet..."
            
            SQLite_Query = putResultInArray(clxRtrn, lngRows, lngColumns)
        End If
    End If
        
    Set clxRtrn = Nothing
    Set sqlManager = Nothing
    
    Application.StatusBar = False
End Function