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
Option Base 0

Public Enum dataTypes
    TYPE_TEXT = 0
    TYPE_NUMERIC = 1
    TYPE_INTEGER = 2
    TYPE_REAL = 3
    TYPE_NONE = 4
End Enum

Public Enum conflictTypes
    TYPE_ABORT = 0
    TYPE_FAIL = 1
    TYPE_IGNORE = 2
    TYPE_REPLACE = 3
    TYPE_ROLLBACK = 4
End Enum
    

Private clxIDs As Collection
Private clxType As Collection
Private clxPK As Collection
Private clxNull As Collection
Private clxUnique As Collection
Private clxConstr As Collection
Private clxDefault As Collection
Private clxConflict As Collection
Private clxCheck As Collection

Public Event itemAdded(ByVal itemID As String, ByVal index As Integer)
Public Event itemRemoved()
Public Event beforeChange(Cancel As Boolean)
Public Event afterChange()
Public Event listCleared()

Private Const upperBound = 8    'indicates the number of (items - 1) required; the array used in base 0

Private Sub Class_Initialize()
    Set clxIDs = New Collection
    Set clxType = New Collection
    Set clxPK = New Collection
    Set clxNull = New Collection
    Set clxUnique = New Collection
    Set clxConstr = New Collection
    Set clxDefault = New Collection
    Set clxConflict = New Collection
    Set clxCheck = New Collection
End Sub

Public Function add(itemParams As String) As Boolean
'This function expects a string containing a semicolon separated list for each attribute. _
The items in the string must be as follows: _
 _
attribute name (m) _
attribute type [0-4] (o) _
whether the attribute is part of the primary key [1/0] (m) _
whether it allows NULLs [1/0] (m) _
whether it must be unique [1/0] (m) _
whether it must have a check constraint [1/0] (m) _
the default value for the attribute (o) _
the conflict clause [0-4] (o) _
the check expression (o) _
 _
e.g.    "column_1;2;1;0;0;0;;2" _
        "column_2;3;0;0;0;0;1;3" _
 _
(o) denotes that the itemParams can be left blank

    Dim varItems As Variant
    Dim inti As Integer
    Dim strID As String
    
    varItems = Split(itemParams, ";")
    
    add = False
    If UBound(varItems) = upperBound Then
    'do not process if the incorrect number of parameters is passed in the string
        strID = varItems(0)
        
        On Error Resume Next
        'try to add the id to the ID collection, if an error is generated it indicates that the column is duplicated - which is wrong
        clxIDs.add strID, strID
        
        If Err.Number = 0 Then
        'the attribute name is unique - adding the other details to the collections
            clxType.add varItems(1), strID
            
            If varItems(2) = "1" Or varItems(2) = "0" Then
                clxPK.add IIf(varItems(2) = "1", True, False), strID
            Else
                clxPK.add False, strID
            End If
            
            If varItems(3) = "1" Or varItems(3) = "0" Then
                clxNull.add IIf(varItems(3) = "1", True, False), strID
            Else
                clxNull.add False, strID
            End If
            
            If varItems(4) = "1" Or varItems(4) = "0" Then
                clxUnique.add IIf(varItems(4) = "1", True, False), strID
            Else
                clxUnique.add False, strID
            End If
            
            If varItems(5) = "1" Or varItems(5) = "0" Then
                clxConstr.add IIf(varItems(2) = "1", True, False), strID
            Else
                clxConstr.add False, strID
            End If
            
            clxDefault.add varItems(6), strID
            
            clxConflict.add varItems(7), strID
            
            clxCheck.add varItems(8), strID
            
            RaiseEvent itemAdded(strID, clxIDs.count)
        Else
        'the attribute name is duplicated
            Err.Raise vbObjectError + 15, "Create SQLite table", "Duplicate attribute found:" & vbCr & vbCr & strID
        End If
        
        On Error GoTo 0
    End If
End Function

Public Sub remove(index As Integer)
'Used to remove an item from the collections.  The item is identified by the index - base 1
    Dim strKey As String

    If index > 0 And index <= clxIDs.count Then
        strKey = clxIDs.item(index)    'get the key to the item
        
        clxIDs.remove index
        clxType.remove strKey
        clxPK.remove strKey
        clxNull.remove strKey
        clxUnique.remove strKey
        clxConstr.remove strKey
        clxDefault.remove strKey
        clxConflict.remove strKey
        clxCheck.remove strKey
        
        RaiseEvent itemRemoved
    End If
End Sub

Public Sub removeKey(itemKey As String)
'Used to remove an item from the collections.  The item is identified by its key

    On Error Resume Next
        clxIDs.remove itemKey
        
        If Err.Number = 0 Then
            clxType.remove itemKey
            clxPK.remove itemKey
            clxNull.remove itemKey
            clxUnique.remove itemKey
            clxConstr.remove itemKey
            clxDefault.remove itemKey
            clxConflict.remove itemKey
            clxCheck.remove itemKey
            
            RaiseEvent itemRemoved
        End If
        
    On Error GoTo 0
End Sub

Public Function count() As Integer
'returns the count of the items held

    count = clxIDs.count
End Function

Public Function item(index As Integer) As Variant
'returns the requested item in a variant array - base 0.  The parameter index is base 1
    Dim varItem(0 To upperBound) As Variant
    
    If index >= 1 And index <= clxIDs.count Then
        varItem(0) = clxIDs.item(index)             'column name
        varItem(1) = clxType.item(varItem(0))       'column type
        varItem(2) = clxPK.item(varItem(0))         'PK true or false
        varItem(3) = clxNull.item(varItem(0))       'Not NULL true or false
        varItem(4) = clxUnique.item(varItem(0))     'Unique true or false
        varItem(5) = clxConstr.item(varItem(0))     'constraint true or false
        varItem(6) = clxDefault.item(varItem(0))    'Default value
        varItem(7) = clxConflict.item(varItem(0))   'conflict clause
        varItem(8) = clxCheck.item(varItem(0))      'check expression
        
        item = varItem
    Else
        Err.Raise vbObjectError + 16, "Create SQLite table", "Index out of bounds."
    End If
End Function

Public Function itemKey(item_Key As String) As Variant
'returns the requested item in a variant array - base 0
    Dim varItem(0 To upperBound) As Variant
    
    On Error Resume Next
        varItem(0) = clxIDs.item(item_Key)
        
        If Err.Number = 0 Then
            varItem(1) = clxType.item(item_Key)
            varItem(2) = clxPK.item(item_Key)
            varItem(3) = clxNull.item(item_Key)
            varItem(4) = clxUnique.item(item_Key)
            varItem(5) = clxConstr.item(item_Key)
            varItem(6) = clxDefault.item(item_Key)
            varItem(7) = clxConflict.item(item_Key)
            varItem(8) = clxCheck.item(item_Key)
            
            itemKey = varItem
        End If
    On Error GoTo 0

End Function

Public Function items(startIndex As Integer, blockSize As Integer) As Variant
'Returns a number of items in a 2dimensional array starting from the item found at startIndex
    Dim varItems() As Variant
    Dim intEndIndex As Integer
    Dim inti As Integer
    
    If blockSize > 0 Then
        If startIndex >= 1 And startIndex <= clxIDs.count Then
            intEndIndex = IIf(startIndex + (blockSize - 1) <= clxIDs.count, startIndex + (blockSize - 1), clxIDs.count)
            ReDim varItems(0 To intEndIndex - startIndex) As Variant
            
            For inti = startIndex To intEndIndex
                varItems(inti - startIndex) = item(inti)
            Next inti
            
            items = varItems
        End If
    End If
End Function

Public Sub clearAll()
    Set clxIDs = Nothing
    Set clxIDs = New Collection
    
    Set clxType = Nothing
    Set clxType = New Collection
    
    Set clxPK = Nothing
    Set clxPK = New Collection
    
    Set clxNull = Nothing
    Set clxNull = New Collection
    
    Set clxUnique = Nothing
    Set clxUnique = New Collection
    
    Set clxConstr = Nothing
    Set clxConstr = New Collection
    
    Set clxDefault = Nothing
    Set clxDefault = New Collection
    
    Set clxConflict = Nothing
    Set clxConflict = New Collection
    
    Set clxCheck = Nothing
    Set clxCheck = New Collection
    
    RaiseEvent listCleared
End Sub

Public Sub setType(index As Integer, listIndex As Integer)
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    RaiseEvent beforeChange(blnCancel)
    
    If Not blnCancel Then
        On Error Resume Next
            strTemp = clxIDs.item(index)
            
            If Err.Number = 0 Then
                clxType.remove strTemp    'the only way to change a value is by removing it...this will jumble the index across the collections
                clxType.add listIndex, strTemp
                
                RaiseEvent afterChange
            End If
        On Error GoTo 0
    End If
End Sub

Public Sub setType_Key(itemKey As String, listIndex As String)
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    RaiseEvent beforeChange(blnCancel)
    
    If Not blnCancel Then
        On Error Resume Next
            strTemp = clxIDs.item(itemKey)
            
            If Err.Number = 0 Then
                clxType.remove itemKey    'the only way to change a value is by removing it...this will jumble the index across the collections
                clxType.add listIndex, itemKey
                
                RaiseEvent afterChange
            End If
        On Error GoTo 0
    End If
End Sub

Public Sub setConflict(index As Integer, listIndex As Integer)
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    RaiseEvent beforeChange(blnCancel)
    
    If Not blnCancel Then
        On Error Resume Next
            strTemp = clxIDs.item(index)
            
            If Err.Number = 0 Then
                clxConflict.remove strTemp    'the only way to change a value is by removing it...this will jumble the index across the collections
                clxConflict.add listIndex, strTemp
                
                RaiseEvent afterChange
            End If
        On Error GoTo 0
    End If
End Sub

Public Sub setConflict_Key(itemKey As String, listIndex As String)
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    RaiseEvent beforeChange(blnCancel)
    
    If Not blnCancel Then
        On Error Resume Next
            strTemp = clxIDs.item(itemKey)
            
            If Err.Number = 0 Then
                clxConflict.remove itemKey    'the only way to change a value is by removing it...this will jumble the index across the collections
                clxConflict.add listIndex, itemKey
                
                RaiseEvent afterChange
            End If
        On Error GoTo 0
    End If
End Sub

Public Sub setPK(index As Integer, Optional Value As Boolean = True)
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    RaiseEvent beforeChange(blnCancel)
    
    If Not blnCancel Then
        On Error Resume Next
            strTemp = clxIDs.item(index)
            
            If Err.Number = 0 Then
                clxPK.remove strTemp    'the only way to change a value is by removing it...this will jumble the index across the collections
                clxPK.add Value, strTemp
                
'                If value Then   'disable allow nulls
'                    clxNull.remove strTemp
'                    clxNull.add False, strTemp
'                    clxConstr.remove strTemp
'                    clxConstr.add False, strTemp
'                    clxUnique.remove strTemp
'                    clxUnique.add False, strTemp
'                End If
                
                RaiseEvent afterChange
            End If
        On Error GoTo 0
    End If
End Sub

Public Sub setPK_Key(itemKey As String, Optional Value As Boolean = True)
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    RaiseEvent beforeChange(blnCancel)
    
    If Not blnCancel Then
        On Error Resume Next
            strTemp = clxIDs.item(itemKey)
            
            If Err.Number = 0 Then
                clxPK.remove itemKey    'the only way to change a value is by removing it...this will jumble the index across the collections
                clxPK.add Value, itemKey
                
'                If value Then   'disable allow nulls
'                    clxNull.remove strTemp
'                    clxNull.add False, strTemp
'                    clxConstr.remove strTemp
'                    clxConstr.add False, strTemp
'                    clxUnique.remove strTemp
'                    clxUnique.add False, strTemp
'                End If
                
                RaiseEvent afterChange
            End If
        On Error GoTo 0
    End If
End Sub

Public Sub setNULL(index As Integer, Optional Value As Boolean = True)
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    RaiseEvent beforeChange(blnCancel)
    
    If Not blnCancel Then
        On Error Resume Next
            strTemp = clxIDs.item(index)
            
            If Err.Number = 0 Then
                clxNull.remove strTemp    'the only way to change a value is by removing it...this will jumble the index across the collections
                clxNull.add Value, strTemp
                
'                If value Then   'disable PK
'                    clxPK.remove strTemp
'                    clxPK.add False, strTemp
'                    clxConstr.remove strTemp
'                    clxConstr.add False, strTemp
'                    clxUnique.remove strTemp
'                    clxUnique.add False, strTemp
'                End If
                
                RaiseEvent afterChange
            End If
        On Error GoTo 0
    End If
End Sub

Public Sub setNULL_Key(itemKey As String, Optional Value As Boolean = True)
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    RaiseEvent beforeChange(blnCancel)
    
    If Not blnCancel Then
        On Error Resume Next
            strTemp = clxIDs.item(itemKey)
            
            If Err.Number = 0 Then
                clxNull.remove itemKey    'the only way to change a value is by removing it...this will jumble the index across the collections
                clxNull.add Value, itemKey
                
'                If value Then   'disable PK
'                    clxPK.remove strTemp
'                    clxPK.add False, strTemp
'                    clxConstr.remove strTemp
'                    clxConstr.add False, strTemp
'                    clxUnique.remove strTemp
'                    clxUnique.add False, strTemp
'                End If
                
                RaiseEvent afterChange
            End If
        On Error GoTo 0
    End If
End Sub

Public Sub setUnique(index As Integer, Optional Value As Boolean = True)
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    RaiseEvent beforeChange(blnCancel)
    
    If Not blnCancel Then
        On Error Resume Next
            strTemp = clxIDs.item(index)
            
            If Err.Number = 0 Then
                clxUnique.remove strTemp    'the only way to change a value is by removing it...this will jumble the index across the collections
                clxUnique.add Value, strTemp
                
'                If value Then   'disable PK
'                    clxPK.remove strTemp
'                    clxPK.add False, strTemp
'                    clxConstr.remove strTemp
'                    clxConstr.add False, strTemp
'                    clxNull.remove strTemp
'                    clxNull.add False, strTemp
'                End If
                
                RaiseEvent afterChange
            End If
        On Error GoTo 0
    End If
End Sub

Public Sub setUnique_Key(itemKey As String, Optional Value As Boolean = True)
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    RaiseEvent beforeChange(blnCancel)
    
    If Not blnCancel Then
        On Error Resume Next
            strTemp = clxIDs.item(itemKey)
            
            If Err.Number = 0 Then
                clxUnique.remove itemKey    'the only way to change a value is by removing it...this will jumble the index across the collections
                clxUnique.add Value, itemKey
                
'                If value Then   'disable PK
'                    clxPK.remove strTemp
'                    clxPK.add False, strTemp
'                    clxConstr.remove strTemp
'                    clxConstr.add False, strTemp
'                    clxNull.remove strTemp
'                    clxNull.add False, strTemp
'                End If
                
                RaiseEvent afterChange
            End If
        On Error GoTo 0
    End If
End Sub

Public Sub setConstraint(index As Integer, Optional Value As Boolean = True)
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    RaiseEvent beforeChange(blnCancel)
    
    If Not blnCancel Then
        On Error Resume Next
            strTemp = clxIDs.item(index)
            
            If Err.Number = 0 Then
                clxConstr.remove strTemp    'the only way to change a value is by removing it...this will jumble the index across the collections
                clxConstr.add Value, strTemp
                
'                If value Then   'disable PK
'                    clxPK.remove strTemp
'                    clxPK.add False, strTemp
'                    clxUnique.remove strTemp
'                    clxUnique.add False, strTemp
'                    clxNull.remove strTemp
'                    clxNull.add False, strTemp
'                End If
                
                RaiseEvent afterChange
            End If
        On Error GoTo 0
    End If
End Sub

Public Sub setConstraint_Key(itemKey As String, Optional Value As Boolean = True)
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    RaiseEvent beforeChange(blnCancel)
    
    If Not blnCancel Then
        On Error Resume Next
            strTemp = clxIDs.item(itemKey)
            
            If Err.Number = 0 Then
                clxConstr.remove itemKey    'the only way to change a value is by removing it...this will jumble the index across the collections
                clxConstr.add Value, itemKey
                
'                If value Then   'disable PK
'                    clxPK.remove strTemp
'                    clxPK.add False, strTemp
'                    clxUnique.remove strTemp
'                    clxUnique.add False, strTemp
'                    clxNull.remove strTemp
'                    clxNull.add False, strTemp
'                End If
                
                RaiseEvent afterChange
            End If
        On Error GoTo 0
    End If
End Sub

Public Sub setDefault(index As Integer, Value As String)
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    RaiseEvent beforeChange(blnCancel)
    
    If Not blnCancel Then
        On Error Resume Next
            strTemp = clxIDs.item(index)
            
            If Err.Number = 0 Then
                clxDefault.remove strTemp    'the only way to change a value is by removing it...this will jumble the index across the collections
                clxDefault.add Value, strTemp
                
                RaiseEvent afterChange
            End If
        On Error GoTo 0
    End If
End Sub

Public Sub setDefault_Key(itemKey As String, Value As String)
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    RaiseEvent beforeChange(blnCancel)
    
    If Not blnCancel Then
        On Error Resume Next
            strTemp = clxIDs.item(itemKey)
            
            If Err.Number = 0 Then
                clxDefault.remove itemKey    'the only way to change a value is by removing it...this will jumble the index across the collections
                clxDefault.add Value, itemKey
                
                RaiseEvent afterChange
            End If
        On Error GoTo 0
    End If
End Sub


Public Sub setCheckExpr(index As Integer, Value As String)
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    RaiseEvent beforeChange(blnCancel)
    
    If Not blnCancel Then
        On Error Resume Next
            strTemp = clxIDs.item(index)
            
            If Err.Number = 0 Then
                clxCheck.remove strTemp    'the only way to change a value is by removing it...this will jumble the index across the collections
                clxCheck.add Value, strTemp
                
                RaiseEvent afterChange
            End If
        On Error GoTo 0
    End If
End Sub

Public Sub setCheckExpr_Key(itemKey As String, Value As String)
    Dim strTemp As String
    Dim blnCancel As Boolean
    
    RaiseEvent beforeChange(blnCancel)
    
    If Not blnCancel Then
        On Error Resume Next
            strTemp = clxIDs.item(itemKey)
            
            If Err.Number = 0 Then
                clxCheck.remove itemKey    'the only way to change a value is by removing it...this will jumble the index across the collections
                clxCheck.add Value, itemKey
                
                RaiseEvent afterChange
            End If
        On Error GoTo 0
    End If
End Sub

Public Sub setAllPK(Value As Boolean)
    Dim varItem As Variant

    Set clxPK = Nothing
    Set clxPK = New Collection
    
    For Each varItem In clxIDs
        clxPK.add Value, varItem
    Next varItem
End Sub

Public Sub setAllNull(Value As Boolean)
    Dim varItem As Variant

    Set clxNull = Nothing
    Set clxNull = New Collection
    
    For Each varItem In clxIDs
        clxNull.add Value, varItem
    Next varItem
End Sub

Public Sub setAllUnique(Value As Boolean)
    Dim varItem As Variant

    Set clxUnique = Nothing
    Set clxUnique = New Collection
    
    For Each varItem In clxIDs
        clxUnique.add Value, varItem
    Next varItem
End Sub

Public Sub setAllConstraint(Value As Boolean)
    Dim varItem As Variant

    Set clxConstr = Nothing
    Set clxConstr = New Collection
    
    For Each varItem In clxIDs
        clxConstr.add Value, varItem
    Next varItem
End Sub

Public Function getPKCount() As Long
    Dim item As Variant
    Dim lngi As Long
    
    lngi = 0
    For Each item In clxPK
        If item = True Then
            lngi = lngi + 1
        End If
    Next item
    
    getPKCount = lngi
End Function

Public Function getPKFields() As Collection
    Dim varItem As Variant
    Dim clxPKFields As Collection
    
    Set clxPKFields = New Collection
    
    For Each varItem In clxIDs
        If clxPK(varItem) Then
            clxPKFields.add varItem
        End If
    Next varItem
    
    Set getPKFields = clxPKFields
End Function

Private Sub Class_Terminate()
    Set clxIDs = Nothing
    Set clxType = Nothing
    Set clxPK = Nothing
    Set clxNull = Nothing
    Set clxUnique = Nothing
    Set clxConstr = Nothing
    Set clxDefault = Nothing
    Set clxConflict = Nothing
    Set clxCheck = Nothing
End Sub