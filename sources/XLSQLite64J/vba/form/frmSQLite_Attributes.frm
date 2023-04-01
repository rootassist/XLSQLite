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

Private clxHeadings As clsAttributes
Private firstRowIndex As Integer
Private database As String
Private table_Name As String
Private database_Path As String
Private closeCode As Integer
Private blnTempTable As Boolean

Private clxReferences As Collection
Private clxFKActions As Collection
Private clxTblConstraintNames As Collection
Private clxTblConstraints As Collection 'the name is not saved with the constraint

Private Enum constraintTypes
    TYPE_NONE = 0
    TYPE_CHECK = 1
    TYPE_FOREIGN_KEY = 2
    TYPE_UNIQUE = 3
End Enum

Private Enum fkActionTypes
    TYPE_NO_ACTION = 0
    TYPE_SET_NULL = 1
    TYPE_SET_DEFAULT = 2
    TYPE_CASCADE = 3
    TYPE_RESTRICT = 4
End Enum

Private Enum fkImm01
    TYPE_NOT_DEFER = 0
    TYPE_DEFER = 1
End Enum

Private Enum fkImm02
    TYPE_NONE = 0
    TYPE_IMMEDIATE = 1
    TYPE_DEFERRED = 2
End Enum

Private Sub addFKImm01(cmbImm As ComboBox)
    With cmbImm
        .AddItem "NOT DEFERRABLE"
        .AddItem "DEFERRABLE"
        .listIndex = 0
    End With
End Sub

Private Sub addFKImm02(cmbImm As ComboBox)
    With cmbImm
        .AddItem "<NONE>"
        .AddItem "INITIALLY IMMEDIATE"
        .AddItem "INITIALLY DEFERRED"
        .listIndex = 0
    End With
End Sub

Private Sub addTypes(cmbDataType As ComboBox)
    With cmbDataType
        .AddItem "TEXT"
        .AddItem "NUMERIC"
        .AddItem "INTEGER"
        .AddItem "REAL"
        .AddItem "NONE"
        .listIndex = 4
    End With
End Sub

Private Function getDataType(index As Byte) As String
    Select Case index
    Case Is = 0
        getDataType = "TEXT"
    Case Is = 1
        getDataType = "NUMERIC"
    Case Is = 2
        getDataType = "INTEGER"
    Case Is = 3
        getDataType = "REAL"
    Case Is = 4
        getDataType = "NONE"
    End Select
End Function

Private Sub addConflicts(cmbConflict As ComboBox)
    With cmbConflict
        .AddItem "ABORT"
        .AddItem "FAIL"
        .AddItem "IGNORE"
        .AddItem "REPLACE"
        .AddItem "ROLLBACK"
        .listIndex = 0
    End With
End Sub

Private Function getConflictType(index As Byte) As String
    Select Case index
    Case Is = 0
        getConflictType = "ABORT"
    Case Is = 1
        getConflictType = "FAIL"
    Case Is = 2
        getConflictType = "IGNORE"
    Case Is = 3
        getConflictType = "REPLACE"
    Case Is = 4
        getConflictType = "ROLLBACK"
    End Select
End Function

Private Sub addTblConstraints(cmbConstraint As ComboBox)
    With cmbConstraint
        .AddItem "<NONE>"
        .AddItem "CHECK"
        .AddItem "FOREIGN KEY"
        .AddItem "UNIQUE"
        .listIndex = 0
    End With
End Sub

Private Sub addFKActions(cmbFKActions As ComboBox)
    With cmbFKActions
        .AddItem "NO ACTION"
        .AddItem "SET NULL"
        .AddItem "SET DEFAULT"
        .AddItem "CASCADE"
        .AddItem "RESTRICT"
        .listIndex = 0
    End With
End Sub

Private Sub btnAddConstraint_Click()
    Dim newConstraint As String
    Dim varItems As Variant
    Dim lngi As Long
    Dim fields01 As String
    Dim fields02 As String
    
    If clxTblConstraintNames Is Nothing Then
        Set clxTblConstraintNames = New Collection
        Set clxTblConstraints = New Collection
    End If
    
    Select Case cmbTblConstraint.listIndex
    Case Is = 0 'none
    
' *********************************************************************************************************

    Case Is = 1 'check
        If Len(Trim(txtTblConstrName.Value)) > 0 And Len(Trim(txtTblCheck.Value)) > 0 Then
        
            If InStr(1, Trim(txtTblConstrName.Value), " ") = 0 Then
                newConstraint = "CONSTRAINT " & Trim(txtTblConstrName.Value) & " CHECK (" & Trim(txtTblCheck.Value) & ")"
                
                On Error Resume Next
                    clxTblConstraintNames.add Trim(txtTblConstrName.Value), Trim(txtTblConstrName.Value)
                    
                    If Err.Number = 0 Then
                        clxTblConstraints.add "tbl_constraint", Replace(newConstraint, " " & Trim(txtTblConstrName.Value) & " ", "")
                        
                        If Err.Number = 0 Then
                            lstTblConstraints.AddItem Trim(newConstraint)
                            
                            resetToDefaults
                        Else
                            clxTblConstraintNames.remove (clxTblConstraintNames.count)
                            MsgBox "制約がすでに設定されています", vbCritical + vbOKOnly, "TITLE"
                        End If
                    Else
                        MsgBox "制約名をすでに使用しています", vbCritical + vbOKOnly, "TITLE"
                    End If
                On Error GoTo 0
            Else
                MsgBox "無効な制約名です", vbCritical + vbOKOnly, "TITLE"
            End If
        End If
        
' *********************************************************************************************************

    Case Is = 2 'foreign key
        If lstFKReferences.ListCount > 0 Then
            If InStr(1, Trim(txtTblConstrName.Value), " ") = 0 And Len(Trim(txtTblConstrName.Value)) > 0 Then
                
                newConstraint = "CONSTRAINT " & Trim(txtTblConstrName.Value) & " FOREIGN KEY ("
                fields01 = ""
                fields02 = ""
                
                For lngi = 0 To lstFKReferences.ListCount - 1
                    varItems = Split(lstFKReferences.List(lngi), " -> ")
                    fields01 = fields01 & ", " & varItems(0)
                    fields02 = fields02 & ", " & Split(varItems(1), ".")(1)
                Next lngi
                
                fields01 = Right(fields01, Len(fields01) - 2)
                fields02 = Right(fields02, Len(fields02) - 2)
                
                newConstraint = newConstraint & fields01 & ") REFERENCES " & Split(varItems(1), ".")(0) & " (" & _
                                fields02 & ") "
                                
                If lstFKs.ListCount > 0 Then
                    For lngi = 0 To lstFKs.ListCount - 1
                        newConstraint = newConstraint & lstFKs.List(lngi) & " "
                    Next lngi
                End If
                
                If Not (cmbImm01.listIndex = 0 And cmbImm02.listIndex = 0) Then
                    newConstraint = newConstraint & cmbImm01.Value & " " & cmbImm02.Value
                End If
                
                On Error Resume Next
                    clxTblConstraintNames.add Trim(txtTblConstrName.Value), Trim(txtTblConstrName.Value)
                    
                    If Err.Number = 0 Then
                        clxTblConstraints.add "tbl_constraint", Replace(newConstraint, " " & Trim(txtTblConstrName.Value) & " ", "")
                        
                        If Err.Number = 0 Then
                            lstTblConstraints.AddItem Trim(newConstraint)
                            
                            resetToDefaults
                        Else
                            clxTblConstraintNames.remove (clxTblConstraintNames.count)
                            MsgBox "制約がすでに設定されています", vbCritical + vbOKOnly, "TITLE"
                        End If
                    Else
                        MsgBox "制約名をすでに使用しています", vbCritical + vbOKOnly, "TITLE"
                    End If
                On Error GoTo 0
                
            Else
                MsgBox "無効な制約名です", vbCritical + vbOKOnly, "TITLE"
            End If
            
        Else
            MsgBox "属性がまだ選択されていません", vbCritical + vbOKOnly, "TITLE"
        End If
        
' *********************************************************************************************************

    Case Is = 3 'unique
        If lstTblAttributes.listIndex > -1 Then
            If InStr(1, Trim(txtTblConstrName.Value), " ") = 0 And Len(Trim(txtTblConstrName.Value)) > 0 Then
            
                newConstraint = "CONSTRAINT " & Trim(txtTblConstrName.Value) & " UNIQUE ("
                
                For lngi = 0 To lstTblAttributes.ListCount - 1
                    If lstTblAttributes.Selected(lngi) Then
                        fields01 = fields01 & ", " & lstTblAttributes.List(lngi)
                    End If
                Next lngi
                
                fields01 = Right(fields01, Len(fields01) - 2)
                
                newConstraint = newConstraint & fields01 & ") ON CONFLICT " & cmbTblConflict.Value
                
                On Error Resume Next
                    clxTblConstraintNames.add Trim(txtTblConstrName.Value), Trim(txtTblConstrName.Value)
                    
                    If Err.Number = 0 Then
                        clxTblConstraints.add "tbl_constraint", Replace(newConstraint, " " & Trim(txtTblConstrName.Value) & " ", "")
                        
                        If Err.Number = 0 Then
                            lstTblConstraints.AddItem Trim(newConstraint)
                            
                            resetToDefaults
                        Else
                            clxTblConstraintNames.remove (clxTblConstraintNames.count)
                            MsgBox "制約がすでに設定されています", vbCritical + vbOKOnly, "TITLE"
                        End If
                    Else
                        MsgBox "制約名をすでに使用しています", vbCritical + vbOKOnly, "TITLE"
                    End If
                On Error GoTo 0
            Else
                MsgBox "無効な制約名です", vbCritical + vbOKOnly, "TITLE"
            End If
                
        Else
            MsgBox "属性がまだ選択されていません", vbCritical + vbOKOnly, "TITLE"
        End If
            
    End Select
End Sub

Private Sub btnAddFK_Click()
    Dim newAction As String
    
    If clxFKActions Is Nothing Then
        Set clxFKActions = New Collection
    End If
    
    If (optOnUpdate.Value Or optOnDelete.Value) Then
        newAction = IIf(optOnUpdate.Value, "ON UPDATE ", "ON DELETE ") & cmbFKAction.Value
        
        On Error Resume Next
            clxFKActions.add newAction, newAction
        
            If Err.Number = 0 Then
                lstFKs.AddItem newAction
            Else
                MsgBox "アクションはすでに設定されています", vbCritical + vbOKOnly, "TITLE"
            End If
        On Error GoTo 0
    Else
        MsgBox "無効なアクションが定義されています", vbOKOnly + vbCritical, "TITLE"
    End If
End Sub

Private Sub btnAddReference_Click()
    Dim newRef As String
    
    If clxReferences Is Nothing Then
        Set clxReferences = New Collection
    End If
    
    If lstTblAttributes.listIndex > -1 And lstRefAttributes.listIndex > -1 Then
        newRef = lstTblAttributes.List(lstTblAttributes.listIndex) & " -> " & _
                 cmbRefTable.Value & "." & lstRefAttributes.List(lstRefAttributes.listIndex)
                 
        On Error Resume Next
            clxReferences.add newRef, newRef
        
            If Err.Number = 0 Then
                lstFKReferences.AddItem newRef
            Else
                MsgBox "リファレンスはすでに設定されています", vbCritical + vbOKOnly, "TITLE"
            End If
        On Error GoTo 0
                                
        cmbRefTable.enabled = False
    Else
        MsgBox "属性への参照が無効です。", vbOKOnly + vbCritical, "TITLE"
    End If
End Sub

Private Sub btnCancel_Click()
    closeCode = 0
    
    Me.Hide
End Sub

Private Sub btnExecute_Click()
    closeCode = 5   'continue with execution of sql onto database
    
    Me.Hide
End Sub

Private Sub btnRemoveConstraint_Click()
    Dim inti As Integer
    
    inti = 0
    While inti < lstTblConstraints.ListCount
        If lstTblConstraints.Selected(inti) Then
            lstTblConstraints.RemoveItem inti
            
            If Not clxTblConstraintNames Is Nothing Then clxTblConstraintNames.remove inti + 1
            If Not clxTblConstraints Is Nothing Then clxTblConstraints.remove inti + 1
            
'            If Not clxFKActions Is Nothing Then clxFKActions.remove inti + 1
'            If Not clxReferences Is Nothing Then clxReferences.remove inti + 1
        Else
            inti = inti + 1
        End If
    Wend
End Sub

Private Sub btnRmvReference_Click()
    Dim inti As Integer
    
    inti = 0
    While inti < lstFKReferences.ListCount
        If lstFKReferences.Selected(inti) Then
            lstFKReferences.RemoveItem inti
            clxReferences.remove inti + 1
        Else
            inti = inti + 1
        End If
    Wend
    
    If inti = 0 Then cmbRefTable.enabled = True
End Sub

Private Sub btnTestSQL_Click()
    Dim st As clsSQLiteManager
    Dim strErr As String
    Dim sql As String
    
    Set st = New clsSQLiteManager
    
    If st.openDB(database_Path) = 0 Then
        st.executeNonQuery "Begin transaction"
    
        sql = txtPreview.Value
        
        If st.executeNonQuery(sql) <> 0 Then
            strErr = "ERROR:" & vbCr & vbCr & st.getError
            st.executeNonQuery "Rollback transaction"
            MsgBox strErr, vbCritical + vbOKOnly, "TITLE"
        Else
            st.executeNonQuery "Rollback transaction"
            MsgBox "SQLは正常にテストされました", vbOKOnly + vbInformation, "Test SQL"
        End If
        
        st.closeDB
    End If
End Sub

Private Sub cmbConflict01_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setConflict firstRowIndex + 0, cmbConflict01.listIndex
    End If
End Sub

Private Sub cmbConflict02_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setConflict firstRowIndex + 1, cmbConflict02.listIndex
    End If
End Sub

Private Sub cmbConflict03_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setConflict firstRowIndex + 2, cmbConflict03.listIndex
    End If
End Sub

Private Sub cmbConflict04_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setConflict firstRowIndex + 3, cmbConflict04.listIndex
    End If
End Sub

Private Sub cmbConflict05_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setConflict firstRowIndex + 4, cmbConflict05.listIndex
    End If
End Sub

Private Sub cmbConflict06_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setConflict firstRowIndex + 5, cmbConflict06.listIndex
    End If
End Sub

Private Sub cmbConflict07_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setConflict firstRowIndex + 6, cmbConflict07.listIndex
    End If
End Sub

Private Sub cmbConflict08_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setConflict firstRowIndex + 7, cmbConflict08.listIndex
    End If
End Sub

Private Sub cmbConflict09_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setConflict firstRowIndex + 8, cmbConflict09.listIndex
    End If
End Sub

Private Sub cmbConflict10_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setConflict firstRowIndex + 9, cmbConflict10.listIndex
    End If
End Sub

Private Sub cmbDataType01_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setType firstRowIndex + 0, cmbDataType01.listIndex
    End If
End Sub

Private Sub cmbDataType02_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setType firstRowIndex + 1, cmbDataType02.listIndex
    End If
End Sub

Private Sub cmbDataType03_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setType firstRowIndex + 2, cmbDataType03.listIndex
    End If
End Sub

Private Sub cmbDataType04_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setType firstRowIndex + 3, cmbDataType04.listIndex
    End If
End Sub

Private Sub cmbDataType05_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setType firstRowIndex + 4, cmbDataType05.listIndex
    End If
End Sub

Private Sub cmbDataType06_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setType firstRowIndex + 5, cmbDataType06.listIndex
    End If
End Sub

Private Sub cmbDataType07_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setType firstRowIndex + 6, cmbDataType07.listIndex
    End If
End Sub

Private Sub cmbDataType08_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setType firstRowIndex + 7, cmbDataType08.listIndex
    End If
End Sub

Private Sub cmbDataType09_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setType firstRowIndex + 8, cmbDataType09.listIndex
    End If
End Sub

Private Sub cmbDataType10_Change()
    If Not clxHeadings Is Nothing Then
        clxHeadings.setType firstRowIndex + 9, cmbDataType10.listIndex
    End If
End Sub

Private Sub cmbRefTable_Change()
    Dim attributes() As String
    Dim lngi As Long
    
    If cmbRefTable.listIndex > -1 Then
        For lngi = 0 To lstRefAttributes.ListCount - 1
            lstRefAttributes.RemoveItem (0)
        Next lngi
        
        attributes() = getTableAttributes(cmbRefTable.List(cmbRefTable.listIndex))
        
        For lngi = 0 To UBound(attributes)
            lstRefAttributes.AddItem attributes(lngi)
        Next lngi
    End If
End Sub

Private Sub cmbRemoveFK_Click()
    Dim inti As Integer
    
    inti = 0
    While inti < lstFKs.ListCount
        If lstFKs.Selected(inti) Then
            lstFKs.RemoveItem inti
            clxFKActions.remove inti + 1
        Else
            inti = inti + 1
        End If
    Wend
End Sub

Private Sub cmbTblConstraint_Change()
    Select Case cmbTblConstraint.listIndex
    Case Is = 0 'none
        lblName.visible = False
        txtTblConstrName.visible = False
        lblTitle.visible = False
        lstTblAttributes.visible = False
        lblMulti.visible = False
        lstTblAttributes.MultiSelect = fmMultiSelectExtended
        txtTblCheck.visible = False
        cmbTblConflict.visible = False
        
        'FK related
        frmFK.visible = False
        optOnDelete.visible = False
        optOnUpdate.visible = False
        btnAddFK.visible = False
        lstFKs.visible = False
        cmbRemoveFK.visible = False
        frmImmediate.visible = False
        cmbImm01.visible = False
        cmbImm02.visible = False
        lblReferences.visible = False
        cmbRefTable.visible = False
        lstRefAttributes.visible = False
        lstFKReferences.visible = False
        btnRmvReference.visible = False
        btnAddReference.visible = False
    
    Case Is = 1 'check
        lblName.visible = True
        txtTblConstrName.visible = True
        lblReferences.Caption = "References"
        lblName.Caption = "Constraint name"
        lblTitle.Caption = "Expression"
        lblTitle.visible = True
        lstTblAttributes.visible = False
        lblMulti.visible = False
        lstTblAttributes.MultiSelect = fmMultiSelectExtended
        txtTblCheck.visible = True
        cmbTblConflict.visible = False
        
        'FK related
        frmFK.visible = False
        optOnDelete.visible = False
        optOnUpdate.visible = False
        btnAddFK.visible = False
        lstFKs.visible = False
        cmbRemoveFK.visible = False
        frmImmediate.visible = False
        cmbImm01.visible = False
        cmbImm02.visible = False
        lblReferences.visible = False
        cmbRefTable.visible = False
        lstRefAttributes.visible = False
        lstFKReferences.visible = False
        btnRmvReference.visible = False
        btnAddReference.visible = False
        
    Case Is = 2 'foreign key
        lblName.visible = True
        txtTblConstrName.visible = True
        lblReferences.Caption = "Parent attribute/column"
        lblName.Caption = "Constraint name"
        lblTitle.Caption = "Child Attribute/Column"
        lblTitle.visible = True
        lstTblAttributes.visible = True
        lblMulti.visible = False
        lstTblAttributes.MultiSelect = fmMultiSelectSingle
        txtTblCheck.visible = False
        cmbTblConflict.visible = False
    
        'FK related
        frmFK.visible = True
        optOnDelete.visible = True
        optOnUpdate.visible = True
        btnAddFK.visible = True
        lstFKs.visible = True
        cmbRemoveFK.visible = True
        frmImmediate.visible = True
        cmbImm01.visible = True
        cmbImm02.visible = True
        lblReferences.visible = True
        cmbRefTable.visible = True
        lstRefAttributes.visible = True
        lstFKReferences.visible = True
        btnRmvReference.visible = True
        btnAddReference.visible = True
    
    Case Is = 3 'unique
        lblName.visible = True
        txtTblConstrName.visible = True
        lblReferences.Caption = "Conflict clause"
        lblTitle.Caption = "Attributes/Columns"
        lblTitle.visible = True
        lstTblAttributes.visible = True
        lblMulti.visible = True
        lstTblAttributes.MultiSelect = fmMultiSelectExtended
        txtTblCheck.visible = False
        cmbTblConflict.visible = True

        'FK related
        frmFK.visible = False
        optOnDelete.visible = False
        optOnUpdate.visible = False
        btnAddFK.visible = False
        lstFKs.visible = False
        cmbRemoveFK.visible = False
        frmImmediate.visible = False
        cmbImm01.visible = False
        cmbImm02.visible = False
        lblReferences.visible = True
        cmbRefTable.visible = False
        lstRefAttributes.visible = False
        lstFKReferences.visible = False
        btnRmvReference.visible = False
        btnAddReference.visible = False

    End Select
End Sub

Private Sub lblCheck_Click()
    clxHeadings.setAllConstraint Not clxHeadings.item(1)(5)
    
    refresh
    refreshPK
End Sub

Private Sub lblNull_Click()
    clxHeadings.setAllNull Not clxHeadings.item(1)(3)
    
    refresh
    refreshPK
End Sub

Private Sub lblPK_Click()
    clxHeadings.setAllPK Not clxHeadings.item(1)(2)
    
    refresh
    refreshPK
End Sub

Private Sub lblUnique_Click()
    clxHeadings.setAllUnique Not clxHeadings.item(1)(4)
    
    refresh
    refreshPK
End Sub


Private Sub mpgAttribs_Change()
    If mpgAttribs.SelectedItem.Name = "Page3" Then
        'todo: update sql
        txtPreview.Value = getCreateTableFullSQL
    End If
End Sub

Private Sub optPK01_Click()
    clxHeadings.setPK firstRowIndex + 0, optPK01.Value
    refreshPK
End Sub

Private Sub optPK02_Click()
    clxHeadings.setPK firstRowIndex + 1, optPK02.Value
    refreshPK
End Sub

Private Sub optPK03_Click()
    clxHeadings.setPK firstRowIndex + 2, optPK03.Value
    refreshPK
End Sub

Private Sub optPK04_Click()
    clxHeadings.setPK firstRowIndex + 3, optPK04.Value
    refreshPK
End Sub

Private Sub optPK05_Click()
    clxHeadings.setPK firstRowIndex + 4, optPK05.Value
    refreshPK
End Sub

Private Sub optPK06_Click()
    clxHeadings.setPK firstRowIndex + 5, optPK06.Value
    refreshPK
End Sub

Private Sub optPK07_Click()
    clxHeadings.setPK firstRowIndex + 6, optPK07.Value
    refreshPK
End Sub

Private Sub optPK08_Click()
    clxHeadings.setPK firstRowIndex + 7, optPK08.Value
    refreshPK
End Sub

Private Sub optPK09_Click()
    clxHeadings.setPK firstRowIndex + 8, optPK09.Value
    refreshPK
End Sub

Private Sub optPK10_Click()
    clxHeadings.setPK firstRowIndex + 9, optPK10.Value
    refreshPK
End Sub

Private Sub optNull01_Click()
    clxHeadings.setNULL firstRowIndex + 0, optNull01.Value
End Sub

Private Sub optNull02_Click()
    clxHeadings.setNULL firstRowIndex + 1, optNull02.Value
End Sub

Private Sub optNull03_Click()
    clxHeadings.setNULL firstRowIndex + 2, optNull03.Value
End Sub

Private Sub optNull04_Click()
    clxHeadings.setNULL firstRowIndex + 3, optNull04.Value
End Sub

Private Sub optNull05_Click()
    clxHeadings.setNULL firstRowIndex + 4, optNull05.Value
End Sub

Private Sub optNull06_Click()
    clxHeadings.setNULL firstRowIndex + 5, optNull06.Value
End Sub

Private Sub optNull07_Click()
    clxHeadings.setNULL firstRowIndex + 6, optNull07.Value
End Sub

Private Sub optNull08_Click()
    clxHeadings.setNULL firstRowIndex + 7, optNull08.Value
End Sub

Private Sub optNull09_Click()
    clxHeadings.setNULL firstRowIndex + 8, optNull09.Value
End Sub

Private Sub optNull10_Click()
    clxHeadings.setNULL firstRowIndex + 9, optNull10.Value
End Sub

Private Sub optUnique01_Click()
    clxHeadings.setUnique firstRowIndex + 0, optUnique01.Value
    refresh
End Sub

Private Sub optUnique02_Click()
    clxHeadings.setUnique firstRowIndex + 1, optUnique02.Value
    refresh
End Sub

Private Sub optUnique03_Click()
    clxHeadings.setUnique firstRowIndex + 2, optUnique03.Value
    refresh
End Sub

Private Sub optUnique04_Click()
    clxHeadings.setUnique firstRowIndex + 3, optUnique04.Value
    refresh
End Sub

Private Sub optUnique05_Click()
    clxHeadings.setUnique firstRowIndex + 4, optUnique05.Value
    refresh
End Sub

Private Sub optUnique06_Click()
    clxHeadings.setUnique firstRowIndex + 5, optUnique06.Value
    refreshPK
End Sub

Private Sub optUnique07_Click()
    clxHeadings.setUnique firstRowIndex + 6, optUnique07.Value
    refreshPK
End Sub

Private Sub optUnique08_Click()
    clxHeadings.setUnique firstRowIndex + 7, optUnique08.Value
    refreshPK
End Sub

Private Sub optUnique09_Click()
    clxHeadings.setUnique firstRowIndex + 8, optUnique09.Value
    refresh
End Sub

Private Sub optUnique10_Click()
    clxHeadings.setUnique firstRowIndex + 9, optUnique10.Value
    refresh
End Sub

Private Sub optCheck01_Click()
    clxHeadings.setConstraint firstRowIndex + 0, optCheck01.Value
    refresh
End Sub

Private Sub optCheck02_Click()
    clxHeadings.setConstraint firstRowIndex + 1, optCheck02.Value
    refresh
End Sub

Private Sub optCheck03_Click()
    clxHeadings.setConstraint firstRowIndex + 2, optCheck03.Value
    refresh
End Sub

Private Sub optCheck04_Click()
    clxHeadings.setConstraint firstRowIndex + 3, optCheck04.Value
    refresh
End Sub

Private Sub optCheck05_Click()
    clxHeadings.setConstraint firstRowIndex + 4, optCheck05.Value
    refresh
End Sub

Private Sub optCheck06_Click()
    clxHeadings.setConstraint firstRowIndex + 5, optCheck06.Value
    refresh
End Sub

Private Sub optCheck07_Click()
    clxHeadings.setConstraint firstRowIndex + 6, optCheck07.Value
    refresh
End Sub

Private Sub optCheck08_Click()
    clxHeadings.setConstraint firstRowIndex + 7, optCheck08.Value
    refresh
End Sub

Private Sub optCheck09_Click()
    clxHeadings.setConstraint firstRowIndex + 8, optCheck09.Value
    refresh
End Sub

Private Sub optCheck10_Click()
    clxHeadings.setConstraint firstRowIndex + 9, optCheck10.Value
    refresh
End Sub

Private Sub ScrollBar_Change()
    firstRowIndex = ScrollBar.Value
    
    refresh
    
    On Error Resume Next
    cmbDataType01.SetFocus
    On Error GoTo 0
End Sub

Private Sub txtDefault01_Change()
    clxHeadings.setDefault firstRowIndex + 0, txtDefault01.Value
End Sub

Private Sub txtDefault02_Change()
    clxHeadings.setDefault firstRowIndex + 1, txtDefault02.Value
End Sub

Private Sub txtDefault03_Change()
    clxHeadings.setDefault firstRowIndex + 2, txtDefault03.Value
End Sub

Private Sub txtDefault04_Change()
    clxHeadings.setDefault firstRowIndex + 3, txtDefault04.Value
End Sub

Private Sub txtDefault05_Change()
    clxHeadings.setDefault firstRowIndex + 4, txtDefault05.Value
End Sub

Private Sub txtDefault06_Change()
    clxHeadings.setDefault firstRowIndex + 5, txtDefault06.Value
End Sub

Private Sub txtDefault07_Change()
    clxHeadings.setDefault firstRowIndex + 6, txtDefault07.Value
End Sub

Private Sub txtDefault08_Change()
    clxHeadings.setDefault firstRowIndex + 7, txtDefault08.Value
End Sub

Private Sub txtDefault09_Change()
    clxHeadings.setDefault firstRowIndex + 8, txtDefault09.Value
End Sub

Private Sub txtDefault10_Change()
    clxHeadings.setDefault firstRowIndex + 9, txtDefault10.Value
End Sub

Private Sub txtCheck01_Change()
    clxHeadings.setCheckExpr firstRowIndex + 0, txtCheck01.Value
End Sub

Private Sub txtCheck02_Change()
    clxHeadings.setCheckExpr firstRowIndex + 1, txtCheck02.Value
End Sub

Private Sub txtCheck03_Change()
    clxHeadings.setCheckExpr firstRowIndex + 2, txtCheck03.Value
End Sub

Private Sub txtCheck04_Change()
    clxHeadings.setCheckExpr firstRowIndex + 3, txtCheck04.Value
End Sub

Private Sub txtCheck05_Change()
    clxHeadings.setCheckExpr firstRowIndex + 4, txtCheck05.Value
End Sub

Private Sub txtCheck06_Change()
    clxHeadings.setCheckExpr firstRowIndex + 5, txtCheck06.Value
End Sub

Private Sub txtCheck07_Change()
    clxHeadings.setCheckExpr firstRowIndex + 6, txtCheck07.Value
End Sub

Private Sub txtCheck08_Change()
    clxHeadings.setCheckExpr firstRowIndex + 7, txtCheck08.Value
End Sub

Private Sub txtCheck09_Change()
    clxHeadings.setCheckExpr firstRowIndex + 8, txtCheck09.Value
End Sub

Private Sub txtCheck10_Change()
    clxHeadings.setCheckExpr firstRowIndex + 9, txtCheck10.Value
End Sub

Private Sub UserForm_Initialize()
    addTypes cmbDataType01
    addTypes cmbDataType02
    addTypes cmbDataType03
    addTypes cmbDataType04
    addTypes cmbDataType05
    addTypes cmbDataType06
    addTypes cmbDataType07
    addTypes cmbDataType08
    addTypes cmbDataType09
    addTypes cmbDataType10
    
    addConflicts cmbConflict01
    addConflicts cmbConflict02
    addConflicts cmbConflict03
    addConflicts cmbConflict04
    addConflicts cmbConflict05
    addConflicts cmbConflict06
    addConflicts cmbConflict07
    addConflicts cmbConflict08
    addConflicts cmbConflict09
    addConflicts cmbConflict10
    
    addConflicts cmbTblConflict
    
    addTblConstraints cmbTblConstraint
    
    addConflicts cmbConflictPK
    
    addFKActions cmbFKAction
    
    addFKImm01 cmbImm01
    
    addFKImm02 cmbImm02
    
    Set clxHeadings = New clsAttributes
    firstRowIndex = 1
    
    closeCode = -1
End Sub

Public Sub setAttributes(tableName As String, rngHeadings As Excel.Range, isTemporary As Boolean)
    Dim rngCell As Range
    Dim strClmnName As String
    
    Me.Caption = Me.Caption & tableName & "]"
    table_Name = tableName
    
    For Each rngCell In rngHeadings
        '[Start:20130907:1]
        '[Cmt:Added to fix issue with column names starting with number, e.g. 01Table, will now be enclosed in square brackets]
        strClmnName = Trim(rngCell.Value)
        If Len(strClmnName) > 0 Then
            If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(strClmnName, 1)), vbTextCompare) = 0 Then
                clxHeadings.add "[" & strClmnName & "];" & dataTypes.TYPE_NONE & ";0;0;0;0;;" & conflictTypes.TYPE_ABORT & ";"
            Else
                clxHeadings.add strClmnName & ";" & dataTypes.TYPE_NONE & ";0;0;0;0;;" & conflictTypes.TYPE_ABORT & ";"
            End If
        End If
        '[End:20130907:1
        
        '[Start:20130907:2]
        '[Cmt:Removed to fix issue with column names starting with number, e.g. 01Table]
        'clxHeadings.add rngCell.value & ";" & dataTypes.TYPE_NONE & ";0;0;0;0;;" & conflictTypes.TYPE_ABORT & ";"
        '[End:20130907:2]
        
        lstTblAttributes.AddItem rngCell.Value
    Next rngCell
    
    ScrollBar.visible = IIf(rngHeadings.columns.count > 10, True, False)
    ScrollBar.Max = rngHeadings.columns.count - 9
    
    If (rngHeadings.columns.count - 10) > 9 Then
        ScrollBar.LargeChange = 10
    Else
        ScrollBar.LargeChange = rngHeadings.columns.count - 10
    End If
    
    blnTempTable = isTemporary
    
    refresh
End Sub

Private Sub refreshPK()
    Dim lngi As Long

    lngi = clxHeadings.getPKCount
    If lngi > 0 Then
        frmPK.enabled = True
        
        If lngi = 1 Then
            chkAutoinc.enabled = True
            optASC.enabled = True
            optDESC.enabled = True
            cmbConflictPK.enabled = False
            lblConflict.enabled = False
        Else
            chkAutoinc.enabled = False
            optASC.enabled = False
            optDESC.enabled = False
            cmbConflictPK.enabled = True
            lblConflict.enabled = True
        End If
    Else
        frmPK.enabled = False
        chkAutoinc.enabled = False
        optASC.enabled = False
        optDESC.enabled = False
        cmbConflictPK.enabled = False
        lblConflict.enabled = False
    End If
End Sub

Private Sub refresh()
    Dim var As Variant
    Dim inti As Integer
    
    var = clxHeadings.items(firstRowIndex, IIf(clxHeadings.count > 10, 10, clxHeadings.count))
    
    For inti = 0 To UBound(var)
        Select Case inti
        Case Is = 0
            lblAttrib01.Caption = var(inti)(0)
            cmbDataType01.listIndex = var(inti)(1)
            optPK01.Value = var(inti)(2)
            optNull01.Value = var(inti)(3)
            optUnique01.Value = var(inti)(4)
            If optUnique01.Value Then cmbConflict01.enabled = True Else cmbConflict01.enabled = False
            optCheck01.Value = var(inti)(5)
            If optCheck01.Value Then txtCheck01.enabled = True Else txtCheck01.enabled = False
            txtDefault01.Value = var(inti)(6)
            cmbConflict01.listIndex = var(inti)(7)
        Case Is = 1
            lblAttrib02.Caption = var(inti)(0)
            cmbDataType02.listIndex = var(inti)(1)
            optPK02.Value = var(inti)(2)
            optNull02.Value = var(inti)(3)
            optUnique02.Value = var(inti)(4)
            If optUnique02.Value Then cmbConflict02.enabled = True Else cmbConflict02.enabled = False
            optCheck02.Value = var(inti)(5)
            If optCheck02.Value Then txtCheck02.enabled = True Else txtCheck02.enabled = False
            txtDefault02.Value = var(inti)(6)
            cmbConflict02.listIndex = var(inti)(7)
        Case Is = 2
            lblAttrib03.Caption = var(inti)(0)
            cmbDataType03.listIndex = var(inti)(1)
            optPK03.Value = var(inti)(2)
            optNull03.Value = var(inti)(3)
            optUnique03.Value = var(inti)(4)
            If optUnique03.Value Then cmbConflict03.enabled = True Else cmbConflict03.enabled = False
            optCheck03.Value = var(inti)(5)
            If optCheck03.Value Then txtCheck03.enabled = True Else txtCheck03.enabled = False
            txtDefault03.Value = var(inti)(6)
            cmbConflict03.listIndex = var(inti)(7)
        Case Is = 3
            lblAttrib04.Caption = var(inti)(0)
            cmbDataType04.listIndex = var(inti)(1)
            optPK04.Value = var(inti)(2)
            optNull04.Value = var(inti)(3)
            optUnique04.Value = var(inti)(4)
            If optUnique04.Value Then cmbConflict04.enabled = True Else cmbConflict04.enabled = False
            optCheck04.Value = var(inti)(5)
            If optCheck04.Value Then txtCheck04.enabled = True Else txtCheck04.enabled = False
            txtDefault04.Value = var(inti)(6)
            cmbConflict04.listIndex = var(inti)(7)
        Case Is = 4
            lblAttrib05.Caption = var(inti)(0)
            cmbDataType05.listIndex = var(inti)(1)
            optPK05.Value = var(inti)(2)
            optNull05.Value = var(inti)(3)
            optUnique05.Value = var(inti)(4)
            If optUnique05.Value Then cmbConflict05.enabled = True Else cmbConflict05.enabled = False
            optCheck05.Value = var(inti)(5)
            If optCheck05.Value Then txtCheck05.enabled = True Else txtCheck05.enabled = False
            txtDefault05.Value = var(inti)(6)
            cmbConflict05.listIndex = var(inti)(7)
        Case Is = 5
            lblAttrib06.Caption = var(inti)(0)
            cmbDataType06.listIndex = var(inti)(1)
            optPK06.Value = var(inti)(2)
            optNull06.Value = var(inti)(3)
            optUnique06.Value = var(inti)(4)
            If optUnique06.Value Then cmbConflict06.enabled = True Else cmbConflict06.enabled = False
            optCheck06.Value = var(inti)(5)
            If optCheck06.Value Then txtCheck06.enabled = True Else txtCheck06.enabled = False
            txtDefault06.Value = var(inti)(6)
            cmbConflict06.listIndex = var(inti)(7)
        Case Is = 6
            lblAttrib07.Caption = var(inti)(0)
            cmbDataType07.listIndex = var(inti)(1)
            optPK07.Value = var(inti)(2)
            optNull07.Value = var(inti)(3)
            optUnique07.Value = var(inti)(4)
            If optUnique07.Value Then cmbConflict07.enabled = True Else cmbConflict07.enabled = False
            optCheck07.Value = var(inti)(5)
            If optCheck07.Value Then txtCheck07.enabled = True Else txtCheck07.enabled = False
            txtDefault07.Value = var(inti)(6)
            cmbConflict07.listIndex = var(inti)(7)
        Case Is = 7
            lblAttrib08.Caption = var(inti)(0)
            cmbDataType08.listIndex = var(inti)(1)
            optPK08.Value = var(inti)(2)
            optNull08.Value = var(inti)(3)
            optUnique08.Value = var(inti)(4)
            If optUnique08.Value Then cmbConflict08.enabled = True Else cmbConflict08.enabled = False
            optCheck08.Value = var(inti)(5)
            If optCheck08.Value Then txtCheck08.enabled = True Else txtCheck08.enabled = False
            txtDefault08.Value = var(inti)(6)
            cmbConflict08.listIndex = var(inti)(7)
        Case Is = 8
            lblAttrib09.Caption = var(inti)(0)
            cmbDataType09.listIndex = var(inti)(1)
            optPK09.Value = var(inti)(2)
            optNull09.Value = var(inti)(3)
            optUnique09.Value = var(inti)(4)
            If optUnique09.Value Then cmbConflict09.enabled = True Else cmbConflict09.enabled = False
            optCheck09.Value = var(inti)(5)
            If optCheck09.Value Then txtCheck09.enabled = True Else txtCheck09.enabled = False
            txtDefault09.Value = var(inti)(6)
            cmbConflict09.listIndex = var(inti)(7)
        Case Is = 9
            lblAttrib10.Caption = var(inti)(0)
            cmbDataType10.listIndex = var(inti)(1)
            optPK10.Value = var(inti)(2)
            optNull10.Value = var(inti)(3)
            optUnique10.Value = var(inti)(4)
            If optUnique10.Value Then cmbConflict10.enabled = True Else cmbConflict10.enabled = False
            optCheck10.Value = var(inti)(5)
            If optCheck10.Value Then txtCheck10.enabled = True Else txtCheck10.enabled = False
            txtDefault10.Value = var(inti)(6)
            cmbConflict10.listIndex = var(inti)(7)
        End Select
    Next inti
    
    For inti = UBound(var) + 1 To 10
        Select Case inti
            Case Is = 0
                frmAttribs01.visible = False
            Case Is = 1
                frmAttribs02.visible = False
            Case Is = 2
                frmAttribs03.visible = False
            Case Is = 3
                frmAttribs04.visible = False
            Case Is = 4
                frmAttribs05.visible = False
            Case Is = 5
                frmAttribs06.visible = False
            Case Is = 6
                frmAttribs07.visible = False
            Case Is = 7
                frmAttribs08.visible = False
            Case Is = 8
                frmAttribs09.visible = False
            Case Is = 9
                frmAttribs10.visible = False
        End Select
    Next inti
    
End Sub

Private Sub refreshTableList()
    Dim s As clsSQLiteManager
    Dim r As Long
    Dim clxRtrn As Collection
    Dim item As Variant
    Dim inti As Integer

    For inti = 0 To cmbRefTable.ListCount - 1
        cmbRefTable.RemoveItem (0)
    Next inti

    If fileExists(database) Then
        Set s = New clsSQLiteManager
    
        r = s.openDB(database)
    
        If r = 0 Then
            Set clxRtrn = New Collection
    
            s.executeQuery "Select name from sqlite_master where type in (""table"", ""view"") UNION ALL Select ""[T] "" || name from sqlite_temp_master where type in (""table"", ""view"") order by name", clxRtrn
    
            If clxRtrn.count > 0 Then
                inti = 0
                For Each item In clxRtrn
                    If inti > 0 Then
                        cmbRefTable.AddItem item(0)
                    End If
                    inti = inti + 1
                Next item
            End If
        End If
    
        s.closeDB
    
        'house keeping
        Set s = Nothing
        Set clxRtrn = Nothing
    End If
End Sub

Public Sub setDB(dbFullPath As String)
    If fileExists(dbFullPath) Then
        database = dbFullPath
        database_Path = dbFullPath

        refreshTableList
    End If
End Sub

Private Function getTableAttributes(tableName As String) As String()
    Dim s As clsSQLiteManager
    Dim r As Long
    Dim clxRtrn As Collection
    Dim tableDef As String
    Dim Query As String
    Dim firstBracket As Long
    Dim columnDeclaration As String
    Dim columns As Variant
    Dim lngi As Long
    
    If fileExists(database) Then
        Set s = New clsSQLiteManager
        Set clxRtrn = New Collection
        
        Query = "PRAGMA table_info(" & tableName & ")"
        
        r = s.openDB(database)
        
        If r <> 0 Then
            Exit Function
        End If
        
        r = s.executeQuery(Query, clxRtrn)
        
        If r <> 0 Then
            Exit Function
        End If
        
        lngi = 0
        If clxRtrn.count > 0 Then
            ReDim columns(0 To clxRtrn.count - 2) As String
            
            For lngi = 0 To clxRtrn.count - 2
                columns(lngi) = clxRtrn.item(lngi + 2)(1)
            Next lngi
        End If
        
        getTableAttributes = columns
        
        r = s.closeDB
        
        Set s = Nothing
        Set clxRtrn = Nothing
    End If
    
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, closeMode As Integer)
    closeCode = closeMode
    
    If closeMode = 0 Then
        Cancel = True
        Me.Hide
    End If
End Sub

Public Function closeMode() As Integer
    closeMode = closeCode
End Function

Public Function getCreateTableFullSQL() As String
    Dim sql As String
    Dim inti As Integer
    Dim varItem As Variant
    
    getCreateTableFullSQL = ""
    
    If Len(Trim(table_Name)) > 0 Then
        sql = "CREATE " & IIf(blnTempTable, "TEMPORARY ", "") & "TABLE IF NOT EXISTS " & Split(table_Name, " in ")(0) & " (" & vbCr
    Else
        Exit Function
    End If
    
    For inti = 1 To clxHeadings.count
        sql = sql & clxHeadings.item(inti)(0) & _
              IIf(clxHeadings.item(inti)(1) <> 4, " " & getDataType(CByte(clxHeadings.item(inti)(1))), "")
        
        If clxHeadings.item(inti)(2) And clxHeadings.getPKCount = 1 Then 'PK
            sql = sql & " PRIMARY KEY "
            
            If optASC.Value Then sql = sql & "ASC" Else sql = sql & "DESC"
            
            If chkAutoinc.Value Then sql = sql & " AUTOINCREMENT"
        End If
        
        If clxHeadings.item(inti)(3) Then   'NOT NULL
            sql = sql & " NOT NULL " & "ON CONFLICT ABORT"
        End If
        
        If clxHeadings.item(inti)(4) Then   'Unique
            sql = sql & " UNIQUE " & "ON CONFLICT " & getConflictType(CByte(clxHeadings.item(inti)(7)))
        End If
        
        If clxHeadings.item(inti)(5) Then   'Constraint
            sql = sql & " CHECK (" & clxHeadings.item(inti)(8) & ")"
        End If
        
        If clxHeadings.item(inti)(6) <> "" Then
            sql = sql & " DEFAULT (" & clxHeadings.item(inti)(6) & ")"
        End If
              
        sql = sql & IIf(inti < clxHeadings.count, ", " & vbCr, "")
    Next
        
    If lstTblConstraints.ListCount > 0 Or clxHeadings.getPKCount > 1 Then
        sql = sql & ", " & vbCr
    Else
        sql = sql & vbCr
    End If
    
    'check if there is a primary key
    If clxHeadings.getPKCount > 1 Then
        sql = sql & "CONSTRAINT " & Split(table_Name, " in ")(0) & "_pk PRIMARY KEY ("
        
        inti = 1
        For Each varItem In clxHeadings.getPKFields
            sql = sql & varItem & IIf(inti < clxHeadings.getPKFields.count, ", ", "")
            inti = inti + 1
        Next varItem
        
        sql = sql & ")"
    End If
    
    If clxHeadings.getPKCount > 1 And lstTblConstraints.ListCount > 0 Then
        sql = sql & "," & vbCr
    ElseIf clxHeadings.getPKCount > 1 And lstTblConstraints.ListCount = 0 Then
        sql = sql & vbCr
    End If
    
    For inti = 0 To lstTblConstraints.ListCount - 1
        sql = sql & lstTblConstraints.List(inti, 0) & IIf(inti < lstTblConstraints.ListCount - 1, ", ", "") & vbCr
    Next inti
    
    sql = sql & ") "

    getCreateTableFullSQL = sql
End Function


Private Sub resetToDefaults()
    Dim lngi As Long
    Dim itemX As Variant    'used to loop through collection items
    
    txtTblConstrName.Value = ""
    
    For lngi = 1 To lstTblAttributes.ListCount
        lstTblAttributes.Selected(lngi - 1) = False
    Next lngi
    
    lstTblAttributes.listIndex = -1
    txtTblCheck.Value = ""
    lstFKs.Clear
    cmbTblConflict.listIndex = -1
    cmbRefTable.listIndex = -1
    cmbRefTable.enabled = True
    lstRefAttributes.Clear
    lstFKReferences.Clear
    
    If Not clxFKActions Is Nothing Then
        For Each itemX In clxFKActions
            clxFKActions.remove 1
        Next
    End If

    If Not clxReferences Is Nothing Then
        For Each itemX In clxReferences
            clxReferences.remove 1
        Next
    End If
End Sub

Public Function getSQL() As String
    
    getSQL = getCreateTableFullSQL()
End Function