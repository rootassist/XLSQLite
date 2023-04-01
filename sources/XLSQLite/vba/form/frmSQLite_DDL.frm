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

Private Const previewLimit = 100
Private frmAttribs As frmSQLite_Attributes

Private Function getFile(winTitle As String, Optional startPath As String = "C:\", Optional filter As String = "Text files (*.txt), *.txt") As String
    Dim inti As Integer
    Dim intLastOccur As Integer
    
    ChDir startPath
    
    getFile = Application.GetOpenFilename(filter, 1, winTitle, , False)
    
    'getFile = Application.GetSaveAsFilename(fileFilter:=filter, _
                        InitialFileName:=startPath, _
                        title:=winTitle)
                        
    Application.StatusBar = False
End Function

Public Function getCreateTableSQL(tableName As String, headings As Range) As String
    Dim rngCell As Range
    Dim columns As String
    Dim sql As String
    Dim strClmnName As String
    
    getCreateTableSQL = ""
    
    If Len(Trim(tableName)) > 0 Then
        sql = "CREATE " & IIf(chkTempTable.Value, "TEMPORARY ", "") & "TABLE IF NOT EXISTS " & tableName & " ("
    Else
        Exit Function
    End If
    
    If headings.columns.count > 0 And headings.rows.count = 1 Then
        For Each rngCell In headings
            '[Start:20130907:3]
            '[Cmt:Added to fix issue with column names starting with number, e.g. 01Table, will now be enclosed in square brackets]
            strClmnName = Trim(rngCell.Value)
            If Len(strClmnName) > 0 Then
                If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(strClmnName, 1)), vbTextCompare) = 0 Then
                    columns = columns & "[" & strClmnName & "], "
                Else
                    columns = columns & strClmnName & ", "
                End If
            Else
                Exit Function
            End If
            '[End:20130907:3]
            
            '[Start:20130907:4]
            '[Cmt:Removed to fix issue with column names starting with number, e.g. 01Table]
            'If Len(Trim(rngCell.value)) > 0 Then
            '    columns = columns & Trim(rngCell.value) & ", "
            'Else
            '    Exit Function
            'End If
            '[End:20130907:4]
        Next rngCell
        
        columns = Left(columns, Len(columns) - 2)
        
        sql = sql & columns & ")"
    Else
        Exit Function
    End If
    
    getCreateTableSQL = sql
End Function

Private Sub btnBrowse_Create_Click()
    Dim txtFile As String
    
    Do
        txtFile = getFile("SQLite Editor - [Select Database]", ThisWorkbook.path & "\sqlite", "SQLite DB files (*.sqlite), *.sqlite")
    
        If txtFile <> "False" Then
            txtDB_Create.Caption = txtFile
        Else
            txtDB_Create.Caption = ""
            
            MsgBox "You must choose a database.", vbCritical + vbOKOnly, "SQLite Editor - [Select Database]"
        End If
    Loop Until txtFile <> "False"
    
End Sub

Private Sub btnChangeDB_Click()
    
    If vbYes = MsgBox("If you change or reset your connection you will loose any temporary tables created on the current database." & _
                     vbCr & vbCr & "Are you sure you want to continue?", vbYesNo + vbDefaultButton2 + vbQuestion, "SQLite - [Change connection]") Then
        frmConnection.Show
    
        UserForm_Activate
    End If
End Sub

Private Sub btnClear_Click()
    If Len(txtLog.Value) > 0 Then
        If MsgBox("Are you sure you want to clear the log?", vbQuestion + vbYesNo + vbDefaultButton2, "SQLite Editor") = vbYes Then
            txtLog.Value = ""
        End If
    End If
End Sub

Private Function checkInputs(blnValidName As Boolean, blnValidRange As Boolean) As Boolean
    Dim rngTest As Range
    Dim regex As VBScript_RegExp_55.RegExp
    Dim clxRtrn As Collection

    
    blnValidName = True
    blnValidRange = False
    Set regex = New VBScript_RegExp_55.RegExp
    Set clxRtrn = New Collection
    
    'check that the data range value is valid
    On Error Resume Next
    Set rngTest = Range(rfdTableSource)
    If Err.Number = 0 Then blnValidRange = True
    Err.Clear
    On Error GoTo 0
    
    With regex
        .MultiLine = False
        .Global = False
        .IgnoreCase = True
        .Pattern = "^([a-z|A-Z|0-9|,|\.|;|'|=|-|`|~|\!|@|#|\$|%|\^|&|\(|\)|_|\+|\||}|{|""|>|<]+|'[a-z|A-Z|0-9| |,|\.|;|'|=|-|`|~|\!|@|#|\$|%|\^|&|\(|\)|_|\+|\||}|{|""|>|<]+')\!\$?[A-Z]+\$?[0-9]+:\$?[A-Z]+\$?[0-9]+$"
    End With
    
    blnValidRange = regex.test(rfdTableSource.Value)
    
    'check that the table name is valid
    If UCase(Left(txtTableName, 7)) = "SQLITE_" Then blnValidName = False
    If Len(txtTableName) = 0 Then blnValidName = False
    
    'the name can only be alphanumeric with no special characters, except for underscore
    With regex
        .MultiLine = False
        .Global = True
        .IgnoreCase = True
        .Pattern = "\w"
    End With
    
    If Len(regex.Replace(txtTableName.Value, "")) > 0 Then blnValidName = False
    
    checkInputs = blnValidName And blnValidRange
End Function

Private Function extractTwoToOneDimArray(sourceArray As Variant, elmntInFirstDimension As Integer, targetArray As Variant)
    Dim lngi As Long
    Dim lowerBound As Long
    Dim upperBound As Long
    
    lowerBound = LBound(sourceArray, 2)
    upperBound = UBound(sourceArray, 2)
    
    ReDim targetArray(lowerBound To upperBound) As Variant
    
    For lngi = lowerBound To upperBound
        targetArray(lngi) = sourceArray(elmntInFirstDimension, lngi)
    Next lngi
End Function

Private Function grabValues(cellRange As Range, Optional forPreview As Boolean = False) As Collection
    Dim lngi As Long
    Dim lngj As Long
    Dim varElements2D() As Variant
    Dim varElements1D() As Variant
    Dim clxRows As Collection
    Dim topLeftCell As Range
    Dim rg As RegExp
    Dim rows As Long
    Dim answer As Long
    Dim strTemp As String
    Dim blnRemoveAllDblQuotes
    Dim intColumns As Integer
    
    If cellRange Is Nothing Then Exit Function
    
    Set rg = New RegExp
    Set clxRows = New Collection
    Set topLeftCell = cellRange.Resize(1, 1)
    intColumns = cellRange.columns.count
    ReDim varElements2D(0 To cellRange.columns.count - 1) As Variant
    
    If Not forPreview Then
        rows = cellRange.rows.count - 1
    Else
        rows = IIf(previewLimit < cellRange.rows.count - 1, previewLimit, cellRange.rows.count - 1)
    End If
    
    rg.Pattern = "[a-z]"
    rg.IgnoreCase = True
    rg.Global = True
    
    For lngi = 0 To rows
    
        varElements2D = topLeftCell.Offset(lngi, 0).Resize(1, intColumns)
        
        If rows > 0 Then
            lblStatus.Caption = "Reading data row " & lngi + 1 & " [" & Int(((lngi + 1) / (rows + 1)) * 100) & "%]"
        End If
        
        extractTwoToOneDimArray varElements2D, 1, varElements1D
        
        DoEvents
        
'        For lngj = 0 To cellRange.columns.count - 1
'            varElements2D(lngj) = topLeftCell.Offset(lngi, lngj).value
'
'            If varElements2D(lngj) = "" Then
'                varElements2D(lngj) = "NULL"
'            Else
'                If Not IsNumeric(varElements2D(lngj)) Then    ' Or rg.test(varElements2D(lngj)) Then
'                    lblStatus.Caption = "Reading row " & lngi & ", column " & lngj
'
'                    If InStr(1, varElements2D(lngj), """") > 0 Then   'check if the text contains double quotes
'                    'ask the user what to do
'                        If blnRemoveAllDblQuotes Then
'                            varElements2D(lngj) = """" & Replace(varElements2D(lngj), """", "") & """"
'                        Else
'                            strTemp = varElements2D(lngj)
'                            answer = Assistant.DoAlert("SQLite - [Create table]", "Double quotes found in '" & strTemp & "'" & vbCr & vbCr & _
'                                                    "Do you want to remove?", msoAlertButtonYesAllNoCancel, msoAlertIconQuery, msoAlertDefaultFirst, msoAlertCancelDefault, False)
'                            Select Case answer
'                            Case Is = vbYes
'                                varElements2D(lngj) = """" & Replace(varElements2D(lngj), """", "") & """"
'                            Case Is = 8
'                                varElements2D(lngj) = """" & Replace(varElements2D(lngj), """", "") & """"
'                                blnRemoveAllDblQuotes = True
'                            Case Is = vbCancel
'                                Me.Hide
'                            End Select
'                        End If
'                    Else
'                        varElements2D(lngj) = """" & varElements2D(lngj) & """"
'                    End If
'                End If
'            End If
'
'            DoEvents
'        Next lngj
        
        clxRows.add varElements1D
    Next lngi
    
    Set grabValues = clxRows
    
    lblStatus.Caption = ""
End Function

Public Function populateTable(tableName As String, Data As Range, Optional forPreview As Boolean = False) As String()
    Dim clxData As Collection
    Dim lngi As Long
    Dim lngj As Long
    Dim attributes As Long
    Dim strValues As String
    Dim insertSQL As String
    Dim sql As String
    Dim allRowsSQL() As String
    Dim r As Long
    
    Set clxData = grabValues(Data, forPreview)
    
    If clxData Is Nothing Then
        MsgBox "ERROR: " & vbCr & vbCr & "Invalid data range.", vbOKOnly + vbCritical, "SQLite - [Create table]"
        Exit Function
    End If
    
    If clxData.count = 0 Then
        MsgBox "ERROR: " & vbCr & vbCr & "Invalid data range.", vbOKOnly + vbCritical, "SQLite - [Create table]"
        Exit Function
    Else
        ReDim allRowsSQL(0 To clxData.count - 1) As String
    End If
    
    insertSQL = "INSERT INTO " & tableName & " VALUES ("
    attributes = UBound(clxData.item(1)) + 1    'the array is base 0
    
    For lngi = 1 To clxData.count
        strValues = ""
        lblStatus.Caption = "Generating INSERT for data row " & lngi & " [" & Int(((lngi) / (clxData.count)) * 100) & "%]"
        
        'strValues = "'" & Join(clxData.item(lngi), "', '") & "'"
        
        If Not chkQuotes.Value Then
            If chkDblQuotes.Value Then
                strValues = """" & Join(clxData.item(lngi), """, """) & """"
            Else
                strValues = "'" & Join(clxData.item(lngi), "', '") & "'"
            End If
        Else
            strValues = Join(clxData.item(lngi), ", ")
        End If
                
        sql = insertSQL & strValues & ")"
        allRowsSQL(lngi - 1) = sql
        
        DoEvents
    Next lngi
    
    populateTable = allRowsSQL
    lblStatus.Caption = ""
End Function

Private Sub btnCreate_Click()
    Dim rngHeadings As Range
    Dim rngData As Range
    Dim rngSheetname As String
    Dim blnValidRange As Boolean
    Dim blnValidName As Boolean
    Dim clxRtrn As Collection
'    Dim s As clsSQLiteManager
    Dim r As Long
    Dim sql As String
    Dim sqlm() As String
    Dim lngi As Long
    Dim startRow As Long

    
    Set clxRtrn = New Collection
    
    If Not fileExists(txtDB_Create.Caption) Then
        MsgBox "Invalid Database specified.", vbOKOnly + vbCritical, "Create SQLite table"
    End If
    
    checkInputs blnValidName, blnValidRange
    
    'if either of the inputs is invalid show an error message and ask user to re-input
    If Not blnValidName Or Not blnValidRange Then
        MsgBox "Invalid" & _
               IIf(Not blnValidRange, " 'Table source'", "") & _
               IIf(Not blnValidRange And Not blnValidName, " and", "") & _
               IIf(Not blnValidName, " 'Table name'", "") & _
               " specified.", vbCritical + vbOKOnly, "Create SQLite table"
    Else
    'both inputs are ok, proceed with table definition
        rngSheetname = Split(rfdTableSource, "!")(0)
        rngSheetname = WorksheetFunction.Substitute(rngSheetname, "'", "")
        
        If chkDropTable.Value Then
            Set rngHeadings = ActiveWorkbook.Worksheets(rngSheetname).Range(Split(rfdTableSource, "!")(1))
            Set rngHeadings = rngHeadings.Resize(1, rngHeadings.columns.count)
            
            If Not checkAttributeNames(rngHeadings) Then
                MsgBox "NULL values are not allowed for table attribute names.", vbCritical + vbOKOnly, "SQLite - [Create table]"
                lblStatus.Caption = ""
                Exit Sub
            End If
            
            If chkTypes.Value Then
                If frmAttribs Is Nothing Then
                    Set frmAttribs = New frmSQLite_Attributes
                    
                    Load frmAttribs
                    frmAttribs.setDB txtDB_Create.Caption
                    frmAttribs.setAttributes txtTableName.Value & " in " & Split(txtDB_Create.Caption, Application.PathSeparator)(UBound(Split(txtDB_Create.Caption, Application.PathSeparator))), _
                                             rngHeadings, _
                                             chkTempTable.Value
                Else
                    frmAttribs.setDB txtDB_Create.Caption
                End If
                
                'Me.Hide
                frmAttribs.Show
                
                Select Case frmAttribs.closeMode
                Case Is = 0  'user cancelled
                    On Error Resume Next
                    Me.Show
                    On Error GoTo 0
                    
                    lblStatus.Caption = ""
                    Exit Sub
                    
                End Select
                
            End If
            
            Set rngData = ActiveWorkbook.Worksheets(rngSheetname).Range(Split(rfdTableSource, "!")(1))
            
            If rngData.rows.count > 1 Then
                Set rngData = rngData.Offset(1, 0).Resize(rngData.rows.count - 1, rngData.columns.count)
            Else
                Set rngData = Nothing
            End If
        Else
            Set rngData = ActiveWorkbook.Worksheets(rngSheetname).Range(Split(rfdTableSource, "!")(1))
        End If
        
'        Set s = New clsSQLiteManager
        Set clxRtrn = New Collection
        
        If folderExists(txtDB_Create.Caption) Then
            If Not sqlManager.isDBOpen Then
                r = sqlManager.openDB(txtDB_Create.Caption)
            End If
            
            If r <> 0 Then
                appendLog sqlManager.getError
                MsgBox "Error: " & sqlManager.getError, vbOKOnly + vbCritical
                lblStatus.Caption = ""
                Exit Sub
            End If
            
            If chkDropTable.Value Then
                
                If chkTypes.Value Then
                    sql = frmAttribs.getSQL
                    
                    'unload frmattribs and set it to nothing
                    Unload frmAttribs
                    Set frmAttribs = Nothing
                Else
                    sql = getCreateTableSQL(txtTableName.Value, ActiveWorkbook.Worksheets(rngSheetname).Range(rngHeadings.Address))
                End If
                
                If chkDropTable.Value Then
                    r = sqlManager.executeNonQuery("DROP TABLE " & txtTableName.Value)
                    appendLog "----------------------------------------------------"
                    appendLog "DROP TABLE " & txtTableName.Value, True
                    
                    If r <> 0 Then
                        appendLog sqlManager.getError
                    Else
                        appendLog "Table dropped"
                    End If
                End If
                
                appendLog "----------------------------------------------------"
                appendLog sql, True
                r = sqlManager.executeNonQuery(sql)
        
                If r <> 0 Then
                    appendLog sqlManager.getError
                    MsgBox "Error: " & sqlManager.getError, vbOKOnly + vbCritical
                    lblStatus.Caption = ""
                    Exit Sub
                Else
                    appendLog "Table created successfully" & vbCr
                End If
            End If
            
            If Not rngData Is Nothing Then
                sqlm = populateTable(txtTableName.Value, ActiveWorkbook.Worksheets(rngSheetname).Range(rngData.Address))
                startRow = ActiveWorkbook.Worksheets(rngSheetname).Range(rngData.Address).Resize(1, 1).row
                
                For lngi = 0 To UBound(sqlm)
                    lblStatus.Caption = "Executing INSERT for row " & lngi & " [" & Int(((lngi + 1) / (UBound(sqlm) + 1)) * 100) & "%]"
                    
                    'appendLog "[Row " & startRow + lngi & "]  " & sqlm(lngi)   -- time-hog!
                    r = sqlManager.executeNonQuery(sqlm(lngi))
                    
                    If r <> 0 Then
                        appendLog sqlManager.getError
                        MsgBox "Error: " & sqlManager.getError, vbOKOnly + vbCritical
                        lblStatus.Caption = ""
                        Exit Sub
                    End If
                    
                    DoEvents
                Next lngi
                
                appendLog lngi & " rows inserted" & vbCr
            End If
            
'            r = sqlManager.closeDB
            
'            Set s = Nothing
            Set clxRtrn = Nothing
            Set rngHeadings = Nothing
            Set rngData = Nothing
            
            txtTableName.Value = ""
            rfdTableSource.Value = ""
            txtPreview.Value = ""
            
            refreshTableList    'to show any newly created table in the list immediately
            
            MsgBox "SQL executed.", vbOKOnly + vbInformation, "XL-SQLite"
        Else
            MsgBox "Unable to acces the given location.", vbCritical + vbOKOnly, "SQLite - [Create table]"
        End If
    End If

    refreshTableList
    lblStatus.Caption = ""
End Sub

Private Sub btnRefreshPreview_Click()
    updateSQLPreview
End Sub

Private Sub btnRefresh_Click()
    refreshTableList
End Sub

Private Sub chkDblQuotes_Click()
    '[Start:20130907:7]
    '[Cmt:Added to fix issue with single quotes in field values]
    updateSQLPreview
    '[End:20130907:7]
End Sub

Private Sub chkDropTable_Click()
    '[Start:20130907:8]
    '[Cmt:Added to fix issue with single quotes in field values]
    lblFirstRow.visible = chkDropTable.Value
    chkTypes.enabled = chkDropTable.Value
    chkTempTable.enabled = True
    btnCreate.Caption = "Next..."
    
    updateSQLPreview
    '[End:20130907:8]
End Sub

Private Sub chkInsertData_Click()
    lblFirstRow.visible = chkInsertData.Value
    chkTypes.enabled = chkInsertData.Value
    chkTempTable.enabled = False
    btnCreate.Caption = "Execute"
    
    updateSQLPreview
End Sub

Private Sub chkQuotes_Click()
    '[Start:20130907:9]
    '[Cmt:Added to fix issue with single quotes in field values]
    If chkQuotes Then
        chkDblQuotes.enabled = False
        chkDblQuotes.Value = False
    Else
        chkDblQuotes.enabled = True
    End If
    
    updateSQLPreview
    '[End:20130907:9]
End Sub

Private Sub chkTempTable_Click()
    updateSQLPreview
End Sub

Private Sub chkTypes_Click()
    If Not chkTypes.Value Then
        btnCreate.Caption = "Execute"
    Else
        btnCreate.Caption = "Next..."
    End If
End Sub

Private Sub txtDB_Create_Change()
    refreshTableList
    
    On Error Resume Next
    txtDB_Create.SetFocus
    On Error GoTo 0
End Sub

Private Sub txtTableName_Change()
    updateSQLPreview
End Sub

Private Sub UserForm_Activate()
    If sqlManager Is Nothing Then
        Set sqlManager = New clsSQLiteManager
    End If
    
    If Not sqlManager.isDBOpen Then
        MsgBox "You must connect to a database first.", vbOKOnly + vbInformation, "SQLite"
        
        frmConnection.Show
        
        If Not sqlManager.isDBOpen Then End
    End If
    
    txtDB_Create.Caption = sqlManager.getCurrentDBPath
    
    refreshTableList
    
    On Error Resume Next
    txtTableName.SetFocus
    On Error GoTo 0
End Sub

Private Sub UserForm_Initialize()
    If sqlManager Is Nothing Then
        Set sqlManager = New clsSQLiteManager
    End If
    
    If Not sqlManager.isDBOpen Then
        MsgBox "You must connect to a database first.", vbOKOnly + vbInformation, "SQLite"
        
        frmConnection.Show
        
        If Not sqlManager.isDBOpen Then End
    End If
    
    txtDB_Create.Caption = sqlManager.getCurrentDBPath
End Sub


Private Sub refreshTableList()
    Dim r As Long
    Dim clxRtrn As Collection
    Dim item As Variant
    Dim inti As Integer

    lbxTables.Value = ""

    If fileExists(txtDB_Create.Caption) Then
        r = 0
        
        If Not sqlManager.isDBOpen Then
            r = sqlManager.openDB(txtDB_Create.Caption)
        End If
    
        If r = 0 Then
            Set clxRtrn = New Collection
    
            sqlManager.executeQuery "Select name from sqlite_master where type in (""table"", ""view"") UNION ALL Select ""[T] "" || name from sqlite_temp_master where type in (""table"", ""view"") order by name", clxRtrn
    
            If clxRtrn.count > 0 Then
                inti = 0
                For Each item In clxRtrn
                    If inti > 0 Then lbxTables.Value = lbxTables.Value & item(0) & vbCr 'skip the header row
                    inti = inti + 1
                Next item
            End If
        End If
    
        On Error Resume Next
        lbxTables.SetFocus
        On Error GoTo 0
    
        'house keeping
        Set clxRtrn = Nothing
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, closeMode As Integer)
    On Error Resume Next
    txtDB_Create.SetFocus   'just in case the other page is showing and this is hidded
    On Error GoTo 0
    
    Cancel = True
    
    txtDB_Create.Caption = "ERROR: Connection lost!"
    
    Me.Hide
End Sub

Private Sub appendLog(text As String, Optional cmd As Boolean = False)
    txtLog.Value = txtLog.Value & IIf(Len(txtLog.Value) = 0, "", vbCr) & IIf(cmd, ">> ", "") & text
End Sub

Private Sub appendPreview(text As String)
    txtPreview.Value = txtPreview.Value & IIf(Len(txtPreview.Value) = 0, "", vbCr) & text
End Sub

Private Function checkAttributeNames(headings As Range) As Boolean
    Dim rngCell As Range
    
    checkAttributeNames = True
    For Each rngCell In headings
        If IsEmpty(rngCell.Value) Then
            checkAttributeNames = False
            Exit For
        End If
    Next rngCell
End Function

Private Function updateSQLPreview()
    Dim rngHeadings As Range
    Dim rngData As Range
    Dim rngSheetname As String
    Dim blnValidRange As Boolean
    Dim blnValidName As Boolean
    Dim regex As VBScript_RegExp_55.RegExp
    Dim sql As String
    Dim sqlm() As String
    Dim lngi As Long
    Dim startRow As Long

    
    txtPreview.Value = ""
    checkInputs blnValidName, blnValidRange
    
    'if either of the inputs is invalid show an error message and ask user to re-input
    If blnValidName And blnValidRange Then
    'both inputs are ok, proceed with table definition
        rngSheetname = Split(rfdTableSource, "!")(0)
        rngSheetname = WorksheetFunction.Substitute(rngSheetname, "'", "")
        
        If chkDropTable.Value Then
            Set rngHeadings = ActiveWorkbook.Worksheets(rngSheetname).Range(Split(rfdTableSource, "!")(1))
            Set rngHeadings = rngHeadings.Resize(1, rngHeadings.columns.count)
            
            If Not checkAttributeNames(rngHeadings) Then
                MsgBox "NULL values are not allowed for table attribute names.", vbCritical + vbOKOnly, "SQLite - [Create table]"
                Exit Function
            End If
            
            Set rngData = ActiveWorkbook.Worksheets(rngSheetname).Range(Split(rfdTableSource, "!")(1))
            
            If rngData.rows.count > 1 Then
                Set rngData = rngData.Offset(1, 0).Resize(rngData.rows.count - 1, rngData.columns.count)
            Else
                Set rngData = Nothing
            End If
        Else
            Set rngData = ActiveWorkbook.Worksheets(rngSheetname).Range(Split(rfdTableSource, "!")(1))
        End If
        
        If chkDropTable.Value Then
            appendPreview "DROP TABLE " & txtTableName.Value & ";" & vbCr
            appendPreview getCreateTableSQL(txtTableName.Value, ActiveWorkbook.Worksheets(rngSheetname).Range(rngHeadings.Address)) & ";" & vbCr
        End If
        
        If Not rngData Is Nothing Then
            sqlm = populateTable(txtTableName.Value, ActiveWorkbook.Worksheets(rngSheetname).Range(rngData.Address), True)
            
            For lngi = 0 To UBound(sqlm)
                appendPreview sqlm(lngi) & ";"
                If lngi = previewLimit - 1 Then
                    appendPreview vbCr & "Showing the first " & previewLimit & " inserts only."
                    Exit For
                End If
            Next lngi
        End If
    End If
End Function