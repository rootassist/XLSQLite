'XLSQlite
'
'LICENSE: The MIT License (MIT)
'
'Copyright ｩ 2014 Mark Camilleri
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

Private blnSelectedTemplate As Boolean
Private strSQLFileFullPath As String

Private Declare Function FindWindowA Lib "user32" _
(ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long
 
Private Declare Function GetWindowLongA Lib "user32" _
(ByVal hwnd As Long, _
ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLongA Lib "user32" _
(ByVal hwnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Private Function getFile(winTitle As String, Optional startPath As String = "C:\", Optional filter As String = "Text files (*.txt), *.txt") As String
    Dim inti As Integer
    Dim intLastOccur As Integer
    
    ChDir startPath
    
    getFile = Application.GetOpenFilename(filter, 1, winTitle, , False)
                        
    Application.StatusBar = False
End Function

Private Function saveFile(winTitle As String, Optional startPath As String = "C:\", Optional filter As String = "Text files (*.txt), *.txt") As String
    Dim inti As Integer
    Dim intLastOccur As Integer
    
    saveFile = Application.GetSaveAsFilename(fileFilter:=filter, _
                        InitialFileName:=startPath, _
                        title:=winTitle)
                        
    Application.StatusBar = False
End Function

Private Sub saveResultToFile(resultset As Collection, fullPath As String, createFile As Boolean)
    Dim intFreeFile As Integer
    Dim lines As Long
    Dim lngi As Long
    Dim lngj As Long
    Dim lngk As Long
    Dim strDelimeter As String
    Dim strBuffer As String
    Dim bufferLines As Long
    Dim strSeparator
    
    bufferLines = 5000
    
    lines = resultset.count
    
    If lines > 0 Then
        appendMessage "Data fetched @" & Now()
        
        intFreeFile = FreeFile
        
        On Error Resume Next
        If createFile Then
            Open fullPath For Output Access Write Lock Write As #intFreeFile
            
            If Err.Number <> 0 Then
                appendLog Err.description & " [" & txtResultFile.Value & "]"
                MsgBox Err.description, vbOKOnly + vbCritical, "SQLite - [Save to file]"
                On Error GoTo 0
                Exit Sub
            End If
        Else
            Open fullPath For Append Access Write Lock Write As #intFreeFile
        
            If Err.Number <> 0 Then
                appendLog Err.description & " [" & txtResultFile.Value & "]"
                MsgBox Err.description, vbOKOnly + vbCritical, "SQLite - [Save to file]"
                On Error GoTo 0
                Exit Sub
            End If
        End If
        On Error GoTo 0
        
        strDelimeter = txtDelimeter.Value
        strSeparator = txtTextSeparator.Value
        lngj = 0
        
        For lngi = 1 To lines
            lblStatus.Caption = "Buffering data row " & lngi - 1 & " [" & Int(((lngi) / (lines)) * 100) & "%]"
            
            If lngj > bufferLines Then
                lblStatus.Caption = "Writing to file..."
                Print #intFreeFile, strBuffer   'write the buffer to disk
                
                strBuffer = strSeparator & Join(resultset.item(lngi), strSeparator & strDelimeter & strSeparator) & strSeparator    'empty the buffer and put the next line in
                lngj = 1    'restart the counter
            Else
                lngj = lngj + 1
                
                Select Case lngj
                Case Is = 1
                    strBuffer = strSeparator & Join(resultset.item(lngi), strSeparator & strDelimeter & strSeparator) & strSeparator    'the first row returned; result attributes
                Case Else
                    On Error Resume Next
                    strBuffer = strBuffer & vbCr & strSeparator & Join(resultset.item(lngi), strSeparator & strDelimeter & strSeparator) & strSeparator
                    
                    If Err.Number <> 0 Then
                        appendMessage "ERROR! unhandled error during buffering [" & Err.description & "]"
                    End If
                    
                    On Error GoTo 0
                End Select
            End If
            
            DoEvents
        Next lngi
        
        If Len(strBuffer) > 0 Then
        'write the last data in the buffer to file
            lblStatus.Caption = "Writing to file..."
            
            Print #intFreeFile, strBuffer
        End If
        
        appendMessage lines - 1 & IIf(lines = 2, " row", " rows") & " returned.  [" & fullPath & "]"
        
        Close #intFreeFile
    End If
End Sub

Private Sub showResult(resultset As Collection, Optional workbookName As String = "", Optional worksheetName As String = "")
    Dim lngi As Long
    Dim lngj As Long
    Dim rows As Long
    Dim columns As Long
    Dim varElements As Variant
    Dim topLeftCell As Variant
    
    If resultset.count > 0 Then
        appendMessage "Data fetched @" & Now()
        
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
        
        Workbooks(workbookName).Worksheets(worksheetName).Cells.ClearContents
        
        Set topLeftCell = Workbooks(workbookName).Worksheets(worksheetName).Range("A1")
        
        If resultset Is Nothing Then
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
        
        If resultset.count = 0 Then
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
            
        rows = resultset.count
        
        varElements = resultset.item(1)
        If Not IsArray(varElements) Then
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
        
        columns = UBound(resultset.item(1)) + 1
        
        'Application.StatusBar = "Populating " & rows - 1 & " row(s)."
        For lngi = 0 To rows - 1
            topLeftCell.Offset(lngi, 0).Resize(1, columns).Value = resultset.item(lngi + 1)
            
            lblStatus.Caption = "Populating data row " & lngi & " [" & Int(((lngi + 1) / (rows + 1)) * 100) & "%]"
            
            DoEvents
        Next lngi
        
        Application.StatusBar = False
        
        Workbooks(workbookName).Worksheets(worksheetName).Select
        
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        
        appendMessage lngi - 1 & IIf(lngi = 2, " row", " rows") & " returned.  [" & cmbWB_SQL & "!" & cmbWS_SQL & "]"
    Else
        appendMessage "No data fetched."
    End If
End Sub

Public Sub SQLite_Query()
    Dim r As Long
    Dim clxRtrn As Collection
    Dim lines As Long
    Dim tableName As String
    Dim WBName As String
    Dim Sheetname As String
    Dim DataRange As String
    Dim Query As String
    Dim varQueries As Variant
    Dim lngi As Long
    Dim lngj As Long
    Dim ws As Worksheet
    Dim blnCreateFile As Boolean
    Dim blnClean As Boolean
    
    If chkSaveToFile.Value Then
        If Trim(txtResultFile.Value) = "" Then
            appendLog "Invalid file name specified. [" & txtResultFile.Value & "]"
            MsgBox "Invalid file name specified.", vbOKOnly + vbCritical, "SQLite - [Save to file]"
            Exit Sub
        End If
        
        If Not folderExists(txtResultFile.Value) Then
            appendLog "Invalid folder name specified. [" & txtResultFile.Value & "]"
            MsgBox "Invalid folder name specified.", vbOKOnly + vbCritical, "SQLite - [Save to file]"
            Exit Sub
        Else
            If Not fileExists(txtResultFile.Value) Then
                blnCreateFile = True
            Else
                blnCreateFile = False
                If chkOverwrite Then
                    On Error Resume Next
                    Kill txtResultFile.Value
                    
                    If Err.Number <> 0 Then
                        appendLog Err.description & " [" & txtResultFile.Value & "]"
                        MsgBox Err.description, vbOKOnly + vbCritical, "SQLite - [Save to file]"
                        On Error GoTo 0
                        Exit Sub
                    End If
                    
                    On Error GoTo 0
                End If
            End If
        End If
    End If
    
    If fileExists(txtDB_SQL.Caption) Then
        Set clxRtrn = New Collection
        
        DoEvents
        
        Query = WorksheetFunction.Trim(txtEditor.Value)
        
        If Len(Query) = 0 Then
            MsgBox "No query found!", vbExclamation
            Exit Sub
        End If
        
        'cleaning the query text
        Do
            blnClean = True
            
            Select Case Asc(Left(Query, 1))
            Case 10, 13, 32
                Query = Right(Query, Len(Query) - 1)
                blnClean = False
            End Select
        
            Select Case Asc(Right(Query, 1))
            Case 10, 13, 32
                Query = Left(Query, Len(Query) - 1)
                blnClean = False
            End Select
            
        Loop Until blnClean
        
        If Not sqlManager.isDBOpen Then
            r = sqlManager.openDB(txtDB_SQL.Caption)
        End If
        
        If r <> 0 Then
            appendMessage "ERROR: " & sqlManager.getError
            Exit Sub
        End If
        
        If Right(Query, Len(txtSQLSeparator.Value)) = txtSQLSeparator.Value Then
            Query = Left(Query, Len(Query) - 1)
        End If
        
        appendLog "<query>"
        appendLog Query
        appendLog "</query>"
        
        varQueries = Split(Query, txtSQLSeparator.Value)  'TODO: this will give bad results if the ; is used in the query itself
        
        For lngi = 0 To UBound(varQueries)
            If Trim(CStr(varQueries(lngi))) <> "" Then
                appendLog ""
                appendLog "<result id=" & 1 + lngi & ">"
                appendMessage "Executing query " & lngi + 1 & " of " & UBound(varQueries) + 1, False
                DoEvents
                
                r = sqlManager.executeQuery(Trim(CStr(varQueries(lngi))), clxRtrn)
                
                If r <> 0 Then
                    appendMessage "ERROR: " & sqlManager.getError
                    appendLog "</result>"
                    appendLog "</execution>"
                    Exit Sub
                End If
                
                If clxRtrn.count = 0 Then
                    appendMessage sqlManager.getLinesChanged & IIf(sqlManager.getLinesChanged = 1, " row", " rows") & " affected."
                Else
                    If Not chkSaveToFile.Value Then
                        If chkNewWB.Value Then
                            If resultWB Is Nothing Then
                                Set resultWB = Application.Workbooks.add
                            End If
                            
                            On Error Resume Next
                            resultWB.Activate
                            
                            If Err.Number <> 0 Then 'the workbook must have been closed...create another one
                                Set resultWB = Nothing
                                Set resultWB = Application.Workbooks.add
                            End If
                            On Error GoTo 0
                            
                            cmbWB_SQL.AddItem resultWB.Name
                            cmbWB_SQL.Value = resultWB.Name
                            
                            For lngj = 1 To cmbWS_SQL.ListCount
                                cmbWS_SQL.RemoveItem 0
                            Next lngj
                            
                            For lngj = 1 To resultWB.Worksheets.count
                                cmbWS_SQL.AddItem resultWB.Worksheets(lngj).Name
                            Next lngj
                            
                            cmbWS_SQL.Value = resultWB.Worksheets(1).Name
                        Else
                            On Error Resume Next
                            Workbooks(cmbWB_SQL.Value).Activate
                            Set ws = ActiveWorkbook.Worksheets(cmbWS_SQL.Value)
                            
                            'if an error was generate the input workbook/worksheet does not exist
                            If Err.Number <> 0 Then
                                Err.Clear
                                
                                Workbooks(cmbWB_SQL.Value).Activate
                                
                                If Err.Number <> 0 Then
                                'the workbook does not exist; open a new workbook
                                    Set resultWB = Nothing
                                    Set resultWB = Application.Workbooks.add
                                    
                                    cmbWB_SQL.AddItem resultWB.Name
                                    cmbWB_SQL.Value = resultWB.Name
                                    
                                    For lngj = 1 To cmbWS_SQL.ListCount
                                        cmbWS_SQL.RemoveItem 0
                                    Next lngj
                                    
                                    For lngj = 1 To resultWB.Worksheets.count
                                        cmbWS_SQL.AddItem resultWB.Worksheets(lngj).Name
                                    Next lngj
                            
                                    cmbWS_SQL.Value = resultWB.Worksheets(1)
                                Else
                                'the worksheet must be missing
                                    Set ws = Workbooks(cmbWB_SQL.Value).Worksheets.add
                                    Err.Clear
                                    
                                    ws.Name = cmbWS_SQL.Value
                                    
                                    'invalid name
                                    If Err.Number <> 0 Then
                                        cmbWS_SQL.Value = ws.Name
                                    End If
                                End If
                            End If
                            On Error GoTo 0
                        End If
                    End If
                    
                    If Not chkSaveToFile.Value Then
                        showResult clxRtrn, cmbWB_SQL.Value, cmbWS_SQL.Value
                    Else
                        saveResultToFile clxRtrn, txtResultFile.Value, blnCreateFile
                    End If
                End If
            End If
            
            appendMessage "Done @" & Now()
            appendLog "</result>"
        Next lngi
                
        Set clxRtrn = Nothing
        
        appendLog "</execution>"
'        MsgBox "SQL executed.", vbOKOnly + vbInformation, "XL-SQLite"
    Else
        appendMessage "Database file not found."
        MsgBox "Database file not found.", vbOKOnly + vbCritical, "SQLite - SQLEditor"
    End If
End Sub

Private Sub btnBrowse_Click()
    Dim txtFile As String
    
    Do
        txtFile = getFile("SQLite Editor - [Select Database]", ThisWorkbook.path & "\sqlite\", "SQLite DB files (*.sqlite), *.sqlite")
        
        If txtFile <> "False" Then
            txtDB_SQL.Caption = txtFile
        Else
            txtDB_SQL.Caption = ""
            txtFile = "Cancel"
            MsgBox "No database was chosen.", vbCritical + vbOKOnly, "SQLite Editor - [Select Database]"
        End If
    Loop Until txtFile <> "False"
    
    If txtFile <> "Cancel" Then refreshTableList
    
End Sub

Private Sub btnBrowseRF_Click()
    Dim txtFile As String
    
    Do
        txtFile = saveFile("SQLite Editor - [Select result file]", ThisWorkbook.path & "\sqlite\", "Pipe delimited (*.sqf), *.sqf")
    
        If txtFile <> "False" Then
            txtResultFile.Value = txtFile
        Else
            txtResultFile.Value = ""
            txtFile = "Cancel"
        End If
    Loop Until txtFile <> "False"
    
End Sub

Private Sub btnChangeDB_Click()
    
    If vbYes = MsgBox("If you change or reset your connection you will loose any temporary tables created on the current database." & _
                     vbCr & vbCr & "Are you sure you want to continue?", vbYesNo + vbDefaultButton2 + vbQuestion, "SQLite - [Change connection]") Then
        frmConnection.Show
        
        UserForm_Initialize
    End If
End Sub

Private Sub btnClear_Click()
    If Len(txtLog.Value) > 0 Then
        If MsgBox("Are you sure you want to clear the log?", vbQuestion + vbYesNo + vbDefaultButton2, "SQLite Editor") = vbYes Then
            txtLog.Value = ""
            lblStatus.Caption = ""
        End If
    End If
End Sub

Private Sub btnExecute_Click()
    Dim blnConfirm As Boolean
    
    blnConfirm = True
    
    If chkConfirm.Value And Not chkNewWB.Value Then
    'TODO: check that workbook and worksheet names are valid
        blnConfirm = IIf(MsgBox("Do you want to paste the result in " & cmbWB_SQL.Value & "!" & cmbWS_SQL.Value & "?", _
                                vbYesNo + vbQuestion + vbDefaultButton2, "SQLite - [SQL Editor]") = vbYes, True, False)
    End If
    
    If blnConfirm Then
        clearMessage
        
        appendLog "----------------------------------------------------"
        appendLog "<execution timestamp=""" & Now() & """>"
        appendLog "Run on: " & txtDB_SQL & vbCr
        
        SQLite_Query
        
        On Error Resume Next
        txtEditor.SetFocus
        On Error GoTo 0
    Else
        clearMessage
        appendMessage "SQL execution aborted!"
    End If
End Sub

Private Sub btnLoad_Click()
    Dim txtFile As String
    Dim txtLine As String
    Dim intFreeFile As Integer
    Dim varItems As Variant
    
    'if there are unsaved changes warn the user
    If Left(Me.Caption, 1) = "*" Then
        If MsgBox("You have unsaved changes which will be lost." & vbCr & vbCr & "Do you want to continue?", vbYesNo + vbQuestion + vbDefaultButton2, "SQLite Editor - [Load SQL]") = vbNo Then
            Exit Sub
        End If
    End If
    
    txtFile = getFile("SQLite Editor - [Load SQL]", ThisWorkbook.path & "\sqlite\scripts\", "SQL files (*.sqlf), *.sqlf")

    If txtFile = "False" Then
        txtFile = "Cancel"
        
        MsgBox "No SQL file was chosen.", vbCritical + vbOKOnly, "SQLite Editor - [Load SQL]"
    End If
    
    If txtFile <> "Cancel" Then
        strSQLFileFullPath = txtFile
        varItems = Split(strSQLFileFullPath, Application.PathSeparator)
        intFreeFile = FreeFile
        txtEditor.text = ""
        
        Open txtFile For Input Access Read As #intFreeFile
        While Not EOF(intFreeFile)
            Line Input #intFreeFile, txtLine
            
            txtEditor.text = txtEditor.text & txtLine & vbCr
        Wend
        
        Close #intFreeFile
        
        'after loading remove the leading asterisk from the window caption
        If Left(Me.Caption, 1) = "*" Then
            Me.Caption = "SQLite - [SQL Editor] {" & varItems(UBound(varItems)) & "}"
            addMinimizeButton
        End If
        
        sqlEditorOrigValue = txtEditor.Value
        lblStatus.Caption = ""
    End If
    
End Sub

Private Sub btnNew_Click()
    If Left(Me.Caption, 1) = "*" Then
        If MsgBox("If you continue all your unsaved changes will be lost." & vbCr & vbCr & "Do you want to continue?", vbYesNo + vbDefaultButton2 + vbQuestion, "SQLite Editor - [New]") = vbYes Then
            txtEditor.Value = ""
            strSQLFileFullPath = ""
            Me.Caption = "SQLite - [SQL Editor]"
            
            txtEditor.SetFocus
            lblStatus.Caption = ""
        End If
    Else
        txtEditor.Value = ""
        strSQLFileFullPath = ""
        Me.Caption = "SQLite - [SQL Editor]"
        
        txtEditor.SetFocus
        lblStatus.Caption = ""
    End If
End Sub

Private Sub btnRefresh_Click()
    refreshTableList
End Sub

Private Sub btnSaveAs_Click()
    Call saveAsFile
End Sub

Private Sub saveAsFile(Optional fileName As String = "")
    Dim txtLine As String
    Dim intFreeFile As Integer
    
    If fileName = "" Then
        fileName = saveFile("SQLite Editor - [Save As]", ThisWorkbook.path & "\sqlite\scripts\", "SQL files (*.sqlf), *.sqlf")
        
        If fileExists(fileName) Then
            If MsgBox("File already exists, do you want to overwrite?", vbYesNo + vbExclamation + vbDefaultButton2, "SQLite Editor - [Save As]") = vbNo Then
                Exit Sub    'abort save as
            End If
        End If
    End If

    If fileName = "False" Then
        fileName = "Cancel"
        
        MsgBox "Save cancelled.", vbCritical + vbOKOnly, "SQLite Editor - [Save As]"
    End If
    
    If fileName <> "Cancel" Then
        intFreeFile = FreeFile
        
        On Error Resume Next
        Kill fileName    'delete any existing file
        On Error GoTo 0
        
        Open fileName For Output Access Write As #intFreeFile
        Print #intFreeFile, txtEditor.text
        
        Close #intFreeFile
        
        'update the filename in the form's caption
        Me.Caption = "SQLite - [SQL Editor] {" & Split(fileName, Application.PathSeparator)(UBound(Split(fileName, Application.PathSeparator))) & "}"
        addMinimizeButton
        
        strSQLFileFullPath = fileName
        sqlEditorOrigValue = txtEditor.Value
        
        appendLog "----------------------------------------------------"
        appendMessage "'" & fileName & "' saved @" & Now()
        appendLog ""
        
        MsgBox "File saved successfully.", vbOKOnly + vbInformation
    End If
    
End Sub

Private Sub btnSave_Click()
    If Left(Me.Caption, 1) = "*" Then  'if the file has not changed there's no need to save it
        If strSQLFileFullPath <> "" Then
            Call saveAsFile(strSQLFileFullPath)
        Else
            Call saveAsFile
        End If
    End If
End Sub

Private Sub chkNewWB_Click()
    chkConfirm.enabled = Not chkNewWB.Value
    Label7.enabled = Not chkNewWB.Value
    Label8.enabled = Not chkNewWB.Value
    cmbWB_SQL.enabled = Not chkNewWB.Value
    cmbWS_SQL.enabled = Not chkNewWB.Value
End Sub

Private Sub chkSaveToFile_Click()
    frmWorkbook.visible = Not chkSaveToFile.Value
    frmResultFile.visible = chkSaveToFile.Value
End Sub

Private Sub cmbTemplates_Change()
    If Not blnSelectedTemplate Then
        txtEditor.SelText = cmbTemplates.Value
        blnSelectedTemplate = True
    Else
        blnSelectedTemplate = False
    End If
    
    cmbTemplates.listIndex = -1
    
    On Error Resume Next
    txtEditor.SetFocus
    On Error GoTo 0
End Sub

Private Function getTableAttributes(tableName As String) As String()
    Dim r As Long
    Dim clxRtrn As Collection
    Dim tableDef As String
    Dim Query As String
    Dim firstBracket As Long
    Dim columnDeclaration As String
    Dim columns As Variant
    Dim lngi As Long
    
    If fileExists(txtDB_SQL.Caption) Then
        Set clxRtrn = New Collection
        
        If Left(tableName, 4) = "[T] " Then
            tableName = Replace(tableName, "[T] ", "", 1, 1)
        End If
        
        Query = "PRAGMA table_info(" & tableName & ")"
        
        r = 0
        If Not sqlManager.isDBOpen Then
            r = sqlManager.openDB(txtDB_SQL.Caption)
        End If
        
        If r <> 0 Then
            Exit Function
        End If
        
        r = sqlManager.executeQuery(Query, clxRtrn)
        
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
        
        Set clxRtrn = Nothing
    End If
    
End Function

Private Sub cmbWB_SQL_Change()
    Dim wksht As Worksheet
    Dim lngi As Long
    
    For lngi = 1 To cmbWS_SQL.ListCount
        cmbWS_SQL.RemoveItem (0)
    Next lngi
    
    On Error Resume Next
        For Each wksht In Workbooks(cmbWB_SQL.Value).Worksheets
            If Err.Number = 0 Then cmbWS_SQL.AddItem wksht.Name Else cmbWS_SQL.Value = ""
        Next
        
        cmbWS_SQL.listIndex = 0
    On Error GoTo 0
End Sub

Private Sub Label11_Click()
    refreshTableList
End Sub

Private Sub lbxAttributes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim strClmnName As String
    
    On Error Resume Next
        If lbxAttributes.List(lbxAttributes.listIndex) <> "<Click on a table/view>" Then
            '[Start:20130907:5]
            '[Cmt:Added to fix issue with column names starting with number, e.g. 01Table, will now be enclosed in square brackets]
            strClmnName = Replace(lbxAttributes.List(lbxAttributes.listIndex), "[T] ", "", 1, 1)
            If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(strClmnName, 1)), vbTextCompare) = 0 Then
                strClmnName = "[" & strClmnName & "]"
            End If
            
            txtEditor.SelText = IIf(chkIncludeTableName.Value, lbxt.List(lbxt.listIndex) & ".", "") & strClmnName
            '[End:20130907:5]
            
            '[Start:20130907:6]
            '[Cmt:Removed to fix issue with column names starting with number, e.g. 01Table, will now be enclosed in square brackets]
            'txtEditor.SelText = IIf(chkIncludeTableName.value, lbxt.List(lbxt.listIndex) & ".", "") & _
                                Replace(lbxAttributes.List(lbxAttributes.listIndex), "[T] ", "", 1, 1)
            '[End:20130907:6]
        End If
        
        txtEditor.SetFocus
    On Error GoTo 0
End Sub

Private Sub lbxt_Click()
    Dim attributes() As String
    Dim lngi As Long
    
    If lbxt.List(lbxt.listIndex) <> "<Select a database>" Then
        lbxAttributes.Clear
        
        attributes() = getTableAttributes(lbxt.List(lbxt.listIndex))
        
        For lngi = 0 To UBound(attributes)
            lbxAttributes.AddItem attributes(lngi)
        Next lngi
    End If
End Sub

Private Sub lbxt_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim fieldName As String
    
    If lbxt.List(lbxt.listIndex) <> "<Select a database>" Then
    
        If Left(lbxt.List(lbxt.listIndex), 4) = "[T] " Then
            fieldName = Replace(lbxt.List(lbxt.listIndex), "[T] ", "", 1, 1)
        Else
            fieldName = lbxt.List(lbxt.listIndex)
        End If

        txtEditor.SelText = fieldName
    End If
    
    On Error Resume Next
    txtEditor.SetFocus
    On Error GoTo 0
End Sub

Private Sub txtEditor_Change()
    If Left(Me.Caption, 1) <> "*" Then
        Me.Caption = "*" & Me.Caption
        addMinimizeButton
    Else
        If sqlEditorOrigValue = txtEditor.Value Then
            Me.Caption = Right(Me.Caption, Len(Me.Caption) - 1)
            addMinimizeButton
        End If
    End If
End Sub

Private Sub addMinimizeButton()
    Dim hwnd As Long
    Dim exLong As Long

    'adding the minimize button to the form
    'the minimize button has to be added each time the form's caption is changed
    
    hwnd = FindWindowA(vbNullString, Me.Caption)
    exLong = GetWindowLongA(hwnd, -16)
    If (exLong And &H20000) = 0 Then
        SetWindowLongA hwnd, -16, exLong Or &H20000
    End If
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
    
    txtDB_SQL.Caption = sqlManager.getCurrentDBPath
    
    refreshTableList

End Sub

Private Sub UserForm_Initialize()
    Dim wbk As Workbook
    Dim lngi As Long
    
    If sqlManager Is Nothing Then
        Set sqlManager = New clsSQLiteManager
    End If
    
    If Not sqlManager.isDBOpen Then
        MsgBox "You must connect to a database first.", vbOKOnly + vbInformation, "SQLite"
        
        frmConnection.Show
        
        If Not sqlManager.isDBOpen Then End
    End If

    txtDB_SQL.Caption = sqlManager.getCurrentDBPath
    
    addMinimizeButton
    
    For lngi = 1 To cmbWB_SQL.ListCount
        cmbWB_SQL.RemoveItem 0
    Next lngi
    
    For Each wbk In Application.Workbooks
        cmbWB_SQL.AddItem wbk.Name
    Next
    
    If fileExists(WorksheetFunction.Substitute(ThisWorkbook.path & "\SQLite\" & ThisWorkbook.Name, ".xlsm", ".sqlite")) Then
        txtDB_SQL.Caption = WorksheetFunction.Substitute(ThisWorkbook.path & "\SQLite\" & ThisWorkbook.Name, ".xlsm", ".sqlite")
    End If
    
    cmbTemplates.AddItem "SELECT * FROM sqlite_master;"
    cmbTemplates.AddItem "SELECT * FROM sqlite_temp_master;"
    cmbTemplates.AddItem "SELECT type, name FROM sqlite_master;"
    cmbTemplates.AddItem "SELECT type, name FROM sqlite_temp_master;"
    cmbTemplates.AddItem "SELECT * FROM tablename;"
    cmbTemplates.AddItem "VACUUM;"
    
    If Not chkNewWB.Value Then
        cmbWB_SQL.Value = ActiveWorkbook.Name
        
        cmbWS_SQL.Value = ActiveWorkbook.ActiveSheet.Name
    End If
    
    lbxAttributes.AddItem "<Click on a table/view>"
    lbxt.AddItem "<Select a database>"
    
    If fileExists(txtDB_SQL.Caption) Then refreshTableList
    
    strSQLFileFullPath = ""
    sqlEditorOrigValue = ""
End Sub

Private Sub appendMessage(text As String, Optional addToLog As Boolean = True)
    If addToLog Then
        txtLog.Value = txtLog.Value & IIf(Len(txtLog.Value) = 0, "", vbCr) & text
    End If

    lblStatus.Caption = text
End Sub

Private Sub clearMessage()
    txtLog.Value = IIf(Len(txtLog.Value) = 0, "", txtLog.Value & vbCr)

    lblStatus.Caption = ""
End Sub

Private Sub appendLog(text As String)
    txtLog.Value = txtLog.Value & IIf(Len(txtLog.Value) = 0, "", vbCr) & text
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, closeMode As Integer)
    Cancel = True
    
    txtDB_SQL.Caption = "ERROR: Connection lost!"
    
    Me.Hide
End Sub

Private Sub refreshTableList()
    Dim r As Long
    Dim clxRtrn As Collection
    Dim item As Variant
    Dim inti As Integer

    lbxt.Clear

    If fileExists(txtDB_SQL.Caption) Then
        If Not sqlManager.isDBOpen Then
            r = sqlManager.openDB(txtDB_SQL.Caption)
        End If
    
        If r = 0 Then
            Set clxRtrn = New Collection
    
            sqlManager.executeQuery "Select name from sqlite_master where type in (""table"", ""view"") UNION ALL Select ""[T] "" || name from sqlite_temp_master where type in (""table"", ""view"") order by name", clxRtrn
    
            If clxRtrn.count > 0 Then
                inti = 0
                For Each item In clxRtrn
                    If inti > 0 Then
                        lbxt.AddItem item(0)
                    End If
                    inti = inti + 1
                Next item
            Else
                lbxAttributes.Clear
                lbxAttributes.AddItem "<Click on a table/view>"
            End If
        End If
    
        'house keeping
        Set clxRtrn = Nothing
    Else
        lbxt.AddItem "<Select a database>"
        lbxAttributes.Clear
        lbxAttributes.AddItem "<Click on a table/view>"
        MsgBox "Database file not found.", vbOKOnly + vbCritical, "SQLite - [SQLEditor]"
    End If
End Sub