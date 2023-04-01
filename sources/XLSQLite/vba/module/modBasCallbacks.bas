Option Explicit

Public gobjRibbon As IRibbonUI
Public bolEnabled As Boolean
Public bolVisible As Boolean


Public Sub OnRibbonLoad_XLSQLite(ribbon As IRibbonUI)
    Set gobjRibbon = ribbon
End Sub

Public Sub OnActionButton_XLSQLite(control As IRibbonControl)
    Select Case control.id
        Case "btnSQLDDL"
            Call SQLiteDDL
        Case "btnSQLEditor"
            Call SQLiteEditor
        Case "btnAbout"
            Call XLSQLiteAbout
        Case Else
    End Select
End Sub

Public Sub GetVisible_XLSQLite(control As IRibbonControl, ByRef visible)
    Select Case control.id
        Case Else
            visible = True
    End Select
End Sub

Public Sub GetEnabled_XLSQLite(control As IRibbonControl, ByRef enabled)
    If Workbooks.count = 0 Then
        enabled = False
    Else
        enabled = True
    End If
End Sub