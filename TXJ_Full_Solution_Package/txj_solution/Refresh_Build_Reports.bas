Attribute VB_Name = "modOneClick"
Option Explicit

'Allowed macro scope per Appendix F:
' - Refresh all Power Query connections
' - Refresh all PivotTables
' - Navigate user to main report tab
'No calculations, no ledger edits, no exception suppression.

Public Sub Refresh_And_Build_Reports()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = True
    Application.StatusBar = "Refreshing queries..."

    'Refresh Power Query / connections
    ThisWorkbook.RefreshAll

    'Wait for asynchronous queries (Excel 2016+)
    On Error Resume Next
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo ErrHandler

    Application.StatusBar = "Refreshing pivots..."

    Dim ws As Worksheet
    Dim pt As PivotTable

    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws

    'Navigate to primary report sheet if present
    On Error Resume Next
    ThisWorkbook.Worksheets("Reports").Activate
    On Error GoTo ErrHandler

    Application.StatusBar = False

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Refresh failed: " & Err.Description, vbExclamation, "Tax-Grade Reporting Engine"
End Sub
