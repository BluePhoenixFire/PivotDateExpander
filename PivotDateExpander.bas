Attribute VB_Name = "PivotDateExpander"
Option Explicit
Sub PivotDateExpander(WB As Workbook) ' The Macro Name
    Dim ws As Worksheet
    For Each ws In WB.Worksheets
        Dim pt As PivotTable
        Dim ptLocation As String
        For Each pt In ws.PivotTables
            On Error GoTo ErrorHandler
            ptLocation = Left(pt.TableRange1.Address, 5)
            Debug.Print "Pivot table location: " & ptLocation
            
            If ptLocation = "$A$6:" Then
                pt.PivotFields("[Accounting Date].[CalendarYearMonthDay].[Calendar Year]").VisibleItemsList = Array("[Accounting Date].[CalendarYearMonthDay].[Calendar Year].&[2023]")
                pt.PivotFields("[Accounting Date].[CalendarYearMonthDay].[Calendar Year]").PivotItems("[Accounting Date].[CalendarYearMonthDay].[Calendar Year].&[2023]").DrilledDown = True
            Else
                pt.PivotFields("[Accounting Date].[CalendarYearMonthDay].[Calendar Year]").VisibleItemsList = Array("[Accounting Date].[CalendarYearMonthDay].[Calendar Year].&[2024]")
                pt.PivotFields("[Accounting Date].[CalendarYearMonthDay].[Calendar Year]").PivotItems("[Accounting Date].[CalendarYearMonthDay].[Calendar Year].&[2024]").DrilledDown = True
            End If
        Next pt
    Next ws
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error: " & Err.Description
    Resume Next

End Sub

Sub FileSelector()
    Dim WB As Workbook
    Dim FileDialog As FileDialog
    Dim SelectedFiles As FileDialogSelectedItems
    Dim FileName As Variant

    ' Create a File Dialog object
    Set FileDialog = Application.FileDialog(msoFileDialogFilePicker)
    FileDialog.Title = "Select files to update"

    ' Disable alerts and screen updating
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    On Error GoTo ErrorHandler

    ' Show the dialog box and get the selected files
    If FileDialog.Show = -1 Then
        Set SelectedFiles = FileDialog.SelectedItems
        If SelectedFiles.Count > 0 Then
            For Each FileName In SelectedFiles
                Set WB = Workbooks.Open(FileName)
                
                PivotDateExpander WB
                
                WB.Save
                WB.Close
            Next FileName
        Else
            MsgBox "No files selected.", vbInformation
        End If
    End If

    ' Re-enable alerts and screen updating
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Update Complete"

    Exit Sub

ErrorHandler:
    MsgBox "An error has occurred: " & Err.Description
    ' Re-enable alerts and screen updating
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

