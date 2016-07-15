Attribute VB_Name = "AutoOpen"
Sub Auto_Open()
    'Call CountVisibleWorksheets
    'Call AddNewWorkbookIfNoneExist
End Sub

'Private Sub Workbook_Activate()
'
'    Call CountVisibleWorksheets
'    Call AddNewWorkbookIfNoneExist
'
'End Sub

Private Function CountVisibleWorksheets() As Integer
    'Return number of non-hidden workbooks
    For Each Sheet In ThisWorkbook.Sheets
        If Sheet.visible = True Then i = i + 1
        Debug.Print "Sheet: " & Sheet.Name
    Next
    CountVisibleWorksheets = i
End Function

Private Function AddNewWorkbookIfNoneExist()
    HIDDENCOUNT = 3
    Dim wb As Workbook
    For Each wb In Workbooks
        Debug.Print ("wb.Name: " & wb.Name)
        If InStr(wb.Name, "xlsx") Then xlsx = True 'If xlsx is found, then assume user doesn't need a blank book1
        If wb.Name = "Book1" Then blankBookExists = True
        wbcount = wbcount + 1
    Next wb
    'If Not xlsx And Not blankBookExists And wbcount < 4 Then Workbooks.Add 'Doesn't work, sheet opens too late
    Debug.Print ("wbcount: " & Str(wbcount))
End Function

'WorkbooksOpenCount = CountVisibleWorkbooks
'If WorkbooksOpenCount < 2 Then Workbooks.Add

'Debug.Print WorkbooksOpenCount


