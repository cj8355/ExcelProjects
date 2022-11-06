export const vbaCode = (`
    
Sub Table2_a()

'// Creating the worksheets

    Sheets.Add after:=Sheets(Sheets.Count)
    Sheets(ActiveSheet.Name).Name = "T2"
    Sheets.Add after:=Sheets(Sheets.Count)
    Sheets(ActiveSheet.Name).Name = "T3"
    Sheets.Add after:=Sheets(Sheets.Count)
    Sheets(ActiveSheet.Name).Name = "T4"
    Sheets.Add after:=Sheets(Sheets.Count)
    Sheets(ActiveSheet.Name).Name = "T5"
    Sheets.Add after:=Sheets(Sheets.Count)
    Sheets(ActiveSheet.Name).Name = "T6"


Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_vlookup As PivotField


On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T2")
wsTarget.Select
wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Sales")
With wsSource
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
    LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
    

'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("B5"), "T2A")
With PT
'// Show GrandTotals
    .ColumnGrand = True
    .RowGrand = True
    
    .RowAxisLayout xlTabularRow
    
    .TableStyle2 = "PivotStyleDark16"
    
    
    
    '// Add Pivot Fields
    
    '// Filters
    
    'If (ActiveSheet.PivotTables("T2A").PivotFields("vlookup").CurrentPage = "#N/A") Then
    
'    On Error Resume Next
'  pf_vlookup = ActiveSheet.PivotTable("T2A").PivotFields("vlookup").PivotItems("#N/A")
'  If Err = 0 Then ActiveSheet.PivotTables("T2A").PivotFields("vlookup").CurrentPage = "#N/A"
'  Else: ActiveSheet.PivotTables(T2A).PivotFields("vlookup").CurrentPage = "blank"
'    Err.Clear
'
'    End If
    
    With .PivotFields("vlookup")
    .Orientation = xlPageField
    .EnableMultiplePageItems = True
    End With

    Set pf_vlookup = PT.PivotFields("vlookup")

    pf_vlookup.ClearAllFilters

    '// Enable multiple filters selection
    pf_vlookup.EnableMultiplePageItems = True
    
    
    pf_vlookup.PivotItems("#N/A").Visible = False
    
    'Else
    
    
    'End If
    

    
    
    '// Rows Section
    
    With .PivotFields("Group")
    .Orientation = xlRowField
    .Subtotals(1) = False
    End With
    
    '// Columns
    'With .PivotFields("Group")
    '.Orientation = xlColumnField
    'End With
    
    '// Values
     With .PivotFields("Sales Units")
    .Orientation = xlDataField
    .Function = xlSum
    .NumberFormat = "#,###"
    End With




End With

CleanUp:
    Set PT = Nothing
    Set PTCache = Nothing
    Set SourceDataRange = Nothing
    Set wsSource = Nothing
    Set wsTarget = Nothing
    Set wb = Nothing
    Set pf_vlookup = Nothing
    
Exit Sub

errHandler:
    MsgBox "Error: " & Err.Description, vbExclamation
    GoTo CleanUp
End Sub



`)