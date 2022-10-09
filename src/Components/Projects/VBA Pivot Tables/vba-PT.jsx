import './vba-PT.scss';
import VBAImg from './vba.png';
import * as React from 'react';
import ExcelFile from './vba-PT.xlsm';
import vbaVid from './vba-vid.webm';

export default function VBA_PT() { 

    const [showResults, setShowResults] = React.useState(false);
    const onClick = () => setShowResults(!showResults);




    const Code3 = `
    <div>
    <code> 
    Sub test2_b()

    Dim wb As Workbook
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim LastRow As Long, LastColumn As Long
    Dim SourceDataRange As Range
    Dim PTCache As PivotCache
    Dim PT As PivotTable
    Dim pf_vlookup As PivotField
    Dim pf_duplicate As PivotField
    Dim vlookupArray(1) As String
    Dim numberOfElements As Integer
    Dim i As Integer
    Dim j As Integer
    
    vlookupArray(1) = "#N/A"
    
    
    On Error GoTo errHandler
    Set wb = ActiveWorkbook
    Set wsTarget = wb.Worksheets("T2")
    wsTarget.Select
    'wsTarget.Cells.Clear
    
    
    '// Step 1. Define data soruce
    Set wsSource = wb.Worksheets("Complaints")
    With wsSource
        LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
        LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
        
    
    '// Define source data object
    Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
    End With
    
    '// 2. Set PT Cache
    Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)
    
    '// 3. Create Pivot Table
    Set PT = PTCache.CreatePivotTable(wsTarget.Range("G5"), "T2B")
    With PT
    '// Show GrandTotals
        .ColumnGrand = True
        .RowGrand = True
        
        .RowAxisLayout xlTabularRow
        
        .TableStyle2 = "PivotStyleDark16"
        
        
        
        '// Add Pivot Fields
        
        '// Filters
        With .PivotFields("vlookup")
        .Orientation = xlPageField
    '    .EnableMultiplePageItems = True
        End With
        
        With .PivotFields("Dup?")
        .Orientation = xlPageField
        .EnableMultiplePageItems = True
        End With
        
        Set pf_vlookup = PT.PivotFields("vlookup")
        Set pf_duplicate = PT.PivotFields("Dup?")
        
    '    pf_vlookup.ClearAllFilters
        pf_duplicate.ClearAllFilters
        
        '// Enable multiple filters selection
        pf_vlookup.EnableMultiplePageItems = True
        pf_duplicate.EnableMultiplePageItems = True
            
    '    pf_vlookup.PivotItems("#N/A").Visible = False
        pf_duplicate.CurrentPage = ""
            
               '// only apply filter if present in data
       
        numberOfElements = UBound(vlookupArray) - LBound(vlookupArray) + 1
    
    If numberOfElements > 0 Then
        With pf_vlookup
            For i = 1 To pf_vlookup.PivotItems.Count
            j = 0
    '            MsgBox pf_vlookup.PivotItems.Count
    '            MsgBox pf_vlookup.PivotItems(i).Name
    '            MsgBox vlookupArray(j)
                
            Do While j < numberOfElements
                
                
                If pf_vlookup.PivotItems(i).Name = vlookupArray(j) Then
                
                    pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = False
                    Exit Do
                Else
                    pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = True
                End If
                j = j + 1
            Loop
            Next i
        End With
    End If
        
        
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
         With .PivotFields("Complaint ID")
        .Orientation = xlDataField
        .Function = xlCount
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
        Set pf_duplicate = Nothing
        
    Exit Sub
    
    errHandler:
        MsgBox "Error: " & Err.Description, vbExclamation
        GoTo CleanUp
    End Sub

    Sub test2_c()

Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_vlookup As PivotField
Dim pf_arCode As PivotField
Dim pf_ARC As PivotField
Dim vlookupArray(1) As String
Dim arcArray(0 To 10) As String
Dim numberOfElements As Integer
Dim numberOfElementsTwo As Integer
Dim i As Integer
Dim j As Integer
Dim c As Integer
Dim d As Integer


vlookupArray(1) = "#N/A"

arcArray(0) = "No Consequences or Impact to Patient"
arcArray(1) = "No Known Impact Or Consequence To Patient"
arcArray(2) = ""
arcArray(3) = "Device No Known Device Problem"
arcArray(4) = "Device No Reported Allegation"
arcArray(5) = "Insufficient Information"
arcArray(6) = "No Clinical Signs, Symptoms or Conditions"
arcArray(7) = "No Code Available"
arcArray(8) = "No Health Consequences or Impact"
arcArray(9) = "No Information"
arcArray(10) = "No Patient Involvement"


On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T2")
wsTarget.Select
'wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Complaints")
With wsSource
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
    LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
    

'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("K5"), "T2C")
With PT
'// Show GrandTotals
    .ColumnGrand = True
    .RowGrand = True
    
    .RowAxisLayout xlTabularRow
    
    .TableStyle2 = "PivotStyleDark16"
    
    
    
    '// Add Pivot Fields
    
    '// Filters
    With .PivotFields("vlookup")
    .Orientation = xlPageField
'    .EnableMultiplePageItems = True
    End With
    
'    With .PivotFields("AR Code Description (GCMS)")
'    .Orientation = xlPageField
'    .EnableMultiplePageItems = True
'    End With
    
    With .PivotFields("AR Code Description (GCMS)2")
    .Orientation = xlPageField
'    .EnableMultiplePageItems = True
    End With
    
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
     With .PivotFields("Complaint ID")
    .Orientation = xlDataField
    .Function = xlCount
    .NumberFormat = "#,###"
    End With
    
        '// ARC filter
    Set pf_ARC = PT.PivotFields("AR Code Description (GCMS)2")
    
   
    Set pf_vlookup = PT.PivotFields("vlookup")
'    Set pf_arCode = PT.PivotFields("AR Code Description (GCMS)")
    
    
    pf_vlookup.ClearAllFilters
'    pf_arCode.ClearAllFilters
    
    '// Enable multiple filters selection
    
    pf_vlookup.EnableMultiplePageItems = True
'    pf_arCode.EnableMultiplePageItems = True
    
       pf_ARC.EnableMultiplePageItems = True
   

   
   '// only apply filter if present in data
   
    numberOfElementsTwo = UBound(arcArray) - LBound(arcArray) + 1
'    MsgBox numberOfElementsTwo

If numberOfElementsTwo > 0 Then
    With pf_ARC
        For c = 1 To pf_ARC.PivotItems.Count

        d = 0
'            MsgBox pf_ARC.PivotItems.Count
'            MsgBox "Arc Array" + pf_ARC.PivotItems(c).Name
'            MsgBox d
            
        Do While d < numberOfElementsTwo
            
            

            If pf_ARC.PivotItems(c).Name = arcArray(d) Then
            
                pf_ARC.PivotItems(pf_ARC.PivotItems(c).Name).Visible = False
                Exit Do
            Else
                pf_ARC.PivotItems(pf_ARC.PivotItems(c).Name).Visible = True
            End If
            d = d + 1
        Loop
        Next c
    End With
End If

               '// only apply filter if present in data
   
    numberOfElements = UBound(vlookupArray) - LBound(vlookupArray) + 1

If numberOfElements > 0 Then
    With pf_vlookup
        For i = 1 To pf_vlookup.PivotItems.Count
        j = 0
'            MsgBox pf_vlookup.PivotItems.Count
'            MsgBox pf_vlookup.PivotItems(i).Name
'            MsgBox vlookupArray(j)
            
        Do While j < numberOfElements
            
            
            If pf_vlookup.PivotItems(i).Name = vlookupArray(j) Then
            
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = False
                Exit Do
            Else
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = True
            End If
            j = j + 1
        Loop
        Next i
    End With
End If
    
    
    
'    pf_vlookup.PivotItems("#N/A").Visible = False
'    pf_arCode.PivotItems("").Visible = False
'    pf_arCode.PivotItems("No Patient Involvement").Visible = False
'    pf_arCode.PivotItems("No Health Consequences or Impact").Visible = False
'    pf_arCode.PivotItems("Insufficient Information").Visible = False
'    pf_arCode.PivotItems("No Known Impact Or Consequence To Patient").Visible = False
'    pf_arCode.PivotItems("No Consequences or Impact to Patient").Visible = False
'    pf_arCode.PivotItems("No Clinical Signs, Symptoms or Conditions").Visible = False
'    pf_arCode.PivotItems("Device No Reported Allegation").Visible = False
'    pf_arCode.PivotItems("Device No Known Device Problem").Visible = False
    
    
    
    





End With

CleanUp:
    Set PT = Nothing
    Set PTCache = Nothing
    Set SourceDataRange = Nothing
    Set wsSource = Nothing
    Set wsTarget = Nothing
    Set wb = Nothing
    Set pf_vlookup = Nothing
    Set pf_arCode = Nothing
    
Exit Sub

errHandler:
    MsgBox "Error: " & Err.Description, vbExclamation
    GoTo CleanUp
End Sub

Sub test2_d()

Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_globalRei As PivotField
Dim pf_vlookup As PivotField
Dim pf_duplicate As PivotField
Dim vlookupArray(1) As String
Dim numberOfElements As Integer
Dim i As Integer
Dim j As Integer

vlookupArray(1) = "#N/A"


On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T2")
wsTarget.Select
'wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Complaints")
With wsSource
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
    LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
    

'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("B35"), "T2D")
With PT
'// Show GrandTotals
    .ColumnGrand = True
    .RowGrand = True
    
    .RowAxisLayout xlTabularRow
    
    .TableStyle2 = "PivotStyleDark16"
    
    
    
    '// Add Pivot Fields
    
    '// Filters
    With .PivotFields("vlookup")
    .Orientation = xlPageField
'    .EnableMultiplePageItems = True
    End With
    
    With .PivotFields("Dup?")
    .Orientation = xlPageField
    .EnableMultiplePageItems = True
    End With
    
    With .PivotFields("Global REI")
    .Orientation = xlPageField
    .EnableMultiplePageItems = True
    End With
    
    
    Set pf_globalRei = PT.PivotFields("Global REI")
    Set pf_vlookup = PT.PivotFields("vlookup")
    Set pf_duplicate = PT.PivotFields("Dup?")
    
    pf_globalRei.ClearAllFilters
    pf_vlookup.ClearAllFilters
    pf_duplicate.ClearAllFilters
    
    '// Enable multiple filters selection
    pf_globalRei.EnableMultiplePageItems = True
    pf_vlookup.EnableMultiplePageItems = True
    pf_duplicate.EnableMultiplePageItems = True
    
    
    pf_globalRei.CurrentPage = "Yes"
'    pf_vlookup.PivotItems("#N/A").Visible = False
    pf_duplicate.CurrentPage = ""
    
               '// only apply filter if present in data
   
    numberOfElements = UBound(vlookupArray) - LBound(vlookupArray) + 1

If numberOfElements > 0 Then
    With pf_vlookup
        For i = 1 To pf_vlookup.PivotItems.Count
        j = 0
'            MsgBox pf_vlookup.PivotItems.Count
'            MsgBox pf_vlookup.PivotItems(i).Name
'            MsgBox vlookupArray(j)
            
        Do While j < numberOfElements
            
            
            If pf_vlookup.PivotItems(i).Name = vlookupArray(j) Then
            
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = False
                Exit Do
            Else
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = True
            End If
            j = j + 1
        Loop
        Next i
    End With
End If
    
    
    
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
     With .PivotFields("Complaint ID")
    .Orientation = xlDataField
    .Function = xlCount
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
    Set pf_globalRei = Nothing
    Set pf_vlookup = Nothing
    Set pf_duplicate = Nothing
    
Exit Sub

errHandler:
    MsgBox "Error: " & Err.Description, vbExclamation
    GoTo CleanUp
End Sub


    </code></div>
    
    `
    

    const Results = () => (
        <div id="results" className="search-results">
            <div>
          <code>
          Sub Run_All2()<br/><br/>

Call Table2_a.Table2_a<br/>
Call test2_b.test2_b<br/>
Call test2_c.test2_c<br/>
Call test2_d.test2_d<br/><br/>

Call test3_a1.test3_a1<br/>
Call Table3_a2.Table3_a2<br/>
Call test3_b1.test3_b1<br/>
Call Table3_b2.Table3_b2<br/><br/>

Call test_4A.test_4A<br/>
Call test_4B.test_4B<br/><br/>

Call test5_a1.test5_a1<br/>
Call test5_a2.test5_a2<br/>
Call test5_a3.test5_a3<br/>
Call test5_b.test5_b<br/><br/>

Call test6_a.test6_a<br/>
Call test6_b.test6_b<br/><br/>



End Sub<br/>
-------------------<br/>

Sub Table2_a()<br/><br/>

'// Creating the worksheets<br/><br/>

    Sheets.Add after:=Sheets(Sheets.Count)<br/>
    Sheets(ActiveSheet.Name).Name = "T2"<br/>
    Sheets.Add after:=Sheets(Sheets.Count)<br/>
    Sheets(ActiveSheet.Name).Name = "T3"<br/>
    Sheets.Add after:=Sheets(Sheets.Count)<br/>
    Sheets(ActiveSheet.Name).Name = "T4"<br/>
    Sheets.Add after:=Sheets(Sheets.Count)<br/>
    Sheets(ActiveSheet.Name).Name = "T5"<br/>
    Sheets.Add after:=Sheets(Sheets.Count)<br/>
    Sheets(ActiveSheet.Name).Name = "T6"<br/><br/>


Dim wb As Workbook<br/>
Dim wsSource As Worksheet, wsTarget As Worksheet<br/>
Dim LastRow As Long, LastColumn As Long<br/>
Dim SourceDataRange As Range<br/>
Dim PTCache As PivotCache<br/>
Dim PT As PivotTable<br/>
Dim pf_vlookup As PivotField<br/><br/>


On Error GoTo errHandler<br/>
Set wb = ActiveWorkbook<br/>
Set wsTarget = wb.Worksheets("T2")<br/>
wsTarget.Select<br/>
wsTarget.Cells.Clear<br/><br/>


'// Step 1. Define data soruce<br/>
Set wsSource = wb.Worksheets("Sales")<br/>
With wsSource<br/>
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row<br/>
    LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column<br/><br/>
    

'// Define source data object<br/>
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))<br/>
End With<br/><br/>

'// 2. Set PT Cache<br/>
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)<br/><br/>

'// 3. Create Pivot Table<br/>
Set PT = PTCache.CreatePivotTable(wsTarget.Range("B5"), "T2A")<br/>
With PT<br/>
'// Show GrandTotals<br/>
    .ColumnGrand = True<br/>
    .RowGrand = True<br/><br/>
    
    .RowAxisLayout xlTabularRow<br/><br/>
    
    .TableStyle2 = "PivotStyleDark16"<br/><br/>
    
    
    
    '// Add Pivot Fields<br/><br/>
    
    '// Filters<br/><br/>
    
    'If (ActiveSheet.PivotTables("T2A").PivotFields("vlookup").CurrentPage = "#N/A") Then<br/><br/>
    
'    On Error Resume Next<br/>
'  pf_vlookup = ActiveSheet.PivotTable("T2A").PivotFields("vlookup").PivotItems("#N/A")<br/>
'  If Err = 0 Then ActiveSheet.PivotTables("T2A").PivotFields("vlookup").CurrentPage = "#N/A"<br/>
'  Else: ActiveSheet.PivotTables(T2A).PivotFields("vlookup").CurrentPage = "blank"<br/>
'    Err.Clear<br/><br/>
'
'    End If<br/><br/>
    
    With .PivotFields("vlookup")<br/>
    .Orientation = xlPageField<br/>
    .EnableMultiplePageItems = True<br/>
    End With<br/><br/>

    Set pf_vlookup = PT.PivotFields("vlookup")<br/><br/>

    pf_vlookup.ClearAllFilters<br/><br/>

    '// Enable multiple filters selection<br/>
    pf_vlookup.EnableMultiplePageItems = True<br/><br/>
    
    
    pf_vlookup.PivotItems("#N/A").Visible = False<br/><br/>
    
    'Else<br/>
    
    
    'End If<br/><br/>
    

    
    
    '// Rows Section<br/><br/>
    
    With .PivotFields("Group")<br/>
    .Orientation = xlRowField<br/>
    .Subtotals(1) = False<br/>
    End With<br/><br/>
    
    '// Columns<br/>
    'With .PivotFields("Group")<br/>
    '.Orientation = xlColumnField<br/>
    'End With<br/><br/>
    
    '// Values<br/>
     With .PivotFields("Sales Units")<br/>
    .Orientation = xlDataField<br/>
    .Function = xlSum<br/>
    .NumberFormat = "#,###"<br/>
    End With<br/><br/>




End With<br/><br/>

CleanUp:<br/>
    Set PT = Nothing<br/>
    Set PTCache = Nothing<br/>
    Set SourceDataRange = Nothing<br/>
    Set wsSource = Nothing<br/>
    Set wsTarget = Nothing<br/>
    Set wb = Nothing<br/>
    Set pf_vlookup = Nothing<br/><br/>
    
Exit Sub<br/><br/>

errHandler:<br/>
    MsgBox "Error: " & Err.Description, vbExclamation<br/>
    GoTo CleanUp<br/>
End Sub<br/><br/>
-------------------<br/>



<Code3/>
    


</code>
</div>


        </div>
      )



    return (
        <div className='vbaContainer'>
            <img src={VBAImg}></img>

            <div className='textContainer'>
            <h2>VBA Pivot Tables</h2>
            <p>Automatically create Pivot Tables using VBA.</p>

            <div className='codeContainer'>
            <input type="submit" value="View Code" onClick={onClick} />
            { showResults ? <Results /> : null }
            <a href={ExcelFile} target="_blank">Download</a>
            <video src={vbaVid} height="250" width="250" controls></video>

            </div>
            

            </div>
        </div>
    )

}