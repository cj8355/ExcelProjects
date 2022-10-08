import './vba-PT.scss';
import VBAImg from './vba.png';
import * as React from 'react';
import ExcelFile from './vba-PT.xlsm';
import vbaVid from './vba-vid.webm';

export default function VBA_PT() { 

    const [showResults, setShowResults] = React.useState(false);
    const onClick = () => setShowResults(!showResults);






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