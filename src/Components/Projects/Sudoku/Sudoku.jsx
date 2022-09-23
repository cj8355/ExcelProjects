import './sudoku.scss';
import sudokuImg from './sudoku.png';
import * as React from 'react';
import ExcelFile from './sudoku-game.xlsm';
import sudokuVid from './sudoku-vid.webm';

export default function Sudoku() {

    const [showResults, setShowResults] = React.useState(false);
    const [showGif, setShowGif] = React.useState(true);

    const onClick = () => setShowResults(!showResults);
    const onClick2 = () => setShowGif(!showGif);

    const code1 = `
        Sub PlayGame()

Dim number As Integer, LoopCount As Integer
Dim CellRow As Integer, CellCol As Integer
Dim uCell As Integer, Comp As Boolean

Do
    Randomize
    CellRow = Int((11 - 3 + 1) * Rnd + 2)
    CellCol = Int((11 - 3 + 1) * Rnd + 2)
    number = Val(Cells(CellRow, CellCol).Value)
    If number <> 0 Then
        Comp = False
        For uCell = 3 To 11
        'Rows
            If Cells(uCell, CellCol).Value = "" Then
                If WorksheetFunction.CountIf(Rows(uCell), num) = 0 And _
                WorksheetFunction.CountIf(Range(qRng(uCell, CellCol)), num) = 0 Then Comp = True
                End If
                
        'Cols
            If Cells(CellRow, uCell).Value = "" Then
                If WorksheetFunction.CountIf(Columns(uCell), num) = 0 And _
                WorksheetFunction.CountIf(Range(qRng(CellRow, uCell)), num) = 0 Then Comp = True
                End If
            Next uCell
            
            If Comp = False Then
                Cells(CellRow, CellCol).Value = ""
                End If
            End If
            
            LoopCount = LoopCount + 1
            If LoopCount > 299 Then Exit Do
            Loop

End Sub

Function qRng(r As Integer, c As Integer) As String
    If c < 6 Then
        If r < 6 Then
            qRng = "C3:E5"
        ElseIf r < 9 Then
            qRng = "C6:E8"
        Else
            qRng = "C9:E11"
        End If
        
    ElseIf c < 9 Then
            If r < 6 Then
                qRng = "F3:H5"
            ElseIf r < 9 Then
            qRng = "F6:H8"
        Else
            qRng = "F9:H11"
        End If
    Else
            If r < 5 Then
                qRng = "I3:K5"
            ElseIf r < 8 Then
            qRng = "I6:K8"
        Else
            qRng = "I9:K11"
            End If
        End If

End Function

    `

    const Results = () => (
        <div id="results" className="search-results">
            <div>
          <code><pre>Sub AddNumbers()<br/><br/>

Dim number As Integer, LoopCount As Integer<br/>
Dim GridRow As Integer, GridCol As Integer<br/>
Dim CellRow As Integer, CellCol As Integer<br/>
Sheets("Sudoku").Select<br/><br/>

StartOver:<br/>
Range("C3:K11").Value = ""<br/><br/>

For number = 1 To 9<br/>
For GridRow = 3 To 11 Step 3<br/>
For GridCol = 3 To 11 Step 3<br/>
LoopCount = 0<br/><br/>

    Do<br/>
    Randomize<br/>
    CellRow = Int((GridRow + 2 - GridRow + 1) * Rnd + GridRow)<br/>
    CellCol = Int((GridCol + 2 - GridCol + 1) * Rnd + GridCol)<br/><br/>
    
    If Cells(CellRow, CellCol) = "" Then<br/>
        If WorksheetFunction.CountIf(Rows(CellRow), number) = 0 And _<br/>
        WorksheetFunction.CountIf(Columns(CellCol), number) = 0 Then<br/>
        Cells(CellRow, CellCol) = number<br/>
        Exit Do<br/>
        End If<br/>
        End If<br/><br/>
        
        LoopCount = LoopCount + 1<br/>
        If LoopCount > 99 Then GoTo StartOver<br/>
        Loop<br/>
        Next GridCol<br/>
        Next GridRow<br/>
        Next number<br/><br/>

End Sub <br/>
----------------------<br/>
</pre>
</code>
</div>

<div>
<code><pre>
    
Sub CreateLayout()<br/><br/>

Sheets("Sudoku").Select<br/><br/>

Cells.Clear<br/><br/>

With Range("C3:K11")<br/>
    .ColumnWidth = 10<br/>
    .RowHeight = 29<br/>
    .Font.Size = 20<br/>
    .HorizontalAlignment = xlCenter<br/>
    .VerticalAlignment = xlCenter<br/>
    .Borders.LineStyle = xlContinuous<br/>
    .Borders.Weight = xlThin<br/>
    End With<br/><br/>
    
    Range("F3:H5").Borders(xlEdgeBottom).LineStyle = xlContinuous<br/>
    Range("F3:H5").Borders(xlEdgeBottom).Weight = xlThick<br/>
    Range("F3:H5").Borders(xlEdgeRight).LineStyle = xlContinuous<br/>
    Range("F3:H5").Borders(xlEdgeRight).Weight = xlThick<br/><br/>
        
   Range("I3:K5").Borders(xlEdgeBottom).LineStyle = xlContinuous<br/>
    Range("I3:K5").Borders(xlEdgeBottom).Weight = xlThick<br/><br/>
    
    Range("C6:E8").Borders(xlEdgeBottom).LineStyle = xlContinuous<br/>
    Range("C6:E8").Borders(xlEdgeBottom).Weight = xlThick<br/>
    Range("C6:E8").Borders(xlEdgeRight).LineStyle = xlContinuous<br/>
    Range("C6:E8").Borders(xlEdgeRight).Weight = xlThick<br/><br/>
    
    Range("F6:H8").Borders(xlEdgeBottom).LineStyle = xlContinuous<br/>
    Range("F6:H8").Borders(xlEdgeBottom).Weight = xlThick<br/>
    Range("F6:H8").Borders(xlEdgeRight).LineStyle = xlContinuous<br/>
    Range("F6:H8").Borders(xlEdgeRight).Weight = xlThick<br/><br/>
    
    Range("C3:E5").Borders(xlEdgeBottom).LineStyle = xlContinuous<br/>
    Range("C3:E5").Borders(xlEdgeBottom).Weight = xlThick<br/>
    Range("C3:E5").Borders(xlEdgeRight).LineStyle = xlContinuous<br/>
    Range("C3:E5").Borders(xlEdgeRight).Weight = xlThick<br/><br/>
    
    Range("I6:K8").Borders(xlEdgeBottom).LineStyle = xlContinuous<br/>
    Range("I6:K8").Borders(xlEdgeBottom).Weight = xlThick<br/><br/>
    
    Range("C9:E11").Borders(xlEdgeRight).LineStyle = xlContinuous<br/>
    Range("C9:E11").Borders(xlEdgeRight).Weight = xlThick<br/><br/>
    
    Range("F9:H11").Borders(xlEdgeRight).LineStyle = xlContinuous<br/>
    Range("F9:H11").Borders(xlEdgeRight).Weight = xlThick<br/><br/><br/>


End Sub<br/>
----------------------<br/>
</pre>
</code>
</div>

<div className='codeOne'>
<pre><code>

{code1}

</code>
</pre>
</div>
        </div>
      )

      const SudokuGif = () => {
        <div id="gif" className='gif-box'>
           
      <video src={sudokuVid} height="300" width="300" controls></video>
     
    
        </div>
      }
  
    return (
        <div className="SudokuContainer">
            <img src={sudokuImg}></img>

            <div className='textContainer'>
            <h2>Sudoku</h2>
            <p>Sudoku game built in Excel using VBA.</p>

            <div className='codeContainer'>
      <input type="submit" value="View Code" onClick={onClick} />
      { showResults ? <Results /> : null }
      <a href={ExcelFile} target="_blank">Download</a>
      <video src={sudokuVid} height="250" width="250" controls></video>
            
    </div>
    {/* <input type="submit" value="Gif" onClick={onClick2} />
      { showGif ? <SudokuGif /> : null } */}
            
            </div>
            
            
        </div>
        
    )
}