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

    const Results = () => (
        <div id="results" className="search-results">
          <code>Sub AddNumbers()

Dim number As Integer, LoopCount As Integer
Dim GridRow As Integer, GridCol As Integer
Dim CellRow As Integer, CellCol As Integer
Sheets("Sudoku").Select

StartOver:
Range("C3:K11").Value = ""

For number = 1 To 9
For GridRow = 3 To 11 Step 3
For GridCol = 3 To 11 Step 3
LoopCount = 0

    Do
    Randomize
    CellRow = Int((GridRow + 2 - GridRow + 1) * Rnd + GridRow)
    CellCol = Int((GridCol + 2 - GridCol + 1) * Rnd + GridCol)
    
    If Cells(CellRow, CellCol) = "" Then
        If WorksheetFunction.CountIf(Rows(CellRow), number) = 0 And _
        WorksheetFunction.CountIf(Columns(CellCol), number) = 0 Then
        Cells(CellRow, CellCol) = number
        Exit Do
        End If
        End If
        
        LoopCount = LoopCount + 1
        If LoopCount > 99 Then GoTo StartOver
        Loop
        Next GridCol
        Next GridRow
        Next number

End Sub 
</code>
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