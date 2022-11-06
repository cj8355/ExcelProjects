export const code = ( `
Sub CreateLayout()

Sheets("Sudoku").Select

Cells.Clear

With Range("C3:K11")
    .ColumnWidth = 10
    .RowHeight = 29
    .Font.Size = 20
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
    End With
    
    Range("F3:H5").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("F3:H5").Borders(xlEdgeBottom).Weight = xlThick
    Range("F3:H5").Borders(xlEdgeRight).LineStyle = xlContinuous
    Range("F3:H5").Borders(xlEdgeRight).Weight = xlThick
        
   Range("I3:K5").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("I3:K5").Borders(xlEdgeBottom).Weight = xlThick
    
    Range("C6:E8").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("C6:E8").Borders(xlEdgeBottom).Weight = xlThick
    Range("C6:E8").Borders(xlEdgeRight).LineStyle = xlContinuous
    Range("C6:E8").Borders(xlEdgeRight).Weight = xlThick
    
    Range("F6:H8").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("F6:H8").Borders(xlEdgeBottom).Weight = xlThick
    Range("F6:H8").Borders(xlEdgeRight).LineStyle = xlContinuous
    Range("F6:H8").Borders(xlEdgeRight).Weight = xlThick
    
    Range("C3:E5").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("C3:E5").Borders(xlEdgeBottom).Weight = xlThick
    Range("C3:E5").Borders(xlEdgeRight).LineStyle = xlContinuous
    Range("C3:E5").Borders(xlEdgeRight).Weight = xlThick
    
    Range("I6:K8").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("I6:K8").Borders(xlEdgeBottom).Weight = xlThick
    
    Range("C9:E11").Borders(xlEdgeRight).LineStyle = xlContinuous
    Range("C9:E11").Borders(xlEdgeRight).Weight = xlThick
    
    Range("F9:H11").Borders(xlEdgeRight).LineStyle = xlContinuous
    Range("F9:H11").Borders(xlEdgeRight).Weight = xlThick


End Sub
`
)