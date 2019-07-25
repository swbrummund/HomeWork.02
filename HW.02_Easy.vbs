Sub Stock_Data_Easy()

For Each ws In Worksheets
ws.Activate
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Volume"
    
    Dim SumTabRow As Integer
    SumTabRow = 2
    Dim Ticker As String
    Dim VolTot As Double
    VolTot = 0
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For I = 2 To lastrow
    
        If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
            Ticker = Cells(I, 1).Value
            VolTot = VolTot + Cells(I, 7).Value
            Cells(SumTabRow, 9).Value = Ticker
            Cells(SumTabRow, 10).Value = VolTot
            SumTabRow = SumTabRow + 1
            VolTot = 0
        Else
            VolTot = VolTot + Cells(I, 7).Value
        End If
            
    Next I
    
    Columns("A:J").AutoFit
    
Next ws
    
End Sub

