Sub Stock_Data_Moderate()

For Each ws In Worksheets
ws.Activate

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
    
    Dim SumTabRow As Integer
    SumTabRow = 2
    Dim Ticker As String
    Dim VolTot, First, Last, YC, PC As Double
    VolTot = 0
    First = Cells(2, 3).Value
    
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For I = 2 To lastrow
    
        If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
            Ticker = Cells(I, 1).Value
            VolTot = VolTot + Cells(I, 7).Value
            Last = Cells(I, 6).Value
            YC = First - Last
            If First <> 0 Then
                PC = (Last / First) - 1
            Else
                PC = 0
            End If
            
            Cells(SumTabRow, 9).Value = Ticker
            Cells(SumTabRow, 10).Value = YC
                If YC > 0 Then
                    Cells(SumTabRow, 10).Interior.ColorIndex = 4
                ElseIf YC < 0 Then
                    Cells(SumTabRow, 10).Interior.ColorIndex = 3
                Else: Cells(SumTabRow, 10).Interior.ColorIndex = 20
                End If
            Cells(SumTabRow, 11).Value = PC
            Cells(SumTabRow, 12).Value = VolTot
            
            First = Cells(I + 1, 3).Value
            SumTabRow = SumTabRow + 1
            VolTot = 0
        Else
            VolTot = VolTot + Cells(I, 7).Value
        End If
            
    Next I
    
    Columns("A:P").AutoFit
    
Next ws
    
End Sub



