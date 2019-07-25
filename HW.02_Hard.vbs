Sub Stock_Data_Hard()

For Each ws In Worksheets
ws.Activate

'Create Headers for Summary Tables
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
    
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    
'Set variables for finding the total volume, percent change, and yearly change
    Dim SumTabRow As Integer
    SumTabRow = 2
    Dim Ticker As String
    Dim VolTot, First, Last, YC, PC As Double
    VolTot = 0
    First = Cells(2, 3).Value
    
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
'Find the total volume, yearly change, and percent change for each stock
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
                'Format cells to be highlighted green for positive, red for negative, and light blue for 0
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
    
'set variables to find the greatest percent increase, greatest percent decrease, and greatest total volume
    Dim TickInc, TickDec, TickVol As String
    TickInc = Cells(2, 9).Value
    TickDec = Cells(2, 9).Value
    TickVol = Cells(2, 9).Value
    
    Dim GPI, GPD, GTV As Double
    GPI = Cells(2, 11).Value
    GPD = Cells(2, 11).Value
    GTV = Cells(2, 12).Value
    
'find the greatest percent increase
    For J = 2 To SumTabRow
        If GPI > Cells(J + 1, 11).Value Then
            GPI = GPI
        Else:
            GPI = Cells(J + 1, 11).Value
            TickInc = Cells(J + 1, 9).Value
        End If
        
        If GPD < Cells(J + 1, 11).Value Then
            GPD = GPD
        Else:
            GPD = Cells(J + 1, 11).Value
            TickDec = Cells(J + 1, 9).Value
        End If
        
        If GTV > Cells(J + 1, 12).Value Then
            GTV = GTV
        Else:
            GTV = Cells(J + 1, 12).Value
            TickVol = Cells(J + 1, 9).Value
        End If
    Next J
    
'Assign Greatest Values to summary table
    Cells(2, 15).Value = TickInc
    Cells(2, 16).Value = GPI
    Cells(3, 15).Value = TickDec
    Cells(3, 16).Value = GPD
    Cells(4, 15).Value = TickVol
    Cells(4, 16).Value = GTV
    
'Format cells
    Columns("K").NumberFormat = "0.00%"
    Range("P2:P3").NumberFormat = "0.00%"
    Columns("A:P").AutoFit
    
Next ws
    
End Sub

