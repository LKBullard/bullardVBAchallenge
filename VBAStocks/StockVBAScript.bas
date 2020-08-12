Attribute VB_Name = "Module1"
Sub Stocks():

    ' Declare all the variables being used
    Dim Ticker As String
    Dim TotalVol As LongLong
    Dim OpenP As Double
    Dim CloseP As Double
    Dim PercentChange As Double
    ' These variables are specifically for Challenge 1
    Dim GreatestIncTicker As String
    Dim GreatestInc As Double
    Dim GreatestDecTicker As String
    Dim GreatestDec As Double
    Dim GreatestTotTicker As String
    Dim GreatestTot As LongLong
    
    ' Looping through each Worksheet
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Set the headers of the new columns being created
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        ' This sets the initial Ticker and Open Price for each worksheet
        TickerCount = 1
        Ticker = ws.Cells(2, 1).Value
        OpenP = ws.Cells(2, 3).Value

        ' Loop through every row
        For i = 2 To LastRow
            ' Checks if the ticker in the next row is different
           If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                ' Updates the variables before outputing
                TickerCount = TickerCount + 1
                CloseP = ws.Cells(i, 6).Value
                TotalVol = TotalVol + ws.Cells(i, 7).Value
                PercentChange = (CloseP - OpenP) / (OpenP)
                
                ' Outputs the information into the new columns
                ws.Cells(TickerCount, 9).Value = Ticker
                ws.Cells(TickerCount, 10).Value = CloseP - OpenP
                ws.Cells(TickerCount, 11).Value = PercentChange
                ws.Cells(TickerCount, 12).Value = TotalVol
                
                ' Formats Yearly Change by color and Percent Change to percents
                If CloseP - OpenP < 0 Then
                    ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
                ElseIf CloseP - OpenP > 0 Then
                    ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
                End If
                ws.Cells(TickerCount, 11).NumberFormat = "0.00%"
                
                ' Conditionals for Challenge 1
                If PercentChange > GreatestInc Then
                    GreatestIncTicker = Ticker
                    GreatestInc = PercentChange
                End If
                If PercentChange < GreatestDec Then
                    GreatestDecTicker = Ticker
                    GreatestDec = PercentChange
                End If
                If TotalVol > GreatestTot Then
                    GreatestTotTicker = Ticker
                    GreatestTot = TotalVol
                End If
                
                ' Resets the variables for the next loop
                Ticker = ws.Cells(i + 1, 1).Value
                OpenP = ws.Cells(i + 1, 3).Value
                TotalVol = 0
            Else
                TotalVol = TotalVol + ws.Cells(i, 7).Value
            End If
            
        Next i
        
        ' Returns the greatest increase, decrease, and total from all stocks
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        ws.Range("O2").Value = GreatestIncTicker
        ws.Range("O3").Value = GreatestDecTicker
        ws.Range("O4").Value = GreatestTotTicker
        ws.Range("P2").Value = GreatestInc
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P3").Value = GreatestDec
        ws.Range("P3").NumberFormat = "0.00%"
        ws.Range("P4").Value = GreatestTot
        
        GreatestInc = 0
        GreatestDec = 0
        GreatestTot = 0
        
    Next ws
    
    MsgBox ("All done.")
End Sub

