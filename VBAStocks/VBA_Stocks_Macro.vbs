Sub VBA_Stocks_Macro()

    ' Define variables:
    ' ticker is a variable to hold the ticker symbol.
    ' openPrice is a variable to hold the opening price at the beginning of a given year.
    ' closingPrice is a variable to hold the closing price at the end of a given year.
    ' yearChange is a variable to hold the difference between closing and opening prices.
    ' percentChange is a variable to hold the percent change value from the opening/closing prices.
    ' greatestInc is a variable to hold the greatest percent increase in percent change.
    ' greatestDec is a variable to hold the greatest percent decrease in percent change.
    ' greatestTot is a variable to hold the greatest total volume from total stock volume.
    ' totalVolume is a variable to hold the total stock volume.
    ' rowCounter is a variable that will go down to the next row as it fills in the data table.
    
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearChange As Double
    Dim percentChange As Double
    Dim greatestInc As Double
    Dim greatestDec As Double
    Dim greatestTot As LongLong
    Dim totalVolume As LongLong
    Dim rowCounter As Integer
    
    '-------------------------'
    ' LOOP THROUGH ALL SHEETS '
    '-------------------------'
    For Each ws In Worksheets
    
        ' Initialize totalVolume with 0.
        totalVolume = 0
    
        ' Initialize row counter with 2 because row 1 has headers.
        rowCounter = 2

        ' Find the last row in the worksheet.
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Create the data table headers.
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Set initial value for the opening price.
        openPrice = ws.Range("C2").Value
        
        '-----------------------'
        ' LOOP THROUGH EACH ROW '
        '-----------------------'
        For x = 2 To lastrow
        
            '----------------------'
            ' CALCULATE THE VALUES '
            '----------------------'
        
            ' If the ticker in the following row is not the same as the current row, do the following:
            If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then

                ' Set the ticker name.
                ticker = ws.Cells(x, 1).Value
                
                ' Find the closing price.
                closePrice = ws.Cells(x, 6).Value

                ' Calculate the yearly price change.
                yearChange = closePrice - openPrice
                
                    ' If the open price is 0, do the following:
                    If openPrice = 0 Then
                    
                        ' Set percent change to 0.
                        percentChange = 0
                        
                        ' Get the next value for open price.
                        openPrice = ws.Range("C" & x + 1).Value
                        
                    ' If the following row is not empty and the open price is not zero, do the following:
                    Else
                        
                        ' Calculate the percent change.
                        percentChange = yearChange / openPrice
                        
                        ' Get the next value for open price.
                        openPrice = ws.Range("C" & x + 1).Value

                    End If
                
                '------------'
                ' DATA TABLE '
                '------------'
                
                ' Print the ticker name.
                ws.Range("I" & rowCounter).Value = ticker
                
                    ' If the year change is positive, do the following:
                    If yearChange >= 0 Then
                
                        ' Print the percent change.
                        ws.Range("J" & rowCounter).Value = yearChange
                    
                        ' Make the cell color green.
                        ws.Range("J" & rowCounter).Interior.ColorIndex = 4
                    
                    ' If the year change is negative, do the following:
                    Else
                    
                        ' Print the percent change.
                        ws.Range("J" & rowCounter).Value = yearChange
                    
                        ' Make the cell color red.
                        ws.Range("J" & rowCounter).Interior.ColorIndex = 3
                        
                    End If
                    
                ' Print the percent change.
                ws.Range("K" & rowCounter).Value = percentChange
                    
                ' Add row to the total stock volume.
                totalVolume = totalVolume + ws.Cells(x, 7).Value
                
                ' Print the total stock volume.
                ws.Range("L" & rowCounter).Value = totalVolume
                
                ' Add one to the row counter.
                rowCounter = rowCounter + 1
                
                ' Reset the total volume.
                totalVolume = 0
            
            ' If the ticker is still the same, do the following:
            Else
                ' Add each row to the total stock volume.
                totalVolume = totalVolume + ws.Cells(x, 7).Value

            End If
            
        Next x
        
        '-----------------------'
        ' GREATEST CHANGE TABLE '
        '-----------------------'

        ' Create the greatest change table headers.
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' Find the last row in the percent change column.
        lastRowGreatest = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        ' Find the greatest percent increase, get the ticker value, and print them to correct cells.
        greatestInc = WorksheetFunction.Max(ws.Range("K2:K" & lastRowGreatest))
        greatestIncIndex = WorksheetFunction.Match(greatestInc, ws.Range("K2:K" & lastRowGreatest), 0)
        ws.Range("P2").Value = ws.Cells(greatestIncIndex + 1, 9).Value
        ws.Range("Q2").Value = greatestInc
        
        ' Find the greatest percent decrease, get the ticker value, and print them to correct cells.
        greatestDec = WorksheetFunction.Min(ws.Range("K2:K" & lastRowGreatest))
        greatestDecIndex = WorksheetFunction.Match(greatestDec, ws.Range("K2:K" & lastRowGreatest), 0)
        ws.Range("P3").Value = ws.Cells(greatestDecIndex + 1, 9).Value
        ws.Range("Q3").Value = greatestDec
        
        ' Find the greatest volume, get the ticker value, and print them to correct cells.
        greatestTot = WorksheetFunction.Max(ws.Range("L2:L" & lastRowGreatest))
        greatestTotIndex = WorksheetFunction.Match(greatestTot, ws.Range("L2:L" & lastRowGreatest), 0)
        ws.Range("P4").Value = ws.Cells(greatestTotIndex + 1, 9).Value
        ws.Range("Q4").Value = greatestTot
        
        ' Adjust the percent change to be in the format 0.00%
        ws.Range("K2:K" & lastRowGreatest).NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        ' Adjust the column width so that data is visible.
        ws.Columns("A:Q").EntireColumn.AutoFit
        
    Next ws

End Sub