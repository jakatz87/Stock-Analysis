# VBA For Stock Analysis

## Project Overview

Our friend Steve is working on giving financial advice for investing in a select group of green energy stocks.  He has been given access to data for 12 different stocks and their performance in both 2017 and 2018 and he needs to have it analyzed.  We are working with the various tools within VBA to display the daily volume and the annual return for each stock in user-friendly format to allow Steve to give the best advice possible.

## Results

### Raw Data
The files we were able to use contained over 3000 rows of the daily activity for each stock in question: the opening price, daily high, daily low, closing price, adjusted close, and volume.  

![Image](https://github.com/jakatz87/stock-analysis/blob/main/resources/Raw%20Data%20Sample.png)

For Steve’s purposes, we are focusing only on the total daily volume to determine the visibility of the stock and the annual return to determine its performance.  

### Plan
Although the math is simple (adding each stock’s daily volume and subtracting the final close of the year to the opening close of the year for a percentage), the coding in VBA was much more complex.

#### Arrays for Variables
The most important aspect to this project was the use of arrays.  Since we are working with 12 different stocks, we had to create arrays for our variables to fit that number of outputs:
```
Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
```
and
```
Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```

#### For Loops

In order to calculate what we need, we had to run through each stock ticker’s data with For Loops and to do that, we needed another variable:  `tickerIndex`.
We first created the `tickerVolumes` loop with an initial `i` and a new `i`:
```
For i = 0 To 11
         tickerVolumes(i) = 0
        
    Next i
```
```
For i = 2 To RowCount
    
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
      If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
      End If
```
Then the Start Price and End Price loops within the next `i`:
```
      If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
      End If
      
      If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
      End If
```

And then we needed to increase the `tickerIndex` variable to repeat the process for all the stocks before we end the `i`:
```
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerIndex = tickerIndex + 1
        End If
        
    Next i
```

#### Display
Once all our math was able to be coded for the appropriate data sets, it was time to ensure the data was displayed on a seperate sheet and in a visually intuitive format.
Cells had to be populated with appropriate data of `tickers(i)`, `tickerVolumes(i)`, and the calculated "Return":
```
For i = 0 To 11
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
Next i
```
Number formats and color coding was included:
```
Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
```

### Conclusion
From our new sheet, we can see that green stocks had a fantastic year in 2017, but not so in 2018.  For staying power in a bad year, Steve will be able to see two stocks that stand out:  RUN and ENPH.  
![image](https://github.com/jakatz87/stock-analysis/blob/main/resources/Advice%202017.png)   ![image](https://github.com/jakatz87/stock-analysis/blob/main/resources/Advice%202018.png)

Both had high trading volumes in both years and both had outstanding performances to give Steve some valuable information for good advice.

## Summary
Refactoring Code
The code we ended with refactored from an original code.  The benefits of using this method to create solutions is that the simple parts, like creating the ticker array and formatting the cells, are taken care of already.  When the outline of what we need is created, it is quite a relief to know that the simple pieces are already accounted for.  The disadvantage of using refactored code is the creative limit it imposes.  When using refactored code, we are limited to the broad programming direction in the code itself.

For this particular refactored code, we had to fully understand the general direction and then ensure the details, variables, and appropriate loops and conditionals can fit the direction.  When attempting to imagine different methods to solve this problem, I ended up making things more complicated than I needed.  For example, I was thinking of nested loops at one point.  I needed help a few times, but was able to make this reformatted concept work for me and Steve.


