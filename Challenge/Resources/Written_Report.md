# Stock Analysis Challenge 

## Overview of Project

Steve (Client) has requested an analysis of multiple stocks over 2017 and 2018 and would like to have the Visual Basic (VBA) module that was previously submitted optimized to complete the analysis of thousands of stocks efficiently.

## Refactoring the Code

The original VBA module was refactored by creating a ticker Index, three output arrays for volumes, starting prices and ending prices.  The Ticker Index was created and set to zero. The Ticker Index will allow each ticker in the ticker array to be accessed by three output arrays. Each output array was set to 12 to indicate the number of elements in the arrays to match the number of elements is in the ticker array. Setting up the output arrays in this way will allow the code to loop through the data for each ticker to retrieve the volume, starting and ending prices more efficiently, therefore excluding the need for a nested loop.  A For Loop was created to initialize the ticker Volumes to zero. A non-nested for loop was created to loop through the data worksheet. Iterations were set to loop from the first row of data through the last row of data.  The loop evaluates the retrieves total volume for the current ticker and then the loop retrieves the start price and ending price of the current ticker second then lastly the loop increases the tickerIndex by one. This loop will continue until all twelve tickers have been looped through.

### Code

Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    Cells(3, 4).Value = "Starting Price"
    Cells(3, 5).Value = "EndingPrice"

    'Initialize array of all tickers
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Create a ticker Index
    Dim tickerIndex As Single
    tickerIndex = 0
    
    '1b) Create three output arrays
   Dim tickerVolumes(12) As Long
   Dim tickerStartingPrices(12) As Single
   Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For I = 0 To 11 'setup iterations from 0 to 11
        tickerVolumes(I) = 0  'setting total starting volume to 0
    Next I
        
    ''2b) Loop over all the rows in the spreadsheet.
    For I = 2 To RowCount  'loops from second row to last row of data
    
        '3a) Increase volume for current ticker
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(I, 8).Value 'gets total ticker volumes
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
             If Cells(I - 1, 1).Value <> tickers(tickerIndex) And Cells(I, 1).Value = tickers(tickerIndex) Then   'gets first price of ticker
                tickerStartingPrices(tickerIndex) = Cells(I, 6).Value  'set the value
                
            End If
        
        '3c) check if the current row is the last row with the selected tickerx.
             If Cells(I + 1, 1).Value <> tickers(tickerIndex) And Cells(I, 1).Value = tickers(tickerIndex) Then 'gest the last price of ticker
             
                tickerEndingPrices(tickerIndex) = Cells(I, 6).Value  'sets the value    

            '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1 'next ticker                
           
           End If
    
    Next I
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For I = 0 To 11  'setup iterations from 0 to 11
        
        Worksheets("All Stocks Analysis").Activate
        'Out put data for the current ticker
        tickerIndex = I
        Cells(4 + I, 1).Value = tickers(tickerIndex)  'outputs ticker
        Cells(4 + I, 2).Value = tickerVolumes(tickerIndex)  'outputs volume
        Cells(4 + I, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1   'calculates and outputs the percentage
        Cells(4 + I, 4).Value = tickerStartingPrices(tickerIndex)  'outputs starting price
        Cells(4 + I, 5).Value = tickerEndingPrices(tickerIndex)   'outputs ending price
        
    Next I
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Font.FontStyle = "Bold"
    Range("A3:E3").Font.FontStyle = "Bold"
    Range("A3:E3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Range("D4:D15").NumberFormat = "$#,##0.00"
    Range("E4:E15").NumberFormat = "$#,##0.00"
    Columns("B:E").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15

    For I = dataRowStart To dataRowEnd
        
        If Cells(I, 3) > 0 Then
            
            Cells(I, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(I, 3).Interior.Color = vbRed
            
        End If
        
    Next I
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


## Results 

The refactored code run the analysis in approximately ¼ of a second where the original code run the analysis in about 1 second, significantly faster than the original code.  This is due to the code requiring a nested For Loop and using the three output arrays. 

## Advantages of Refactoring

The advantage of refactoring the was making the code significantly faster, which will make the code more efficient when analyzing a larger number of stocks.  Another advantage is make the code cleaner and more readable. Personally, I found the advantage of refactoring was being able to better understand how the code works by reworking the code.

## Disadvantages of Refactoring

There really are no disadvantages in refactoring the code for the client, with the exception of paying to do the refactoring, however the cost would likely be offset by the more efficient code.  I did find it difficult to do at first to do the refactoring, so it could be a disadvantage for an analysis that does not have syntax experience with VBA to do the refactoring.

## Run Times for Original and Refactored Code

### Original Code Run Time for 2017 Stock Data
![2017 Original Code Run Time](https://github.com/blueschistrocks/stock-analysis/blob/8876a6657153bf9fcba45d1d8449595bcbff474a/Challenge/Resources/2017_runtime_original.png)

### Original Code Run Time for 2018 Stock Data
![2018 Original Code Run Time](https://github.com/blueschistrocks/stock-analysis/blob/8876a6657153bf9fcba45d1d8449595bcbff474a/Challenge/Resources/2018_runtime_original.png)

### Refactored Code Run Time for 2017 Stock Data
![2017 Refactored Code Run Time]( https://github.com/blueschistrocks/stock-analysis/blob/8876a6657153bf9fcba45d1d8449595bcbff474a/Challenge/Resources/VBA_Challenge_2017.png)

### Refactored Code Run Time for 2018 Stock Data
![2018 Refactored Code Run Time]( https://github.com/blueschistrocks/stock-analysis/blob/8876a6657153bf9fcba45d1d8449595bcbff474a/Challenge/Resources/VBA_Challenge_2018.png)

## Refactored Code Excel Analysis Output 

### Output for the 2017 Stock Analysis
![2017 Refactored Code Excel Output ](https://github.com/blueschistrocks/stock-analysis/blob/8876a6657153bf9fcba45d1d8449595bcbff474a/Challenge/Resources/2017_Excel.png)

### Output for the 2018 Stock Analysis
![2018 Refactored Code Excel Output ](https://github.com/blueschistrocks/stock-analysis/blob/8876a6657153bf9fcba45d1d8449595bcbff474a/Challenge/Resources/2018_Excel.png)




