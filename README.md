# Stock-Analysis

## Overview of Project 

### The prupose of this analysis was to help Steve analyze the stock data his parents were interested in investing into. We used Microsoft VBA tool to create, edit, alter, and fine tune code to properly run macros to generate the correct data needed. We wanted to analyze 12 specifics stocks to gather their total volume of gains or loss from the beginning of the year to the end of the year. 

## Results

### The 2017 return had a more positive outcome for all the assigned stock tickers except for TERP which show a negative return percentage. The 2018 return did poorly for mostly all stock tickers coming in at a negative return except for ENPH & RUN tickers coming in above 80%. 

## Stock Performance Between 2017 and 2018
![VBA_Challenge_2017](https://user-images.githubusercontent.com/118647523/207748628-a24ee3c0-7c89-4767-be64-25c616e506f0.png)![VBA_Challenge_2018](https://user-images.githubusercontent.com/118647523/207748643-a60f7747-82f7-4429-a1bf-5cda2650789c.png)

## Execution Times
### The execution times for 2017 & 2018 came out slightly faster after refactoring the original code and restructering each line item. As seen in the image below the original execution time for 2017 came in at 0.7265 and 2018 came in at 0.7578.

![Original VBA_Challenge_2017 png](https://user-images.githubusercontent.com/118647523/207751258-9d0eafed-89d7-426e-aa2c-5066646ef573.png)![Original VBA_Challenge_2018 png](https://user-images.githubusercontent.com/118647523/207751279-eba39c97-1bd9-43a6-9d3f-59502189406e.png)

### After refactoring the code as seen in the images below the 2017 came in at 0.1328 and 2018 at 0.1171 showing the efficiency of how fast the new code genrate results.

![VBA_Challenge_2017 png](https://user-images.githubusercontent.com/118647523/207751335-6b8aa068-7f47-45ee-8e3c-a81ee3bcdf7d.png)![VBA_Challenge_2018 png](https://user-images.githubusercontent.com/118647523/207751345-8ec64566-80d2-427a-88a7-716ede8917db.png)


## Summary 

- The advantages of refactoring the codes index, array, loops, and format is to have the code compile the data needed and process the information more efficiently.
- The cons of refactoring the original VBA script is that the code becomes more complex and more difficult to comprehend during execution which can create more frustration. The pros for refactoring the original VBA script is receiving more automated information and more organized data for the required conclusion.

## Refactored Script 

Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("AllStocksAnalysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    Range("A1").Value = "All Stocks (2018)"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

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
    
    '1a) Create a ticker Index
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(tickerIndex) = 0
    tickerStartingPrices(tickerIndex) = 0
    tickerEndingPrices(tickerIndex) = 0
    
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3d Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("AllStocksAnalysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("AllStocksAnalysis").Activate
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
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
