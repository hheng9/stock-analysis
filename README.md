# Stock-Analysis

## Overview of Project 

### The prupose of this analysis was to help Steve analyze the stock data his parents invested into. We used Microsoft VBA tool to create, edit, alter, and fine tune code to properly run macros for the correct data needed.

## Results

### The 2017 return had a more positive outcome for all the assigned tickers whereas the 2018 total daily volume had a more negative return percentage. The execution times came out slightly faster after refactoring the original code and restructering each line item.

## Summary 

- The advantages of refactoring the codes index, array, loops, and format is to have the code compile the data needed and process the information more efficiently.
- The cons of refactoring the original VBA script is that the code becomes more complex and more difficult to comprehend. The pros for refactoring the original VBA script is receiving more automated information and more organized data for the required conclusion.


![VBA_Challenge_2017](https://user-images.githubusercontent.com/118647523/207214830-b023659e-2aff-4e57-8830-03a00800d504.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/118647523/207214846-1d111f3f-6598-4192-b352-9fc5adfb9084.png)

![2017](https://user-images.githubusercontent.com/118647523/207214858-e1293e66-254e-45fa-9a21-f1659a019046.png)

![2018](https://user-images.githubusercontent.com/118647523/207214869-6519bd8f-2e74-495a-9492-941371f3ac4f.png)

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
