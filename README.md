# stock_analysis
VBA practice


## Overview of Project
### Purpose
The purpose of this project was to refactor a Microsoft Excel VBA code that we worked on during the Module 2 about stock information for the year 2017 and 2018 and determine whether or not the stocks are worth investing. The purpose of this challenge was to increase the efficiency of the original code so that it would run faster in case we need to use this macros for large data. 

## Results

### The Given Data
The data included in the challenge is two lists of stocks for the year 2017 and 2018 which included the following information Ticker name, date, the Open and Close prices, the adjusted price, the high and low prices and finally the Volume.
I created two tables with stock information on 12 different stocks. The tables include information such as the ticker name, Total daily Volume, and Yearly Return.

### Analysis

The refactoring process was done was follow

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
    
    
    Dim tickerVolumes(12) As Double

    Dim tickerStartingPrices(12) As Single

    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
      
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        
    
        tickerStartingPrices(i) = 0
        
        
        tickerEndingPrices(i) = 0
        
        
        
    Next i
      
      
      
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        
    End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
            
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            
    End If
     

            '3d Increase the tickerIndex.
            
      If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerIndex = tickerIndex + 1
            
            
    End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        
    Cells(4 + i, 1).Value = tickers(i)
    
    Cells(4 + i, 2).Value = tickerVolumes(i)
    
    Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
        
    Next i
    
    'Formatting
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
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


![](https://github.com/nadiezhdamhb/stock_analysis/blob/main/run%20time%202017.png)


![](https://github.com/nadiezhdamhb/stock_analysis/blob/main/run%20time%202018.png)




1. What are the advantages or disadvantages of refactoring code?

The advantages to refactoring the code is that it gives you an opportunity to review your code and make it more efficient by finding ways to make it shorter, reuse code and try to solve problems or type code more elegantly.

The disadvantages are that in reviewning the code to refactor you can encounter other issues like what happened to me with the Error 9 out of range message I kept getting becuase I changed something in the code while working on the refactoring. Therefore I believe there could be a room for errors when reviewing the code. Nevertheless, refactoring is a good way to make your code better for any data given in the future. 


2. How do these pros and cons apply to refactoring the original VBA script?
They apply to the original script as the code run time was shorter after refactoring it. 

Like I mentioned on the previous question, I also run into an issue with the code due to a missing piece but I was able to fix it on time. 
