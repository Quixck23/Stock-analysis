# **VBA Challenge Stock Analysis**

## Overview of Project
   After the analysis of Steve's Parennts DAQO New Energy Corp (DQ) stock and also other 12 stocks in the look for a better alternative. We have decided to analize the whole market but in order for us to do it. We need to refactor our code so we can excute the analisys on a timely manner. 
   
## Results
  ***On our first Analysis for Module 1***
  
    '3a)Initialize variables for the starting price and ending price.
        Dim StartingPrice As Double
        Dim endingPrice As Double
        
    '3b)Activate the data worksheet.
        Worksheets(yearValue).Activate
    '3c)Find the number of rows to loop over.
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
    '4)Loop through the tickers.
        For i = 0 To 11
        Ticker = tickers(i)
        totalVolume = 0
    '5)Loop through rows in the data.
     Worksheets(yearValue).Activate
        For j = 2 To RowCount
     '5a)Find the total volume for the current ticker.
     If Cells(j, 1).Value = Ticker Then
        totalVolume = totalVolume + Cells(j, 8).Value
        End If
     
    '5b)Find the starting price for the current ticker.
        If Cells(j - 1, 1).Value <> Ticker And Cells(j, 1).Value = Ticker Then
            StartingPrice = Cells(j, 6).Value
            End If
    '5c)Find the ending price for the current ticker.
        If Cells(j + 1, 1).Value <> Ticker And Cells(j, 1).Value = Ticker Then
        endingPrice = Cells(j, 6).Value
        
        End If
    Next j
    '6)Output the data for the current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = Ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / StartingPrice - 1
        
        Next i
 
        
![VBA_Challenge2017Module1](https://user-images.githubusercontent.com/88118587/134945761-e7639ea2-af2d-42f2-825e-348100391ae6.PNG)![VBA_Challenge2018Module1](https://user-images.githubusercontent.com/88118587/134945769-f6113d35-e8ae-4de7-995e-5d77b9cfa094.PNG)

        
   ***On our improved refactored Analysis for Module 2***     
   
     'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    For i = 0 To 11
    TickerIndex = tickers(i)

    '1b) Create three output arrays
    Dim TickerVolumes As Long
    Dim TickerStartingPrices, TickerEndingPrices As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
           For j = 2 To RowCount
            TickerVolumes = 0
    ''2b) Loop over all the rows in the spreadsheet.
  
     '3a) Increase volume for current ticker
        TickerVolumes = TickerVolumes + Range("H" & j).Value
      
        
     '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If TickerIndex = Range("A" & j - 1).Value Then
        Else
        TickerStartingPrices = Range("F" & j).Value
        End If
     
        'End If
        
     '3c) check if the current row is the last row with the selected ticker
         'If the next row????????s ticker doesn????????t match, increase the tickerIndex.
          If TickerIndex = Range("A" & j + 1).Value Then
        Else
        TickerEndingPrices = Range("F" & j).Value
        End If
    Next j
          '3d Increase the tickerIndex.
         Next i
         
![VBA_Challenge2018](https://user-images.githubusercontent.com/88118587/134946922-c1d05f6f-13df-487b-bedb-e5bf5e7b9151.PNG)
![VBA_Challenge2017](https://user-images.githubusercontent.com/88118587/134946933-8d0565bd-ad6a-43f6-8933-15205228f31a.PNG)
   
***Comparison***
   From 2017, in which most of the stocks were on the green until 2018, the stocks researched took a deep dive.
   DQ had as much as 3 times more volume traded than the previous year but its price value went down %62.9.
   -Since most of these stocks are green energy companies, it possibly could have been a shift in the market or a new technology was discovered. The only stock that remained on the positive was ENPH and RUN; This could be due to a Technology patent that made them resilient to market changes. Either ENPH or RUN would be a much better option for Steven's Parents.

 ## Summary
   1. Advantages of refactoring our code improve the time and we could use for future stock analysis with a larger data. 
   2. Our pros of refactoring have a more useful code to be used even in other templates beside stocks. Cons it could be time consuming but its worth the work.
  
            
        
   
   
