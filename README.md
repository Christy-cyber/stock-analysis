# Analysis of Green Stocks

## Overview of Project

Steve investigated the performance of DQ stock, a green stock, for 2018 to determine if it would be a good investment for his parents.  He found that it did not perform well in 2018 (-62.6% return) so inquired about the performance of other green stocks that would be good potential investments.  Steve analyzed 12 green stocks, including DQ, from 2018 and found that two, ENPH and RUN, had positive returns (81.9% and 84.0%, respectively).

### Purpose

The purpose of this exercise was to
  1) further investigate the return on the same 12 green stocks in 2017
  2) refactor VBA code that would result in faster processing of data output for these 12 stocks for 2017 and 2018

## Results
### Performance of Green Stocks in 2017 Compared to 2018
Eleven of 12 green stocks had positive returns in 2017, with four having over 100% returns (DQ, ENPH, FSLR, and SEDG).  Only one stock, TERP, showed a negative return (-7.2%).

![VBA_Challenge_2017--Stock Performance Only](https://user-images.githubusercontent.com/95387273/147701968-3d2acc71-6e47-4447-b046-092668f5ecb7.png)

Only two of the same twelve stocks, ENPH and RUN, had positive returns in 2018 (81.9% and 84.0%, respectively).

![VBA_Challenge_2018--Stock Performance Only](https://user-images.githubusercontent.com/95387273/147702207-8fc37f9d-7173-433a-a0f7-d090814ce0f3.png)

### Performance of Refactored VBA Code
The original code, "Year Value Analysis," for 2017 and 2018 took 0.609 seconds and 0.625 seconds, respectively, to run.  The original code also did not include formatting for either table, which was applied through a separate subroutine (See "Format Table" button in each screenshot below.).  

![Year_Value_Analysis_Time_2017](https://user-images.githubusercontent.com/95387273/147707079-61968fce-18ad-4e65-9f78-2aa0d3d55845.png)

![Year_Value_Analysis_Time_2018](https://user-images.githubusercontent.com/95387273/147707107-e60c49c0-6b02-40f5-9ea8-dd2c972b3081.png)


The reason the original code took longer to process was because the data for total volume and return for each ticker were being processed before moving on to the next ticker (See Comment '5 and '5a below which indicate a loop within a loop.).

    '3b) Activate Worksheet
    Sheets(yearValue).Activate
    
    '3c) Establish number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    '4) Loop through the array
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
                
       '5) Loop through rows in the data for each ticker
        Sheets(yearValue).Activate
        For j = 2 To RowCount
            
            '5a) Get total volume for current ticker
            If Cells(j, 1).Value = ticker Then
            
            'increase totalVolume by the value in the current row for "tickers"
            totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            '5b) Get starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
            'set starting price
            startingPrice = Cells(j, 6).Value
            End If
            
            '5c) Get ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
            'set ending price
            endingPrice = Cells(j, 6).Value
            End If
            
        Next j
        
        '6) Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1


The refactored code, "All Stocks Analysis Refactored," for 2017 and 2018 took 0.109 seconds and 0.105 seconds, respectively, to run and also included formatting for the table written into the code.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/95387273/147707881-f49fb12a-fea0-45bf-b770-ca5cef4171c8.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/95387273/147707919-22af815c-7491-4f5f-8ce0-263abecdef67.png)

The reason these took less time to process was because a ticker index and single loop were created to read the entire sheet instead of processing a single ticker at a time (See 1a, 2b and 3a-c).  Note that code for the formatting of the table is included under "'Formatting" at the bottom of the code shown below, which also increased efficiency.  The entirety of the original "Year Value Analysis" code and "All Stocks Analysis Refactored" code can be seen in the VBA Challenge Excel spreadsheet.

    
    'Activate data worksheet
    Sheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
    
        
    '2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(i) = 0
        
        Next i
    
        
    '2b) Loop over all the rows in the spreadsheet.
        
            For i = 2 To RowCount
        
                '3a) Increase volume for current ticker
                If Cells(i, 1).Value = tickers(tickerIndex) Then
                    'increase totalVolume by the value in the current row for "tickers"
                    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                End If
        
                '3b) Check if the current row is the first row with the selected tickerIndex.
                If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
                    'set starting price
                    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                End If
                 
                '3c) check if the current row is the last row with the selected ticker
                If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                
                    'set ending price
                    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                    
                    '3d) progress to next ticker
                    tickerIndex = tickerIndex + 1
                    
                End If
                 
        
            Next i
        
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    Columns("B:C").HorizontalAlignment = xlCenter

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i


## Summary
### Advantages and Disadvantages of Refactoring Code in General
Refactoring code can save a lot of time because not all of the code has to be rewritten.  The most useful sections can be copied and pasted, and, in some cases, only a few lines may have to be added or deleted. Ultimately, by improving its logic, code becomes more efficient. One of the disadvantages is that, in some cases, it is easier to write the code from scratch than try to follow someone else's logic, especially when notations have not been included in the code. Debugging can also be an issue and may require more time, especially when significant changes to the original code are being made.

### Advantages and Disadvantages of the Original and Refactored VBA Script
The original VBA script worked well on a small number of stocks for both 2017 and 2018.  Writing code and initializing an array of 12 ticker symbols did not take much time and the output was provided in about half of a second.  The advantage of the refactored script, though, was that the time to analyze the data was much faster and final output was provided in a tenth of a second.  This may not seem like so much of a difference for 12 stocks as one waits for output, but as the number of stocks increases in the analysis, the wait time would also increase and be noticeable.  A major disadvantage of the original code is that it would not work well for a large data set.  It did not take much time to initialize an array for 12 tickers, but one would not want to do this for many ticker symbols.  Why would someone want to initialize 100 ticker symbols if they did not have to?  It would be easier to rewrite the code to automatically harvest ticker symbols as part of the output.  This was also the same disadvantage that I have for the refactored VBA script.  While the overall time to create output decreased with the refactored script, the code would not be useful for a large data set and would need to be refactored to harvest ticker symbols automatically.

### Recommendations for Selecting Green Stocks
Steve should recommend the green stocks ENPH and RUN to his parents because of their high returns over the course of two years.
