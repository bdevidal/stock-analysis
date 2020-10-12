# stock-analysis

##Overview of Project: 

During the initial phases of this project, we were tasked by Steve to develop a VBA script to do some basic analytics of a set of stock data. The data set is a years worth of daily entries for a number of different stocks, containing basic information like opening price, closing price, trading volume, etc. A basic brute force script was developed to calculate the total trading volume and yearly return for each of the 12 listed stocks. While functional, there was concern that it was inefficient and processor heavy, and that there might be issues scaling it beyond the given data set. To that end, after establishing that our initial approach was sound, we decided to see if we could refactor our script to reduce overhead and increase efficiency 

##Results: 

Our initial approach used a series of heavy loops. There are 3012 rows in our data set. Our initial code required a full pass evaluation through that data set for each stock entry, with the output written at the end of each pass, as shown in this code:

```

    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        
'Loop through rows in the data.

        Worksheets(yearValue).Activate
        
            For j = 2 To RowCount

'Find the total volume for the current ticker.

                If Cells(j, 1).Value = ticker Then
                    
                    totalVolume = totalVolume + Cells(j, 8).Value
                    
                End If
                

'Find the starting price for the current ticker.

                If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then

                    startingPrice = Cells(j, 6).Value

                End If
'Find the ending price for the current ticker.

                If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then

                    endingPrice = Cells(j, 6).Value

                End If
                
        Next j
        
'Output the data for the current ticker.

        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
    Next i

    ```

Across 12 stocks and three different evaluations (volume, beginning, and end), that 108,432 touches in a single run, plus writes at the end. 

To increase efficiency, we changed over to an array model, as shown here:

```

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
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    
    For i = 0 To 11
        tickerVolumes(i) = 0
        
    Next i
    
    
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        
                If Cells(i, 1).Value = tickers(tickerIndex) Then
                    
                    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                    
                End If
                

        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then

                    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

                End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
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
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

        
    Next i

    ```

As you can see, by establishing a series of output arrays (tickerVolumes, tickerStartingPrices, and tickerEndingPrices), we were able to write the relevant data into these arrays in a single pass. We do this by establishing the variable tickerIndex to unify the data across arrays, then increment it at each transition between individual stocks. Once the arrays are full, we then do a final output to our presentation sheet. This drastically reduces the overhead per run, as can be seen in the following images;

![2017 Original](/Resources/yearValueAnalysisRuntime2017.png)
![2018 Original](/Resources/yearValueAnalysisRuntime2018.png)

![2017 Refactored](/Resources/VBA_Challenge_2017.png)
![2018 Refactored](/Resources/VBA_Challenge_2017.png)

##Summary: 

Given the above results, the refactored script should be much more scalable than the initial approach. However, it did require extensive effort above and beyond the initial code, which did work. It is also more complicated, may introduce additional points of failure, and may be harder to maintain. 
