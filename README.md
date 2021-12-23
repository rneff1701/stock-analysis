# Stock Analysis Refactoring
## Analyze stock data of green energy companies

### Project Overview
The purpose of this project was to review stock information from 2017 and 2018 on a selection of green energy companies.  Analysis was initially done on an individual company, DQ, which had the requestor's client was interested in.  The analysis consisted of calculating the yearly volume of trades and the yearly return of the stock.  After determining that the stock in question had a negative return in 2018 the analysis was expanded to other companies.

## Results

While all of the stocks except one had a positive yearly return in 2017, in 2018 the majority of the stocks had a negative yearly return.  Only ENPH and RUN had positive performance in both years.

The original code:
```
Sub AllStocksAnalysis()

'add code for timer
    Dim startTime As Single
    Dim endTime  As Single

   '1 Format the output sheet on All Stocks Analysis worksheet
   Worksheets("All Stocks Analysis").Activate
   ' change code to remove hard code for year from Range("A1").Value = "All Stocks (2018)" and added Call to call other subroutine as per https://powerspreadsheets.com/excel-vba-inputbox/ and https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/calling-sub-and-function-procedures
   Dim yearValue As String
   'Call yearValueAnalysis
   yearValue = InputBox("What year would you like to run the analysis on?")
   
   'add code to start timer
   startTime = timer
   
   Range("A1").Value = "All Stocks (" + yearValue + ")"
   'Create header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2 Initialize array of all tickers
   Dim tickers(11) As String
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
   '3a Initialize variables for starting price and ending price
   Dim startingPrice As Single
   Dim endingPrice As Single
   '3b Activate data worksheet
   ' change to remove hard code for year
   Sheets(yearValue).Activate
   '3c Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4 Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5 loop through rows in the data
   ' change to remove hard code for year
   Sheets(yearValue).Activate
       For j = 2 To RowCount
       '5a Get total volume for current ticker
        If Cells(j, 1).Value = ticker Then

        totalVolume = totalVolume + Cells(j, 8).Value

         End If
         '5b get starting price for current ticker
         If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

           startingPrice = Cells(j, 6).Value

          End If

           '5c get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

           endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
   
   'add code to end timer
   endTime = timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    

```

The original code execution time is below.

![original code 2017](https://user-images.githubusercontent.com/95188079/147279838-fffc19fe-06f6-4227-b64f-f272b9bcd395.png)

![original code 2018](https://user-images.githubusercontent.com/95188079/147279844-58d486c4-b5fb-4f46-bbb2-7d733366de78.png)

The refactored code:

```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = timer
    
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
    Dim tickerIndex As Integer
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    
    For i = 0 To 11
        tickerVolumes(i) = 0
        'Activate data worksheet
      Worksheets(yearValue).Activate
    
    Next i

    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
         'End If
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         
          'End If
         End If
    
            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            
             'End If
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

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

```

The refactored script ran much faster.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/95188079/147279147-b57f436a-2447-486e-8da2-8e657ec8f263.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/95188079/147279159-8db08a56-9815-4da1-b449-19911494b3db.png)


### Summary
What are the advantages or disadvantages of refactoring code?

Advantage 1: Refactoring can speed up execution time.

Advantage 2: If the code has not been reviewed in some time, refactoring could identify an issue in the code that was not caught in testing.  It's also an opportunity to add enhancements or update it if the audience of the output has changed.

Disadvantage 1: The time spent refactoring for a potentially minimal gain.


How do these pros and cons apply to refactoring the original VBA script?

Advantage 1: The original time was 2.7 seconds and the refactored time was 0.4 seconds.

Advantage 2: N/A for this project

Disadvantage 1: The time spent refactoring and troubleshooting issues was about 4 hours for me.  The first day I worked on it for about 3 hours and could not figure out why I kept getting a next without for error.  I got frustrated and eventually I saved my code to a text document, deleted it from the workbook, and then started over again the next day.  I did find it easier to work on the code in Visual Studio instead of Excel.  Although the script did run faster by a few seconds in a professional setting it would probably not be the best use of time for files this size.

Also, while not a disadvantage some questions to consider are who determines when the code is efficient enough?  Is that the requestor, the client, the developer?  At what point does the time you spend refactoring outweigh the time you would gain?  Are new results being compared to the old to confirmt the code is working as expected?  What is the frequency at which code should be reviewed to see if it can be refactored?
