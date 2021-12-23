# Stock Analysis Refactoring
## Analyze stock data of green energy companies

### Project Overview
The purpose of this project was to review stock information from 2017 and 2018 on a selection of green energy companies.  Analysis was initially done on an individual company, DQ, which had the requestor's client was interested in.  The analysis consisted of calculating the yearly volume of trades and the yearly return of the stock.  After determining that the stock in question had a negative return in 2018 the analysis was expanded to other companies.

## Results
The original code:
`Sub AllStocksAnalysis()

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
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)`
    


The original code execution time is below.

![original code 2017](https://user-images.githubusercontent.com/95188079/147279838-fffc19fe-06f6-4227-b64f-f272b9bcd395.png)

![original code 2018](https://user-images.githubusercontent.com/95188079/147279844-58d486c4-b5fb-4f46-bbb2-7d733366de78.png)

The refactored code:

The refactored script ran much faster.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/95188079/147279147-b57f436a-2447-486e-8da2-8e657ec8f263.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/95188079/147279159-8db08a56-9815-4da1-b449-19911494b3db.png)


### Summary
What are the advantages or disadvantages of refactoring code?
Advantage 1: Refactoring the code did speed up the execution code.  The original time was 2.7 seconds and the refactored time was 0.4 seconds.

Advantage 2: If the code has not been reviewed in some time, refactoring could identify an issue in the code that was not caught in testing.  It's also an opportunity to add enhancements or update it if the audience of the output has changed.

Disadvantage 1: The time spent refactoring and troubleshooting issues was about 4 hours for me.  The first day I worked on it for about 3 hours and could not figure out why I kept getting a next without for error.  I got frustrated and eventually I saved my code to a text document, deleted it from the workbook, and then started over again the next day.

Not strictly a disadvantage some questions to consider are who determines when the code is efficient enough?  Is that the requestor, the client, the developer?  At what point does the time you spend refactoring outweigh the time you would gain?  Are new results being compared to the old to confirmt the code is working as expected?  


How do these pros and cons apply to refactoring the original VBA script?
