# Stock Analysis Refactoring
## Analyze stock data of green energy companies

### Project Overview
The purpose of this project was to review stock information from 2017 and 2018 on a selection of green energy companies.  Analysis was initially done on an individual company, DQ, which had the requestor's client was interested in.  The analysis consisted of calculating the yearly volume of trades and the yearly return of the stock.  After determining that the stock in question had a negative return in 2018 the analysis was expanded to other companies.

## Results

While all of the stocks except one had a positive yearly return in 2017, in 2018 the majority of the stocks had a negative yearly return.  Only ENPH and RUN had positive performance in both years.

The original code:
```
test test 
```

The original code execution time is below.

![original code 2017](https://user-images.githubusercontent.com/95188079/147279838-fffc19fe-06f6-4227-b64f-f272b9bcd395.png)

![original code 2018](https://user-images.githubusercontent.com/95188079/147279844-58d486c4-b5fb-4f46-bbb2-7d733366de78.png)

The refactored code:

```
test test
```

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
