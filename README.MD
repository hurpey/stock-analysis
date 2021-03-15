Overview of Project: Explain the purpose of this analysis.
Summary
The purpose of this analysis is to determine if refactoring the VBA code previously created for Steve’s Stock Market Analysis will make the VBA script run faster run. If the code is refactored successfully, it will make the script more efficient and easier to read. It will also help Steve process the data sets for more years in a shorter period of time.
Background
The script and workbook created is to help Steve stock market research for his parents. The current dataset analyed includes information for 2017 and 2018. However, Steve will like to carry out an analysis for even more data within a shorter time frame by refactoring the VBA code used for the current data analysis.
This challenge will help determine whether refactoring the current code successfully makes the VBA script run faster. A written analysis and explanation of findings is also included to help Steve better understand the refactoring process. 
Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be the entry point to working with the existing code at a job.

Results
The result of the analysis shows that the return on ENPH and RUN stocks increased year to year from 2017 to 2018, compared to other stocks which declined. 
To refactor the code, the following steps were performed
'1a) Created a ticker Index
For i = 0 To 11
tickerIndex = tickers(i)


    '1b) Created three output arrays
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single, tickerEndingPrices As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
        Worksheets(yearValue).Activate
        tickerVolumes = 0
        
    ''2b) Loop over all the rows in the spreadsheet.
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(j, 1).Value = tickerIndex Then
            
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
        If Cells(j, 1).Value = tickerIndex Then
            
        'End If
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
        tickerVolumes = tickerVolumes + Cells(j, 8).Value
        
        End If

Summary:
Advantages of refactoring code
1. It helps to fix bugs especially when the existing codes have been badly written and find bugs easily
2. It helps to detect bad patterns in the code like duplicate codes, long methods and tight coupling, thereby making a code more efficient
Cons of refactoring codes
1. It can introduce bugs into the existing code.
2. The cost for refactoring a code may be more expensive than writing a code from scratch

How do these pros and cons apply to refactoring the original VBA script?
