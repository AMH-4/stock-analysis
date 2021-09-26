# **stock-analysis**
*Explore green energy stock perfomance by analyzing financial data using VBA*

## Overview of Project:

#### *Purpose:*
Throughout Module 2 the goal was to help Steve, a Finance graduate, find a way he could quickly analyze an entire dataset at the click of a button. Steve would use this information to advise his clients/parents on the stock market as an investment. The purpose of this challenge is to refactor the Module 2 solution code, so that the VBA script could run faster. A refactored code would enable Steve to analyze a larger data set (the entire stock market) in a more efficient manner.

#### *Background:*
Steve, a Finance graduate, was in search of a way he could quickly analyze an entire dataset of stocks at the click of a button. As he begins his financial advisory practice, his parents became his first clients. Steve’s parents have been seeking to invest in green energy stocks, however they need more information. Using the VBA extension, I was able to build a workbook that automated task interacting with Excel. This project allowed Steve to quickly access 12 green energy stocks and provide a potential diversified portfolio for his parents. 

## Results 

#### *Process:*
With the purpose of this analysis definded, one main deliverable be neccessary to perform. 
- Refactor VBA code and measure performance
This deliverable will include an updated workbook and a folder with PNGs of the pop-ups with script run time.

Once the given files had been downloaded and named appropriately, the following septs were taken to perfomr the above necessary deliverables chronologically :

1. (Step 1a.)Create a tickerIndex variable and set it equal to zero before iterating over all the rows. You will use this tickerIndex to access the correct index across the four different arrays you’ll be using: the tickers array and the three output arrays you’ll create in Step 1b.

```
Dim tickerIndex As Integer
tickerIndex = 0
```

2. (Setp 1b.)Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices. The tickerVolumes array should be a Long data type. The tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.

```
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```

3. (Step 2a.)Create a for loop to initialize the tickerVolumes to zero.

```
For i = 0 To 11

ticker = tickers(i)
totalVolume(i) = 0

Next i

```

4. (Step 2b.)Create a for loop that will loop over all the rows in the spreadsheet.

```
For i = 2 To RowCount
```

5. (Step 3a.)Inside the for loop in Step 2b, write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker. Use the tickerIndex variable as the index.

```
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i,8).Value
```

6. (Step 3b.)Write an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current starting price to the tickerStartingPrices variable.

```
If Cells(i-1,1).Value <> tickers(tickerIndex) Then

tickerStartingPrices(tickerIndex) = Cells(i,6).Value

End If
```

7. (Step 3c.)Write an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable.

```
If Cells(i + 1,1).Value <> tickers(tickerIndex) Then

tickerEndingPrices(tickerIndex) = Cells(i,6).Value

```

8. (Step 3d.)Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.

```
tickerIndex = tickerIndex + 1

End If 
Next i
```

9. (Step 4.)Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.

```
For i = 0 To 11

Worksheets("All Stocks Analysis").Activate
tickerIndex = 1
Cells(i + 4, 1).Value = ticker(tickerIndex)
Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1

Next i 
```

10. Finally, run the stock analysis, then confirm that your stock analysis outputs for 2017 and 2018 are the same as they were in the module. Savethe pop-up messages showing elapsed run time for the refactored code for each year. 

*Below are the screen-shot images of the run times from the origninal code script ran in Modlule 2 for both 2017 and 2018.*
![VBA_Module2_2017OriginalRun](VBA_Module2_2017OriginalRun.png)
![VBA_Module2_2018OriginalRun](VBA_Module2_2018OriginalRun.png)

*Below are the screen-shot images of the run times from the refactored code script ran in this Challenge for both 2017 & 2018.*
![VBA_Challenge_2017](VBA_Challenge_2017.png)
![VBA_Challenge_2018](VBA_Challenge_2018.png)

*Below are the screen-shot images of the reults of the code script for both 2017 and 2018.*
![VBA Challenge_2017Results](VBA_Challenge_2017Results.png)
![VBA Challenge_2018Results](VBA_Challenge_2018Results.png)

#### *Analysis:*
The focus of this challenge project was to determine and compare the effectivness of the orginal code script to the refactored code script. A refactored code script allows the user to automate analysis and reduce the chance of accidents and/or errors. As seen in the screen-shot images below, the orgianal code script did run faster than the refactored code script. However, faster does not always mean more efficient. The orginal code script, while good for analysis of 12 stocks, was not built to analyze a robust amount of data. Refactoring the code script means that the user/client/Steve can perform analysis of a larger data set (the entire stock market). 










## Summary:

  ### What are the advantages or disadvantages of refactoring code?
  ### How do these pros and cons apply to refactoring the original VBA script?
