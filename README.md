# The VBA Refactor Challenge

## Overview of Project

### Purpose
The purpose of this project was to refactor the code so it could loop through all of the data once but in a shorter amount of time. The refactored code collects the same information in the earlier module, but the run time is shorter.
### Background
Steve is analyzing this dataset to explore the best choice for where his parents should invest. He wants to expand the dataset to include the entire stock market over the last few years. Although the existing code works well for a dozen stocks, it might not work well for thousands of stocks. But even if it does, it may take a long time to execute. By refactoring the code, the run time for the expanded dataset makes the VBA script run faster and more efficiently.



## Original script of 2017 Run Time
The execution time of the original script for the 2017 dataset was 0.79s as seen below.
<img width="260" alt="Screen Shot 2022-06-29 at 10 40 55 PM" src="https://user-images.githubusercontent.com/107584891/176594232-227065f4-b2f4-47e3-979b-2d7644642d58.png">

## Original script of 2018 Run Time
The execution time of the original script for the 2018 dataset was 0.61s as seen below.
<img width="258" alt="Screen Shot 2022-06-29 at 10 43 05 PM" src="https://user-images.githubusercontent.com/107584891/176594545-44a834ea-8ab3-4229-bf95-34b20368219b.png">

## Refactored script of 2017 Run Time
The execution time of the refactored script for the 2017 dataset was 0.57s as seen below.
<img width="258" alt="Screen Shot 2022-06-29 at 10 48 07 PM" src="https://user-images.githubusercontent.com/107584891/176594989-e7a1b7f4-2bb8-4110-97ec-489ddc52e1cb.png">

## Refactored script of 2018 Run Time
The execution time of the refactored script for the 2018 dataset was 0.57s as seen below.
<img width="266" alt="Screen Shot 2022-06-29 at 10 49 10 PM" src="https://user-images.githubusercontent.com/107584891/176595146-1d2e93ad-83e5-41bc-849f-fa4156978245.png">
 

Sub AllStocksAnalysisRefactored()
    
    Dim startTime As Single

    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("AllStocksAnalysis").Activate
    
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
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single

''2a) Create a for loop to initialize the tickerVolumes to zero.
' If the next row’s ticker doesn’t match, increase the tickerIndex.
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
    
    '3c) check if the current row is the last row with the selected ticker
    'If  Then
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



### Results
In terms of stock performance, there were more favorable returns in 2017 compared to 2018, as seen in the images below. In 2017, there were quite a few stocks that had high returns, i.e. DQ(199.4%), ENPH(129.5%), FSLR(101.3%), SEDQ(184.5%), and only had one negative return, i.e. TERP(-7.2%) whereas in 2018 the stocks had a negative return on all of them except ENPH (81.9%) and Run( 84%). The total daily volumes varied a bit and had a decline in 2018. The stocks "FSLR," "RUN," and "SEDQ" had an increase in Total Daily Volume, but among these, only "ENPH" and RUN" had a positive return. However, it is difficult to see a direct correlation between the total daily volume change and each stock's positive or negative return.

### 2017 
<img width="491" alt="Screen Shot 2022-06-30 at 9 21 08 AM" src="https://user-images.githubusercontent.com/107584891/176720487-1a317ae6-363e-4f02-a4b2-73301ba84d6a.png">

### 2018
<img width="548" alt="Screen Shot 2022-06-30 at 9 19 31 AM" src="https://user-images.githubusercontent.com/107584891/176721072-4d159f66-a302-4bce-9e8d-8f47a4747dcf.png">



## Summary
What are the advantages or disadvantages of refactoring code?
The advantages of refactoring code are that there are fewer steps, it is easier to read even for future users unfamiliar with the code, fills less memory on a computer, and delivers a faster runtime. It also allows for more adaptability as it can handle larger datasets more efficiently. The disadvantages of refactoring code would be the time and money spent having to go back to the original code to make these changes. It also requires a good understanding of the original code to rework it. 

By refactoring the "All stocks analysis" macro, the run time was reduced by half.  


How do these pros and cons apply to refactoring the original VBA script?
The refactored code allowed for a faster runtime, as seen by the calculated execution times. The execution time of the original script for the 2017 dataset was 0.79s and .84s for the 2018 one. However, the refactored script execution time for both the 2017 and 2018 datasets was 0.57s. The code's lower execution time and heightened efficiency are definite pros. The cons to refactoring would be the time spent working to optimize the code. In this case, it involved creating three other arrays for tickerVolumes, tickerStartingPrices, and tickerEndingPrices aside from the tickers array already existing in the original code. It also involved creating an additional variable, "tickerIndex," to access the stock ticker index in the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays. Though this was not too time-consuming to add to the refactored code, it did require a thorough understanding of the original code and how to incorporate arrays for efficiency.
