# Stock_Analysis

## Overview of Project
This project uses VBA to analyze a handful of green energy stock data and help a client decide if investing in a DQ stock (DAQO New Energy Corp) is a good decision.
Using VBA with the analysis helps automate it and reduces the chances of errors and the time needed to run the it. 

### Purpose
1. Analyze the total daily volume and yearly return for each stock. If a stock is traded often, then the price will accurately reflect its value. The yearly volume gives a rough idea of how often DQ stock gets traded. 


## Analysis
#### 1. Writing a Sub to compute DQ stock yearly volume and yearly return: 

![yearly return and volume](https://user-images.githubusercontent.com/79415699/109164731-c0be9c00-7748-11eb-8f38-d62330e281da.JPG)

#### 2. Adjusting the code using loops to run through all stock types and return yearly volume and yearly return in the output sheet:

```
Sub AllStocksAnalysis()
   '1) Format the output sheet on All Stocks Analysis worksheet
   
   Worksheets("All Stocks Analysis").Activate
   Range("A1").Value = "All Stocks (2018)"
   
   'Create a header row
   
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers
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
   '3a) Initialize variables for starting price and ending price
   Dim startingPrice As Single
   Dim endingPrice As Single
   '3b) Activate data worksheet
   Worksheets("2018").Activate
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets("2018").Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
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

End Sub
```

#### 3. Static and conditional Formatting
* Set up a loop that goes through every row and changes the cells color based on positive/negative values to better see and interpret the results. 

![color conditional formatting](https://user-images.githubusercontent.com/79415699/109344764-2f7d2180-783d-11eb-8065-f6e69301b073.JPG)

#### 4. Create Buttons to run, format and clear analysis
* Created a buttons for the end user to use to run the analysis, which makes the program more interactive and accessible. 

![buttons](https://user-images.githubusercontent.com/79415699/109347149-acf66100-7840-11eb-8bc6-a2b6798c83fa.JPG)

* The first button clears the worksheet, it's linked to the following subroutine:
```
Sub ClearWorksheet()

Cells.Clear

End Sub
```
* The second button runs the analysis for all stocks as per (2)
* The Third button formats the table to make sense of the results (green=good return, red=bad return), as per (3).

#### 5.Running analysis for any year.
Currently the Macros only analysis stocks in 2018, to modify it and make allow the end-user to pick a year for analysis, the following changes were made:
* Added an inputbox at the beginning of the code in (2)
```yearValue = InputBox("What year would you like to run the analysis on?")```
* Replaced hard-coded "2018" string ```Worksheets("2018").Activate``` with the dynamic value stored in yearValue ``````Worksheets(yearValue).Activate```

Below is an example of the input box and the 2017 Stock Analysis results:

![userinput](https://user-images.githubusercontent.com/79415699/109350488-8f77c600-7845-11eb-9d48-ad92276fe9f1.JPG)

![2017results](https://user-images.githubusercontent.com/79415699/109350494-91418980-7845-11eb-9d72-5f1dd72fe0a9.JPG)

#### 6. Code Performance
To measure how long the code took to perform the analysis, a timer was added at the beginning and end of the sub as follows:

```Sub AllStocksAnalysis()

'Set Timer
Dim startTime As Single
Dim endTime As Single

'User inputs value for year

yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
```
 
```Next i

    endTime = Timer
    MsgBox " This code ran in " & (endTime - startTime) & "seconds for the year" & (yearValue)
    
End Sub
```
The runtime for 2017 analysis is displayed below:

![runtime](https://user-images.githubusercontent.com/79415699/109351576-7ec84f80-7847-11eb-8989-067a72237bcf.JPG)

#### 7. Refactored code
The corde was refactored to improve its performance - reduce execution time. The refactored code combined the formatting and stock analysis calculations into one sub statement allowing for processing with one sub statement rather than two. Execution time went from 0.88 to 0.22 seconds for the year 2017. Below is the new runtime for the refactored code:

![refactored runtime](https://user-images.githubusercontent.com/79415699/109434931-a05f3d80-79e5-11eb-97ae-4a6d354c69e8.JPG)


## Summary

#### Pros and cons of refactoring code.
* Pros: What are the advantages or disadvantages of refactoring code? The advantages of refactoring code are that it can improve efficiency in the script, take fewer steps, use less memory, and allow the script to be better understood with improved logic. Disadvantages with refactoring code can be that since it was not created by the new user, it may be hard to understand the original intent of parts of the script if the script was not commented well.
* Cons: Since refactoring uses the original code that may have been created by another user, it may be challenging to understand the purpose of the original script if it was not well commented. 

#### How do these pros and cons apply to refactoring the original VBA script? 
These pros and cons directly apply to refactoring the orignal VBA script, the refactored code was more efficient, took less memory and was executed quicker than the original code. The origial script was well commented and both original and refactored code were made by the same user, which made the task a little easier to perform, but there might still be room for improvement. 

