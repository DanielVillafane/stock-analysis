# Stock Analysis using VBA

## Overview: 
### Leveraging VBA,  we created an Excel worksheet capable of analyzing all stocks. We tested the run times of our code using prevoius 2017 and 2018 stock information. 

## Results:
### We refactored the code we had previously developed by inserting arrays. This allowed us to run the code on more data at a faster rate. The refactored code and completion times are below:
###  Sub AllStocksAnalysisRefactored()
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
   
     For h = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
   Next h
    
    
   
    
    
    
   
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
             If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    
             End If
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
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
        ticker = tickers(i)
        totalVolume = tickerVolumes(i)
        endingPrice = tickerEndingPrices(i)
        StartingPrice = tickerStartingPrices(i)
        
        
        
        Worksheets("AllStocksAnalysis").Activate
        Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / StartingPrice - 1
        
        
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

![2017 run time](VBA_Challeneg_2017.png)
![2018 run time](VBA_Challeneg_2018.png)
Using the Kickstarter dataset:
1. I created a pivot table and filtered the table based on 
parent category: Theater. 
2. I then combined all campaign creation dates by year and displayed the combined data by month. 
3. I isolated the outcomes data and filtered to show only successful, failed, or canceled campaigns. 

## Analysis 2: Outcomes based on goals. 
### Description of Analysis 2:
####
Using Kickstarter dataset:
1. I created a new worksheet to capture percentage of successful, failed, and canceled campaignes by goal range.
2. Filtered the data to pull only outcome results for sub category: plays

## Results:
### Analysis 1: The analysis of outcomes based on launch date reveals the following:
1. The month's of May, June, July appear to be the best time to launch a campaign. The produced both the most campiagns launched and the highest number of successful campaigns.
2. The Month of Decmeber apears to be the least favorable month to launch a campaign. It produced both the least amount of campaigns with abbout 50% of these campaigns failing. 
[Analysis data](Kickstarter_Challenge.xlsx)
![Outcomes based on Launch Date](Theater_Outcomes_vs_Launch.png)

### Analysis 2: THe analysis of outcomes based on goals reveals the following:
1. Campaigns with a goal of $15,000.00 or less had a success rate greater than 50%
[Analysis data](Kickstarter_Challenge.xlsx)
![Outcomes based on Goals](Outcomes_vs_Goals.png)

### A more in depth analysis of the kicksatrter data would provide a more accurate analysis. For example: Identifying outliers or adding additional data to our analysis, such as pledged amounts and number of backers, may tell a more complete story. 

