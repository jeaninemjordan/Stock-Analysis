# Stock-Analysis

## An analysis of 2017 and 2018 green energy stock data: refactored. 

### Submitted by: Jeanine Jordan, Class: Bootcamp: UCF-VIRT-DATA-PT-06-2022-U-B-TTH, Module 2 Challenge

#### Overview, Background and Purpose

This project analyzes the data for twelve different stocks to determine the most lucrative options for potential investors. The potential investors in question want to invest in green stocks and have expressed a specific interest in the DAQO New Energy Group (NYSE: DQ). An analysis was conducted to weigh DQ against other green energy stock options using collected data the investors supplied. After an initial analysis was conducted, the investors requested that the runtime of the analysis be reduced, resulting in the need to refactor the code written for this analysis.  

The workbook analysis created is interactive for the user using Visual Basic Application (VBA) within Excel. It provides each stock option’s annual volume and return on investment for two years, 2017 and 2018, with the click of a button. The runtime for the analysis is also reported at the end of each analysis. 

**Results**

Three new arrays were created: tVolume to contain volume, tStartPrice to contain the starting price and tEndPrice to contain the ending price. These arrays hold data for each stock when a for loop analyzes them.  

The original code established a ticker symbol that could be called on for each of the stocks. The three arrays were matched with the ticker array using a variable called the tickerIndex. 

These arrays allow for the use of Nested For Loops and variables that can loop through the collected data and complete the analysis.

Below are depictions of the Refactored and Original code for comparison. 

**Original Code**

    Sub AllStocksAnalysis()

    Dim startTime As Single
    
    Dim endTime  As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer

    '1) Format the output sheet on All Stocks Analysis worksheet
   
     Worksheets("All Stocks Analysis").Activate
    
     Range("A1").Value = "All Stocks (2018)"
   
    'Create a header row
   
    Cells(3, 1).Value = "Ticker"
   
    Cells(3, 2).Value = "Total Daily Volume"
   
    Cells(3, 3).Value = "Return"

    '2) Initialize array of all tickers
   
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
       
       TotalVolume = 0
       
       '5) loop through rows in the data
       
    Worksheets("2018").Activate
       
       For j = 2 To RowCount
       
           '5a) Get total volume for current ticker
           
           If Cells(j, 1).Value = ticker Then

               TotalVolume = TotalVolume + Cells(j, 8).Value

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
       
       Cells(4 + i, 2).Value = TotalVolume
       
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

     Next i

    endTime = Timer
    
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub

**Refactored Code**

    Sub AllStocksAnalysisRefactored()
    
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
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
    
    '1a) Create a ticker index
    
    Dim tickerIndex As Integer
    
    'Initiate tickerIndex at zero.
    
    tickerIndex = 0
    
    '1b) Create three output arrays
    
    Dim tVolume(12) As Long
    Dim tStartPrice(12) As Single
    Dim tEndPrice(12) As Single
    
    
    '2a) Create for loop to initialize the ticker Volumes to zero.
    
    For tickerIndex = 0 To 11
    
    
    tVolume(tickerIndex) = 0
    
    'Activate data worksheet
    
    Worksheets(yearValue).Activate
        
        '2b) Loop over all the rows in the spreadsheet.
        
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker.
            
            tVolume(tickerIndex) = tVolume(tickerIndex) + Cells(i, 8).Value
    
        
            '3b) Check if the current row is the first row with the current ticker.
                    
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
                'if it is the first row for current ticker, set starting price.
                
                tStartPrice(tickerIndex) = Cells(i, 6).Value
            
            'End If
            
            End If
            
            
        '3c) Check if the current row is the last row with the current ticker.

            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                'if the next row's ticker doesn't match, increase the tickerIndex
                
                tEndPrice(tickerIndex) = Cells(i, 6).Value
            
            'End if
            
            End If
            
        '3d) Increase the tickerIndex
        
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                
                
                tickerIndex = tickerIndex + 1
    
            
            'End If
            
            End If
    
        Next i
        
     Next tickerIndex
    
      '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
      For i = 0 To 11
        
        'Activate Output Worksheet
        
        Worksheets("All Stocks Analysis").Activate
        
        'Ticker Row Label
        
        Cells(4 + i, 1).Value = tickers(i)
        
        'Sum of Volume
        
        Cells(4 + i, 2).Value = tVolume(i)
        
        'ReturnValue
        
        Cells(4 + i, 3).Value = tEndPrice(i) / tStartPrice(i) - 1
        
    
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
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

     End Sub

**Execution Time**

The runtime of the analysis was reduced in the refactored code at the request of the potential investors. The original code ran in 1.421997seconds in 2017 and 1.436035 seconds in 2018. The refactored code ran 1.1570434 seconds faster at 0.2649536 seconds in 2017 and ran 1.1960448 seconds faster in 2018 at 0.2399902 seconds. This means there was an 81% reduction of runtime in 2017 and an 83% reduction of runtime in 2018.

Below are depictions of the runtime reports for each year, 2017 and 2018, in both the original and refactored code. 

**2017 Run Time With Original Code**

![VBA_Challenge_Runtime_OGCODE_2017.PNG](https://github.com/jeaninemjordan/Stock-Analysis/blob/main/Resources/VBA_Challenge_Runtime_OGCODE_2017.PNG)

**2018 Run Time With Original Code**

![VBA_Challenge_Runtime_OGCODE_2018.PNG](https://github.com/jeaninemjordan/Stock-Analysis/blob/main/Resources/VBA_Challenge_Runtime_OGCODE_2018.PNG)

**2017 Run Time with Refactored Code**

![VBA_Challenge_Runtime_REFACTORED_2017.PNG](https://github.com/jeaninemjordan/Stock-Analysis/blob/main/Resources/VBA_Challenge_Runtime_REFACTORED_2017.PNG)

**2018 Run Time with Refactored Code**

![VBA_Challenge_Runtime_REFACTORED_2018.PNG](https://github.com/jeaninemjordan/Stock-Analysis/blob/main/Resources/VBA_Challenge_Runtime_REFACTORED_2018.PNG)

**Stock Performance: 2017 versus 2018**

As pictured below, there is a significant change in return on investment for all stocks from 2017 to 2018. It is evident that DAQO New Energy Group’s (NYSE: DQ) ROI underwent a significant decline in one year and would not likely be a lucrative choice to invest in at this time. Alternatively, we can see Enphase Energy Inc. (NASDAQ: ENPH) continues to have a positive return indicating it may be a safe choice for the potential investors to choose. Lastly, the most profitable choice appears to be the Sunrun Inc. (NASDAQ: RUN) stock option. With its dramatic increase in only one year, it is indicative that this stock’s ROI could continue to rise. 

![VBA_Challenge_2017.PNG](https://github.com/jeaninemjordan/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

![VBA_Challenge_2017.PNG](https://github.com/jeaninemjordan/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

**Advantages and Disadvantages of Refactoring Code**

This analysis allows the potential investors a fast, interactive, and easily understandable method of weighing their options so they can choose the most lucrative green stock option. 

In software, it is necessary to refactor code to make it more adaptive, can help investigate the presence of bugs, improve speed, improve readability, and alter the design without changing the internal structure of the code. 

While refactoring code can optimize performance, there are also possible disadvantages. New bugs within the code could be introduced during the refactoring process, and the act of refactoring code alone can be very time consuming. There is a significant amount of time spent testing before code can be refactored successfully. 

**Conclusion**

By refactoring the code during this project, the analysis was optimized and became more efficient. Unnecessary clutter was removed and the code itself was made easier to read. The runtime was reduced by up to 1.1960448 seconds. While this may seem like a small number, the execution time was reduced by 83%. This is a significant difference, especially as the data set increases in size. 

