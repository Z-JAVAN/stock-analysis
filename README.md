# Green stock-analysis
Background:<br/>

Steve wants to analyze a group of 12 green stocks to support his parents' investment decisions. We can analyze the annual stock volume and return on investment (ROI).
We are able to analyze each share with charts, and check information accurately, and now we can expand our research from 12 green shares. 


Purpose:</br>
Steve wants to look at more stocks so he can get more accurate information. He also knows that he must try to update his information and be able to avoid damage by careful analysis and make a good offer. He also can not spend much time on this and must use vbs to get information quickly and accurately. He knows that charts can help him better understand and process faster.</br>

Results:</br>
 
 We will created 3 new arrays: -tickerVolumes(12) to hold volume -tickerStartingPrices(12) to hold starting price -tickerEndingPrices(12) to hold ending price this information help us to check better our information 

Matching the 3 performance arrays with the ticker array is done by using a variable called the tickerIndex.

Now that I have created these arrays, I can use Nested For Loops and variables to loop through the data and complete the analysis.


Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime  As Single
    
    yearValue = InputBox("What year would you like to run the analysis on - 2017 or 2018?")
    
    
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate

    
    'Create a header row
    Range("A1").Value = "All Stocks (" + yearValue + ")"
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
    Dim rowIndex As Integer


    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

            
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For tickerIndex = 0 To 11
    
        tickerVolumes(tickerIndex) = 0
    
        '2b) Loop over all the rows in the spreadsheet.
        Worksheets(yearValue).Activate
        
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
            '3c) check if the current row is the last row with the selected ticker
            'If the next row’s ticker doesn’t match, increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
            End If
        Next i
        
    Next tickerIndex
    
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

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub



Sub ClearWorksheet()

    Cells.Clear

End Sub

2017 vs 2018 Stock Performance:</br>

![image_name](path/to/Untitled.png)g)<br/> <br/>


![outcomes_vs_picture.png](/resources/outcomes_vs_picture.png)<br/>
 </br>
Steve should look at the chart and see if the industry works before advising his parents on his investment decision. According to the chart and information, we see that many stocks have decreased in volume, so it is not a good choice for his parents in investing. And invest in other stocks. In the performance of green stocks in 2017 compared to 2018, we are witnessing a large decrease in volume, so it is better for them to invest more carefully.</br>


Execution time:</br>
Execution time improved from 0.9433594 seconds to 0.1708984 seconds for 2017, and, 1.066406 to 0.1894531 for 2018. That’s an improvement  82% for each year.</br>
