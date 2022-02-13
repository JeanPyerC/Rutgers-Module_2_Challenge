Attribute VB_Name = "Module1"
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
    
    '1a) Create a ticker Index and set ot equal to zero before iterating over all the rows
    'JP - When this starts, it's starting from 0 and replacing the 'i' to equal that value. That value correspond to
    'to the tickers string, example, tickers(0) = "AY"
   For I = 0 To 11
       tickerIndex = tickers(I)
       
       
    '1b) Create three output arrays
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single
    Dim tickerEndingPrices As Single
       
       
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    'JP: Acitvating yearvalue and setting the tickerVolume to 0.
       Worksheets(yearValue).Activate
       tickerVolumes = 0
       
       ''2b) Loop over all the rows in the spreadsheet.
       'JP: As of now, the vba is focusing on the year sheet we selected, it's going to start at 2 to the value of Rowcount
       For j = 2 To RowCount
               
           'JP: If the selected Cell matches the tickerindex.....
           If Cells(j, 1).Value = tickerIndex Then
           
              '3a) Increase volume for current ticker
              'JP: TickerVolume is set to 0, if the statment is true it will add the value from (j,8)
              tickerVolumes = tickerVolumes + Cells(j, 8).Value
        
           End If
           
           
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'JP:If  the value does not equal to the current tickerindex AND if the second value equals to the current index.....
           If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               tickerStartingPrices = Cells(j, 6).Value
               
          'End If
           End If

        '3c) check if the current row is the last row with the selected ticker
        'JP:If  the value does not equal to the current ticker index AND if the second value equals to the current index.....
           If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               tickerEndingPrices = Cells(j, 6).Value
               
          'End If
           End If
           
       Next j
       
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.

           Worksheets("All Stocks Analysis").Activate
           
           Cells(4 + I, 1).Value = tickerIndex
           Cells(4 + I, 2).Value = tickerVolumes
           Cells(4 + I, 3).Value = tickerEndingPrices / tickerStartingPrices - 1
    
            

   Next I
 
   
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Range("A4:C15").Font.Size = 14
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For I = dataRowStart To dataRowEnd
        
        If Cells(I, 3) > 0 Then
            
            Cells(I, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(I, 3).Interior.Color = vbRed
            
        End If
        
    Next I
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

