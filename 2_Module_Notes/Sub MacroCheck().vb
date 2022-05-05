Sub MacroCheck()

    Dim testMessage As String
    
    testMessage = "Hello World!"
    
    MsgBox (testMessage)
    
End Sub


Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime  As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer

      Worksheets("All Stocks Analysis").Activate
      
       Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
'Step 1a:Create a tickerIndex variable and set it equal to zero
    Worksheets(yearValue).Activate
    tickerVolume = 0
      
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
    
'Step 1b:Create three output arrays
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single
    Dim tickerEndingPrices As Single

    Worksheets(yearValue).Activate
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    
'Step 2a:Create a for loop to initialize the tickerVolumes to zero
    For i = 0 To 11
    
        ticker = tickers(i)
        tickerVolume = 0
    
'Step 2b:Create a for loop that will loop over all the rows in the spreadsheet.
    Worksheets(yearValue).Activate
    For j = 2 To RowCount
    
'Step 3a: increases the current tickerVolume
    If Cells(j, 1).Value = ticker Then
        
        tickerVolume = tickerVolume + Cells(j, 8).Value
    
    End If
    
  'Step 3b:Write an if-then statement to check if the current row is the first row with the selected tickerIndex
    
    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

            tickerStartingPrice = Cells(j, 6).Value
            
'Step 3c:Write an if-then statement to check if the current row is the last row with the selected tickerIndex

        End If
        
'Step 3d:Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker
        
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

            tickerEndingPrice = Cells(j, 6).Value

        End If
        
    Next j

'Step 4:Use a for loop to loop through your arrays
        
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = tickerVolume
    Cells(4 + i, 3).Value = tickerEndingPrice / tickerStartingPrice - 1

Next i

endTime = Timer

MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

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

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i

End Sub

