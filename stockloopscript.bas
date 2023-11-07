Attribute VB_Name = "Module1"
Sub stockloop():
    
    ' Create a script that loops through all the stocks for one year and outputs the following information:
    ' - The ticker symbol
    ' - Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
    ' - The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
    ' - Conditional formatting is applied correctly and appropriately to the percent change column and yearly change column
    ' - The total stock volume of the stock.
    ' - Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.
    
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
    
    ' Step 1: set up variables and their starting points.
    Dim tickername As String
    Dim tickervolume As Double
    ' to keep track of each ticker name in the summary table
    Dim summary_ticker_rows As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim annualChange As Double
    ' will need to format the values in the column to percentage
    Dim percentChange As Double
    Dim lastrow As Double
    
    
    ' declaring starting points:
    tickervolume = 0
    summary_ticker_rows = 2
    'set starting point for opening price and then increment later in the loop
    openPrice = ws.Cells(2, 3).Value
    'close price will be declared later within the loop after the last row for each ticker is found.
    annualChange = (closePrice - openPrice)
    percentChange = (annualChange / openPrice)
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    
    'Step 2: label the headers for summary table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Annual Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Step 3: Start the for loops, first layer to indent through worksheets. Then address what will
    '  happen within each sheet. Make sure ticker column is sorted on all sheets.
        
        For i = 2 To lastrow
    
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
        
                'identifying the ticker symbol
                tickername = Cells(i, 1).Value
            
                'add ticker name to summary table
                ws.Range("I" & summary_ticker_rows).Value = tickername
            
            
                'add  total stock volume for each ticker name
                tickervolume = tickervolume + ws.Cells(i, 7).Value
            
                'add stock volume to summary table
                ws.Range("L" & summary_ticker_rows).Value = tickervolume
            
                'Pull the close info from that last row
                closePrice = ws.Cells(i, 6).Value
                
                'initiate the annual change calculation
                annualChange = (closePrice - openPrice)
                
                'add annual change calculation result to summary table
                ws.Range("J" & summary_ticker_rows).Value = annualChange
                
                ' Make sure to use conditional formatting that will highlight positive change in green and negative change in red.
                    If ws.Range("J" & summary_ticker_rows).Value < 0 Then
                    
                        ws.Range("J" & summary_ticker_rows).Interior.ColorIndex = 3
                        
                    Else
                    
                        ws.Range("J" & summary_ticker_rows).Interior.ColorIndex = 4
                        
                    End If
                
                'initiate the percent change calculation
                percentChange = (annualChange / openPrice)
                
                'add percent change to summary table (need to convert it to percentage format
                ws.Range("K" & summary_ticker_rows).Value = percentChange
                ws.Range("K" & summary_ticker_rows).NumberFormat = "0.00%"
            
                ' add one to the ticker rows
                summary_ticker_rows = summary_ticker_rows + 1
                
                'Reset stock volume
                tickervolume = 0

                'Reset the open price so that it grabs the first one for each new ticker name for the next loop.
                openPrice = ws.Cells(i + 1, 3)
            
            
            Else ' if ticker name is the same
                'add the stock volume
                tickervolume = tickervolume + ws.Cells(i, 7).Value
            
            End If
    
        Next i
        
    'After collecting all the data for each unique ticker, we will indentify the greatest % increase, decrease, and greatest total volume, along
    '   with identifying the ticker name associated with it.
    'I need to look through column K (or 11) to find the minimum and maximum values and print them to cells(3,16) and cells(2,16) respectively
    '   also look in the L (12) column for the maximum value and print it to cells(4,15). Also need to collect ticker name associated with each and place in row 15
    
    ' Declare a lastrow value for the summary table
    lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row

    'Create headers for new table
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    
        'Create a for loop to search for mins and maxs + print their value and ticker name on the appropriate cells. Calling to fill specific cells
        '  in this instance because I know how many values I'm looking for vs. the first summary table I built without knowing how many rows were
        '   going to be filled.
    
        For i = 2 To lastrow_summary_table
            
            'Use a conditional to find the maximum percent change, pull and print ticker name + value, and convert value to percentage
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.max(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 16).NumberFormat = "0.00%"
    
            'if the value isn't the max but is the min - pull and print ticker name + value, and convert value to percentage
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.min(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 16).NumberFormat = "0.00%"
            
            'if the value isn't the max percent or min percent, but it is the max of the stock volume then pull and print ticker name + value
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.max(ws.Range("L2:L" & lastrow_summary_table)) Then
                ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i

    Next ws
    
    
End Sub

