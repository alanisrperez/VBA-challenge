Sub tickerloop():
Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        'Set variable for name
        Dim tickername As String
    
        'Set varable for total count on the total volume
        Dim tickervolume As Double
        tickervolume = 0

        'Keep track of each ticker name
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
        'Yearly price = Close Price rice at end of yr - Open Price at start of yr)
        'Percent change =((Close - Open)/Open)*100
        Dim open_price As Double
        open_price = Cells(2, 3).Value
        
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double

        'Label headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

        'Count # of rows in first column
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        'Loop
        For i = 2 To lastrow

            'Searches for when value of next cell changes from current
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
              'Set ticker name
              tickername = Cells(i, 1).Value

              'Add volume
              tickervolume = tickervolume + Cells(i, 7).Value

              'Add ticker name
              Range("I" & summary_ticker_row).Value = tickername

              'Add volume for each ticker
              Range("L" & summary_ticker_row).Value = tickervolume

              'Closing price
              close_price = Cells(i, 6).Value

              'Yearly change
               yearly_change = (close_price - open_price)
              
              'Add yearly change
              Range("J" & summary_ticker_row).Value = yearly_change

              'Check non-divisibilty when calculating the percent
                If open_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = yearly_change / open_price
                End If

              'Add yearly change
              Range("K" & summary_ticker_row).Value = percent_change
              Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              'Reset the row counter, add to summary_ticker_row
              summary_ticker_row = summary_ticker_row + 1

              'Reset volume to zero
              tickervolume = 0

              'Reset opening price
              open_price = Cells(i + 1, 3)
            
            Else
              
               'Add the volume
              tickervolume = tickervolume + Cells(i, 7).Value
            
            End If
        
        Next i

    'Conditional formatting (green is 10, red is 3)
    'Find the last row

    lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code yearly change
        For i = 2 To lastrow_summary_table
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 10
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i

    'Highlight stock price changes
    'Label cells

        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"

    'Find max and min values in Percent Change and max in Total Stock Volume
    'Collect ticker name and values for percent change and total volume for that ticker

        For i = 2 To lastrow_summary_table
            'max percent change
            If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
                Cells(2, 16).Value = Cells(i, 9).Value
                Cells(2, 17).Value = Cells(i, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"

            'min percent change
            ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
                Cells(3, 16).Value = Cells(i, 9).Value
                Cells(3, 17).Value = Cells(i, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            
            'max volume
            ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
                Cells(4, 16).Value = Cells(i, 9).Value
                Cells(4, 17).Value = Cells(i, 12).Value
            
            End If
        
        Next i
    Next ws
End Sub