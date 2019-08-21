Attribute VB_Name = "Module1"
Sub stock_summary_generator()
    Dim ws As Worksheet
    For Each ws In Worksheets

        Dim nrow As Long
        Dim output_row As Long
        Dim i As Long
        Dim total_volume As Double
        Dim close_price As Double
        Dim open_price As Double
        Dim column As Double
        Dim yearly_change As Double
        Dim percent_change As String
        
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Opening Price"
        ws.Range("K1") = "Closing Price"
        ws.Range("L1") = "Yearly change"
        ws.Range("M1") = "% change"
        ws.Range("N1") = "total volume"
        
        nrow = Cells(Rows.Count, 1).End(xlUp).Row 'count the number of rows in first column
        
        i = 2
        summary_table_row = 2
        column = 0
        total_volume = 0
        
        'row 1 = ticker
        'row 2 = date
        'row 3 = open
        'row 4 = high
        'row 5 = low
        'row 6 = close
        'row 7 = volume
        
        open_price = ws.Cells(2, 3).Value 'define open price befofre for loop; open price is in 3rd column (1+2)
        
        For i = 2 To nrow
             If ws.Cells(i, 1).Value = ws.Cells((i + 1), 1).Value Then  'If ticker same between row i and row(i+1) then keep adding total volume
                 total_volume = total_volume + ws.Cells(i, 7).Value
                     
             Else 'if there's a mismatch between ticker at row i and row (i+1) that means the ticker switched but we'd still want the volume from the last row to be added before the switch
                 total_volume = total_volume + ws.Cells(i, 7).Value
                 ws.Range("I" & summary_table_row).Value = ws.Cells(i, 1).Value  ' Print ticker symbol in column I, starting at row 2
                 ws.Range("N" & summary_table_row).Value = total_volume 'Print the total volume in column J, starting at row 2
                 total_volume = 0 'reset the total at 0 afeter printing the ticker symbol and total volume
                 
                 close_price = ws.Cells(i, 6).Value 'the closing price for the current ticker will be the last row before the switch
             
             If open_price <> 0 Then
                 percent_change = (close_price - open_price) / open_price 'calculate the percentage change from opening price to closing price
                 percent_change = FormatPercent(percent_change, 2) 'convert the percentage change to % with 2 decimal places
                 yearly_change = close_price - open_price 'calculate yearly change
                 yearly_change = Round(yearly_change, 2) 'round the yearly change 2 decimal places
                 ws.Range("J" & summary_table_row).Value = open_price 'print the opening price to column J
                 ws.Range("K" & summary_table_row).Value = close_price 'print the closing price to column K
                 ws.Range("L" & summary_table_row).Value = yearly_change 'print the yearly change to column L
                 ws.Range("M" & summary_table_row).Value = percent_change 'print the percentage change to column M
                 open_price = ws.Cells((i + 1), 3).Value 'the 'new' open_price will be the open_price for the next ticker, which is row (i+1)
                 summary_table_row = summary_table_row + 1 ' Move to next row
             
             Else
                percent_change = 0
                percent_change = FormatPercent(percent_change, 2) 'convert the percentage change to % with 2 decimal places
                yearly_change = close_price - open_price 'calculate yearly change
                yearly_change = Round(yearly_change, 2) 'round the yearly change to 2 decimal places
                ws.Range("J" & summary_table_row).Value = open_price 'print the opening price to column J
                ws.Range("K" & summary_table_row).Value = close_price 'print the closing price to column K
                ws.Range("L" & summary_table_row).Value = yearly_change 'print the yearly change to column L
                ws.Range("M" & summary_table_row).Value = percent_change 'print the percentage change to column M
                open_price = ws.Cells((i + 1), 3).Value 'the 'new' open_price will be the open_price for the next ticker, which is row (i+1)
                summary_table_row = summary_table_row + 1 ' Move to next row

            End If
        End If
        Next i
        
        Dim j As Long
        Dim greatest_increase As Double
        Dim greatest_decrease As Double
        Dim summary_table_nrow As Double
        Dim greatest_increase_ticker As String
        Dim greatest_decrease_ticker As String
        Dim greatest_volume_ticker As String
        
        greatest_increase = ws.Range("M2") 'set the greatest_increase starting point comparison at L2
        greatest_increase = ws.Range("M2") 'set the greatest_decrease starting point comparison at L2
        greatest_volume = ws.Range("N2") 'set the greatest volume
        
        summary_table_nrow = ws.Cells(Rows.Count, 9).End(xlUp).Row 'count the number of rows in the 9th column (column I) which has all the unique tickers
        
        For j = 2 To summary_table_nrow
             If ws.Cells(j, 13).Value > greatest_increase Then 'find the greatest % increase
                 greatest_increase = ws.Cells(j, 13).Value
                 greatest_increase_ticker = ws.Cells(j, 9).Value '9th column is where the tickers are for the summary table
             End If
             'when j = 2, if cell(2,13), i.e., L2 > L2, then greatest_increase is at L2,
             'when j = 3, if cells (3,13), i.e., L3 > the greatest_increase (L2), then greatest increase is now L3
             'when j = 4, if cells (4,13), i.e., L4 > the great_increase (L3), then the greatest increase is now L4, etc
             
             If ws.Cells(j, 13).Value < greatest_decrease Then
                 greatest_decrease = ws.Cells(j, 13).Value
                 greatest_decrease_ticker = ws.Cells(j, 9).Value '9th column is where the tickers are for the summary table
             'when j = 2, if cell(2,13), i.e., L2 < L2, then greatest_decrease is at L2,
             'when j = 3, if cells (3,13), i.e., L3 < the greatest_decrease (L2), then greatest decrease is now L3
             'when j = 4, if cells (4,13), i.e., L4 < the great_decrease (L3), then the greatest decrease is now L4, etc
             End If
             
             If ws.Cells(j, 14).Value > greatest_volume Then
                 greatest_volume = ws.Cells(j, 14).Value '14th column is where the volumes are located
                 greatest_volume_ticker = ws.Cells(j, 9).Value '9th column is where the tickers are for the summary table
             End If
             
            'coloring the cells based on positive or negative yearly changes
             If ws.Cells(j, 12).Value < 0 Then 'if % change is less than 0, i.e., negative
                 ws.Cells(j, 12).Interior.ColorIndex = 3 'then color the % change as red
             End If
             
             If ws.Cells(j, 12).Value > 0 Then 'if % change is greater than 0, i.e., positive
                 ws.Cells(j, 12).Interior.ColorIndex = 4 'then color the % change as green
             End If
             
             If ws.Cells(j, 12).Value = 0 Then 'if % change is =0
                 ws.Cells(j, 12).Interior.ColorIndex = 2 ' then color the % change as white
             End If
             
             'coloring the cells based on positive or negative % changes
             If ws.Cells(j, 13).Value < 0 Then 'if % change is less than 0, i.e., negative
                 ws.Cells(j, 13).Interior.ColorIndex = 3 'then color the % change as red
             End If
             
             If ws.Cells(j, 13).Value > 0 Then 'if % change is greater than 0, i.e., positive
                 ws.Cells(j, 13).Interior.ColorIndex = 4 'then color the % change as green
             End If
             
             If ws.Cells(j, 13).Value = 0 Then 'if % change is =0
                 ws.Cells(j, 13).Interior.ColorIndex = 2 ' then color the % change as white
             End If
             
         Next j
         
         ws.Range("P2").Value = "Greatest % increase"
         ws.Range("Q2").Value = greatest_increase_ticker
         ws.Range("R2").Value = FormatPercent(greatest_increase, 2)
         
         ws.Range("P3").Value = "Greatest % decrease"
         ws.Range("Q3").Value = greatest_decrease_ticker
         ws.Range("R3").Value = FormatPercent(greatest_decrease, 2)
        
         ws.Range("P4").Value = "Greatest volume"
         ws.Range("Q4").Value = greatest_volume_ticker
         ws.Range("R4").Value = greatest_volume
         
         ws.Columns("I:R").AutoFit
         
    Next ws
End Sub
