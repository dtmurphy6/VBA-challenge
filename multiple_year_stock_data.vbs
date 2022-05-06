Sub stonks()
    
    'Loop through all sheets
    For Each ws In Worksheets
    
        'Create a variable to hold the file name
        Dim work_sheet_name As String
        
        'Grab the worksheet name
        work_sheet_name = ws.Name
    
        'Declare initial variable for holding the ticker
        Dim ticker As String
        
        'Declare open price and close price
        Dim open_price As Variant
        Dim close_price As Variant
        open_price = 0
        close_price = 0
        
        'Set an initial variable for holding the yearly change
        Dim yearly_change As Variant
        yearly_change = 0
        
        'Set an initial variable for the percent change
        Dim percent_change As Variant
        percent_change = 0
        
        'Set an initial variable for the volume of stock
        Dim stock_volume As Variant
        stock_volume = 0
        
        'Set initial variables for greatest % increase, decrease, and greatest total volume
        Dim g_increase As Variant
        Dim g_decrease As Variant
        Dim g_volume As Variant
        Dim g_ticker_inc As String
        Dim g_ticker_dec As String
        Dim g_ticker_vol As String
        g_increase = 0
        g_decrease = 0
        g_volume = 0
        
        
        'Keep track of the location of each ticker in the summary table
        Dim output_table_row As Integer
        output_table_row = 2
        
        'Determine the last row
        Dim last_row As Variant
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox last_row
            
        'Set Headers
        ws.Range("I" & 1).Value = "Ticker"
        ws.Range("J" & 1).Value = "Yearly Change"
        ws.Range("K" & 1).Value = "Percent Change"
        ws.Range("L" & 1).Value = "Total Stock Volume"
        ws.Range("P" & 1).Value = "Ticker"
        ws.Range("Q" & 1).Value = "Value"
        ws.Range("O" & 2).Value = "Greatest % Increase"
        ws.Range("O" & 3).Value = "Greatest % Decrease"
        ws.Range("O" & 4).Value = "Greatest Total Volume"
        
        'Loop through all tickers
        For i = 2 To last_row
            
            'Check to see if we are still within the same ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Set the ticker name
                ticker = ws.Cells(i, 1).Value
                
                'Add to the total stock volume
                stock_volume = stock_volume + ws.Cells(i, 7).Value
                
                'Print the ticker in Column I
                ws.Range("I" & output_table_row).Value = ticker
                
                'Determine close price
                close_price = ws.Cells(i, 6).Value
                
                'Determine yearly change
                yearly_change = close_price - open_price
                
                'Print the yearly change in Column J
                ws.Range("J" & output_table_row).Value = yearly_change
                
                'Change color for positive or negative
                If yearly_change < 0 Then
                    ws.Range("J" & output_table_row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & output_table_row).Interior.ColorIndex = 4
                End If
                
                'Determine the percent change
                percent_change = yearly_change / open_price
                
                'Print the percent change in Column K
                ws.Range("K" & output_table_row).Value = percent_change
                     
                'Print the total stock volume in Column L
                ws.Range("L" & output_table_row).Value = stock_volume
                
                'Add one to the output table row
                output_table_row = output_table_row + 1
                
                'Reset the total stock volume
                stock_volume = 0
            
            'if the cell immediately following a row is the same ticker...
            Else
                
                'Determine open price
                If stock_volume = 0 Then
                    open_price = ws.Cells(i, 3).Value
                End If
                
                'Add to the total stock volume
                stock_volume = stock_volume + ws.Cells(i, 7).Value
                
            End If
        
        Next i
            
        'Determine the last row for percent change
        last_row = ws.Cells(Rows.Count, 11).End(xlUp).Row
        'MsgBox last_row
        
        'Loop through all percent change rows
        For i = 2 To last_row
            
            'Check for greatest increase
            If ws.Range("K" & i).Value > g_increase Then
                g_increase = ws.Range("K" & i).Value
                g_ticker_inc = ws.Range("I" & i).Value
            End If
        
            'Check for greatest decrease
            If ws.Range("K" & i).Value < g_decrease Then
                g_decrease = ws.Range("K" & i).Value
                g_ticker_dec = ws.Range("I" & i).Value
            End If
        
            'Check for greatest total volume
            If ws.Range("L" & i).Value > g_volume Then
                g_volume = ws.Range("L" & i).Value
                g_ticker_vol = ws.Range("I" & i).Value
            End If
            
            'Format percent change column to percentage
            ws.Range("K" & i).NumberFormat = "0.00%"
                
        Next i
                
        
        'Set greatest % increase, decrease, and greatest total volume
        ws.Range("P" & 2).Value = g_ticker_inc
        ws.Range("P" & 3).Value = g_ticker_dec
        ws.Range("P" & 4).Value = g_ticker_vol
        ws.Range("Q" & 2).Value = g_increase
        ws.Range("Q" & 2).NumberFormat = "0.00%"
        ws.Range("Q" & 3).Value = g_decrease
        ws.Range("Q" & 3).NumberFormat = "0.00%"
        ws.Range("Q" & 4).Value = g_volume
        
    Next ws
            
End Sub
