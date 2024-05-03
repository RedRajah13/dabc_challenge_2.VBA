Multiple Year Stock Data

Option Explicit
Sub Stocks():
    
    Dim i As Long 'row #
    Dim vol As LongLong 'column G
    Dim ticker As String 'column I
    Dim total_vol As LongLong 'column L
    
    Dim k As Long ' leaderboard row
    
    Dim rate_close As Double
    Dim rate_open As Double
    Dim price_change As Double
    Dim percent_change As Double
    
    Dim ws As Worksheet
    Dim lastRow As Long
    
    For Each ws In ThisWorkbook.Worksheets
    
        lastRow = ws.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
        ' asked Xpert
        
        total_vol = 0
        k = 2
        
        ' column titles
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
                   
        ' assign open for first ticker
        rate_open = ws.Cells(2, 3).Value
        
        For i = 2 To lastRow
           vol = ws.Cells(i, 7).Value
           ticker = ws.Cells(i, 1).Value
        
           ' loop rows 2 to lastRow
           ' check if next row ticker is different
           ' if not different, then add to total_vol
           ' if different, then add total_vol to leaderboard
           ' reset total_vol to 0
           
           If (ws.Cells(i + 1, 1).Value <> ticker) Then
               total_vol = total_vol + vol
                
               ' Get closing price of ticker
               rate_close = ws.Cells(i, 6).Value
               price_change = rate_close - rate_open
                
               ws.Cells(k, 9).Value = ticker
               ws.Cells(k, 10).Value = price_change
               ws.Cells(k, 12).Value = total_vol
               
               ' formatting
               If (price_change > 0) Then
                   ws.Cells(k, 10).Interior.ColorIndex = 4 'green
                   ws.Cells(k, 11).Interior.ColorIndex = 4 'green
               ElseIf (price_change < 0) Then
                   ws.Cells(k, 10).Interior.ColorIndex = 3 'red
                   ws.Cells(k, 11).Interior.ColorIndex = 3 'red
               Else:
                   ws.Cells(k, 10).Interior.ColorIndex = 2 'white
                   ws.Cells(k, 11).Interior.ColorIndex = 2 'white
               End If
               
               If (rate_open = 0) Then
                   percent_change = 0
                   
                Else: percent_change = price_change / rate_open
                   ws.Cells(k, 11).Value = percent_change
               End If
                
           'reset
           total_vol = 0
           k = k + 1
           rate_open = ws.Cells(i + 1, 3).Value 'look ahead to get next ticker open
        
           Else
               total_vol = total_vol + vol
           End If
        Next i
        
        ws.Columns("K:K").NumberFormat = "0.00%"
        ws.Columns("I:L").AutoFit
        
    Next ws
End Sub
        
Sub Leaderboard()
        
    Dim ws As Worksheet
    Dim i As Integer
    
    Dim col_total_vol As Range
    Dim col_percent_change As Range
    Dim cell_total_vol As Range
    Dim cell_percent As Range
    
    Dim gr_increase As Double
    Dim gr_decrease As Double
    Dim gr_total_vol As LongLong
    
    Dim gr_inc_ticker As String
    Dim gr_dec_ticker As String
    Dim gr_vol_ticker As String
    
    For Each ws In ThisWorkbook.Worksheets
    
        ' titles
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
    
        ' set ranges to search
        Set col_total_vol = ws.Range("L2:L1501")
        Set col_percent_change = ws.Range("K2:K1501")
        
        gr_total_vol = col_total_vol.Cells(1, 1).Value
        gr_increase = col_percent_change.Cells(1, 1).Value
        gr_decrease = col_percent_change.Cells(1, 1).Value
        
        ' loop through ranges to find leaderboard stats
        For Each cell_percent In col_percent_change
           If cell_percent.Value > gr_increase Then
               gr_increase = cell_percent.Value
           End If
        Next cell_percent
        
        For Each cell_percent In col_percent_change
           If cell_percent.Value < gr_decrease Then
               gr_decrease = cell_percent.Value
           End If
        Next cell_percent
        
        For Each cell_total_vol In col_total_vol
           If cell_total_vol.Value > gr_total_vol Then
               gr_total_vol = cell_total_vol.Value
           End If
        Next cell_total_vol
        
    ' Match tickers to values
        For i = 2 To 1501
             If gr_increase = ws.Cells(i, 11).Value Then
                gr_inc_ticker = ws.Cells(i, 9).Value
            End If
            If gr_decrease = ws.Cells(i, 11).Value Then
                gr_dec_ticker = ws.Cells(i, 9).Value
            End If
            If gr_total_vol = ws.Cells(i, 12).Value Then
                gr_vol_ticker = ws.Cells(i, 9).Value
            End If
        Next i
        
        ' assign leaderboard cells
        ws.Cells(2, 17).Value = gr_increase
        ws.Cells(3, 17).Value = gr_decrease
        ws.Cells(4, 17).Value = gr_total_vol
        ws.Cells(2, 16).Value = gr_inc_ticker
        ws.Cells(3, 16).Value = gr_dec_ticker
        ws.Cells(4, 16).Value = gr_vol_ticker
        
        ' formatting
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Columns("O:Q").AutoFit
    Next ws
End Sub