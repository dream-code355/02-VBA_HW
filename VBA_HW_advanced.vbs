Sub StockAnalysis()
    ''' Calculates Year Change, Percent Change, Total Vol for each Ticker
    ''' Applies formatting
    For Each ws In Worksheets
        'get number of rows.
        Dim nrows As Long
        nrows = ws.Cells(ws.Cells.Rows.Count, 1).End(xlUp).Row

        'set header for calculated columns
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Year Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Volume"

        'Initialize variables and counters
        Dim Vopen As Double
        Dim Vclose As Double
        Dim change As Double

        Dim cur_row As Long
        Dim cur_open As Long
        Dim total_vol As Double
        cur_row = 2
        cur_open = 2
        total_vol = 0

        'loop through cells to calculate value
        For i = 2 To nrows
            total_vol = total_vol + ws.Cells(i, 7).Value
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ws.Cells(cur_row, 9).Value = ws.Cells(i, 1).Value
                Vclose = ws.Cells(i, 6).Value
                Vopen = ws.Cells(cur_open, 3).Value
                change = Vclose - Vopen
                ws.Cells(cur_row, 10).Value = change
                If change >= 0 Then
                    ws.Cells(cur_row, 10).Interior.ColorIndex = 10
                Else
                    ws.Cells(cur_row, 10).Interior.ColorIndex = 30
                End If
                If Vopen <> 0 Then 'don't divide by 0!
                    ws.Cells(cur_row, 11).Value = Round(change / Vopen, 2)
                Else
                    ws.Cells(cur_row, 11).Value = "Error"
                End If
                ws.Cells(cur_row, 12).Value = total_vol

                cur_row = cur_row + 1
                cur_open = i + 1
                total_vol = 0
            End If
        Next i

        'add another for loop to get max values
        Dim min_pct_change As Variant
        Dim min_pct_change_ticker As String
        Dim max_pct_change As Variant
        Dim max_pct_change_ticker As String
        Dim max_total As Double
        Dim max_total_ticker As String

        min_pct_change = ws.Cells(2, 11).Value
        max_pct_change = ws.Cells(2, 11).Value
        max_total = ws.Cells(2, 12).Value
        nrows = ws.Cells(ws.Cells.Rows.Count, 9).End(xlUp).Row

        ws.Cells(2, 14).Value = "Min Percent Change"
        ws.Cells(3, 14).Value = "Max Percent Change"
        ws.Cells(4, 14).Value = "Max Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"

        For i = 2 To nrows:
            If ws.Cells(i, 11).Value < min_pct_change And ws.Cells(i,11) <> "Error" Then
                min_pct_change = ws.Cells(i, 11).Value
                min_pct_change_ticker = ws.Cells(i, 9).Value
            End If
            If ws.Cells(i, 11).Value > max_pct_change And ws.Cells(i, 11) <> "Error" Then
                max_pct_change = ws.Cells(i, 11).Value
                max_pct_change_ticker = ws.Cells(i, 9).Value
            End If
            If ws.Cells(i, 12).Value > max_total Then
                max_total = ws.Cells(i, 12).Value
                max_total_ticker = ws.Cells(i, 9).Value
            End If
        Next i

        ws.Cells(2, 15).Value = min_pct_change_ticker
        ws.Cells(3, 15).Value = max_pct_change_ticker
        ws.Cells(4, 15).Value = max_total_ticker
        ws.Cells(2, 16).Value = min_pct_change
        ws.Cells(3, 16).Value = max_pct_change
        ws.Cells(4, 16).Value = max_total
    Next ws
End Sub
