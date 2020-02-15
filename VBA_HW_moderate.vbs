Sub StockAnalysis()
    ''' Calculates Year Change, Percent Change, Total Vol for each Ticker
    ''' Applies formatting
    For Each ws In Worksheets
        'get number of rows.
        Dim nrows As Long
        nrows = ws.Cells(ws.Cells.Rows.Count, 1).End(xlUp).Row

        'set header for calculated columns
        ws.Cells(1, 8) = "Ticker"
        ws.Cells(1, 9) = "Year Change"
        ws.Cells(1, 10) = "Percent Change"
        ws.Cells(1, 11) = "Total Volume"

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
                ws.Cells(cur_row, 8).Value = ws.Cells(i, 1).Value
                Vclose = ws.Cells(i, 6).Value
                Vopen = ws.Cells(cur_open, 3).Value
                change = Vclose - Vopen
                ws.Cells(cur_row, 9).Value = change
                If change >= 0 Then
                    ws.Cells(cur_row, 9).Interior.ColorIndex = 10
                Else
                    ws.Cells(cur_row, 9).Interior.ColorIndex = 30
                End If
                If Vopen <> 0 Then 'don't divide by 0!
                    ws.Cells(cur_row, 10).Value = Round(change / Vopen, 2)
                Else
                    ws.Cells(cur_row, 10).Value = "Error"
                End If
                ws.Cells(cur_row, 11).Value = total_vol

                cur_row = cur_row + 1
                cur_open = i + 1
                total_vol = 0
            End If
        Next i
    Next ws
End Sub