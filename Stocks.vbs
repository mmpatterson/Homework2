Sub stocks()

'Code should account for everything in the Hard section

For Each ws In Worksheets
    '******Initialize Variables ********
    
    'Grab sum for each ticker
    Dim numrows As Long
    numrows = Cells(Rows.Count, 1).End(xlUp).Row
    
    'The row that the values for each ticker will be printed on
    Dim print_row As Integer
    print_row = 2
    Dim sum_ticker As Double
    
    'Placeholder for ticker name
    Dim ticker As String
    
    'Opening day value
    Dim first_day As Double
    
    'Closing value
    Dim last_day As Double
    
    'Closing value minus opening day value
    Dim year_change As Double
    
    '(Closing value - Opening Value)/Opening Value
    Dim percent_change As Double
    
    'Number of tickers researched
    Dim num_tickers As Integer
    
    'Ticker row value with greatest % Increase
    Dim num_increase As Integer
    
    'Ticker row value with greatest % decrease
    Dim num_decrease As Integer
    
    'Ticker Value with greatest % increase
    Dim ticker_increase As String
    
    'Ticker value with greatest % decrease
    Dim ticker_decrease As String
    
    'Ticker value with greatest total volume
    Dim ticker_volume As String
    
    'Ticker row value with greatest total volume
    Dim num_volume As Integer
    
    '*********Add Cell Names********
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    
    '*********8BEGIN SCRIPT*********
        'sum of total volume for stock
        sum_ticker = 0
                
        For i = 2 To numrows
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                first_day = Cells(i, 3)
            ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ticker = Cells(i, 1).Value
                sum_ticker = sum_ticker + Cells(i, 7).Value
                Cells(print_row, 9).Value = ticker
                Cells(print_row, 12).Value = sum_ticker
                last_day = Cells(i, 6)
                'Year Change
                year_change = last_day - first_day
                Cells(print_row, 10).Value = year_change
                If (year_change > 0) Then
                    Cells(print_row, 10).Interior.Color = vbGreen
                ElseIf (year_change < 0) Then
                    Cells(print_row, 10).Interior.Color = vbRed
                End If
                'Percent Change
                If (first_day <> 0) Then
                    percent_change = ((last_day - first_day) / first_day)
                End If
                Cells(print_row, 11).Value = percent_change
                Cells(print_row, 11).NumberFormat = "0.00%"
                If (percent_change > 0) Then
                    Cells(print_row, 11).Interior.Color = vbGreen
                ElseIf (percent_change < 0) Then
                    Cells(print_row, 11).Interior.Color = vbRed
                End If
                'update print row and reset ticker sum
                print_row = print_row + 1
                sum_ticker = 0
            Else
                sum_ticker = sum_ticker + Cells(i, 7).Value
            End If
        Next i
        
                        
    'Number of stocks researched
    num_tickers = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Initialize greatest increase, decrease, volume ticker (A, AA, etc.)
    ticker_increase = Cells(2, 9).Value
    ticker_decrease = Cells(2, 9).Value
    ticker_volume = Cells(2, 9).Value
    
    'Initialize row for greatest increase, decrease, volume
    num_increase = 2
    num_decrease = 2
    num_volume = 2
    
    'Go through tickers than have been extracted
    For m = 3 To num_tickers
        If Cells(m, 11).Value > Cells(num_increase, 11).Value Then
            ticker_increase = Cells(m, 9).Value
            num_increase = m
        End If
        If Cells(m, 11).Value < Cells(num_decrease, 11).Value Then
            ticker_decrease = Cells(m, 9).Value
            num_decrease = m
        End If
        If Cells(m, 12).Value > Cells(num_volume, 12).Value Then
            ticker_volume = Cells(m, 9).Value
            num_volume = m
        End If
    Next m
        
    'ticker name for greatest % increase
    Cells(2, 15).Value = Cells(num_increase, 9).Value
    
    'ticker name for greatest % decrease
    Cells(3, 15).Value = Cells(num_decrease, 9).Value
    
    'ticker name for greatest stock volume
    Cells(4, 15).Value = Cells(num_volume, 9).Value
    
    'value of greatest % increase
    Cells(2, 16).Value = Cells(num_increase, 11).Value
    Cells(2, 16).NumberFormat = "0.00%"
    
    'value of greatest % decrease
    Cells(3, 16).Value = Cells(num_decrease, 11).Value
    Cells(3, 16).NumberFormat = "0.00%"
    
    'value of greatest stock volume
    Cells(4, 16).Value = Cells(num_volume, 12).Value
Next ws
End Sub



