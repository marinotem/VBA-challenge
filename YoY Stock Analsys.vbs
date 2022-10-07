Sub go()
    ' making the ticker name a string'
    Dim ticker As String
    
    'assigning percentages as doubles'
    Dim yr_chng As Double
    Dim pct_chng As Double
    
    'this is a long one so.... long'
    Dim ttl_vol As Long
    
    'setting a counter up to keep track of sums for tickers'
    Dim counter As Double
    counter = 0
    
    'keeping track of tickers'
    Dim ticker_table_row As Integer
    ticker_table_row = 2
    
    'specifying this as this worksheet'
    'Dim ws As Worksheet'
    
    'setting up r to store the total rows til the end'
    Dim r As Long
    
    'setting i as my variable to loop through'
    Dim i As Long
    
    'creating space for the end and beginning prices'
    Dim yr_beg As Double
    Dim yr_close As Double
    
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    'this essentially sets up r as the last nonblank row value for column 1'
    r = Cells(Rows.Count, 1).End(xlUp).Row
    
    'set year open value'
    yr_beg = Cells(2, 3).Value
    
    For i = 2 To r
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            counter = counter + Cells(i, 7).Value
            ticker = Cells(i, 1).Value
            yr_close = Cells(i, 6).Value
            yr_chng = (yr_beg - yr_close)
            pct_chng = (yr_chng / yr_beg)
            Range("I" & ticker_table_row).Value = ticker
            Range("J" & ticker_table_row).Value = yr_chng
            Range("K" & ticker_table_row).Value = FormatPercent(pct_chng)
                If Range("J" & ticker_table_row).Value > 0 Then
                    Range("J" & ticker_table_row).Interior.ColorIndex = 4
                Else
                    Range("J" & ticker_table_row).Interior.ColorIndex = 3
                End If
            Range("L" & ticker_table_row).Value = counter
            ticker_table_row = ticker_table_row + 1
            counter = 0
            yr_beg = Cells(i + 1, 3).Value
        Else
            counter = counter + Cells(i, 7).Value()
        End If
    Next i
End Sub
