Sub ticker_summary()

    Dim ticker_col, date_col, open_col, high_col, low_col, close_col, vol_col, unigue_ticker_col As Integer
    ticker_col = 1
    date_col = 2
    open_col = 3
    high_col = 4
    low_col = 5
    close_col = 6
    vol_col = 7
    unique_ticker_col = 9
    Dim array_temp() As Long
    Dim open_prc() As Single
    Dim close_prc() As Single
    Dim tot_vol As Double
    Dim open_non_zero As Single

For Each Sheet In Worksheets

    'Get Unique Ticker Values
    Sheet.Range("A1", Sheet.Range("A1").End(xlDown)).Copy Destination:=Sheet.Range("I1")
    Sheet.Range("I1", Sheet.Range("I1").End(xlDown)).RemoveDuplicates Columns:=1, Header:=xlYes
    
    'Get Ticker row count for column A & I
    ticker_row_counter = Sheet.Cells(Rows.Count, ticker_col).End(xlUp).Row
    unique_ticker_row_counter = Sheet.Cells(Rows.Count, unique_ticker_col).End(xlUp).Row
    
    MsgBox (ticker_row_counter)
    MsgBox (unique_ticker_row_counter)
    
    'Titles
    Sheet.Range("I1").Value = "Ticker"
    Sheet.Range("J1").Value = "Yearly Change"
    Sheet.Range("K1").Value = "Percent Change"
    Sheet.Range("L1").Value = "Total Stock Volume"
    Sheet.Range("P1").Value = "Ticker"
    Sheet.Range("Q1").Value = "Value"
    Sheet.Range("O2").Value = "Greatest % Increase"
    Sheet.Range("O3").Value = "Greatest % Decrease"
    Sheet.Range("O4").Value = "Greatest Total Volume"
 
    i = 2
    k = 0
    l = 0
    tot_vol = 0
    open_non_zero = 0
    
    'Loop over main data
    For j = 2 To ticker_row_counter + 1
        'Check for conntinous ticker
        If Sheet.Cells(j, ticker_col).Value = Sheet.Cells(i, unique_ticker_col).Value Then
            ReDim Preserve open_prc(k)
            ReDim Preserve close_prc(k)
            open_prc(k) = Sheet.Cells(j, open_col).Value
            'Handle non zero open price and note array index
            If open_prc(0) <> 0 Then
                open_non_zero = open_prc(0)
                l = 0
            ElseIf open_non_zero = 0 Then
                open_non_zero = open_prc(k)
                l = k
            End If
            close_prc(k) = Sheet.Cells(j, close_col).Value
            tot_vol = Sheet.Cells(j, vol_col).Value + tot_vol
            k = k + 1
        'Add values to unique column when continous ticker breaks
        Else
            Sheet.Cells(i, unique_ticker_col + 1) = close_prc(k - 1) - open_prc(l)
            'Handle if all open prices are zero
            If open_prc(l) <> 0 Then
                Sheet.Cells(i, unique_ticker_col + 2) = (close_prc(k - 1) - open_prc(l)) / open_prc(l)
            Else
                Sheet.Cells(i, unique_ticker_col + 2) = 0
            End If
            
            Sheet.Cells(i, unique_ticker_col + 1).NumberFormat = "#####0.#########"
            'Color for increase or decrease
            If Sheet.Cells(i, unique_ticker_col + 1).Value < 0 Then
            Sheet.Cells(i, unique_ticker_col + 1).Interior.Color = RGB(255, 0, 0)
            Else
            Sheet.Cells(i, unique_ticker_col + 1).Interior.Color = RGB(0, 255, 0)
            End If
            Sheet.Cells(i, unique_ticker_col + 2).NumberFormat = "####.00%"
            Sheet.Cells(i, unique_ticker_col + 3) = tot_vol
            
            'Note the values for current row to be compared for continous ticker if statement
            k = 0
            ReDim Preserve open_prc(k)
            ReDim Preserve close_prc(k)
            tot_vol = 0
            open_prc(k) = Sheet.Cells(j, open_col).Value
            If open_prc(0) <> 0 Then
                open_non_zero = open_prc(0)
            ElseIf open_non_zero = 0 Then
                open_non_zero = open_prc(k)
            End If
            close_prc(k) = Sheet.Cells(j, close_col).Value
            tot_vol = Sheet.Cells(j, vol_col).Value + tot_vol
            k = k + 1
            i = i + 1
            open_non_zero = 0
            
        End If

    Next j
        ' Max increase%, Max decrease% and Max total stock volumes
        Sheet.Range("Q2") = "%" & WorksheetFunction.Max(Sheet.Range("K2:K" & unique_ticker_row_counter)) * 100
        Sheet.Range("Q3") = "%" & WorksheetFunction.Min(Sheet.Range("K2:K" & unique_ticker_row_counter)) * 100
        Sheet.Range("Q4") = WorksheetFunction.Max(Sheet.Range("L2:L" & unique_ticker_row_counter))
        
        'Find the indexes
        inc_index = WorksheetFunction.Match(WorksheetFunction.Max(Sheet.Range("K2:K" & unique_ticker_row_counter)), Sheet.Range("K2:K" & unique_ticker_row_counter), 0)
        dec_index = WorksheetFunction.Match(WorksheetFunction.Min(Sheet.Range("K2:K" & unique_ticker_row_counter)), Sheet.Range("K2:K" & unique_ticker_row_counter), 0)
        max_vol_index = WorksheetFunction.Match(WorksheetFunction.Max(Sheet.Range("L2:L" & unique_ticker_row_counter)), Sheet.Range("L2:L" & unique_ticker_row_counter), 0)
        
        'Find the ticker based on indexes
        Sheet.Range("P2") = Sheet.Cells(inc_index + 1, 9)
        Sheet.Range("P3") = Sheet.Cells(dec_index + 1, 9)
        Sheet.Range("P4") = Sheet.Cells(max_vol_index + 1, 9)
Next Sheet
    
End Sub
