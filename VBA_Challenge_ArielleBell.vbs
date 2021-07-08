Sub VBA_Challenge()

Dim w As Worksheet
    For Each w In ActiveWorkbook.Worksheets
    w.Activate
    
    'Set initial variables for the ticker name, opening price, closing price, and volume
    Dim ticker_name As String
    Dim opening_price As Double
    Dim closing_price As Double
    Dim yearly_change As Double
    Dim percentage_change As Double
    Dim total_stock_volume As Double
    Dim lastrow As Long
    
    'Set the values for initial variables
    opening_price = Cells(2, 3).Value
    closing_price = 0
    total_stock_volume = 0
    
    'Create summary box
    Dim summary_box As Integer
    summary_box = 2
    
    'Determine the last row
    lastrow = Range("A2", Range("A2").End(xlDown)).Rows.Count
    
    'Establish the for loop to go through it all now
    For i = 2 To lastrow
        
        'Move on if it's not the same value
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'Set the ticker name
        ticker_name = Cells(i, 1).Value
        
        'Add to the closing total
        closing_price = closing_price + Cells(i, 6).Value
        
        'Add to the volume total
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        'Write the ticker and its associated opening price, closing price, yearly difference, and percent change
        Range("I" & summary_box).Value = ticker_name
        Range("J" & summary_box).Value = opening_price
        Range("K" & summary_box).Value = closing_price
        Range("L" & summary_box).Value = closing_price - opening_price
        
        If opening_price > 0 Then
        
            Range("M" & summary_box).Value = (closing_price - opening_price) / opening_price
        
            Else
            
            End If
            
        'Write the total stock volume
        Range("N" & summary_box).Value = total_stock_volume
        
        'Add one to the ticker summary table and reset open and closing totals for new loop
        summary_box = summary_box + 1
        opening_price = Cells(i + 1, 3).Value
        closing_price = 0
        total_stock_volume = 0

    Else
    
        'Reach last number of the closing price
        closing_price = closing_price
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
    
    End If
    
'Set conditional formatting

    'Create if statement for positive change to be green, else negative change is red
    If Range("L" & summary_box).Value > 0 Then
        Range("L" & summary_box).Interior.ColorIndex = 4
        
    Else: Range("L" & summary_box).Interior.ColorIndex = 3
    
    End If

    Next i

Next w

End Sub


