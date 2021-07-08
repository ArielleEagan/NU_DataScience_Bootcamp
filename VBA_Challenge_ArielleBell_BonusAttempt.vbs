'Bonus: Your solution will also be able to return the stock with:

Sub Bonus()
    
    'Create bonus box
    Dim bonus_box As Integer
    summary_box = 2
    
    'Create bonus variables
    Dim greatest_percentage_increase As Double
    Dim greatest_percentage_decrease As Double
    Dim greatest_total_volume As Long
    
    ' Declare Current as a worksheet object variable.
    Dim ws As Worksheet

    ' Loop through all of the worksheets in the active workbook.
    For Each ws In Worksheets

        ' Loop through all of the rows in each worksheet
        For i = 2 To lastrow
        
            'If Cells(i, 13).Value > max_value Then
            'max_value = greatest_percentage_increase
            
            'Find min value
            ElseIf Cells(i, 13).Value < min_value Then
            min_value = greatest_percentage_decrease
            
            'Find greatest total volume
            ElseIf Cells(i, 14).Value > max_value Then
            max_value = greatest_total_volume
            
            End If
            
            'Write in to bonus box
            ws.Range("Q" & bonus_box).Name = Value = greatest_percentage_increase"
            ws.Range("R" & bonus_box).Value = greatest_percentage_decrease
            ws.Range("S" & bonus_box).Value = greatest_total_volume
        
        Next
        
    Next ws

End Sub
