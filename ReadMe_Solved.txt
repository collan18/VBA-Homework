Sub Stock_Data():

For Each ws In Worksheets
    'Declare Variables
    Dim yearly_change As Double
    Dim open_value As Long
    Dim closing_value As Long
    
    open_row = 2
    Total_Volume = 0
    Summary_table_row = 2
       
   'Set heading row
   ws.Range("I1").Value = "Ticker"
   ws.Range("J1").Value = "Yearly Change"
   ws.Range("K1").Value = "Percentage Change"
   ws.Range("L1").Value = "Total Volume"
   ws.Range("P1").Value = "Ticker"
   ws.Range("Q1").Value = "Value"
   ws.Range("O2").Value = "Greatest % Increase"
   ws.Range("O3").Value = "Greatest % Decrease"
   ws.Range("O4").Value = "Greatest Total Volume"

   'loop through and grab values for a ticker
   LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
   For Row = 2 To LastRow
   
   'Calculate Total Volume
   
   Total_Volume = Total_Volume + ws.Cells(Row, 7).Value
   ws.Cells(Summary_table_row, 12).Value = Total_Volume
   
    If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
    
        'Get aand paste ticker
        Ticker = ws.Cells(Row, 1).Value
        ws.Cells(Summary_table_row, 9).Value = Ticker
        
        'Calculate Yearly Change
        open_value = ws.Cells(open_row, 3).Value
        closing_value = ws.Cells(Row, 6).Value
        yearly_change = closing_value - open_value
        ws.Cells(Summary_table_row, 10).Value = yearly_change
        
        'Calculate Percentage Change
        'Account for overflow error
        If open_value = 0 Then
          ws.Cells(Summary_table_row, 11).Value = 0
        Else
            Percentage_change = yearly_change / open_value
            ws.Cells(Summary_table_row, 11).Value = Percentage_change
        End If
        
        'Format color for yearly change Column so that pos values are diff from neg veg values
        
        If yearly_change <= 0 Then
            'Fill cell with red color
            ws.Cells(Summary_table_row, 10).Interior.Color = RGB(255, 0, 0)
            ElseIf yearly_change > 0 Then
            'Fill cell with color green
            ws.Cells(Summary_table_row, 10).Interior.Color = RGB(0, 255, 0)
        End If
        
        'Format Percentage Change from General to Percentage
        ws.Cells(Summary_table_row, 11).NumberFormat = "0.00%"
            
        'Format number type of Total Volume Column
        ws.Cells(Summary_table_row, 12).NumberFormat = "#,###"
        
    'Reset for new ticker
        open_row = Row + 1
        Total_Volume = 0
        Summary_table_row = Summary_table_row + 1
    End If
    
    Next Row
    
    'Calculate Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
    
    Greatest_increase = Application.WorksheetFunction.Max(ws.Range("K:K"))
    ws.Cells(2, 17).Value = Greatest_increase
    
    Greatest_decrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
    ws.Cells(3, 17).Value = Greatest_decrease
    
    Greatest_Volume = Application.WorksheetFunction.Max(ws.Range("L:L"))
    ws.Cells(4, 17).Value = Greatest_Volume
    
        For Row = 2 To LastRow
        Percentage_change = ws.Cells(Row, 11).Value
        Ticker = ws.Cells(Row, 9).Value
        
            If ws.Cells(Row, 11).Value = Greatest_increase Then
                ws.Cells(2, 16).Value = Ticker
            ElseIf ws.Cells(Row, 11).Value = Greatest_decrease Then
                   ws.Cells(3, 16).Value = Ticker
            ElseIf ws.Cells(Row, 12).Value = Greatest_Volume Then
                   ws.Cells(4, 16).Value = Ticker
            End If
    
    Next Row
            
    'Format Column Q as Percentage
    ws.Range("Q2:Q4").NumberFormat = "0.00%"
        
    
Next ws
End Sub
        