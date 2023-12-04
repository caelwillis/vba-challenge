Attribute VB_Name = "Module1"
Sub stocks()

    Dim ticker_name As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim volume As Double
    Dim stock_table_row As Long
    Dim ws As Worksheet
    Dim startprice As Double
    
    
    
    For Each ws In Worksheets
    
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        
        
        volume = 0
        stock_table_row = 2
        startprice = ws.Cells(2, 3).Value
        
        
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        For i = 2 To RowCount
    
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker_name = ws.Cells(i, 1).Value
                volume = volume + ws.Cells(i, 7).Value
                yearly_change = ws.Cells(i, 6).Value - startprice
                percent_change = (ws.Cells(i, 6).Value - startprice) / startprice
                
                
           
                
                'Print output
                ws.Range("I" & stock_table_row).Value = ticker_name
                ws.Range("J" & stock_table_row).Value = yearly_change
                ws.Range("K" & stock_table_row).Value = percent_change
                ws.Range("L" & stock_table_row).Value = volume
            
                 'Format output
                 
                 ws.Range("K" & stock_table_row).NumberFormat = "0.00%"
                 
                 
                 
                 If ws.Range("J" & stock_table_row).Value < 0 Then
                    ws.Range("J" & stock_table_row).Interior.Color = RGB(200, 1, 1)
                
                ElseIf ws.Range("J" & stock_table_row).Value > 0 Then
                    ws.Range("J" & stock_table_row).Interior.Color = RGB(1, 200, 1)
                    
                End If
                
            
            
            
            
                 stock_table_row = stock_table_row + 1
                 startprice = ws.Cells(i + 1, 3).Value
                 volume = 0
            Else
                volume = volume + ws.Cells(i, 7).Value
        
            End If
            
            
            
            
        Next i

    Next ws
    

End Sub
