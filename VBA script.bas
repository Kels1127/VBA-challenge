Attribute VB_Name = "Module3"
Sub multiyr()

    Dim ws As Worksheet
    For Each ws In Worksheets
    
    
    Dim WorkSheetName As String
        
    Dim ticker As String
    Dim summary_table_row As Integer
    Dim open_val As Double
    Dim close_val As Double
    Dim percent_change As Double
    Dim yearly_change As Double
    Dim total_stock_volume As Double
    Dim rowCount As Long
    Dim max_increase_ticker As String
    Dim max_decrease_ticker As String
    Dim max_volume_ticker As String
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim max_volume As Double
  
    summary_table_row = 2
    open_val = ws.Cells(2, 3).Value
    max_increase = 0
    max_decrease = 0
    max_volume = 0
    
    rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    
    For i = 2 To rowCount
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ticker = ws.Cells(i, 1).Value
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            close_val = ws.Cells(i, 6).Value
            
            ws.Range("I" & summary_table_row).Value = ticker
            
            yearly_change = close_val - open_val
            percent_change = (close_val - open_val) / open_val * 100
            
            ws.Range("J" & summary_table_row).Value = yearly_change
            ws.Range("K" & summary_table_row).Value = percent_change
            ws.Range("L" & summary_table_row).Value = total_stock_volume
        
           
        If percent_change > max_increase Then
                max_increase = percent_change
                max_increase_ticker = ticker
        End If
        
         If percent_change < max_decrease Then
                max_decrease = percent_change
                max_decrease_ticker = ticker
        End If
            
         If total_stock_volume > max_volume Then
                max_volume = total_stock_volume
                max_volume_ticker = ticker
         End If
        
            summary_table_row = summary_table_row + 1
        
            open_val = ws.Cells(i + 1, 3).Value
            total_stock_volume = 0
       Else
       
           total_stock_volume = total_stock_volume + Cells(i, 7).Value
                     
        End If
       
        ws.Range("K" & summary_table_row).Style = "Percent"
        ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
        ws.Range("K" & summary_table_row).Value = percent_change
    
    Next i
    
    
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("P2").Value = max_increase_ticker
        ws.Range("Q2").Value = max_increase
        ws.Range("P3").Value = max_decrease_ticker
        ws.Range("Q3").Value = max_decrease
        ws.Range("P4").Value = max_volume_ticker
        ws.Range("Q4").Value = max_volume
        
    
  Next ws
        
End Sub


