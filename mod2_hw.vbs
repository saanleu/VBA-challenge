Attribute VB_Name = "Module1"
Sub multi_test()

    For Each ws In Worksheets
        Dim ticker As String
        Dim total_vol As Double
        Dim open_price, close_price As Double
        Dim yearly_change, percent_change As Double
        
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        wkstName = ws.Name
        
        Dim table_rows As Integer
        table_rows = 2
        ws.Range("I1").EntireColumn.Insert
        ws.Cells(1, 9).Value = "Ticker"
        
        ws.Range("J1").EntireColumn.Insert
        ws.Cells(1, 10).Value = "Yearly Change"
        
        ws.Range("K1").EntireColumn.Insert
        ws.Cells(1, 11).Value = "Percent Change"
        
        ws.Range("L1").EntireColumn.Insert
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        
        open_price = ws.Cells(2, 3).Value
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                ws.Range("I" & table_rows).Value = ticker
                
                close_price = ws.Cells(i, 6).Value
                yearly_change = close_price - open_price
                ws.Range("J" & table_rows).Value = yearly_change
                If yearly_change > 0 Then
                    ws.Range("J" & table_rows).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & table_rows).Interior.ColorIndex = 3
                End If
                
                If open_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = yearly_change / open_price
                End If
                ws.Range("K" & table_rows).Value = percent_change
                ws.Range("K" & table_rows).NumberFormat = "0.00%"
                
                total_vol = total_vol + ws.Cells(i, 7).Value
                ws.Range("L" & table_rows).Value = total_vol
                
                table_rows = table_rows + 1
                yearly_change = 0
                open_price = ws.Cells(i + 1, 3).Value
                total_vol = 0
            Else
                total_vol = total_vol + ws.Cells(i, 7).Value
                
            End If
        Next i
        
        
        ws.Range("P1").EntireColumn.Insert
        ws.Cells(1, 16).Value = "Ticker"
        
        ws.Range("Q1").EntireColumn.Insert
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        Dim op_inc, op_dec As Double
        Dim op_vol As Double
        Dim tick_inc, tick_dec, tick_vol As String
        op_inc = 0
        op_dec = ws.Cells(2, 11).Value
        op_vol = 0
        table_rows = table_rows - 1
        For i = 2 To table_rows
            If ws.Cells(i, 11).Value > op_inc Then
                op_inc = ws.Cells(i, 11).Value
                tick_inc = ws.Cells(i, 9).Value
            Else
            End If
            
            If ws.Cells(i, 11).Value < op_dec Then
                op_dec = ws.Cells(i, 11).Value
                tick_dec = ws.Cells(i, 9).Value
            Else
            End If
            
            If ws.Cells(i, 12).Value > op_vol Then
                op_vol = ws.Cells(i, 12).Value
                tick_vol = ws.Cells(i, 9).Value
            Else
            End If
        Next i
        
        ws.Cells(2, 16).Value = tick_inc
        ws.Cells(2, 17).Value = op_inc
        ws.Range("Q" & 2).NumberFormat = "0.00%"
        
        ws.Cells(3, 16).Value = tick_dec
        ws.Cells(3, 17).Value = op_dec
        ws.Range("Q" & 3).NumberFormat = "0.00%"
        
        ws.Cells(4, 16).Value = tick_vol
        ws.Cells(4, 17).Value = op_vol
                     
                 
     Next ws
      
End Sub

