Sub Alphabetical_Test()
    
    For Each ws In Worksheets
        
        Dim Ticker_Name As String
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Total_Volume As Double
        Dim LastRow As Long
        Dim Summary_Table_Row As Long
        Dim Open_Price As Double
        Dim Close_Price As Double
        
        Dim Greatest_Percent_Increase_Ticker As String
        Dim Greatest_Percent_Decrease_Ticker As String
        Dim Greatest_Total_Volume_Ticker As String
        Dim Greatest_Percent_Increase_Value As Double
        Dim Greatest_Percent_Decrease_Value As Double
        Dim Greatest_Total_Volume As Double
        
        Greatest_Percent_Increase_Value = 0
        Greatest_Percent_Decrease_Value = 0
        Greatest_Total_Volume = 0

   
        Total_Volume = 0
        
        ws.Cells(1, 9).Value = "Ticker Name"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        Summary_Table_Row = 2
        
        Open_Price = ws.Cells(2, 3).Value
        
        For i = 2 To LastRow
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                Ticker_Name = ws.Cells(i, 1).Value
                
                
                Close_Price = ws.Cells(i, 6).Value
                
                
                Yearly_Change = Close_Price - Open_Price
                

                If Open_Price <> 0 Then
                    Percent_Change = (Yearly_Change / Open_Price) * 100
                Else
                    Percent_Change = 0
                End If
                
                ws.Cells(Summary_Table_Row, 9).Value = Ticker_Name
                ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
                ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
                ws.Cells(Summary_Table_Row, 12).Value = Total_Volume
                
                
                ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                
                If Yearly_Change > 0 Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                ElseIf Yearly_Change < 0 Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                End If
                
                Open_Price = ws.Cells(i + 1, 3).Value
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                Total_Volume = 0
            End If
                            
        Next i
        
        Dim volume As Double
        Dim P_change_percent As Double
        
        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        For i = 2 To LastRow
            volume = ws.Cells(i, 12).Value
            If volume > Greatest_Total_Volume Then
                Greatest_Total_Volume = volume
                Greatest_Total_Volume_Ticker = ws.Cells(i, 9).Value
            End If
            
            ws.Cells(4, 16).Value = Greatest_Total_Volume_Ticker
            ws.Cells(4, 17).Value = Greatest_Total_Volume
            
                
            P_change_percent = ws.Cells(i, 11).Value
            If P_change_percent > 0 Then
                
                Greatest_Percent_Increase_Ticker = ws.Cells(i, 9).Value
                Greatest_Percent_Increase_Value = P_change_percent
                
                
            Else
                
                Greatest_Percent_Decrease_Ticker = ws.Cells(i, 9).Value
                Greatest_Percent_Decrease_Value = P_change_percent
                
                
            End If
            
            ws.Cells(2, 16).Value = Greatest_Percent_Increase_Ticker
            ws.Cells(2, 17).Value = Greatest_Percent_Increase_Value
            ws.Cells(3, 16).Value = Greatest_Percent_Decrease_Ticker
            ws.Cells(3, 17).Value = Greatest_Percent_Decrease_Value
            
        Next i
              
        
    Next ws

End Sub


