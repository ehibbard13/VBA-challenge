Attribute VB_Name = "Module1"
 Sub multi_year_stock_data()

For Each ws In Worksheets

    Dim Ticker_Name As String
    Dim Ticker_Row As Integer
    Ticker_Row = 2
    Dim Yearly_Change As Double
    Dim Closing_Price As Double
    Dim Opening_Price As Double
    Dim Percent_Change As Double
    Dim Total_Stock As Double
    Dim Greatest_Increase As Double
    Greatest_Increase = 0
    Dim Greatest_Decrease As Double
    Greastest_Decrease = 0
    Dim Greatest_Total_Volume As Double
    Greatest_Total_Volume = 0
    
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Pecent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("Q1").Value = "Value"
    ws.Range("P1").Value = "Ticker"
    
    
    ws.Range("I1:Q1").EntireColumn.AutoFit
      
      LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      
      
            Opening_Price = ws.Cells(2, 3).Value
            Total_Stock = 0
            
             For i = 2 To LastRow
            
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker_Name = ws.Cells(i, 1).Value
                ws.Range("I" & Ticker_Row).Value = Ticker_Name
               Closing_Price = ws.Cells(i, 6).Value
               Yearly_Change = Closing_Price - Opening_Price
               ws.Range("J" & Ticker_Row).Value = Yearly_Change
               Percent_Change = Yearly_Change / Opening_Price
               ws.Range("K" & Ticker_Row).Value = Percent_Change
               ws.Range("K" & Ticker_Row).NumberFormat = "0.00%"
               ws.Range("Q2").NumberFormat = "0.00%"
               ws.Range("Q3").NumberFormat = "0.00%"
               Total_Stock = Total_Stock + ws.Cells(i, 7).Value
               ws.Range("L" & Ticker_Row).Value = Total_Stock
               Ticker_Row = Ticker_Row + 1
               Opening_Price = ws.Cells(i + 1, 3).Value
                Total_Stock = 0
               
               Else: Total_Stock = Total_Stock + ws.Cells(i, 7).Value
               
            
              End If
              
              If ws.Range("J" & Ticker_Row).Value >= 0 Then
                ws.Range("J" & Ticker_Row).Interior.ColorIndex = 4
              Else
                ws.Range("J" & Ticker_Row).Interior.ColorIndex = 3
                End If
                
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
            ws.Range("Q2").Value = ws.Range("K" & i).Value
            ws.Range("P2").Value = ws.Range("I" & i).Value
            End If
            
            If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
            ws.Range("Q3").Value = ws.Range("K" & i).Value
            ws.Range("P3").Value = ws.Range("I" & i).Value
            End If
             
            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
            ws.Range("Q4").Value = ws.Range("L" & i).Value
            ws.Range("P4").Value = ws.Range("I" & i).Value
            End If
            
            
    Next i
    
Next ws

End Sub

