Sub Module2ChallengePractice()

For Each ws In Worksheets
    Dim Ticker As String
    Dim Vol_Total As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Summary_Table_Row As Integer
    Dim Last_Row As Integer
    Dim New_Last_Row As Integer
    Dim Percent_Change As Double
    Dim Yearly_Change As Double
    Dim Max_Percent_Change_Ticker As String
    Dim Max_Percent_Change As Double
    Dim Min_Percent_Change_Ticker As String
    Dim Min_Percent_Change As Double
    Dim Max_Stock_Change_Ticker As String
    Dim Max_Stock_Volume As Double
  
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"

    Last_Row2 = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Summary_Table_Row = 2
    Vol_Total = 0
    Max_Percent_Change_Ticker = ""
    Max_Percent_Change = 0
    Min_Percent_Change_Ticker = ""
    Min_Percent_Change = 0
    Max_Stock_Change_Ticker = ""
    Max_Stock_Volume = 0
       
        For i = 2 To Last_Row2
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                Open_Price = ws.Cells(i, 3).Value
                Vol_Total = 0
            End If
            
            Vol_Total = Vol_Total + ws.Cells(i, 7).Value

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                Close_Price = ws.Cells(i, 6).Value
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                Yearly_Change = Close_Price - Open_Price
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
      
                If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.Color = vbGreen
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.Color = vbRed
                End If
                
                Percent_Change = (Close_Price - Open_Price) / (Open_Price)
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                If ws.Range("K" & Summary_Table_Row).Value >= 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.Color = vbGreen
                Else
                    ws.Range("K" & Summary_Table_Row).Interior.Color = vbRed
                End If
                
                ws.Range("L" & Summary_Table_Row).Value = Vol_Total
                Summary_Table_Row = Summary_Table_Row + 1
                
                If Vol_Total > Max_Stock_Volume Then
                    Max_Stock_Volume = Vol_Total
                    Max_Stock_Volume_Ticker = ws.Cells(i, 1).Value
                End If
                
                If Percent_Change > Max_Percent_Change Then
                    Max_Percent_Change = Percent_Change
                    Max_Percent_Change_Ticker = ws.Cells(i, 1).Value
                End If
                
                If Percent_Change < Min_Percent_Change Then
                    Min_Percent_Change = Percent_Change
                    Min_Percent_Change_Ticker = ws.Cells(i, 1).Value
                End If
            
            End If
                 
        Next i
         
        ws.Range("P4").Value = Max_Stock_Volume_Ticker
        ws.Range("P2").Value = Max_Percent_Change_Ticker
        ws.Range("P3").Value = Min_Percent_Change_Ticker
        ws.Range("Q4").Value = Max_Stock_Volume
        ws.Range("Q2").Value = Max_Percent_Change
        ws.Range("Q3").Value = Min_Percent_Change
Next ws

End Sub



