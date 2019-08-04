Sub data_1()

    For Each ws In Worksheets

    Dim ticker_name As String
    Dim ticker_row As Double
    Dim total_volume, yearly_change, closing_price, opening_price, percent_change As Double
    total_volume = 0
    ticker_row = 2
    Dim i As Long

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    opening_price = ws.Cells(2, 3).Value

        For i = 2 To LastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ticker_name = ws.Cells(i, 1).Value

            ws.Range("I" & ticker_row).Value = ticker_name

            closing_price = ws.Cells(i, 6).Value

            yearly_change = closing_price - opening_price
            ws.Range("j" & ticker_row).Value = yearly_change

            If (opening_price = 0 And closing_price = 0) Then

                percent_change = 0
            
            ElseIf (opening_price = 0 And closing_price <> 0) Then
            
                percent_change = 1
            
            Else
                percent_change = yearly_change / opening_price
                ws.Range("k" & ticker_row).Value = percent_change
                ws.Range("k" & ticker_row).NumberFormat = "0.00%"
           
            End If
                        

            total_volume = total_volume + ws.Cells(i, 7).Value

            ws.Range("L" & ticker_row).Value = total_volume

            ticker_row = ticker_row + 1

            opening_price = ws.Cells(i + 1, 3).Value

            total_volume = 0

            Else

            total_volume = total_volume + ws.Cells(i, 7).Value

            End If
            
        Next i

        pcLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

        For j = 2 To pcLastRow
        
            If (ws.Cells(j, 10).Value > 0 Or ws.Cells(j, 10).Value = 0) Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
        
        Next j
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
       
      
        For Z = 2 To pcLastRow
            If ws.Cells(Z, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & pcLastRow)) Then
                ws.Cells(2, 16).Value = ws.Cells(Z, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(Z, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(Z, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & pcLastRow)) Then
                ws.Cells(3, 16).Value = ws.Cells(Z, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(Z, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(Z, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & pcLastRow)) Then
                ws.Cells(4, 16).Value = ws.Cells(Z, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(Z, 12).Value
            End If
        Next Z





        


    Next ws

End Sub


