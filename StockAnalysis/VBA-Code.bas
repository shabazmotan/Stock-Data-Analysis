Attribute VB_Name = "Module1"
Sub stocks()

    
    Dim lastrow As Long
    Dim i As Long
    Dim count As Long
    Dim opening As Double
    Dim closing As Double
    Dim volume As Double
    Dim ticker As String
    Dim percent_change As Double
    Dim value As Double
   
   ' Defined variables
   
    For Each ws In Worksheets
        lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
        count = 2
        volume = 0
        opening = ws.Cells(2, 3).value
        ticker = ws.Cells(2, 1).value
   ' Set count and last row
   
        ws.Cells(1, 9).value = "Ticker"
        ws.Cells(1, 10).value = "Yearly Change"
        ws.Cells(1, 11).value = "Percent Change"
        ws.Cells(1, 12).value = "Total Stock Volume"
        ws.Cells(2, 15).value = "Greatest % Increase"
        ws.Cells(3, 15).value = "Greatest % decrease"
        ws.Cells(4, 15).value = "Greatest Total Volume"
        
    
    'Defined column headers
    
    'Iterate throught the rows
            For i = 2 To lastrow
            
            volume = volume + ws.Cells(i, 7).value
            
            If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then
    
                closing = ws.Cells(i, 6).value
                percent_change = (closing - opening) / opening
                
                ws.Cells(count, 9).value = ticker
                ws.Cells(count, 10).value = closing - opening
                
                If (closing - opening > 0) Then
                    ws.Cells(count, 10).Interior.ColorIndex = 4
                ElseIf (closing - opening < 0) Then
                ws.Cells(count, 10).Interior.ColorIndex = 3
                
                End If
                
                 'Greatest % Increase'
                If ws.Cells(2, 17).value <= percent_change Then
                    ws.Cells(2, 17).value = percent_change
                    ws.Cells(2, 17).NumberFormat = "0.00%"
                    ws.Cells(2, 16).value = ticker
                
                End If
                
                 'Greatest % decrease'
                If ws.Cells(3, 17).value >= percent_change Then
                    ws.Cells(3, 17).value = percent_change
                    ws.Cells(3, 17).NumberFormat = "0.00%"
                    ws.Cells(3, 16).value = ticker
    
                End If
                
                'Greatest total volume'
               If ws.Cells(4, 17).value <= volume Then
                    ws.Cells(4, 17).value = volume
                    ws.Cells(4, 16).value = ticker
                End If
                
            
                ws.Cells(count, 11).value = percent_change
                ws.Cells(count, 11).NumberFormat = "0.00%"
                ws.Cells(count, 12).value = volume
                
                opening = ws.Cells(i + 1, 3).value
                ticker = ws.Cells(i + 1, 1).value
                count = count + 1
                volume = 0
            End If
        Next i
    Next ws
End Sub

