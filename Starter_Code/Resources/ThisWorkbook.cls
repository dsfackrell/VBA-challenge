VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub process()
    For Each ws In Worksheets
        Dim tickerSymbol As String
        
        ' Fill out Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        Dim lastRow As Long
        
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Define the current data to help fill out the output
        Dim outputRow As Integer
        Dim currentTicker As String
        Dim openPrice As Double
        Dim closePrice As Double
        Dim totalStock As Double
        Dim yearlyChange As Double
        Dim percentChange As Double
        
        Dim GreatIncTick As String
        Dim GreatIncValue As Double
        Dim GreatDecTick As String
        Dim GreatDecValue As Double
        Dim GreatTotalTick As String
        Dim GreatTotalValue As Double
        
        'Get the first data
        outputRow = 2
        openPrice = ws.Range("C2").Value
        currentTicker = ws.Range("A2").Value
        
        For i = 2 To lastRow
            ' Find if row is a new ticker then add to the output
            If currentTicker <> ws.Cells(i, 1).Value Then
                ' If not first row or last row then get the previous row for close price and get the Yearly Change and Percent Change
                If i > 2 And i < lastRow Then
                    
                    closePrice = ws.Cells(i - 1, 6).Value
                    
                    yearlyChange = closePrice - openPrice
                    percentChange = yearlyChange / openPrice
                    
                    
                    'check for greatest
                    If percentChange > GreatIncValue Then
                        GreatIncTick = currentTicker
                        GreatIncValue = percentChange
                    ElseIf percentChange < GreatDecValue Then
                        GreatDecTick = currentTicker
                        GreatDecValue = percentChange
                    End If
                    
                    If totalStock > GreatTotalValue Then
                        GreatTotalTick = currentTicker
                        GreatTotalValue = totalStock
                    End If
                    
                    ' Fill out output
                    ws.Cells(outputRow, 9).Value = currentTicker
                    ws.Cells(outputRow, 10).Value = yearlyChange
                    
                    If yearlyChange > 0 Then
                        ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                    End If
                    
                    ws.Cells(outputRow, 11).Value = percentChange
                    ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                    
                    ws.Cells(outputRow, 12).Value = totalStock
                    outputRow = outputRow + 1
                    totalStock = 0
                End If
                    
                openPrice = ws.Cells(i, 3).Value
                currentTicker = ws.Cells(i, 1).Value
            ElseIf i = lastRow Then
                totalStock = totalStock + ws.Cells(i, 7)
                
                closePrice = ws.Cells(i, 6).Value
                
                yearlyChange = closePrice - openPrice
                percentChange = yearlyChange / openPrice
                
                'check for greatest
                If percentChange > GreatIncValue Then
                    GreatIncTick = currentTicker
                    GreatIncValue = percentChange
                ElseIf percentChange < GreatDecValue Then
                    GreatDecTick = currentTicker
                    GreatDecValue = percentChange
                End If
                
                If totalStock > GreatTotalValue Then
                    GreatTotalTick = currentTicker
                    GreatTotalValue = totalStock
                End If
                    
                    
                ' Fill out output
                ws.Cells(outputRow, 9).Value = currentTicker
                ws.Cells(outputRow, 10).Value = yearlyChange
                
                If yearlyChange > 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                End If
                
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                ws.Cells(outputRow, 12).Value = totalStock
                
            End If
            
            totalStock = totalStock + ws.Cells(i, 7)
        Next i
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("P2").Value = GreatIncTick
        ws.Range("P3").Value = GreatDecTick
        ws.Range("P4").Value = GreatTotalTick
            
        ws.Range("Q2").Value = GreatIncValue
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = GreatDecValue
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = GreatTotalValue
    
    Next ws
End Sub

