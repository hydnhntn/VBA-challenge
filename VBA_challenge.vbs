Attribute VB_Name = "Module1"
Sub StockLoop()

' Create Loop for BONUS: loops through worksheets and applies script to each

Dim ws As Worksheet

' loop through worksheets
For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
    

    ' ----------------- Begin worksheet level script --------------------------
    
    ' Create variables
    Dim x As Integer
    Dim tickerCount As Integer
    Dim sOpen As Double
    Dim sClose As Double
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    
    ' Get last row number (by counting) that has value
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    ' Set range of ticker values in sheet
    Set tickerRange = Range(Cells(2, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, 1))
    ' Set starting row value for loop
    x = 2
    
    ' loop through whole table
    For i = 2 To lastrow
    
        ' find unique ticker values (column A)
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        
            ' add unique values for ticker to running list in other column
            Cells(x, 9).Value = Cells(i, 1).Value
            
            ' get count of rows with ticker
            tickerCount = Application.WorksheetFunction.CountIf(tickerRange, Cells(i, 1).Value)
            
            ' get <open> value from FIRST row of ticker
            sOpen = Cells(i, 3).Value
            
            ' get <close> value from LAST row of ticker
            sClose = Cells(i + tickerCount - 1, 6).Value
            
            ' get yearly change: Beginning<open> - End<close>
            yChange = sClose - sOpen
            Cells(x, 10).Value = yChange
            
                ' color cell based on positive change (green) or negative change (red) or no change (no color change)
            If yChange < 0 Then
                Cells(x, 10).Interior.ColorIndex = 3
            ElseIf yChange > 0 Then
                Cells(x, 10).Interior.ColorIndex = 4
            End If
            
            'get percent change: 100*((Beginning<open> - End<close>)/Beginning<open>)
            If sOpen > 0 Then
                pChange = 100 * ((sClose - sOpen) / sOpen)
                Cells(x, 11).Value = Str(Round(pChange, 2)) & "%"
            Else
                pChange = 100 * (sClose - sOpen)
                Cells(x, 11).Value = Str(Round(pChange, 2)) & "%"
            End If
            
            'get total stock volume: sum of <vol> for ticker
            Set tickervolumeRange = Range(Cells(i, 7), Cells(i + tickerCount - 1, 7))
            Cells(x, 12).Value = WorksheetFunction.Sum(tickervolumeRange)
            
            x = x + 1
                    
        End If
    Next i
    
    
    
    
    ' --------------------------------** BONUS **-------------------------------------
    '  return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    lastrow_short = Cells(Rows.Count, 9).End(xlUp).Row
    
        ' compare each PERCENT CHANGE row and update the MAX along with the ticker value
    maxP = 0
    For i = 2 To lastrow_short
        If Cells(i, 11).Value > maxP Then
        maxP = Cells(i, 11).Value
        maxP_ticker = Cells(i, 9).Value
        End If
    Next i
    
    
        ' compare each PERCENT CHANGE row and update the MIN along with the ticker value
    minP = 0
    For i = 2 To lastrow_short
        If Cells(i, 11).Value < minP Then
        minP = Cells(i, 11).Value
        minP_ticker = Cells(i, 9).Value
        End If
    Next i
    
    
        ' compare each TOTAL STOCK VOLUME row and update the MAX along with the ticker value
    maxV = 0
    For i = 2 To lastrow_short
        If Cells(i, 12).Value > maxV Then
        maxV = Cells(i, 12).Value
        maxV_ticker = Cells(i, 9).Value
        End If
    Next i
    
    ' set values for min/max after loops to decrease amount of writes
    
    ' set values for Greatest % Increase
    Cells(2, 16).Value = maxP_ticker
    Cells(2, 17).Value = Str(maxP * 100) & "%"
    
    'set values for Greatest % Decrease
    Cells(3, 16).Value = minP_ticker
    Cells(3, 17).Value = Str(minP * 100) & "%"
    
    'set values forGreatest Total Volume
    Cells(4, 16).Value = maxV_ticker
    Cells(4, 17).Value = maxV
    
    ' ----------------- End worksheet specific script --------------------------


Next ws

End Sub
