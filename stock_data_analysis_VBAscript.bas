Attribute VB_Name = "Module1"
Sub stockData():
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets

        'set variables
        Dim rowStart, rowCounter, totalVolume, row, x As Long
        Dim openPrice, closePrice, percentChange, Increase, Decrease As Double
        Dim perChange, totVol As Range
    
        'label columns/cells
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        'set inital values
        totalVolume = 0
        rowStart = 2
        rowCounter = 0
        row = 2
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
        'loop through rows
        For x = 2 To lastRow
            If ws.Cells(x, 1).Value <> ws.Cells(x + 1, 1).Value Then
                
                'save ticker to column "I"
                ws.Cells(row, 9).Value = ws.Cells(x, 1).Value
                
                openPrice = ws.Cells(rowStart, 3).Value
                closePrice = ws.Cells(rowStart + rowCounter, 6).Value
                
                'calculate
                yearlyChange = closePrice - openPrice
                percentChange = yearlyChange / openPrice
                totalVolume = totalVolume + ws.Cells(x, 7).Value
                
                'save values to respective columns
                ws.Cells(row, 10).Value = yearlyChange
                ws.Cells(row, 11).Value = percentChange
                ws.Cells(row, 12).Value = totalVolume
                
                'reset row counters
                rowCounter = 0
                rowStart = x + 1
                row = row + 1
                totalVolume = 0
            
            Else
                'adjust volume and counter values
                totalVolume = totalVolume + ws.Cells(x, 7).Value
                rowCounter = rowCounter + 1
                   
            End If
        Next x
    
    'format cells
        'resize label cells
        ws.Range("A:O").EntireColumn.AutoFit
        
        'Adding percent formatting
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        'define last row for consolidated ticker list
        lastTicker = ws.Cells(Rows.Count, 9).End(xlUp).row

        'set ranges for percent change and total volume
        Set perChange = Range(ws.Cells(2, 11), ws.Cells(lastTicker, 11))
        Set totVol = Range(ws.Cells(2, 12), ws.Cells(lastTicker, 12))

        'calculate greatest % increase, % decrease, and total volume
        Decrease = Application.WorksheetFunction.Min(perChange)
        Increase = Application.WorksheetFunction.Max(perChange)
        greatestVol = Application.WorksheetFunction.Max(totVol)

        'assigning calculated values to cells
        ws.Cells(2, 17).Value = Increase
        ws.Cells(3, 17).Value = Decrease
        ws.Cells(4, 17).Value = greatestVol
        
        For x = 2 To lastTicker
            'add percent formatting to column "K"
            ws.Cells(x, 11).NumberFormat = "0.00%"
            
            'conditional formating for "Yearly Change" column
            If ws.Cells(x, 10).Value < 0 Then
                ws.Cells(x, 10).Interior.ColorIndex = 3
            ElseIf ws.Cells(x, 10).Value >= 0 Then
                ws.Cells(x, 10).Interior.ColorIndex = 4
            End If
            
            'greatest % Increase and % Decrease ticker symbols
            If ws.Cells(x, 11).Value = Increase Then
                ws.Cells(2, 16).Value = ws.Cells(x, 9).Value
            ElseIf ws.Cells(x, 11).Value = Decrease Then
                ws.Cells(3, 16).Value = ws.Cells(x, 9).Value
            End If
            
            'greatest volume ticker symbol
            If ws.Cells(x, 12).Value = greatestVol Then
                ws.Cells(4, 16).Value = ws.Cells(x, 9).Value
            End If
            
        Next x
        
    Next ws
        
End Sub

