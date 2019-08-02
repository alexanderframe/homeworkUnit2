Sub unit2Easy    
    
    'declare variables for storing ticker name, total stock volume, table row, and the last row in sheet
Dim strTicker As String
Dim lngTotalVolume, lngTableRow, lngLastRow As Long

    ' set script to run for each worksheet
For Each ws in Worksheets

    'Find last row in sheet
    lngLastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    'create table for ticker name and total stock volume
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"

    'set initial volume to 0 and table row to 2
    lngTotalVolume = 0
    lngTableRow = 2

    'iterate through entire sheet
    For i = 2 to lngLastRow

        'if next row is not same stock then assign Ticker and add to Total Volume
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            strTicker = ws.Cells(i, 1).Value
            lngTotalVolume = lngTotalVolume + ws.Cells(i, 7).Value
            
            'add stock name and volume to table
            ws.Cells(lngTableRow, 9).Value = strTicker
            ws.Cells(lngTableRow, 10).Value = lngTotalVolume

            'add to table row and reset volume
            lngTableRow = lngTableRow + 1
            lngTotalVolume = 0

        'if next row is same stock then add to total volume
        Else
            lngTotalVolume = lngTotalVolume + ws.Cells(i,7).Value

        End If

    Next i

Next ws



End Sub