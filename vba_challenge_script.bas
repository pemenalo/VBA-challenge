Attribute VB_Name = "Module1"
Sub stocks_checker()

'Variable definitions:
Dim x As Long
Dim ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim totalstockvolume As Double
Dim TickerRow As Integer
Dim opennumber As Double
Dim closenumber As Double
Dim ws As Worksheet

'For Loop:
For Each ws In Worksheets
    ws.Cells(1, 9).Value = "TICKER"
    ws.Cells(1, 10).Value = "YEARLY CHANGE"
    ws.Cells(1, 11).Value = "PERCENT CHANGE"
    ws.Cells(1, 12).Value = "TOTAL STOCK VOLUME"

    TickerRow = 1
    totalstockvolume = 0
    openingprice = ws.Cells(2, 3).Value
    closingprice = 0

'Want to loop through the rows
    lastrow = ActiveSheet.UsedRange.Rows.Count
    For x = 2 To lastrow

'My conditionals
If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
    ticker = ws.Cells(x, 1).Value
    TickerRow = TickerRow + 1
    ws.Cells(TickerRow, 9).Value = ticker
    ws.Cells(TickerRow, 12).Value = totalstockvolume
    closingprice = ws.Cells(x, 6).Value
    YearlyChange = closingprice - openingprice
    ws.Cells(TickerRow, 10).Value = YearlyChange

If openingprice = 0 Then
    ws.Cells(TickerRow, 11).Value = Str(0) + "%"

Else
    ws.Cells(TickerRow, 11).Value = Str(Round((YearlyChange / openingprice) * 100, 2)) + "%"

End If
'To move to next ticker reset at zero(0)
    openingprice = ws.Cells(x + 1, 3).Value
    totalstockvolume = 0
    
'Color coding my yearly change
If (YearlyChange > 0) Then
    ws.Cells(TickerRow, 10).Interior.ColorIndex = 4

Else
    ws.Cells(TickerRow, 10).Interior.ColorIndex = 3

End If

Else
    totalstockvolume = totalstockvolume + ws.Cells(x, 7).Value

End If

Next x

Next

    
End Sub


