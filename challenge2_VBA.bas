Attribute VB_Name = "Module1"

Sub testing():

'Loop through all sheets
For Each ws In Worksheets

'Determining name of the columns
ws.Cells(1, "I").Value = "Ticker"
ws.Cells(1, "J").Value = "Yearly Change"
ws.Cells(1, "K").Value = "Percentage Change"
ws.Cells(1, "L").Value = "Total Stock Volume"

'Determaine the last Row
LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

'Determine variables
Dim SummaryRow As Integer

'Dim volume As Long (not looping when volume is dimmed)
Dim ticker As String

'Determinine values
SummaryRow = 2
volume = 0
Next_Ticker = 2

For i = 2 To LastRow

'If tickers differ
If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
    ticker = ws.Cells(i, 1).Value
    volume = volume + ws.Cells(i, "G").Value
    closingprice = ws.Cells(i, "F").Value
    Change = (ws.Cells(i, "F").Value - ws.Cells(Next_Ticker, "C").Value)
    openingprice = ws.Cells(Next_Ticker, "C").Value
    yearlychange = (ws.Cells(i, "F").Value - ws.Cells(Next_Ticker, "C").Value)
    percentagechange = (yearlychange / openingprice) * 100
    Next_Ticker = i + 1
    
    ws.Cells(SummaryRow, "I").Value = ticker
    ws.Cells(SummaryRow, "J").Value = yearlychange
    ws.Cells(SummaryRow, "K").Value = percentagechange
    ws.Cells(SummaryRow, "L").Value = volume
    
SummaryRow = SummaryRow + 1
volume = 0
openingprice = ws.Cells(i + 1, "C").Value

'If tickers are the same
Else
volume = volume + ws.Cells(i, 7).Value

End If
Next i


'Determining name of the columns
ws.Cells(1, "P").Value = "Ticker"
ws.Cells(1, "Q").Value = "Value"
ws.Cells(2, "O").Value = "Greatest % increase"
ws.Cells(3, "O").Value = "Greatest % decrease"
ws.Cells(4, "O").Value = "Greatst Total Volume"

'Determaine the last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For j = 2 To LastRow

'Colouring column with
If ws.Cells(j, "K") > 0 Then
ws.Cells(j, "K").Interior.ColorIndex = 4
ElseIf ws.Cells(j, "K") < 0 Then
ws.Cells(j, "K").Interior.ColorIndex = 3
End If

'Finding Min/Max values in columns

'Dim percent As Long
If ws.Cells(j, "K").Value = WorksheetFunction.Max(ws.Range("K2:K" & LastRow)) Then
ws.Cells(2, "Q").Value = ws.Cells(j, "K").Value
ws.Cells(2, "P").Value = ws.Cells(j, "I").Value

ElseIf ws.Cells(j, "K").Value = WorksheetFunction.Min(ws.Range("K2:K" & LastRow)) Then
ws.Cells(3, "Q").Value = ws.Cells(j, "K").Value
ws.Cells(3, "P").Value = ws.Cells(j, "I").Value
End If

If ws.Cells(j, "L").Value = WorksheetFunction.Max(ws.Range("L2:L" & LastRow)) Then
ws.Cells(4, "Q").Value = ws.Cells(j, "L").Value
ws.Cells(4, "P").Value = ws.Cells(j, "I").Value
End If

Next j

Next ws
End Sub


