Attribute VB_Name = "Module1"
Sub TickerCount()

Dim ws As Worksheet
On Error Resume Next
For Each ws In ThisWorkbook.Worksheets
Dim Ticker As String
Dim Yearly_Change As Double
Dim Percet_Change As Double
Dim Total_Stock_Volume As Integer
Dim Summary_Table_Row As Integer

Summary_Table_Row = 2

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"


For i = 2 To ws.UsedRange.Rows.Count
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

Ticker = ws.Cells(i, 1).Value
tsv = tsv + ws.Cells(i, 7).Value
year_open = ws.Cells(i, 3).Value
year_close = ws.Cells(i, 6).Value
Yearly_Change = year_close - year_open
percent_change = year_close / year_open

ws.Cells(Summary_Table_Row, 9).Value = Ticker
ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
ws.Cells(Summary_Table_Row, 11).Value = percent_change
ws.Cells(Summary_Table_Row, 12).Value = tsv
Summary_Table_Row = Summary_Table_Row + 1
tsv = 0

ws.Columns("K").NumberFormat = "0.00%"

End If
Next i

Dim rng As Range
Dim lastRow As Long
For j = 2 To lastRow
lastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
Set rng = Range(J2 & lastRow)
If cell.Value >= 0 Then
Range(cell.Address).Interior.Color = vbGreen
ElseIf cell.Value < 0 Then
Range(cell.Address).Interior.Color = vbRed
End If




'code below did not fill color
'YearLastRow = Cells(Rows.Count, "J").End(xlUp).Row
'For j = 2 To YearLastRow
'If YearLastRow >= 0 Then
'YearLastRow.Interrior.Color = vbGreen
'ElseIf YearLastRow < 0 Then
'YearLastRow.Interrior.Color = vbRed
'End If

Next
Next
End Sub
