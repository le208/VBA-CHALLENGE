Attribute VB_Name = "Module1"
Sub ticker()
For Each ws In Worksheets

Dim TicketName As String
Dim YearChange, PercentChange, Volume As Double
Dim LastRow As Long
Dim FirstRangeRow, SummaryTableRow As Integer
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Volume = 0
FirstRangeRow = 2
SummaryTableRow = 2

For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            TickerName = ws.Cells(i, 1).Value
            YearChange = ws.Cells(i, 6).Value - ws.Cells(FirstRangeRow, 3).Value
            ws.PercentChange = YearChange / ws.Cells(FirstRangeRow, 3).Value
            ws.Range("I" & SummaryTableRow).Value = TickerName
            ws.Range("J" & SummaryTableRow).Value = YearChange
            ws.Range("K" & SummaryTableRow).Value = PercentChange
            ws.Range("L" & SummaryTableRow).Value = Volume
            If ws.Cells(SummaryTableRow, 10) > 0 Then
                ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
                Else
                ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
            End If
            SummaryTableRow = SummaryTableRow + 1
            TotalVolume = 0
            FirstRangeRow = i + 1
        Else
            Volume = Volume + ws.Cells(i, 7).Value
            FirstRangeRow = i - (i - FirstRangeRow)
    End If
    Next i
    Next ws
End Sub

