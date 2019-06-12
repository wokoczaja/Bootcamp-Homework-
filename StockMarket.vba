Option Explicit

Sub StockMarket()

Dim Ticker As String
Dim TotalVolume As Double
Dim Sheet As Worksheet
Dim i As Long
'SummaryRow is a tracker for row location of ticker and total volume information in the summary
Dim SummaryRow As Long
Dim LastRow As Long

'Set-up ability to work through worksheets in workbook

For Each Sheet In Worksheets
    Sheet.Cells(1, 9).Value = "Ticker"
    Sheet.Cells(1, 10).Value = "Total Volume"

'Initialize Volume count with volume 0
    TotalVolume = 0
'Set initial counter for row number in summary
    SummaryRow = 2
'Find last row of data on each worksheet
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A") _
        .End(xlUp).Row

'Inserting "i" loop to run through the rows of data to get the total volumes
    For i = 2 To LastRow
          TotalVolume = TotalVolume + Sheet.Cells(i, 7).Value
          
          If Sheet.Cells(i, 1).Value <> Sheet.Cells(i + 1, 1).Value Then
             Sheet.Cells(SummaryRow, 9).Value = Sheet.Cells(i, 1).Value
             Sheet.Cells(SummaryRow, 10).Value = TotalVolume
'Add 1 to SummaryRow to move summary information in summary table
            SummaryRow = SummaryRow + 1
'Reset total volume
            TotalVolume = 0
          End If
    Next i

Next Sheet

End Sub
