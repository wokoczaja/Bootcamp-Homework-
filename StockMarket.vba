Option Explicit

Sub StockMarket()

Dim Ticker As String
Dim TotalVolume As Double
Dim Sheet As Worksheet
Dim i As Long
'SummaryRow is a tracker for row location of ticker and total volume information in the summary
Dim SummaryRow As Long
Dim LastRow As Long
Dim ClosePrice As Double
Dim OpenPrice As Double
Dim YearlyChange As Double

'Set-up ability to work through worksheets in workbook

For Each Sheet In Worksheets
    Sheet.Cells(1, 9).Value = "Ticker"
    Sheet.Cells(1, 10).Value = "Total Volume"
    Sheet.Cells(1, 11).Value = "Yearly Change"


'Initialize Volume count with volume 0
    TotalVolume = 0
'Set initial counter for row number in summary
    SummaryRow = 2
'Find last row of data on each worksheet
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A") _
        .End(xlUp).Row

'Inserting "i" loop to run through the rows of data to get the total volumes and prices
    For i = 2 To LastRow
          If Sheet.Cells(i, 1).Value <> Sheet.Cells(i - 1, 1).Value Then
              OpenPrice = Sheet.Cells(i, 3).Value
          End If
          TotalVolume = TotalVolume + Sheet.Cells(i, 7).Value
          
          If Sheet.Cells(i, 1).Value <> Sheet.Cells(i + 1, 1).Value Then
             Sheet.Cells(SummaryRow, 9).Value = Sheet.Cells(i, 1).Value
             Sheet.Cells(SummaryRow, 10).Value = TotalVolume
             ClosePrice = Sheet.Cells(i, 6).Value
             YearlyChange = ClosePrice - OpenPrice
             Sheet.Cells(SummaryRow, 11).Value = YearlyChange
'Add 1 to SummaryRow to move summary information in summary table
            SummaryRow = SummaryRow + 1
'Reset total volume and prices
            TotalVolume = 0
            OpenPrice = 0
            ClosePrice = 0
            YearlyChange = 0
          End If
    Next i

Next Sheet

End Sub
