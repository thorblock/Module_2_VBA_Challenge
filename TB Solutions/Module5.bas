Attribute VB_Name = "Module5"
Sub SummarySummaryCreator()

Dim LowestYear, HighestYear, BigVolume As Variant
Dim TickerCol, YearCol, VolumeCol As Range

SumlastRow = Cells(Rows.Count, 9).End(xlUp).Row

'summary of summary table setup
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

' Set Range Columns
Set TickerCol = Range("I2:I" & Rows.Count)
Set YearCol = Range("K2:K" & Rows.Count)
Set VolumeCol = Range("L2:L" & Rows.Count)
' I can't believe I forgot about the excel premade functions
LowestYear = Application.WorksheetFunction.Min(YearCol)
HighestYear = Application.WorksheetFunction.Max(YearCol)
BigVolume = Application.WorksheetFunction.Max(VolumeCol)

' loop + incremental; value comparison using previously determined figures^
For i = 2 To SumlastRow
' best ticker with highest year %
  If Cells(i, 11).Value = HighestYear Then
  Range("O2").Value = Cells(i, 9).Value
  Range("P2").Value = HighestYear
  End If
' worst ticker with lowest year %
  If Cells(i, 11).Value = LowestYear Then
  Range("O3").Value = Cells(i, 9).Value
  Range("P3").Value = LowestYear
  End If
' ticker with the largest volume
  If Cells(i, 12).Value = BigVolume Then
  Range("O4").Value = Cells(i, 9).Value
  Range("P4").Value = BigVolume
  End If

Next i

End Sub
