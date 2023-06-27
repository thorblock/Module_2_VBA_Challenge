Attribute VB_Name = "Module2"
Sub SummaryCreator():

' variables
Dim tickerUnique As String
Dim openPrice, closePrice, totalVolume, yearlyChange, yearlyPercent As Double
Dim SummaryRow, lastRow As Long

' id lastrow, assumes table uniformity
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'summary setup on each page
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Year Percent Change"
Range("L1").Value = "Total Volume"
SummaryRow = 2

' For loop + incremental, running If variants so I don't have to haggle with order
    For i = 2 To lastRow
        'Volume increase in same elements
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
            totalVolume = totalVolume + Cells(i, 7).Value
        End If
        ' If to hold unique opening price
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            openPrice = Cells(i, 3).Value
        End If
        ' If to compile and then reset
        If Cells(i, 1).Value <> Cells(i + 1, 1) Then
        ' unique ticker
            tickerUnique = Cells(i, 1).Value
        ' unique close price
            closePrice = Cells(i, 6).Value
        ' add volume of last matching row
            totalVolume = totalVolume + Cells(i, 7).Value
        ' print ticker name
            Range("I" & SummaryRow).Value = tickerUnique
        ' yearlyChange calculation and print, keeping yearlyChange and yearlyPercent as variables for bonus table
            yearlyChange = closePrice - openPrice
            Range("J" & SummaryRow).Value = yearlyChange
            yearlyPercent = (yearlyChange / openPrice)
            Range("K" & SummaryRow).Value = yearlyPercent
        ' print volume total
            Range("L" & SummaryRow).Value = totalVolume
        ' increment loop
            SummaryRow = SummaryRow + 1
        ' clear the volume total, so a new tick symbol begins
            totalVolume = 0
        End If
    Next i
    
End Sub

