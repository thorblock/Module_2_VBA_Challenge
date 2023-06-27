Attribute VB_Name = "Module1"
Sub LongConversionDate()

' We want to go from an 8 digit date to MM/DD/YYYY
Dim Yr, Mo, Da As Variant

'lastrow
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow
    ' grabs 4 character starting from the left
    Yr = Left(Cells(i, 2).Value, 4)
    ' grabs 2 characters starting at character 5; second arg of Mid
    Mo = Mid(Cells(i, 2).Value, 5, 2)
    ' grabs 2 characters starting from the right
    Da = Right(Cells(i, 2).Value, 2)
    
    ' Replaces the cell contents that we just read from with the Month/Day/Year format
    ' DateSerial where u at, probably should add a check to make sure it's an eight digit date
    Cells(i, 2).Value = Mo & "/" & Da & "/" & Yr
    
Next i

End Sub
