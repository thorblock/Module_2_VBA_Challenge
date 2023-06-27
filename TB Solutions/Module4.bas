Attribute VB_Name = "Module4"
Sub CellFormats()


SumlastRow = Cells(Rows.Count, 9).End(xlUp).Row

For f = 2 To SumlastRow
    If Cells(f, 10).Value <= 0 Then
        Cells(f, 10).Interior.ColorIndex = 3
        Cells(f, 11).Interior.ColorIndex = 3
    Else
        Cells(f, 10).Interior.ColorIndex = 4
        Cells(f, 11).Interior.ColorIndex = 4
    End If
Next f

End Sub
