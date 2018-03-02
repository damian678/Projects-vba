Sub tabliczka()
    Workbooks.Add
    Workbooks(2).Activate
    ActiveWorkbook.Worksheets(1).Name = "Tabliczka mnożenia"
    Range("A1:P1").Merge
    Cells(1, 1) = "Tabliczka mnożenia"
    Cells(1, 1).Font.FontStyle = "Italic"
    Cells(1, 1).Font.Underline = True
    Cells(1, 1).Font.Size = 13
    Range("B2:P2").Font.FontStyle = "Bold Italic"
    Range("B2:P2").Font.Size = 11
    Range("B2:P2").Interior.ColorIndex = 6
    Range("A3:A17").Font.FontStyle = "Bold Italic"
    Range("A3:A17").Font.Size = 11
    Range("A3:A17").Interior.Color = rgbOrange
    Range("B3:P17").Font.Size = 12
    Dim j As Integer
    Dim i As Integer
    Rows("1:17").RowHeight = 18
    Columns("A:P").ColumnWidth = 6
    Rows("1:17").HorizontalAlignment = xlCenter
    Rows("1:17").VerticalAlignment = xlCenter
    Columns("A:P").HorizontalAlignment = xlCenter
    Columns("A:P").VerticalAlignment = xlCenter
    Worksheets(1).PageSetup.LeftMargin = Application.CentimetersToPoints(1)
    Worksheets(1).PageSetup.RightMargin = Application.CentimetersToPoints(1)
    Worksheets(1).PageSetup.TopMargin = Application.CentimetersToPoints(1)
    Worksheets(1).PageSetup.BottomMargin = Application.CentimetersToPoints(1)
    Worksheets(1).PageSetup.Orientation = xlLandscape
    For i = 2 To 16
        For j = 1 To 15
            Cells(2, i) = i - 1
            Cells(j + 2, 1) = j
            Cells(i, j).BorderAround LineStyle = xlContinous
            
        Next j
    Next i
    For i = 3 To 17
        For j = 2 To 16
        Cells(i, j) = (i - 2) * (j - 1)
            If j - i = -1 Then
            Cells(i, j).Interior.ColorIndex = 4
            End If
        Next j
    Next i
    For i = 2 To 17
        For j = 1 To 16
        Cells(i, j).BorderAround LineStyle = xlContinous
            Cells(i, j).BorderAround LineStyle = xlContinous
            Cells(2, j).Borders(xlTop).Weight = xlMedium
            Cells(2, j).Borders(xlBottom).Weight = xlMedium
            Cells(i, 1).Borders(xlRight).Weight = xlMedium
            Cells(17, j).Borders(xlBottom).Weight = xlMedium
            Cells(i, 16).Borders(xlRight).Weight = xlMedium
        Next j
    Next i
End Sub



Sub Pascal()
    Workbooks.Add
    Workbooks(2).Activate
    ActiveWorkbook.Worksheets(1).Name = "Trójkąt Pascala"
    Dim i As Integer
    Dim j As Integer
    Rows("1:25").RowHeight = 18
    Columns("A:Y").ColumnWidth = 8
    Rows("1:25").HorizontalAlignment = xlCenter
    Rows("1:25").VerticalAlignment = xlCenter
    Columns("A:Y").HorizontalAlignment = xlCenter
    Columns("A:Y").VerticalAlignment = xlCenter
    For i = 1 To 24
        For j = 1 To 24
           If i >= j Then
            Cells(i, i) = 1
            Cells(i + 1, 1) = 1
            Cells(i + 1, j + 1) = Cells(i, j) + Cells(i, j + 1)
            Cells(1, 1).Interior.ColorIndex = 4
            Cells(i + 1, 1).Interior.ColorIndex = 4
            Cells(i + 1, j + 1).Interior.ColorIndex = 3
            Cells(i + 1, j + 1).BorderAround LineStyle = xlContinous
            Cells(i + 1, 1).BorderAround LineStyle = xlContinous
            Cells(1, 1).BorderAround LineStyle = xlContinous
            Cells(i + 1, i + 1).Interior.ColorIndex = 4
            Else: Cells(i, j) = " "
            End If
        Next j
    Next i
End Sub
