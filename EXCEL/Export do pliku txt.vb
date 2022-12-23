'Export do pliku txt
Sub export_pts()
    myFile = ActiveWorkbook.Path & "\pts.txt" 'filepath
    Open myFile For Output As #1  'otwarcie do edycji
    i = 3 ' punkt startowy dla rows
    While Cells(i, 1) <> ""
        Write #1, Cells(i, 1)
        Write #1, Cells(i, 2)
        Write #1, Cells(i, 5)
        Write #1, Cells(i, 6)
        Write #1, Cells(i, 3)
        Write #1, Cells(i, 4)
        i = i + 1
    Wend
    Close #1
End Sub