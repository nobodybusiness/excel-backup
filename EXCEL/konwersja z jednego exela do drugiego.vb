'konwersja z jednego exela do drugiego
Sub konwersja_z_ZR()
On Error GoTo ErrorHandler
'*************************************
'SELECTING FILE TO CONVERT
'directory current
    curWorkbookPath = Application.ActiveWorkbook.Path
    curWorkbookName = Application.ActiveWorkbook.Name
    Dim fd As Office.FileDialog
    Dim strFile As String
 
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
 
    With fd
 
    .Filters.Clear
    .Filters.Add "Excel Files", "*.xlsx?", 1
    .Title = "Wybierz Zalacznik Rejonizacyjna do konwersji"
    .AllowMultiSelect = False
 
    .InitialFileName = curWorkbookPath
        If .Show = True Then
'directory of ZR
            zrWorkbookPath = .SelectedItems(1)
        End If
    End With

'*************************************
'OPENING FILE ZR AND TAKING DATA
    Workbooks.Open (zrWorkbookPath)
    nameOfZR = ActiveWorkbook.Name
'creating list for each type
    Dim uliceVal, numeryVal, HHVal, LUVal, podVal, PEVal As Object
    Set uliceVal = CreateObject("System.Collections.ArrayList")
    Set numeryVal = CreateObject("System.Collections.ArrayList")
    Set HHVal = CreateObject("System.Collections.ArrayList")
    Set LUVal = CreateObject("System.Collections.ArrayList")
    Set podVal = CreateObject("System.Collections.ArrayList")
    Set PEVal = CreateObject("System.Collections.ArrayList")
    
'add data to arraylists
    Dim multiArray(1 To 100, 1 To 6) As Variant
    n = 22
    i = 1
    While (Cells(n, 1) <> "")
    'ulice
        multiArray(i, 1) = Cells(n, 30)
    'numery
        multiArray(i, 2) = Cells(n, 31)
    'HH and LU
        If (Cells(n, 34) <> "") Then
            multiArray(i, 3) = Cells(n, 41)
            multiArray(i, 4) = 0
        ElseIf (Cells(n, 38) <> "") Then
            multiArray(i, 4) = Cells(n, 41)
            multiArray(i, 3) = 0
        Else
            multiArray(i, 3) = 1
            multiArray(i, 4) = 0
        End If
    'pod
        multiArray(i, 5) = Cells(n, 43)
    'PE
        multiArray(i, 6) = Cells(n, 7)
    i = i + 1
    n = n + 1
    Wend
' zamkniecie arkusza
    Workbooks(nameOfZR).Close SaveChanges:=False
'petla wartosci poczatkowe
    i = 1
    n = 1
'petla
    While (multiArray(i, 1) <> "")
'ulica
        Cells(2 + n, 3) = Left(multiArray(i, 1), Len(multiArray(i, 1)) - 5)
'numer
        Cells(2 + n, 4) = multiArray(i, 2)
'HH
        Cells(2 + n, 5) = multiArray(i, 3)
'LU
        Cells(2 + n, 6) = multiArray(i, 4)
'pod
        threeLeftChar = Left(multiArray(i, 5), 3)
            If (threeLeftChar = "ZJN") Then
                Cells(2 + n, 7) = "instalacja napowietrzna"
            Else
            Cells(2 + n, 7) = "instalacja doziemna"
            End If
'PE
        PENumber = Left(multiArray(i, 6), InStrRev(multiArray(i, 6), "Słup") - 2)
        PEUlica = Mid(multiArray(i, 6), InStrRev(multiArray(i, 6), "ul."))
        Cells(2 + n, 8) = PENumber & " " & PEUlica
'uzg
        typeOfZJ = Mid(multiArray(i, 5), 4)
        If (typeOfZJ = "30") Then
            Cells(2 + n, 9) = "Wymagane dostawienie słupa na działce właściciela w momencie zainteresowania usługą"
        Else
            Cells(2 + n, 9) = "Brak uzgodnienia z właścicielem"
        End If
'warunek jednego adresu
        If (multiArray(i, 1) = multiArray(i + 1, 1) And multiArray(i, 2) = multiArray(i + 1, 2)) Then
            Cells(2 + n, 6) = multiArray(i + 1, 4)
            i = i + 1
        End If
        i = i + 1
        n = n + 1
    Wend
MsgBox "Sprawdz czy liczba HH/LU zgadza się z ZR!"
Exit Sub
ErrorHandler:
    MsgBox "Coś poszło nie tak, spróbuj ponownie lub zapytaj się Piotrka ;)"
End Sub
