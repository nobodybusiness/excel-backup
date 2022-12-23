'pomocnicza funcja do okreslania z ktorej do ktroej apteki przenoscic zaopatrzenie
Sub basicsort()

'liczba aptek
    liczapt = Cells(2, 1)
'wsp. pkt startowego
    pktstx = 1
    pktsty = 6
'przesunięcie względem początku układu
    dy = 0
'potencjal po y
    py = 1
'nazwa po y
    nazwa = 4
'pomocnicza ark2, zczytanie wart po y (numer czy nazwa)
    pomark2 = 5
'przejscie do nastepnych wierszy _
"Główna pętla"
    While Cells(pktsty + dy, pktstx) <> ""
'wstepne sumy zerowe
        sumzerzap = 1
        sumzernad = 1
'pętla sprawdzajaca "sumy zerowe"
        While sumzerzap <> 0 And sumzernad <> 0
'wyszukiwanie i zerownie rownych wartosci
            For j = 5 To 4 + liczapt
                For i = 6 + liczapt To 5 + 2 * liczapt
                    If Cells(pktsty + dy, j) = 0 Or Cells(py, j) = 0 Or Cells(py, i) = 0 Then 'warunek 0 potencjalu
                        'do nothing
                    ElseIf Cells(pktsty + dy, j) = Cells(pktsty + dy, i) Then
'funkcja wpis
                        zmy = pktsty + dy
 'y=5 to nazwa!
                        ark1 = Cells(5, i) 'nazwa komorki z nadmiarem
'tu pomark2=5
                        ark2 = Cells(pomark2, j) 'nazwa/numer komorki z zapot.
                        ark3 = ActiveSheet.Name ' zczytanie nazwy bierzacego arkusza
                        zmn = Cells(pktsty + dy, j) ' liczb do przeniesienia
                        Call wpis(zmy, ark1, ark2, zmn, ark3)
'zerowanie komorek na podst arkuszu
                        Cells(pktsty + dy, j) = 0
                        Cells(pktsty + dy, i) = 0
                    'CALL COS TAM
                    Else
                        'do nothing
                    End If
                Next i
            Next j
'sprawdzenie czy nie wyzerowano wszystkiego
            newsumzerzap = 0 'nowe sumy zerowe
            newsumzernad = 0
            For i = 5 To 4 + liczapt
                newsumzerzap = newsumzerzap + Cells(py, i) * Cells(pktsty + dy, i)
            Next i
            For j = 6 + liczapt To 5 + 2 * liczapt
                newsumzernad = newsumzernad + Cells(py, j) * Cells(pktsty + dy, j)
            Next j
            If newsumzerzap = 0 Or newsumzernad = 0 Then
                'do nothing
            Else
'gdy nie zerowe sumy po I wyrownaniu
'wyszkuanie najwiekszego zapotszebowania
                k = 0 'pomocnicza wartosc przechowujaca dane
                xmaxzap = 0 ' miejsce max zapotszeb
                For j = 5 To 4 + liczapt
                    If Cells(pktsty + dy, j) > k And Cells(py, j) <> 0 Then 'warunek 0 potencjalu
                        k = Cells(pktsty + dy, j)
                        xmaxzap = j
                    Else
                        'do nothing
                    End If
                Next j
'wyszukanie największego nadmiaru
                k = 0 'przechowywanie wartosci
                xmaxnad = 0 'miejsce max nadmiaru
                For i = 6 + liczapt To 5 + 2 * liczapt
                    If Cells(pktsty + dy, i) > k And Cells(py, i) <> 0 Then 'warunek 0 potencjalu
                        k = Cells(pktsty + dy, i)
                        xmaxnad = i
                    Else
                        'do nothing
                    End If
                Next i
'sprawdzenie czy zapotszebowanie w pelni splenione
                If Cells(pktsty + dy, xmaxnad) > Cells(pktsty + dy, xmaxzap) Then
'funkcja wpis
                        zmy = pktsty + dy
 'y=5 to nazwa!
                        ark1 = Cells(5, xmaxnad) 'nazwa komorki z nadmiarem
'tu pomark2=5
                        ark2 = Cells(pomark2, xmaxzap) 'nazwa/numer komorki z zapot.
                        ark3 = ActiveSheet.Name ' zczytanie nazwy bierzacego arkusza
                        zmn = Cells(pktsty + dy, xmaxzap) ' liczb do przeniesienia
                        Call wpis(zmy, ark1, ark2, zmn, ark3)
'zerowanie komorek na podst arkuszu
                    Cells(pktsty + dy, xmaxnad) = Cells(pktsty + dy, xmaxnad) - Cells(pktsty + dy, xmaxzap)
                    Cells(pktsty + dy, xmaxzap) = 0
                
                ElseIf Cells(pktsty + dy, xmaxzap) > Cells(pktsty + dy, xmaxnad) Then
'funkcja wpis
                        zmy = pktsty + dy
 'y=5 to nazwa!
                        ark1 = Cells(5, xmaxnad) 'nazwa komorki z nadmiarem
'tu pomark2=5
                        ark2 = Cells(pomark2, xmaxzap) 'nazwa/numer komorki z zapot.
                        ark3 = ActiveSheet.Name ' zczytanie nazwy bierzacego arkusza
                        zmn = Cells(pktsty + dy, xmaxnad) ' liczb do przeniesienia
                        Call wpis(zmy, ark1, ark2, zmn, ark3)
'zerowanie komorek na podst arkuszu
                    Cells(pktsty + dy, xmaxzap) = Cells(pktsty + dy, xmaxzap) - Cells(pktsty + dy, xmaxnad)
                    Cells(pktsty + dy, xmaxnad) = 0
                
                Else
                    'do nothing
                End If
            End If
'przeliczanie od nowa sum zerowych
            sumzerzap = 0
            sumzernad = 0
            For i = 5 To 4 + liczapt
                sumzerzap = sumzerzap + Cells(py, i) * Cells(pktsty + dy, i)
            Next i
            For j = 6 + liczapt To 5 + 2 * liczapt
                sumzernad = sumzernad + Cells(py, j) * Cells(pktsty + dy, j)
            Next j
        Wend
'bezposredni przeskok do nast. wiersza
        dy = dy + 1
    Wend
    
End Sub
Function wpis(y, ark1, ark2, n, ark3)
'przypiasnie nazw produktu
    a = Cells(y, 1)
    b = Cells(y, 2)
    c = Cells(y, 3)
    Application.Goto (ActiveWorkbook.Sheets(ark1).Range("A1")) ' przejscie do danego arkusza
    p = 6 ' startowy punkt wpisywania po y
    r = 0 ' zmienna nastepnych linijek
    While Cells(p + r, 1) <> ""
        r = r + 1
    Wend
    
    Cells(p + r, 1) = a
    Cells(p + r, 2) = b
    Cells(p + r, 3) = c
    
 ' wspolrzedna boczna
    Z = 4
    Cells(p + r, Z) = n
    Cells(p + r, Z + 1) = ark2
    Application.Goto (ActiveWorkbook.Worksheets(ark3).Range("A1"))
End Function

Sub tworzenie_nowych_arkuszy()
'liczba aptek
    n = Cells(2, 1)
    nazwaarkuszapodst = ActiveSheet.Name
'petla tworzenia arkuszy
    For i = 1 To n
        Sheets.Add.Name = Cells(5, 4 + i)
        Application.Goto (ActiveWorkbook.Worksheets(nazwaarkuszapodst).Range("A1"))
    Next i
End Sub
