'porównanie projektowanych materialow do uzytych
Sub RM()
'przypisanie n dla petli pozniej
n = 1
'przejscie od ZAM
    Worksheets("ZAM").Activate
'wartosci poczatkowe
ileZam = 2
colZamIndeks = 1
stringToNumberFlag = 0 ' flaga do osznaczenia warunku konwersji stringa na number
'petla przejscia po zamowieniach
While (Cells(ileZam, colZamIndeks) <> "")
    Cells(ileZam, colZamIndeks) = Cells(ileZam, colZamIndeks) + 0
    curIndeks = Cells(ileZam, colZamIndeks)
    curZamNaz = Cells(ileZam, colZamIndeks + 1)
    curZamIlo = Cells(ileZam, colZamIndeks + 2)
'przejscie do ZAP
    Worksheets("ZAP").Activate
'wartosci poczatkowe
    ileZap = 2
    colZapIndeks = 1
'petla przypisania numeracji
    If (stringToNumberFlag = 0) Then
        While (Cells(ileZap, colZapIndeks) <> "")
            Cells(ileZap, colZapIndeks) = Cells(ileZap, colZapIndeks) + 0
            ileZap = ileZap + 1
        Wend
        ileZap = 2
        colZapIndeks = 1
        stringToNumberFlag = 1
    End If
'petla przejscia porownan indeksow ZAM-ZAP
    While (Cells(ileZap, colZapIndeks) <> curIndeks And Cells(ileZap, colZapIndeks) <> "")
        ileZap = ileZap + 1
    Wend
    curZapIlo = Cells(ileZap, colZapIndeks + 2)
'przejscie do RM
    Worksheets("RM").Activate
    Cells(ileZam, colZamIndeks) = curIndeks
    Cells(ileZam, colZamIndeks + 1) = curZamNaz
    Cells(ileZam, colZamIndeks + 2) = curZamIlo
    Cells(ileZam, colZamIndeks + 3) = curZapIlo
    deltaIlo = curZamIlo - curZapIlo
'ustalenie rozrzucenie/domowienie
    If (deltaIlo >= 0) Then
         Cells(ileZam, colZamIndeks + 4) = deltaIlo
    Else
        Cells(ileZam, colZamIndeks + 5) = -deltaIlo
    End If
'przejscie do ZAM
    Worksheets("ZAM").Activate
    ileZam = ileZam + 1
Wend
'---------------------------------
'Wyszukanie brakujacych zamowien

'przejscie do ZAP
Worksheets("ZAP").Activate
ileZap = 2
colZapIndeks = 1
'petla przejscia po zapotrzebowaniu
While (Cells(ileZap, colZapIndeks) <> "")
    curIndeks = Cells(ileZap, colZapIndeks)
    curZapNaz = Cells(ileZap, colZapIndeks + 1)
    curZapIlo = Cells(ileZap, colZapIndeks + 2)
'przejscie do ZAM
    Worksheets("ZAM").Activate
    ileZam = 2
    colZamIndeks = 1
    istZam = 0
'petla wyszukania czy istnieje zamowienie
    While (Cells(ileZam, colZamIndeks) <> curIndeks And Cells(ileZam, colZamIndeks) <> "")
        ileZam = ileZam + 1
    Wend
'spr istn zamowienia
    If (Cells(ileZam, colZamIndeks) = curIndeks) Then
        istnZam = 1
    Else
        istnZam = 0
    End If
'jezeli NIE istnieje zamowienie
    If (istnZam = 0) Then
'przejscie do RM
        Worksheets("RM").Activate
'petla wyszukania pierwszego wolnego pola -> zalezne od globalnego n
        While (Cells(n, 1) <> "")
            n = n + 1
        Wend
        
        Cells(n, colZapIndeks) = curIndeks
        Cells(n, colZapIndeks + 1) = curZapNaz
        Cells(n, colZapIndeks + 2) = 0
        Cells(n, colZapIndeks + 3) = curZapIlo
        Cells(n, colZapIndeks + 4) = 0
        Cells(n, colZapIndeks + 5) = curZapIlo
    End If
'przejscie do ZAP
    Worksheets("ZAP").Activate
    ileZap = ileZap + 1
Wend

'---------------------------------
'przejscie do RM - wynikowe
    Worksheets("RM").Activate
End Sub

Sub delete()
Worksheets("RM").Activate
n = 2
While (Cells(n, 1) <> "")
    Cells(n, 1) = ""
    Cells(n, 2) = ""
    Cells(n, 3) = ""
    Cells(n, 4) = ""
    Cells(n, 5) = ""
    Cells(n, 6) = ""
    n = n + 1
Wend
End Sub

Sub TEST()
Cells(2, 1) = Cells(2, 1) + 0
End Sub
