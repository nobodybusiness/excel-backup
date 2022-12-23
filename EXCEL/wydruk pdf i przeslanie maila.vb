'wydruk pdf i przeslanie maila
Sub Wydruk_pdf()
    'przejscie po imionach i nazwiskach
    Sheets(1).Activate 'przeskok do bazy danych
    n = 2 'start w rows
    While Cells(n, 2) <> ""
        
        pesel = Cells(n, 2)
        
    Sheets(2).Activate 'przeskok do karty informacyjnej
    
        Cells(13, 4) = pesel
        
        
    sciezka = ActiveWorkbook.Path & "\" & Cells(16, 4) & " -umowa.pdf" ' przypisanie imienia i nazwiska
    
    Sheets(3).Activate 'przeskok do umowy
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=sciezka, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=True, _
        From:=1, _
        To:=6, _
        OpenAfterPublish:=False
        
    Sheets(2).Activate 'przeskok do karty informacyjnej
    sciezka = ActiveWorkbook.Path & "\" & Cells(16, 4) & " -oswiadczenie.pdf" ' przypisanie imienia i nazwiska
    
    Sheets(4).Activate 'przeskok do oswiadczenia
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=sciezka, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=True, _
        From:=1, _
        To:=1, _
        OpenAfterPublish:=False
    If Sheets(6).Cells(10, 3) = "TAK" Then
        Call email_send
    End If
    Sheets(1).Activate 'przeskok do bazy danych
    n = n + 1
    
    Wend
    If Sheets(6).Cells(10, 3) <> "TAK" Then
        MsgBox ("Nie wysłano maili zgodnie z zaznaczeniem w arkuszu mail.")
    End If

    
End Sub
Sub email_send()
'uruchomic tools->references->microsoft outolook 2012
'On Error GoTo ErrHandler
    ' SET Outlook APPLICATION OBJECT.
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
    
    ' CREATE EMAIL OBJECT.
    Dim objEmail As Object
    Set objEmail = objOutlook.CreateItem(olMailItem)
    recivers = Sheets(2).Cells(42, 2) & ";" & Sheets(2).Cells(42, 6)
    subj = Sheets(6).Cells(2, 2)
    Text_mail = Sheets(6).Cells(3, 2)
    With objEmail
        .To = recivers
        .Subject = subj
        .Body = Text_mail
        .Attachments.Add ActiveWorkbook.Path & "\" & Sheets(2).Cells(16, 4) & " -oswiadczenie.pdf"
        .Attachments.Add ActiveWorkbook.Path & "\" & Sheets(2).Cells(16, 4) & " -umowa.pdf"
        .Send    ' DISPLAY MESSAGE.
    End With
    
    ' CLEAR.
    Set objEmail = Nothing:    Set objOutlook = Nothing
    GoTo Koniec
ErrHandler:
    MsgBox ("Wystapił bład. Maile nie zostały wysłane. Sprawdz uwagi w arkuszu mail")
    End
Koniec:
End Sub