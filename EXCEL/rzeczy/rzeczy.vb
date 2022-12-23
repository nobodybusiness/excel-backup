Sub program()
'   Application.EnableCancelKey = xlDisabled
form.CheckBox1.Caption = ActiveWorkbook.Sheets(22).Range("F5")
form.CheckBox2.Caption = ActiveWorkbook.Sheets(22).Range("F6")
form.CheckBox3.Caption = ActiveWorkbook.Sheets(22).Range("F7")
form.CheckBox4.Caption = ActiveWorkbook.Sheets(22).Range("F8")
form.CheckBox5.Caption = ActiveWorkbook.Sheets(22).Range("F9")
form.CheckBox6.Caption = ActiveWorkbook.Sheets(22).Range("F10")
form.CheckBox7.Caption = ActiveWorkbook.Sheets(22).Range("F11")
form.CheckBox8.Caption = ActiveWorkbook.Sheets(22).Range("F12")
form.CheckBox9.Caption = ActiveWorkbook.Sheets(22).Range("F13")
form.CheckBox10.Caption = ActiveWorkbook.Sheets(22).Range("F14")
form.CheckBox11.Caption = ActiveWorkbook.Sheets(22).Range("F15")
form.CheckBox12.Caption = ActiveWorkbook.Sheets(22).Range("F16")
form.CheckBox13.Caption = ActiveWorkbook.Sheets(22).Range("F17")
form.CheckBox14.Caption = ActiveWorkbook.Sheets(22).Range("F18")
form.CheckBox15.Caption = ActiveWorkbook.Sheets(22).Range("F19")
form.CheckBox16.Caption = ActiveWorkbook.Sheets(22).Range("F20")
form.CheckBox17.Caption = ActiveWorkbook.Sheets(22).Range("F21")
form.CheckBox18.Caption = ActiveWorkbook.Sheets(22).Range("F22")
form.CheckBox19.Caption = ActiveWorkbook.Sheets(22).Range("F23")
form.CheckBox20.Caption = ActiveWorkbook.Sheets(22).Range("F24")

x = 0
''Nazwywanie domyślne arkusza; pętla czyści nazwy arkuszów na domyślne
'               w przypadku usunięcia rekordu z tabeli "Nazwy działów"
While x < 21 - 2
        MyRange = Col_Letter(6) & x + 5
        If ActiveWorkbook.Sheets(22).Range(MyRange) = "-" Or ActiveWorkbook.Sheets(22).Range(MyRange) = "" Then
        Sheets(x + 2).Name = x + 2
    Else
        Sheets(x + 2).Name = ActiveWorkbook.Sheets(22).Range(MyRange)
    End If
    x = x + 1
Wend

form.Aktualizuj = True
form.Aktualizuj = Click
form.dzisiaj = True
form.show
End Sub

Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function


