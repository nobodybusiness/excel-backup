VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form 
   Caption         =   "Okno wyboru"
   ClientHeight    =   11115
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10275
   OleObjectBlob   =   "form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Aktualizuj_click()
Dim n As Integer
n = 2
While Cells(n, 1) <> ""
form.ListBox1.AddItem (Cells(n, 1))
n = n + 1
Wend
End Sub
Private Sub anuluj_Click()
End
End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub data_Change()

End Sub

Private Sub Dodaj_Click()

End Sub

Private Sub ilosc_Change()

End Sub

Private Sub dzisiaj_Click()
If form.dzisiaj.Value = False Then
    form.data.Text = ""
Else
    form.data.Text = Date
End If
End Sub

Private Sub Label4_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub Odejmij_Click()

End Sub

Private Sub ok_Click()
Dim n As Integer
n = 2
m = 0
k = form.ListBox1.Text
If k = "" Then
    MsgBox "Nie wybrano przedmiotu", , "B³¹d"
Else
    If form.data.Text = "" Then
        MsgBox "Nie wprowadzono daty, spróbuj ponownie", , "B³¹d"
    Else
      '  x = 0
       ' While x < 20
    
        If form.CheckBox1.Value = True Then
        m = m + 1
        End If
        
        If form.CheckBox2.Value = True Then
        m = m + 1
        End If
        If form.CheckBox3.Value = True Then
        m = m + 1
        End If
        If form.CheckBox4.Value = True Then
        m = m + 1
        End If
        If form.CheckBox5.Value = True Then
        m = m + 1
        End If
        If form.CheckBox6.Value = True Then
        m = m + 1
        End If
        If form.CheckBox7.Value = True Then
        m = m + 1
        End If
        If form.CheckBox8.Value = True Then
        m = m + 1
        End If
        If form.CheckBox9.Value = True Then
        m = m + 1
        End If
        If form.CheckBox10.Value = True Then
        m = m + 1
        End If
        If form.CheckBox11.Value = True Then
        m = m + 1
        End If
        If form.CheckBox12.Value = True Then
        m = m + 1
        End If
        If form.CheckBox13.Value = True Then
        m = m + 1
        End If
        If form.CheckBox14.Value = True Then
        m = m + 1
        End If
        If form.CheckBox15.Value = True Then
        m = m + 1
        End If
        If form.CheckBox16.Value = True Then
        m = m + 1
        End If
        If form.CheckBox17.Value = True Then
        m = m + 1
        End If
        If form.CheckBox18.Value = True Then
        m = m + 1
        End If
        If form.CheckBox19.Value = True Then
        m = m + 1
        End If
        If form.CheckBox20.Value = True Then
        m = m + 1
        End If
        If m = 1 Then
            While k <> Cells(n, 1)
            n = n + 1
            Wend
            If form.CheckBox1.Value = True Then
                Application.Goto (ActiveWorkbook.Sheets(2).Range("A1"))
                If form.Odejmij = True Then
                    Z = -form.ilosc.Value
                Else
                    Z = form.ilosc.Value
                End If
                Cells(n, 2) = Cells(n, 2) + Z
                k = 1
                While Cells(n, k) <> ""
                k = k + 1
                Wend
                Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
   
    
            ElseIf form.CheckBox2.Value = True Then
                Application.Goto (ActiveWorkbook.Sheets(3).Range("A1"))
            If form.Odejmij = True Then
        Z = -form.ilosc.Value
        Else
        Z = form.ilosc.Value
    End If
    Cells(n, 2) = Cells(n, 2) + Z
    k = 1
    While Cells(n, k) <> ""
        k = k + 1
    Wend
    Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
    
    
ElseIf form.CheckBox3.Value = True Then
    Application.Goto (ActiveWorkbook.Sheets(4).Range("A1"))
    If form.Odejmij = True Then
        Z = -form.ilosc.Value
        Else
        Z = form.ilosc.Value
    End If
    Cells(n, 2) = Cells(n, 2) + Z
    k = 1
    While Cells(n, k) <> ""
        k = k + 1
    Wend
    Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
    
    
ElseIf form.CheckBox4.Value = True Then
    Application.Goto (ActiveWorkbook.Sheets(5).Range("A1"))
    If form.Odejmij = True Then
        Z = -form.ilosc.Value
        Else
        Z = form.ilosc.Value
    End If
    Cells(n, 2) = Cells(n, 2) + Z
    k = 1
    While Cells(n, k) <> ""
        k = k + 1
    Wend
    Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
    
    
ElseIf form.CheckBox5.Value = True Then
    Application.Goto (ActiveWorkbook.Sheets(6).Range("A1"))
    If form.Odejmij = True Then
        Z = -form.ilosc.Value
        Else
        Z = form.ilosc.Value
    End If
    Cells(n, 2) = Cells(n, 2) + Z
    k = 1
    While Cells(n, k) <> ""
        k = k + 1
    Wend
    Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
    
    
ElseIf form.CheckBox6.Value = True Then
    Application.Goto (ActiveWorkbook.Sheets(7).Range("A1"))
    If form.Odejmij = True Then
        Z = -form.ilosc.Value
        Else
        Z = form.ilosc.Value
    End If
    Cells(n, 2) = Cells(n, 2) + Z
    k = 1
    While Cells(n, k) <> ""
        k = k + 1
    Wend
    Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
  
    
ElseIf form.CheckBox7.Value = True Then
    Application.Goto (ActiveWorkbook.Sheets(8).Range("A1"))
    If form.Odejmij = True Then
        Z = -form.ilosc.Value
        Else
        Z = form.ilosc.Value
    End If
    Cells(n, 2) = Cells(n, 2) + Z
    k = 1
    While Cells(n, k) <> ""
        k = k + 1
    Wend
    Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
   
    
ElseIf form.CheckBox8.Value = True Then
    Application.Goto (ActiveWorkbook.Sheets(9).Range("A1"))
    If form.Odejmij = True Then
        Z = -form.ilosc.Value
        Else
        Z = form.ilosc.Value
    End If
    Cells(n, 2) = Cells(n, 2) + Z
    k = 1
    While Cells(n, k) <> ""
        k = k + 1
    Wend
    Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
    
    
ElseIf form.CheckBox9.Value = True Then
    Application.Goto (ActiveWorkbook.Sheets(10).Range("A1"))
    If form.Odejmij = True Then
        Z = -form.ilosc.Value
        Else
        Z = form.ilosc.Value
    End If
    Cells(n, 2) = Cells(n, 2) + Z
    k = 1
    While Cells(n, k) <> ""
        k = k + 1
    Wend
    Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
    
    
ElseIf form.CheckBox10.Value = True Then
    Application.Goto (ActiveWorkbook.Sheets(11).Range("A1"))
    If form.Odejmij = True Then
        Z = -form.ilosc.Value
        Else
        Z = form.ilosc.Value
    End If
    Cells(n, 2) = Cells(n, 2) + Z
    k = 1
    While Cells(n, k) <> ""
        k = k + 1
    Wend
    Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
    
    
ElseIf form.CheckBox11.Value = True Then
    Application.Goto (ActiveWorkbook.Sheets(12).Range("A1"))
    If form.Odejmij = True Then
        Z = -form.ilosc.Value
        Else
        Z = form.ilosc.Value
    End If
    Cells(n, 2) = Cells(n, 2) + Z
    k = 1
    While Cells(n, k) <> ""
        k = k + 1
    Wend
    Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
    
    
ElseIf form.CheckBox12.Value = True Then
    Application.Goto (ActiveWorkbook.Sheets(13).Range("A1"))
    If form.Odejmij = True Then
        Z = -form.ilosc.Value
        Else
        Z = form.ilosc.Value
    End If
    Cells(n, 2) = Cells(n, 2) + Z
    k = 1
    While Cells(n, k) <> ""
        k = k + 1
    Wend
    Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
 
    
ElseIf form.CheckBox13.Value = True Then
    Application.Goto (ActiveWorkbook.Sheets(14).Range("A1"))
    If form.Odejmij = True Then
        Z = -form.ilosc.Value
        Else
        Z = form.ilosc.Value
    End If
    Cells(n, 2) = Cells(n, 2) + Z
    k = 1
    While Cells(n, k) <> ""
        k = k + 1
    Wend
    Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
   
    
ElseIf form.CheckBox14.Value = True Then
    Application.Goto (ActiveWorkbook.Sheets(15).Range("A1"))
    If form.Odejmij = True Then
        Z = -form.ilosc.Value
        Else
        Z = form.ilosc.Value
    End If
    Cells(n, 2) = Cells(n, 2) + Z
    k = 1
    While Cells(n, k) <> ""
        k = k + 1
    Wend
    Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
   
   
ElseIf form.CheckBox15.Value = True Then
    Application.Goto (ActiveWorkbook.Sheets(16).Range("A1"))
    If form.Odejmij = True Then
        Z = -form.ilosc.Value
        Else
        Z = form.ilosc.Value
    End If
    Cells(n, 2) = Cells(n, 2) + Z
    k = 1
    While Cells(n, k) <> ""
        k = k + 1
    Wend
    Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
 
    
ElseIf form.CheckBox16.Value = True Then
    Application.Goto (ActiveWorkbook.Sheets(17).Range("A1"))
    If form.Odejmij = True Then
        Z = -form.ilosc.Value
        Else
        Z = form.ilosc.Value
    End If
    Cells(n, 2) = Cells(n, 2) + Z
    k = 1
    While Cells(n, k) <> ""
        k = k + 1
    Wend
    Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
 
ElseIf form.CheckBox17.Value = True Then
                Application.Goto (ActiveWorkbook.Sheets(18).Range("A1"))
                If form.Odejmij = True Then
                    Z = -form.ilosc.Value
                Else
                    Z = form.ilosc.Value
                End If
                Cells(n, 2) = Cells(n, 2) + Z
                k = 1
                While Cells(n, k) <> ""
                k = k + 1
                Wend
                Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
                
                ElseIf form.CheckBox18.Value = True Then
                Application.Goto (ActiveWorkbook.Sheets(19).Range("A1"))
                If form.Odejmij = True Then
                    Z = -form.ilosc.Value
                Else
                    Z = form.ilosc.Value
                End If
                Cells(n, 2) = Cells(n, 2) + Z
                k = 1
                While Cells(n, k) <> ""
                k = k + 1
                Wend
                Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
                
                ElseIf form.CheckBox19.Value = True Then
                Application.Goto (ActiveWorkbook.Sheets(20).Range("A1"))
                If form.Odejmij = True Then
                    Z = -form.ilosc.Value
                Else
                    Z = form.ilosc.Value
                End If
                Cells(n, 2) = Cells(n, 2) + Z
                k = 1
                While Cells(n, k) <> ""
                k = k + 1
                Wend
                Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
                
                ElseIf form.CheckBox20.Value = True Then
                Application.Goto (ActiveWorkbook.Sheets(21).Range("A1"))
                If form.Odejmij = True Then
                    Z = -form.ilosc.Value
                Else
                    Z = form.ilosc.Value
                End If
                Cells(n, 2) = Cells(n, 2) + Z
                k = 1
                While Cells(n, k) <> ""
                k = k + 1
                Wend
                Cells(n, k) = form.data.Text & " " & "(" & Z & ")"
   
Else
    MsgBox "Wprowadzono nie poprawnie dane, spróbuj raz jeszcze", , "B³¹d"
End If
    Else
        MsgBox "Wybierz 1 grupê", , "B³¹d"
End If
End If
End If
End Sub

