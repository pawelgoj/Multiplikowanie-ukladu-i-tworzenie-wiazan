Attribute VB_Name = "Wypisywanie_wiazan"
Option Explicit
Option Base 1
Sub W1(Nagluwek, lista_wynikow, poczatek, koniec, k, wyniki, IdH, IdO, Idsub)
            Worksheets("Systam-skalowanie duzy").Cells(3 + Nagluwek + (lista_wynikow - k), poczatek) = "id1: " & IdH
            Worksheets("Systam-skalowanie duzy").Cells(3 + Nagluwek + (lista_wynikow - k), poczatek + 1) = "id2: " & IdO
            If UBound(wyniki, 2) = 3 Then
                Worksheets("Systam-skalowanie duzy").Cells(3 + Nagluwek + (lista_wynikow - k), poczatek + 2) = "id3: " & Idsub
            End If
            Range(Cells(4 + Nagluwek + (lista_wynikow - k), poczatek), Cells(3 + Nagluwek + lista_wynikow, koniec)) = wyniki
            Nagluwek = Nagluwek + 1
End Sub

     
    'Wypisuje wyniki
Sub Wypisywanie(k, lista_wynikow, Nagluwek, iteracje_szukanie_wiazan, wyniki, IdH, IdO, Idsub, lista_ID, przelacznik)
    Dim poczatek As Integer
    Dim koniec As Integer
    Dim i As Long
    Dim ID() As Integer
    

        If UBound(wyniki, 2) = 3 Then
            poczatek = 27
            koniec = 29
        Else
            poczatek = 23
            koniec = 24
        End If
        
    If przelacznik = False Then
        W1 Nagluwek, lista_wynikow, poczatek, koniec, k, wyniki, IdH, IdO, Idsub
    ElseIf przelacznik = True Then
      If lista_ID(iteracje_szukanie_wiazan, 4) = 0 Then
        W1 Nagluwek, lista_wynikow, poczatek, koniec, k, wyniki, IdH, IdO, Idsub
      Else
        ReDim ID(k)
        For i = 1 To k
            ID(i) = lista_ID(iteracje_szukanie_wiazan, 4)
        Next i
        Range(Cells(3 + (lista_wynikow - k), poczatek - 1), Cells(2 + lista_wynikow, poczatek - 1)) = ID
        Range(Cells(3 + (lista_wynikow - k), poczatek), Cells(2 + lista_wynikow, koniec)) = wyniki
      End If

    End If
            
End Sub
