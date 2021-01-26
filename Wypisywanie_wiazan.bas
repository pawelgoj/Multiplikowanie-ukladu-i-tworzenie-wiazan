Attribute VB_Name = "Wypisywanie_wiazan"
Option Explicit
    'Wypisuje wyniki
Sub Wypisywanie(k, Nagluwek, iteracje_szukanie_wiazan, wyniki, IdH, IdO, Idsub)
    Dim poczatek As Integer
    Dim koniec As Integer
    

        If UBound(wyniki, 2) = 3 Then
            poczatek = 25
            koniec = 27
        Else
            poczatek = 22
            koniec = 23
        End If
        
            
        Worksheets("Systam-skalowanie duzy").Cells(3 + Nagluwek + k * (iteracje_szukanie_wiazan - 1), poczatek) = "id1: " & IdH
        Worksheets("Systam-skalowanie duzy").Cells(3 + Nagluwek + k * (iteracje_szukanie_wiazan - 1), poczatek + 1) = "id2: " & IdO
        If UBound(wyniki, 2) = 3 Then
            Worksheets("Systam-skalowanie duzy").Cells(3 + Nagluwek + k * (iteracje_szukanie_wiazan - 1), poczatek + 2) = "id3: " & Idsub
        End If
            
        Range(Cells(4 + Nagluwek + k * (iteracje_szukanie_wiazan - 1), poczatek), Cells(3 + Nagluwek + k * iteracje_szukanie_wiazan, koniec)) = wyniki
        Nagluwek = Nagluwek + 1

End Sub
