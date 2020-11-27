Attribute VB_Name = "Katy"
' Autor Pawe³ Goj
' Makro ma Eecel tworzace listê wi¹zañ k¹towych Al-O-H w Montmorylonicie na
' podstawie tablicy Zawieraj¹cej po³ozenia czastek w uk³adzie kartezjañskim.
Option Base 1 'tablice sa numerowane od 1

Private Function modul_wektor(i As Long, j As Long, maly_uklad, xO As Double, yO As Double, zO As Double, wskaznik As Integer) As Double
    If wskaznik = 0 Then
        modul_wektor = _
                ((maly_uklad(j, 5) - maly_uklad(i, 5)) ^ 2 + (maly_uklad(j, 6) - maly_uklad(i, 6)) ^ 2 + (maly_uklad(j, 7) - maly_uklad(i, 7)) ^ 2) _
                ^ (1 / 2)
    ElseIf wskaznik = 1 Then
        modul_wektor = _
                ((xO - maly_uklad(i, 5)) ^ 2 + (yO - maly_uklad(i, 6)) ^ 2 + (zO - maly_uklad(i, 7)) ^ 2) _
                ^ (1 / 2)
    End If
End Function

Private Sub wczytuj_do_tablicy_katy(k, i, j, l, katy, maly_uklad) 'Wczytuje katy do tablicy katy
    k = k + 1
    ReDim Preserve katy(3, k)
    katy(1, k) = maly_uklad(i, 1)
    katy(2, k) = maly_uklad(j, 1)
    katy(3, k) = maly_uklad(l, 1)
End Sub

Sub katy_sub(x, y, z, xz, R1, R2, IdH, IdO, Idsub, maly_uklad, ilosc_wierszy_maly)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Integer
Dim n As Integer
Dim modul_w As Double
Dim wskaznik As Integer
Dim katy() As Double
Dim katy2() As Double
Dim wiersze As Long
Dim iteracja As Integer
Dim iteracjaO As Integer
Dim iteracjaSub As Integer
Dim xO As Double
Dim yO As Double
Dim zO As Double

k = 0
'maly_ukladw kolumny w tablicy: id-1, molekula-2, type-3, charge-4, x-5, y-6, z-7, inne
For i = 1 To ilosc_wierszy_maly
    If maly_uklad(i, 3) = IdH Then
        For j = 1 To ilosc_wierszy_maly
            If maly_uklad(j, 3) = IdO Then
                wskaznik = 0
                modul_w = modul_wektor(i, j, maly_uklad, xO, yO, zO, wskaznik) 'odleg³oœæ H - O
                If modul_w < R1 Then
                    For l = 1 To ilosc_wierszy_maly
                        If maly_uklad(l, 3) = Idsub Then
                            wskaznik = 0
                            modul_w = modul_wektor(j, l, maly_uklad, xO, yO, zO, wskaznik) 'odleg³oœæ O- sub
                            If modul_w < R2 Then
                                wczytuj_do_tablicy_katy k, i, j, l, katy, maly_uklad
                            End If
                            iteracjaSub = 0
                            For n = 1 To 26
                                wskaznik = 1
                                wiersze = l 'musi szukac op Sub
                                iteracjaSub = iteracjaSub + 1
                                iteracja = iteracjaSub
                                warunki_periodyczne_sub x, y, z, xz, xO, yO, zO, maly_uklad, wiersze, iteracja
                                modul_w = modul_wektor(j, l, maly_uklad, xO, yO, zO, wskaznik)
                                If modul_w < R2 Then
                                    wczytuj_do_tablicy_katy k, i, j, l, katy, maly_uklad
                                End If
                            Next n
                        End If
                    Next l
                End If
                iteracjaO = 0
                For m = 1 To 26
                    wskaznik = 1
                    wiersze = j
                    iteracjaO = iteracjaO + 1 'tutaj jest cos nie tak
                    iteracja = iteracjaO
                    warunki_periodyczne_sub x, y, z, xz, xO, yO, zO, maly_uklad, wiersze, iteracja
                    modul_w = modul_wektor(i, j, maly_uklad, xO, yO, zO, wskaznik)
                    If modul_w < R1 Then
                        For l = 1 To ilosc_wierszy_maly
                            If maly_uklad(l, 3) = Idsub Then
                                wskaznik = 0
                                modul_w = modul_wektor(j, l, maly_uklad, xO, yO, zO, wskaznik) 'odleg³oœæ O- sub
                                If modul_w < R2 Then
                                    wczytuj_do_tablicy_katy k, i, j, l, katy, maly_uklad
                                End If
                                iteracjaSub = 0
                                For n = 1 To 26
                                    wskaznik = 1
                                    wiersze = l
                                    iteracjaSub = iteracjaSub + 1
                                    iteracja = iteracjaSub
                                    warunki_periodyczne_sub x, y, z, xz, xO, yO, zO, maly_uklad, wiersze, iteracja
                                    modul_w = modul_wektor(j, l, maly_uklad, xO, yO, zO, wskaznik)
                                    If modul_w < R2 Then
                                        wczytuj_do_tablicy_katy k, i, j, l, katy, maly_uklad
                                    End If
                                Next n
                            End If
                        Next l
                    End If
                Next m
            End If
        Next j
    End If
Next i

If k = 0 Then
    Range(Cells(2, 25), Cells(100, 27)).ClearContents
    MsgBox ("Incorrect ID of atom 1 or 2 else Too low cut radius")
    
Else
    ReDim katy2(k, 3)
    Range(Cells(3, 25), Cells(100 + k, 27)).ClearContents
    For i = 1 To k
        For j = 1 To 3
            katy2(i, j) = katy(j, i)
        Next j
    Next i
    
    Worksheets("Systam-skalowanie duzy").Cells(3, 25) = "id1: " & IdH
    Worksheets("Systam-skalowanie duzy").Cells(3, 26) = "id2: " & IdO
    Worksheets("Systam-skalowanie duzy").Cells(3, 27) = "id3: " & Idsub
    
    Range(Cells(4, 25), Cells(3 + k, 27)) = katy2
End If

End Sub
