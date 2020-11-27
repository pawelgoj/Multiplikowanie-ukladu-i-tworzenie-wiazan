Attribute VB_Name = "Wiazania"
' Autor: Pawe³ Goj
' Jêzyk VBA
' Makro znajduje pary atomów H i O po³¹czone z sob¹ wi¹zaniem
Option Base 1 'tablice sa numerowane od 1

Sub Znajdowanie_wiazan(x, y, z, xz, R1, IdH, IdO, maly_uklad, ilosc_wierszy_maly)

Dim i As Long
Dim j As Long
Dim k As Long
Dim m As Integer
Dim modul_wektor As Double
Dim wiazania() As Double
Dim wiazania2() As Double
Dim wiersze As Long
Dim iteracja As Integer
Dim xO As Double
Dim yO As Double
Dim zO As Double

k = 0

'maly_ukladw kolumny w tablicy: id-1, molekula-2, type-3, charge-4, x-5, y-6, z-7, inne
For i = 1 To ilosc_wierszy_maly
    If maly_uklad(i, 3) = IdH Then
        For j = 1 To ilosc_wierszy_maly
            If maly_uklad(j, 3) = IdO Then
                modul_wektor = _
                ((maly_uklad(j, 5) - maly_uklad(i, 5)) ^ 2 + (maly_uklad(j, 6) - maly_uklad(i, 6)) ^ 2 + (maly_uklad(j, 7) - maly_uklad(i, 7)) ^ 2) _
                ^ (1 / 2)
                If modul_wektor < R1 Then
                    k = k + 1
                    ReDim Preserve wiazania(2, k)
                    wiazania(1, k) = maly_uklad(i, 1)
                    wiazania(2, k) = maly_uklad(j, 1)
                End If
                iteracja = 0
                For m = 1 To 26
                    wiersze = j
                    iteracja = iteracja + 1
                    warunki_periodyczne_sub x, y, z, xz, xO, yO, zO, maly_uklad, wiersze, iteracja
                    modul_wektor = _
                    ((xO - maly_uklad(i, 5)) ^ 2 + (yO - maly_uklad(i, 6)) ^ 2 + (zO - maly_uklad(i, 7)) ^ 2) ^ (1 / 2)
                    If modul_wektor < R1 Then
                        k = k + 1
                        ReDim Preserve wiazania(2, k)
                        wiazania(1, k) = maly_uklad(i, 1)
                        wiazania(2, k) = maly_uklad(j, 1)
                    End If
                Next m
            End If
        Next j
    End If
Next i

If k = 0 Then
    Range(Cells(2, 22), Cells(100, 23)).ClearContents
    MsgBox ("Incorrect ID of atom 1 or 2 else Too low cut radius")
    
Else
    ReDim wiazania2(k, 2)
    Range(Cells(3, 22), Cells(100 + k, 23)).ClearContents
    For i = 1 To k
        For j = 1 To 2
            wiazania2(i, j) = wiazania(j, i)
        Next j
    Next i
    
    Worksheets("Systam-skalowanie duzy").Cells(3, 22) = "id1: " & IdH
    Worksheets("Systam-skalowanie duzy").Cells(3, 23) = "id2: " & IdO
    
    Range(Cells(4, 22), Cells(3 + k, 23)) = wiazania2
End If
End Sub
