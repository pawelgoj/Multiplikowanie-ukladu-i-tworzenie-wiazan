Attribute VB_Name = "Warunki_periodyczne"

'Autor Pawe³ Goj
'Warunki periodyczne dla ukladu jednoskoœnego i uk³adów o wy¿szej symetrii
Option Base 1
Sub warunki_periodyczne_sub(x, y, z, xz, xO, yO, zO, maly_uklad, wiersze, iteracja)
    
    
    'x y z xz - wymiary uk³adu
    'xH yH zH - koordynaty atomu 2 w parze
    'maly_uklad to tablica z danymi
    'wiersze, kolumny to wiersze i kolumny tablicy
    
    'maly_ukladw kolumny w tablicy: id-1, molekula-2, type-3, charge-4, x-5, y-6, z-7, inne
     
                'przekszta³cenia dnaego atomu
                'œciany 6
            Select Case iteracja 'inny if
                Case 1
                    xO = maly_uklad(wiersze, 5) + x
                    yO = maly_uklad(wiersze, 6)
                    zO = maly_uklad(wiersze, 7)
                Case 2
                    xO = maly_uklad(wiersze, 5) - x
                    yO = maly_uklad(wiersze, 6)
                    zO = maly_uklad(wiersze, 7)
                Case 3
                    xO = maly_uklad(wiersze, 5)
                    yO = maly_uklad(wiersze, 6) + y
                    zO = maly_uklad(wiersze, 7)
                Case 4
                    xO = maly_uklad(wiersze, 5)
                    yO = maly_uklad(wiersze, 6) - y
                    zO = maly_uklad(wiersze, 7)
                Case 5
                    xO = maly_uklad(wiersze, 5) + xz 'musi przesun¹æ siê o wektor [xz, 0, z]
                    yO = maly_uklad(wiersze, 6)
                    zO = maly_uklad(wiersze, 7) + z
                Case 6
                    xO = maly_uklad(wiersze, 5) - xz 'musi przesun¹æ siê o wektor -[xz, 0, z]
                    yO = maly_uklad(wiersze, 6)
                    zO = maly_uklad(wiersze, 7) - z
                'Krawêdzie 12
                Case 7
                    xO = maly_uklad(wiersze, 5) + x
                    yO = maly_uklad(wiersze, 6) + y
                    zO = maly_uklad(wiersze, 7)
                Case 8
                    xO = maly_uklad(wiersze, 5) + x
                    yO = maly_uklad(wiersze, 6) + y
                    zO = maly_uklad(wiersze, 7)
                Case 9
                    xO = maly_uklad(wiersze, 5) - x
                    yO = maly_uklad(wiersze, 6) + y
                    zO = maly_uklad(wiersze, 7)
                Case 10
                    xO = maly_uklad(wiersze, 5) + x
                    yO = maly_uklad(wiersze, 6) - y
                    zO = maly_uklad(wiersze, 7)
                Case 11
                    xO = maly_uklad(wiersze, 5) + x + xz
                    yO = maly_uklad(wiersze, 6)
                    zO = maly_uklad(wiersze, 7) + z
                Case 12
                    xO = maly_uklad(wiersze, 5) - x - xz
                    yO = maly_uklad(wiersze, 6)
                    zO = maly_uklad(wiersze, 7) - z
                Case 13
                    xO = maly_uklad(wiersze, 5) - x + xz
                    yO = maly_uklad(wiersze, 6)
                    zO = maly_uklad(wiersze, 7) + z
                Case 14
                    xO = maly_uklad(wiersze, 5) + x - xz
                    yO = maly_uklad(wiersze, 6)
                    zO = maly_uklad(wiersze, 7) - z
                Case 15
                    xO = maly_uklad(wiersze, 5) + xz
                    yO = maly_uklad(wiersze, 6) + y
                    zO = maly_uklad(wiersze, 7) + z
                Case 16
                    xO = maly_uklad(wiersze, 5) - xz
                    yO = maly_uklad(wiersze, 6) + y
                    zO = maly_uklad(wiersze, 7) - z
                Case 17
                    xO = maly_uklad(wiersze, 5) + xz
                    yO = maly_uklad(wiersze, 6) - y
                    zO = maly_uklad(wiersze, 7) + z
                Case 18
                    xO = maly_uklad(wiersze, 5) - xz
                    yO = maly_uklad(wiersze, 6) - y
                    zO = maly_uklad(wiersze, 7) - z
                    
                'wieszcho³ki 8
                Case 19
                    xO = maly_uklad(wiersze, 5) + x + xz
                    yO = maly_uklad(wiersze, 6) + y
                    zO = maly_uklad(wiersze, 7) + z
                Case 20
                    xO = maly_uklad(wiersze, 5) + x + xz
                    yO = maly_uklad(wiersze, 6) - y
                    zO = maly_uklad(wiersze, 7) + z
                Case 21
                    xO = maly_uklad(wiersze, 5) - x - xz
                    yO = maly_uklad(wiersze, 6) + y
                    zO = maly_uklad(wiersze, 7) - z
                Case 22
                    xO = maly_uklad(wiersze, 5) - x - xz
                    yO = maly_uklad(wiersze, 6) - y
                    zO = maly_uklad(wiersze, 7) - z
                Case 23
                    xO = maly_uklad(wiersze, 5) - x + xz
                    yO = maly_uklad(wiersze, 6) + y
                    zO = maly_uklad(wiersze, 7) + z
                Case 24
                    xO = maly_uklad(wiersze, 5) - x + xz
                    yO = maly_uklad(wiersze, 6) - y
                    zO = maly_uklad(wiersze, 7) + z
                Case 25
                    xO = maly_uklad(wiersze, 5) + x - xz
                    yO = maly_uklad(wiersze, 6) + y
                    zO = maly_uklad(wiersze, 7) - z
                Case 26
                    xO = maly_uklad(wiersze, 5) + x - xz
                    yO = maly_uklad(wiersze, 6) - y
                    zO = maly_uklad(wiersze, 7) - z
            End Select
End Sub
