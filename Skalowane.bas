Attribute VB_Name = "Skalowane"
' Autor Pawe³ Goj
' Makro skalujace uk³ad
Option Explicit 'gdy jest wypisana ta opcja wszystkie zmienne musza byc zdefiniowane
Option Base 1 'tablice sa numerowane od 1


Sub skalowanie(H, n, m, x, y, z, xz, ilosc_wierszy_maly, ilosc_kolumn_maly, maly_uklad)
'skaluje uklad do durzego ukladu procedura musi byæ publiczna by mo¿na by³o sie
'do niej odwolaæ w innym arkuszu
Dim i As Long
Dim j As Integer
Dim k As Long
Dim l As Integer
Dim i_na_wysokosc As Long
Dim duzy_uklad() As Double 'tablica o zmiannej wielkoœci
Dim zmienna_pomocnicza As Long

'Worksheets("Systam-skalowanie duzy").Cells(3, 10) = _
'LBound(maly_uklad, 1) 'dolna granica tablicy wymiar x, "_ to ³amanie lini"
'Worksheets("Systam-skalowanie duzy").Cells(4, 10) = _
'UBound(maly_uklad, 1) 'górna granica tablicy wymiar x
'Worksheets("Systam-skalowanie duzy").Cells(5, 10) = _
'LBound(maly_uklad, 2) 'dolna granica tablicy wymiar y
'Worksheets("Systam-skalowanie duzy").Cells(6, 10) = UBound(maly_uklad, 2) 'górna granica tablicy wymiar y

i = 0
j = 0

i = ilosc_wierszy_maly * (n * m) * H
j = ilosc_kolumn_maly

Range(Cells(2, 13), Cells(1000 + (n * m * H) * ilosc_wierszy_maly, 12 + ilosc_kolumn_maly)).ClearContents
ReDim duzy_uklad(i, j)
                'ReDim Preserve -zmiana wymiarów tablicy (Preserve -oznacza, ¿e maj¹ byæ zachowane
                'wszystkie wrtoœci znajdujace sie dotychczas w tablicy i mozna tylko zmieniac ostatni
                'wymiar tablicy )


'Wartoœci j od 4- 6 to polozenia (x- 4, y-5)
For i_na_wysokosc = 0 To H - 1
    For l = 1 To n
        For k = 0 To m - 1
            For i = 1 To ilosc_wierszy_maly
                For j = 1 To ilosc_kolumn_maly
                
                        zmienna_pomocnicza = k + m * (l - 1) + (i_na_wysokosc) * m * n
                        
                        If l = 1 And k = 0 And i_na_wysokosc = 0 Then 'kopia ukladu
                            duzy_uklad(i, j) = maly_uklad(i, j)
                        ElseIf j = 1 Then 'id atmu
                            duzy_uklad(i + (zmienna_pomocnicza * ilosc_wierszy_maly), j) = (zmienna_pomocnicza * ilosc_wierszy_maly) + maly_uklad(i, j)
                        ElseIf j > 1 And j < 5 Then 'molekula, ladunek, type
                            duzy_uklad(i + (zmienna_pomocnicza * ilosc_wierszy_maly), j) = maly_uklad(i, j)
                        ElseIf j = 5 Then  'x
                            duzy_uklad(i + (zmienna_pomocnicza * ilosc_wierszy_maly), j) = i_na_wysokosc * xz + (l - 1) * x + maly_uklad(i, j)
                        ElseIf j = 6 Then 'y
                            duzy_uklad(i + (zmienna_pomocnicza * ilosc_wierszy_maly), j) = k * y + maly_uklad(i, j)
                        ElseIf j = 7 Then 'z
                            duzy_uklad(i + (zmienna_pomocnicza * ilosc_wierszy_maly), j) = i_na_wysokosc * z + maly_uklad(i, j)
                        Else ' i inne
                            duzy_uklad(i + (zmienna_pomocnicza * ilosc_wierszy_maly), j) = maly_uklad(i, j)
                        End If
                Next j
            Next i
        Next k
    Next l
Next i_na_wysokosc

Range(Cells(3, 13), Cells(2 + (n * m * H) * ilosc_wierszy_maly, 12 + ilosc_kolumn_maly)) = duzy_uklad 'efektywne wyrzucanie danych
End Sub
