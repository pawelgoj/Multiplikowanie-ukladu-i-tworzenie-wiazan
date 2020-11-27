Attribute VB_Name = "Wczytywanie_ukladu"
'Autor Pawe³ Goj
'Wczytuj uklad
Option Base 1 'tablice sa numerowane od 1
Public Sub Wczytywaj_uklad(a, b, ilosc_wierszy_maly, ilosc_kolumn_maly, maly_uklad, x, y, z, xz)

Dim i As Long
Dim j As Integer
 

i = 0
j = 0
    
Do While Worksheets("Systam-skalowanie duzy").Cells(a + i, b) > 0
    j = 0
    Do While Worksheets("Systam-skalowanie duzy").Cells(a + i, b + j) <> 0
        j = j + 1
    Loop
    i = i + 1
Loop

ilosc_wierszy_maly = i
ilosc_kolumn_maly = j


maly_uklad = Range(Cells(a, b), Cells(a + i, b + j))
'efektywne wczytywanie danych maly uklad wczesniej zdefiniowany jako Variant

End Sub

