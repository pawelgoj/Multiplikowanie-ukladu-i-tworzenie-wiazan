VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Base 1 'tablice sa numerowane od 1
'Funkcja mieniaj¹ca kropki na przecinki
Function zamien_kropi(arg As String) As Double
   arg = Replace(arg, ".", ",")
   zamien_kropi = CDbl(arg) 'konwertuje dane do typu double
End Function

'Funkcja skaluj uklad
Private Sub CommandButton1_Click()
    On Error GoTo Error_handler
    Dim n As Integer
    Dim m As Integer
    Dim H As Integer
    
    Dim x As String
    Dim y As String
    Dim z As String
    Dim xz As String
    
    Dim ilosc_wierszy_maly As Long
    Dim ilosc_kolumn_maly As Integer
    Dim maly_uklad As Variant 'Efektywne wczytywanie za pomoca zmiannej variant
    
    'Wczytaj zmienne
    n = TextBox1  'x
    m = TextBox4  'y
    H = TextBox7  'z
    x = TextBox2
    y = TextBox3
    z = TextBox5
    xz = TextBox6
    x = zamien_kropi(x)
    y = zamien_kropi(y)
    z = zamien_kropi(z)
    xz = zamien_kropi(xz)
    
    'Wczytaj uklad
    Const a As Integer = 3 'Komórki w których zaczyna siê tabela
    Const b As Integer = 1
    Wczytywaj_uklad a, b, ilosc_wierszy_maly, ilosc_kolumn_maly, maly_uklad, x, y, z, xz
    'procedura musi mieæ
    'zapisane wczystkie zmienne które wchodz¹ i wychodz¹ z niej
    
    'Skaluj uklad
    skalowanie H, n, m, x, y, z, xz, ilosc_wierszy_maly, ilosc_kolumn_maly, maly_uklad
    'wywo³ywanie procedury i argumenty
    
MsgBox "Koniec"

Exit Sub 'Koñczy procedurê

Error_handler:
MsgBox "Z³e typy wprowadzanych danny, id atomu nie mo¿e równac sie 0" 'komunikat przy bledzie
Err.Clear 'czyœci b³êdy

End Sub
'Funkcja twórz wi¹zania
Private Sub CommandButton2_Click()
    'On Error GoTo Error_handler
    Dim IdO As Variant
    Dim IdH As Variant
    Dim x As String
    Dim y As String
    Dim z As String
    Dim xz As String
    Dim R1 As String
    Dim Wczytywanie_id_z_listy As Boolean
    Dim ilosc_wierszy_maly As Long
    Dim ilosc_kolumn_maly As Integer
    Dim maly_uklad As Variant
    Dim lista_ID As Variant
    Dim wiersze_listy_ID As Integer
    
    x = TextBox2
    y = TextBox3
    z = TextBox5
    xz = TextBox6
    R1 = TextBox8
    
    Wczytywanie_id_z_listy = Range("J31")
    
    If Wczytywanie_id_z_listy = True Then
        MsgBox ("Wczytujê z listy")
        i = 0
        Do While Worksheets("Systam-skalowanie duzy").Cells(56 + i, 9) > 0
            i = i + 1
        Loop
        
        If i = 0 Then
            MsgBox (" Pusta lista")
        Else
            wiersze_listy_ID = i
            lista_ID = Range(Cells(56, 9), Cells(55 + wiersze_listy_ID, 12))
        End If
    Else
        IdH = TextBox10
        IdO = TextBox11
        IdH = CDbl(IdH)
        IdO = CDbl(IdO)
    End If
    
        x = zamien_kropi(x)
        y = zamien_kropi(y)
        z = zamien_kropi(z)
        xz = zamien_kropi(xz)
        R1 = zamien_kropi(R1)

    'Wczytaj uklad- Komórki w których zaczyna siê tabela
    Const a As Integer = 3 'sta³a, nie da siê przypisaæ wartoœci do sta³ej
    Const b As Integer = 13 'sta³a
    
    'maly_uklad = uk³ad wczytywany- obszar du¿ego uk³adu.
    Wczytywaj_uklad a, b, ilosc_wierszy_maly, ilosc_kolumn_maly, maly_uklad, x, y, z, xz
    
    'ZnajdŸ wi¹zania
    Znajdowanie_wiazan x, y, z, xz, R1, IdH, IdO, maly_uklad, ilosc_wierszy_maly, lista_ID, wiersze_listy_ID
    
    MsgBox "Koniec"
    
    Exit Sub 'Koñczy procedurê

'Error_handler:
'MsgBox "Z³e typy wprowadzanych danny, id atomu nie mo¿e równac sie 0" 'komunikat przy bledzie
'Err.Clear 'czyœci b³êdy

End Sub
'Funkcja twórz wiazania katowe
Private Sub CommandButton3_Click()
    'On Error GoTo Error_handler 'gdy jest blad przekierowuje do Error_handler
    Dim R1 As String
    Dim R2 As String
    Dim IdO As Variant
    Dim IdH As Variant
    Dim Idsub As Variant
    Dim Wczytywanie_id_z_listy As Boolean
    Dim x As String
    Dim y As String
    Dim z As String
    Dim xz As String
    Dim lista_ID As Variant
    Dim wiersze_listy_ID As Integer
    
    x = TextBox2
    y = TextBox3
    z = TextBox5
    xz = TextBox6
    
    R1 = TextBox8
    R2 = TextBox9
    
    Wczytywanie_id_z_listy = Range("J31")
    
    If Wczytywanie_id_z_listy = True Then
        MsgBox ("Wczytujê z listy")
        i = 0
        Do While Worksheets("Systam-skalowanie duzy").Cells(56 + i, 9) > 0
            i = i + 1
        Loop
        
        If i = 0 Then
            MsgBox (" Pusta lista")
        Else
            wiersze_listy_ID = i
            lista_ID = Range(Cells(56, 9), Cells(55 + wiersze_listy_ID, 12))
        End If
    Else
        IdH = TextBox10
        IdO = TextBox11
        Idsub = TextBox12
        IdH = CDbl(IdH)
        IdO = CDbl(IdO)
        Idsub = CDbl(Idsub)
    End If
    
    x = zamien_kropi(x)
    y = zamien_kropi(y)
    z = zamien_kropi(z)
    xz = zamien_kropi(xz)
    R1 = zamien_kropi(R1)
    R2 = zamien_kropi(R2)
    
    'Wczytaj uklad- Komórki w których zaczyna siê tabela
    Const a As Integer = 3 'sta³a, nie da siê przypisaæ wartoœci do sta³ej
    Const b As Integer = 13 'sta³a
    
    'maly_uklad = uk³ad wczytywany- obszar du¿ego uk³adu.
    Wczytywaj_uklad a, b, ilosc_wierszy_maly, ilosc_kolumn_maly, maly_uklad, x, y, z, xz
    
    'ZnajdŸ wi¹zania wiazania katowe
    katy_sub x, y, z, xz, R1, R2, IdH, IdO, Idsub, maly_uklad, ilosc_wierszy_maly, lista_ID, wiersze_listy_ID
    
    MsgBox "Koniec"
Exit Sub 'Koñczy procedurê

'Error_handler:
'MsgBox "Z³e typy wprowadzanych danny, id atomu nie mo¿e równac sie 0" 'komunikat przy bledzie
'Err.Clear 'czyœci b³êdy
    
End Sub

Private Sub Label7_Click()

End Sub
