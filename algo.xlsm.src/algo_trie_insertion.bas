Attribute VB_Name = "algo_trie_insertion"
'---------------------------------------------------------------------------------------
' Module        : algo_trie_insertion
' Author        : misterneo
' Date          : 12/02/2016
' Description   :
'---------------------------------------------------------------------------------------


Option Base 0



Sub test()

Dim tabdata(6) As Integer

tabdata(1) = 5
tabdata(2) = 2
tabdata(3) = 4
tabdata(4) = 6
tabdata(5) = 1
tabdata(6) = 3

Call TriSelection(tabdata)

End Sub


'http://www.giacomazzi.fr/infor/Tri/PgmVB4.htm#TSelection
'http://www.cmdvb.fr/5-methodes-pour-trier-un-tableau-en-visual-basic-selection-insertion-bulles-shell-rapide/

Sub TriSelection(Tableau As Variant)
'************************************************************
' Tri d'un tableau selon l'algorithme du tri par sélection
' Tableau       Tableau à trier
' Ipos_min_Tableau  Indice de l'échelon pos_mini à trier
' IMax_Tableau  Indice de l'échelon maxi à trier
'************************************************************
Dim W_Long As Long
Dim i As Long
Dim pos_min As Long
Dim pos As Long


pos = LBound(Tableau)
While (pos < UBound(Tableau))
    ' Recherche du plus petit élément dans le reste du tableau
    pos_min = pos
    For i = (pos + 1) To IMax_Tableau
        If Tableau(pos_min) > Tableau(i) Then
            pos_min = i
        End If
    Next i
    ' Echange de T(pos_min) et T(pos)
    W_Long = Tableau(pos)
    Tableau(pos) = Tableau(pos_min)
    Tableau(pos_min) = W_Long
    pos = pos + 1
Wend

Debug.Print Join(Tableau, ";")
End Sub
