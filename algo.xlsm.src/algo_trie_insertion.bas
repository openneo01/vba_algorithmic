Attribute VB_Name = "algo_trie_insertion"
'---------------------------------------------------------------------------------------
' Module        : algo_trie_insertion
' Author        : misterneo
' Date          : 12/02/2016
' Description   :
'---------------------------------------------------------------------------------------


Option Base 1



Sub test()

Dim tabdata(6) As Variant

tabdata(1) = 5
tabdata(2) = 2
tabdata(3) = 4
tabdata(4) = 6
tabdata(5) = 1
tabdata(6) = 3

Call TriSelection2(tabdata)


End Sub


'http://www.giacomazzi.fr/infor/Tri/PgmVB4.htm#TSelection
'http://www.cmdvb.fr/5-methodes-pour-trier-un-tableau-en-visual-basic-selection-insertion-bulles-shell-rapide/



Sub TriSelection2(oarray() As Variant)

poscurr = LBound(oarray)

Do While poscurr < UBound(oarray)
    min = poscurr
    
    For i = poscurr + 1 To UBound(oarray)
    
        If oarray(min) > oarray(i) Then
            min = i
        End If
    
    Next

    tmp = oarray(poscurr)
    oarray(poscurr) = oarray(min)
    oarray(min) = tmp
    
    poscurr = poscurr + 1

Loop

End Sub

