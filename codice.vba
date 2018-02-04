' Rilasciato con Licenza GPL 3.0
'
' Autore: Ing. Marcello Sonaglia
'
' email: assistenza@comuneortona.ch.it
'
Public candidati As Integer
Public estratti As Integer
Public titolari As Integer
Public max_titolari As Integer
Public max_riserve As Integer
Public riserve As Integer

Public Function memorizza_titolare(estratto)
    Dim indice_titolari As Integer
    
    indice_titolari = titolari + 4
    Worksheets("Scrutatori").Cells(indice_titolari, 2) = estratti + 1
    Worksheets("Scrutatori").Cells(indice_titolari, 3) = estratto
    Worksheets("Estrazione").Cells(5, 18) = titolari + 1
    Worksheets("Estrazione").Cells(3, 18) = estratti + 1
End Function
Public Function memorizza_riserva(estratto)
    Dim indice_riserve As Integer
    
    indice_riserve = riserve + 4
    Worksheets("Riserve").Cells(indice_riserve, 2) = estratti + 1
    Worksheets("Riserve").Cells(indice_riserve, 3) = estratti + 1 - max_titolari
    Worksheets("Riserve").Cells(indice_riserve, 4) = estratto
    Worksheets("Estrazione").Cells(7, 18) = riserve + 1
    Worksheets("Estrazione").Cells(3, 18) = estratti + 1
End Function
Public Function carica_dati()
    candidati = Worksheets("Estrazione").Cells(3, 8)
    estratti = Worksheets("Estrazione").Cells(3, 18)
    max_titolari = Worksheets("Estrazione").Cells(5, 8)
    titolari = Worksheets("Estrazione").Cells(5, 18)
    max_riserve = Worksheets("Estrazione").Cells(7, 8)
    riserve = Worksheets("Estrazione").Cells(7, 18)
End Function
Public Function non_presente(estratto)
    Dim test As Object

    non_presente = True
    
    If Not Worksheets("Scrutatori").Range("C4:C111").Find(estratto, LookIn:=xlValues, After:=Worksheets("Scrutatori").Range("C111"), LookAt:=xlWhole) Is Nothing Then
        non_presente = False
    ElseIf Not Worksheets("Riserve").Range("D4:D111").Find(estratto, LookIn:=xlValues, After:=Worksheets("Scrutatori").Range("D111"), LookAt:=xlWhole) Is Nothing Then
            non_presente = False
    End If
End Function
Sub estrai_scrutatore()
    Dim estrazione As Integer
    
    carica_dati
        
    estrazione = Int((Rnd * candidati) + 1)
    Do While Not non_presente(estrazione)
        estrazione = Int((Rnd * candidati) + 1)
    Loop
    If titolari < max_titolari Then
        memorizza_titolare (estrazione)
        Worksheets("Estrazione").Cells(10, 12) = estrazione
    ElseIf riserve < max_riserve Then
        memorizza_riserva (estrazione)
        Worksheets("Estrazione").Cells(10, 12) = estrazione
    Else
        MsgBox ("ESTRAZIONI TERMINATE")
    End If
    If estratti = max_titolari - 1 Then MsgBox ("ESTRAZIONE DEI TITOLARI TERMINATA")
End Sub
Sub reset_click()
    Worksheets("Estrazione").Cells(3, 8) = 2097
    Worksheets("Estrazione").Cells(3, 18) = 0
    Worksheets("Estrazione").Cells(5, 8) = 108
    Worksheets("Estrazione").Cells(5, 18) = 0
    Worksheets("Estrazione").Cells(7, 8) = 108
    Worksheets("Estrazione").Cells(7, 18) = 0
    Worksheets("Estrazione").Cells(10, 12) = ""
    
    Worksheets("Scrutatori").Range("B4:C111").ClearContents
    Worksheets("Riserve").Range("B4:D111").ClearContents
    
    Randomize
End Sub
