Sub rechercher_valeur()
    'Déclaration des variables
    Dim chemin_fichier As String
    Dim valeur_recherchee As String
    Dim derniere_ligne As Long
    Dim i As Long
    Dim trouve As Boolean
    
    'Définition du chemin d'accès du fichier contenant la valeur recherchée
    chemin_fichier = "C:\Chemin\vers\fichier.xlsx" 'Modifier le chemin d'accès en fonction de votre fichier
    
    'Affectation de la valeur sélectionnée
    valeur_recherchee = Sheets("feuil1").Range("A1").Value 'Modifier la plage de cellules en fonction de votre besoin
    'Affectation de la valeur sélectionnée
    valeur_recherchee = Selection.Value
    
    'Ouverture du fichier contenant la colonne de recherche
    Dim fichier_source As Workbook
    Set fichier_source = Workbooks.Open(chemin_fichier)
    
    'Détermination de la dernière ligne de la colonne de recherche
    derniere_ligne = fichier_source.Sheets("Feuil1").Cells(Rows.Count, "A").End(xlUp).Row
    
    'Recherche de la valeur dans la colonne A du fichier source
    For i = 1 To derniere_ligne
        If fichier_source.Sheets("Feuil1").Cells(i, 1).Value = valeur_recherchee Then
            ligne_trouvee = i 'La valeur a été trouvée, on mémorise la ligne
            Exit For
        End If
    Next i
    
    'Affichage du message si la valeur est trouvée ou non
    If ligne_trouvee > 0 Then
        MsgBox "La valeur " & valeur_recherchee & " a été trouvée à la ligne " & ligne_trouvee & " du fichier source."
    Else
        MsgBox "La valeur " & valeur_recherchee & " n'a pas été trouvée dans le fichier source."
    End If
    
    'Fermeture du fichier source
    fichier_source.Close
    
    'Libération de la mémoire
    Set fichier_source = Nothing
End Sub
