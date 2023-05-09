Sub rechercher_et_copier()
    'Déclaration des variables
    Dim feuille_source As Worksheet
    Dim feuille_dest As Worksheet
    Dim chemin_fichier As String
    Dim derniere_ligne As Long
    Dim i As Long, j As Long
    Dim valeur_recherchee As String
    
    'Affectation de la feuille source
    Set feuille_source = ThisWorkbook.Sheets("Feuille2")
    
    'Affectation de la feuille de destination
    Set feuille_dest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    feuille_dest.Name = "Resultats"
    
    'Définition du chemin d'accès du fichier contenant les données recherchées
    chemin_fichier = "C:\Chemin\vers\fichier.xlsx" 'Modifier le chemin d'accès en fonction de votre fichier
    
    'Ouverture du fichier contenant les données recherchées
    Dim fichier_source As Workbook
    Set fichier_source = Workbooks.Open(chemin_fichier)
    
    'Détermination de la dernière ligne de la feuille source
    derniere_ligne = feuille_source.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Boucle sur les lignes de la feuille source
    For i = 1 To derniere_ligne
        valeur_recherchee = feuille_source.Cells(i, 1).Value 'Valeur à rechercher dans le fichier source
        j = 1 'Initialisation de l'index de ligne pour la feuille de destination
        
        'Boucle sur les lignes du fichier source
        Do Until fichier_source.Sheets("Feuil1").Cells(j, 1).Value = ""
            'Si la valeur recherchée est trouvée dans la colonne A du fichier source, copier la ligne correspondante
            If fichier_source.Sheets("Feuil1").Cells(j, 1).Value = valeur_recherchee Then
                fichier_source.Sheets("Feuil1").Range("A" & j & ":D" & j).Copy Destination:=feuille_dest.Range("A" & Rows.Count).End(xlUp).Offset(1)
            End If
            j = j + 1
        Loop
    Next i
    
    'Fermeture du fichier source
    fichier_source.Close
    
    'Libération de la mémoire
    Set feuille_source = Nothing
    Set feuille_dest = Nothing
    Set fichier_source = Nothing
End Sub
