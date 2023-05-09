Sub copier_colonne()
    'Déclaration des variables
    Dim feuille_source As Worksheet
    Dim feuille_dest As Worksheet
    Dim derniere_ligne As Long
    
    'Affectation de la feuille source
    Set feuille_source = ThisWorkbook.Sheets("LISTE PROJETS RAA")
    
    'Affectation de la feuille de destination
    Set feuille_dest = ThisWorkbook.Sheets("Feuille2")
    
    'Détermination de la dernière ligne de la colonne à copier
    derniere_ligne = feuille_source.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Copie de la colonne A de la feuille source à la colonne A de la feuille de destination
    feuille_source.Range("A1:A" & derniere_ligne).Copy Destination:=feuille_dest.Range("A1")
    
    'Libération de la mémoire
    Set feuille_source = Nothing
    Set feuille_dest = Nothing
End Sub
