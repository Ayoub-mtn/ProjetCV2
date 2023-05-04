Sub CopierLignes()
    Dim Feuil1 As Worksheet
    Dim Feuil2 As Worksheet
    Dim Donnees() As Variant
    Dim Trouve As Range
    Dim i As Long
    Dim Fichier1 As Workbook
    
    ' Ouvre le fichier Excel contenant la feuille "Feuil1"
    Set Fichier1 = Workbooks.Open("C:\Chemin\vers\Fichier1.xlsx")
    
    ' Déclaration des variables pour les feuilles de calcul
    Set Feuil1 = Fichier1.Worksheets("Feuil1")
    Set Feuil2 = Worksheets("LISTE JALONS")
    
    ' Récupération des données de la plage D2:D50 de la feuille "LISTE PROJETS RAA"
    Donnees = Feuil2.Range("D2:D50").Value
    
    ' Parcours des données récupérées
    For i = 1 To UBound(Donnees, 1)
        ' Recherche de la donnée dans la colonne A de la feuille "Feuil1"
        Set Trouve = Feuil1.Range("A:A").Find(What:=Donnees(i, 1), LookIn:=xlValues, LookAt:=xlWhole)
        If Not Trouve Is Nothing Then
            ' Copie de la ligne correspondante dans la feuille "LISTE JALONS"
            Feuil1.Rows(Trouve.Row).Copy Feuil2.Cells(i + 1, 1)
        End If
    Next i
    
    ' Ferme le fichier Excel contenant la feuille "Feuil1"
    Fichier1.Close SaveChanges:=False
End Sub
