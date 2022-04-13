Private Sub Rapport_Click()
    Dim strSQL As String
 '   Dim nomFichier As String
    Dim dbs As Database
    Dim qdf As QueryDef
    Dim thisDate As Date

    thisDate = Today
    strQry = "assainnissement"
    nomFichier = "t"
    
    Dim cheminFichier As String
    cheminFichier = "C:\Users\Resoudrien\"
    cheminFichier = cheminFichier & "rapport_assainissement_" & nomFichier & ".xls"
    
    MsgBox ("Excel exporte ici : " + cheminFichier)
   
    'Dim repertoire As String, nomFichier As String, extension As String
    'repertoire = "c:\"
    'nomFichier = "fichierTest"
    'extension = ".xls"
    'Application.Dialogs(xlDialogSaveAs).Show repertoire & nomFichier & extension

    Set dbs = CurrentDb
    Set qdf = dbs.CreateQueryDef(strQry)

    ' REQUETE ASSAINISSEMENT SEXE
    strSQL = "SELECT Adherents.Sexe_adh, Adherents.Nom_adh, Adherents.Prenom_adh "
    strSQL = strSQL & "FROM Adherents INNER JOIN Cotisations ON Adherents.Num_adh = Cotisations.Num_adh "
    strSQL = strSQL & "WHERE Cotisations.DateCotis>#31/12/2020# AND Sexe_adh IS NULL;"
    qdf.Sql = strSQL
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, strQry, cheminFichier, True, "Sexe_adh Invalide"
        
    ' REQUETE ASSAINISSEMENT ANNEE
    strSQL = "SELECT Adherents.Sexe_adh, Adherents.Nom_adh, Adherents.Prenom_adh, Adherents.AnneeNaissance "
    strSQL = strSQL & "FROM Adherents INNER JOIN Cotisations ON Adherents.Num_adh = Cotisations.Num_adh "
    strSQL = strSQL & "WHERE Cotisations.DateCotis>#31/12/2020# AND AnneeNaissance NOT LIKE '[0-9][0-9][0-9][0-9]';"
    qdf.Sql = strSQL
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, strQry, cheminFichier, True, "Annee Naissance Invalide"
    
    ' REQUETE ASSAINISSEMENT NATIONALITE
    strSQL = "SELECT Adherents.Sexe_adh, Adherents.Nom_adh, Adherents.Prenom_adh, Adherents.Nationalite_adh "
    strSQL = strSQL & "FROM Adherents INNER JOIN Cotisations ON Adherents.Num_adh = Cotisations.Num_adh "
    strSQL = strSQL & "WHERE Cotisations.DateCotis>#31/12/2020# AND Nationalite_adh IS NULL;"
    qdf.Sql = strSQL
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, strQry, cheminFichier, True, "Nationalite Invalide"
        
   ' REQUETE ASSAINISSEMENT Code Postal
    strSQL = "SELECT COUNT(*), Adherents.Nom_adh, Adherents.Prenom_adh, Adherents.Nationalite_adh, Adherents.CP_adh "
    strSQL = strSQL & "FROM Adherents INNER JOIN Cotisations ON Adherents.Num_adh = Cotisations.Num_adh "
    strSQL = strSQL & "WHERE Cotisations.DateCotis>#31/12/2020# AND CP_adh NOT LIKE '[0-9][0-9][0-9][0-9][0-9]'"
    strSQL = strSQL & "GROUP BY Adherents.Nom_adh, Adherents.Prenom_adh, Adherents.Nationalite_adh, Adherents.CP_adh;"
    qdf.Sql = strSQL
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, strQry, cheminFichier, True, "CP Invalide"
            
   ' REQUETE ASSAINISSEMENT Ville
    strSQL = "SELECT COUNT(*), Adherents.Nom_adh, Adherents.Prenom_adh, Adherents.Ville_adh "
    strSQL = strSQL & "FROM Adherents INNER JOIN Cotisations ON Adherents.Num_adh = Cotisations.Num_adh "
    strSQL = strSQL & "WHERE Cotisations.DateCotis>#31/12/2020# AND Ville_adh IS NULL OR Ville_adh LIKE '[a-zA-Z0-9 ]' "
    strSQL = strSQL & "GROUP BY Adherents.Nom_adh, Adherents.Prenom_adh, Adherents.Ville_adh;"
    qdf.Sql = strSQL
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, strQry, cheminFichier, True, "Ville Invalide"
    
    ' REQUETE ASSAINISSEMENT Adresse
    strSQL = "SELECT COUNT(*), Adherents.Nom_adh, Adherents.Prenom_adh, Adherents.Adresse_adh "
    strSQL = strSQL & "FROM Adherents INNER JOIN Cotisations ON Adherents.Num_adh = Cotisations.Num_adh "
    strSQL = strSQL & "WHERE Cotisations.DateCotis>#31/12/2020# AND Adresse_adh LIKE '[a-zA-Z0-9 ]' "
    strSQL = strSQL & "GROUP BY Adherents.Nom_adh, Adherents.Prenom_adh, Adherents.Adresse_adh;"
    qdf.Sql = strSQL
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, strQry, cheminFichier, True, "Adresse Invalide"
            
   ' REQUETE ASSAINISSEMENT Statut
    strSQL = "SELECT COUNT(*), Adherents.Nom_adh, Adherents.Prenom_adh, Adherents.Statut_adh "
    strSQL = strSQL & "FROM Adherents INNER JOIN Cotisations ON Adherents.Num_adh = Cotisations.Num_adh "
    strSQL = strSQL & "WHERE Cotisations.DateCotis>#31/12/2020# AND Adherents.Statut_adh = 0 "
    strSQL = strSQL & "GROUP BY Adherents.Nom_adh, Adherents.Prenom_adh, Adherents.Statut_adh;"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, strQry, cheminFichier, True, "Statut Invalide"
            
    ' REQUETE ASSAINISSEMENT Niveau Formation
    strSQL = "SELECT COUNT(*), Adherents.Nom_adh, Adherents.Prenom_adh, Adherents.NiveauF_adh "
    strSQL = strSQL & "FROM Adherents INNER JOIN Cotisations ON Adherents.Num_adh = Cotisations.Num_adh "
    strSQL = strSQL & "WHERE Cotisations.DateCotis>#31/12/2020# AND Adherents.NiveauF_adh = 0 "
    strSQL = strSQL & "GROUP BY Adherents.Nom_adh, Adherents.Prenom_adh, Adherents.NiveauF_adh;"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, strQry, cheminFichier, True, "Niveau Formation Invalide"
           
    qdf.Sql = strSQL

    DoCmd.DeleteObject acQuery, strQry
    Application.FollowHyperlink cheminFichier
End Sub


