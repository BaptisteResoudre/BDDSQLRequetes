Private Sub Rapport_Click()
    Dim strSQL As String
    Dim nomFichier As String
    Dim dbs As Database
    Dim qdf As QueryDef
    Dim thisDate As Date

    thisDate = Today
    strQry = "fe"
    nomFichier = "covfefe"
    
    Dim cheminFichier As String
    cheminFichier = "C:\Users\Resoudrien\"
    cheminFichier = cheminFichier & "rapport_assainissement_" & nomFichier & ".xls"
    
    MsgBox ("Excel exporte ici : " + cheminFichier)

    Set dbs = CurrentDb
    Set qdf = dbs.CreateQueryDef(strQry)

    ' REQUETE ASSAINISSEMENT SEXE
    strSQL = "SELECT Adherents.Sexe_adh, Adherents.Nom_adh, Adherents.Prenom_adh FROM Adherents INNER JOIN Cotisations ON Adherents.Num_adh = Cotisations.Num_adh WHERE Cotisations.DateCotis>#31/12/2020# AND Sexe_adh IS NULL;"
    qdf.Sql = strSQL
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, strQry, cheminFichier, True, "Sexe_adh Invalide"
        
    ' REQUETE ASSAINISSEMENT ANNEE
    strSQL = "SELECT Adherents.Sexe_adh, Adherents.Nom_adh, Adherents.Prenom_adh, Adherents.AnneeNaissance, Passages.Date_Pass from Adherents INNER JOIN Passages ON Adherents.Num_adh = Passages.Num_adh WHERE AnneeNaissance NOT LIKE '[0-9][0-9][0-9][0-9]' and Passages.Date_Pass>#31/12/2020#;"
    qdf.Sql = strSQL
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, strQry, cheminFichier, True, "Annee Naissance Invalide"
    
    ' REQUETE ASSAINISSEMENT NATIONALITE
    strSQL = "SELECT Adherents.Sexe_adh, Adherents.Nom_adh, Adherents.Prenom_adh, Adherents.Nationalite_adh FROM Adherents INNER JOIN Passages ON Adherents.Num_adh = Passages.Num_adh WHERE Passages.Date_Pass>#31/12/2020# AND Nationalite_adh IS NULL;"
    qdf.Sql = strSQL
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, strQry, cheminFichier, True, "Nationalite Invalide"
        
   ' REQUETE ASSAINISSEMENT Code Postal
    strSQL = "SELECT Adherents.Sexe_adh, Adherents.Nom_adh, Adherents.Prenom_adh, Adherents.Nationalite_adh, Adherents.CP_adh FROM Adherents INNER JOIN Passages ON Adherents.Num_adh = Passages.Num_adh WHERE Passages.Date_Pass>#31/12/2020# AND CP_adh NOT LIKE '[0-9][0-9][0-9][0-9][0-9]';"
    qdf.Sql = strSQL
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, strQry, cheminFichier, True, "Code Postal Invalide"

    DoCmd.DeleteObject acQuery, strQry
    Application.FollowHyperlink cheminFichier
End Sub
