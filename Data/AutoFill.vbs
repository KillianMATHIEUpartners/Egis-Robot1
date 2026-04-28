Sub AutoFillFormulas()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Récupérer la feuille Workpaper
    Set ws = ThisWorkbook.Sheets("Workpaper for migration")
    
    ' Trouver la dernière ligne avec des données (colonne E)
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    
    ' AutoFill colonne G (ProjectName) de G6 jusqu'à lastRow
    ws.Range("G6").AutoFill Destination:=ws.Range("G6:G" & lastRow)
    
    ' AutoFill colonne J (LegalEntityName) de J6 jusqu'à lastRow
    ws.Range("J6").AutoFill Destination:=ws.Range("J6:J" & lastRow)
    
    ' AutoFill colonne L (OwningOrganizationName) de L6 jusqu'à lastRow
    ws.Range("L6").AutoFill Destination:=ws.Range("L6:L" & lastRow)
    
    ' AutoFill colonne N (SourceTemplateName) de N6 jusqu'à lastRow
    ws.Range("N6").AutoFill Destination:=ws.Range("N6:N" & lastRow)
    
End Sub
