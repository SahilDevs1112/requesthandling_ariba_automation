Attribute VB_Name = "VBA_LookupPlantCode"
Sub LookupPlantCode()
        
    'Plant and Location code
    plantCode = formulaeSheet.Range("B50").Value
Start:
    On Error Resume Next
    Set plantCodeFound = plantCodeSheet.UsedRange.Find(what:=plantCode, LookIn:=xlValues, lookat:=xlWhole)
    
    If Not plantCodeFound Is Nothing Then
        
        DataSheet.Range("M" & CurrRow).Value = plantCodeFound.Value
        DataSheet.Range("N" & CurrRow).Value = plantCodeSheet.Cells(plantCodeFound.Row, plantCodeFound.Column + 1)
        DataSheet.Range("J" & CurrRow).Value = "LIVE"
        
    Else
    
        Set plantCodeFound = AUCPlantCodeSheet.UsedRange.Find(what:=plantCode, LookIn:=xlValues, lookat:=xlWhole)
        
        On Error Resume Next
        If Not plantCodeFound Is Nothing Then
            
            DataSheet.Range("M" & CurrRow).Value = plantCodeFound.Value
            DataSheet.Range("N" & CurrRow).Value = AUCPlantCodeSheet.Cells(plantCodeFound.Row, plantCodeFound.Column + 1)
            DataSheet.Range("J" & CurrRow).Value = "AUC"
        
        Else
        
            'Here common Plant Codes will be updated and later needs to be reviewed
            
            plantCode = formulaeSheet.Cells(formulaeSheet.Range("D50:D57").Find(what:=cityAriba, LookIn:=xlValues, lookat:=xlPart).Row, _
            formulaeSheet.Range("D50:K50").Find(what:=companyCode, LookIn:=xlValues, lookat:=xlPart).Column)
            
            GoTo Start
        
        End If
    
    End If
    
    DataSheet.Range("O" & CurrRow).Value = "NA"
    
End Sub
