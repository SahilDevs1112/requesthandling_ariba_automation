Attribute VB_Name = "VBA_UpdateDivisionCode"
Sub UpdateDivisionCode()

    'Update Division Code based on Asset Class
    
    If Left(Trim(DataSheet.Range("AI" & CurrRow).Value), 3) = "TPX" Then
    
        DataSheet.Range("AH" & CurrRow).Value = Right(Trim(DataSheet.Range("AI" & CurrRow).Value), 4)
    
    Else
    
        On Error Resume Next
        Set DivCodeFound = DecisionSheet.Range("M:M").Find(what:=AssetType, LookIn:=xlValues, lookat:=xlPart)
        
        If Not DivCodeFound Is Nothing Then
            
            DataSheet.Range("AH" & CurrRow).Value = DecisionSheet.Range("N" & DivCodeFound.Row).Value
            
        Else
            
            DecisionSheet.Range("M" & lastRow(DecisionSheet, 13) + 1).Value = AssetType
            DecisionSheet.Range("N" & lastRow(DecisionSheet, 13)).Formula = "=Data!AH" & CurrRow
            
        End If
    
    
    End If

End Sub
