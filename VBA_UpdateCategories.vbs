Attribute VB_Name = "VBA_UpdateCategories"
Sub UpdateCategories()
    
    'Update Category
    On Error Resume Next
    Set CategoryFound = DecisionSheet.Range("D:D").Find(what:=AssetType, LookIn:=xlValues, lookat:=xlPart)
    
    If Not CategoryFound Is Nothing Then
        
        DataSheet.Range("Q" & CurrRow).Value = DecisionSheet.Range("E" & CategoryFound.Row).Value
        DataSheet.Range("R" & CurrRow).Value = DecisionSheet.Range("F" & CategoryFound.Row).Value
        
    Else
        
        DecisionSheet.Range("D" & lastRow(DecisionSheet, 4) + 1).Value = AssetType
        DecisionSheet.Range("E" & lastRow(DecisionSheet, 4)).Formula = "=Data!Q" & CurrRow
        DecisionSheet.Range("F" & lastRow(DecisionSheet, 4)).Formula = "=Data!R" & CurrRow
        
    End If

End Sub
