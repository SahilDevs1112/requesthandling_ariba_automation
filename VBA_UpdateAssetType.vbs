Attribute VB_Name = "VBA_UpdateAssetType"
Sub UpdateAssetType()
    
    On Error Resume Next
    Set DescriptionFound = DecisionSheet.Range("A:A").Find(what:=Description, LookIn:=xlValues, lookat:=xlPart)
    
    If Not DescriptionFound Is Nothing Then
    
        DataSheet.Range("L" & CurrRow).Value = DecisionSheet.Range("B" & DescriptionFound.Row).Value
    
    Else
    
        DecisionSheet.Range("A" & lastRow(DecisionSheet, 1) + 1) = Description
        AssetTypeFromUser = InputBox("What will be the Asset Class for Description - " & Description, "Give a Asset Class")
        DecisionSheet.Range("B" & lastRow(DecisionSheet, 1)).Value = AssetTypeFromUser
        DataSheet.Range("L" & CurrRow).Value = AssetTypeFromUser
    
    End If
    
End Sub
