Attribute VB_Name = "VBA_UpdateAssetClass"
Sub UpdateAssetClass()
    
    On Error Resume Next
    Set AssetClassFound = DecisionSheet.Range("H:H").Find(what:=AssetType, LookIn:=xlValues, lookat:=xlPart)
    
    If Not AssetClassFound Is Nothing Then
    
        If PR_Nature = "LIVE" Then
        
            If LvaNlva = "LVA" Then
            
                If DecisionSheet.Range("J" & AssetClassFound.Row).Value Is Empty Then
                
                    DecisionSheet.Range("J" & AssetClassFound.Row).Formula = "=Data!U" & CurrRow
                    
                Else
                
                    DataSheet.Range("U" & CurrRow).Value = DecisionSheet.Range("J" & AssetClassFound.Row).Value
                    
                End If
            
            ElseIf LvaNlva = "NLVA" Then
            
                If DecisionSheet.Range("I" & AssetClassFound.Row).Value Is Empty Then
                
                    DecisionSheet.Range("I" & AssetClassFound.Row).Formula = "=Data!U" & CurrRow
                    
                Else
                
                    DataSheet.Range("U" & CurrRow).Value = DecisionSheet.Range("I" & AssetClassFound.Row).Value
                    
                End If
                                
            End If
            
            DataSheet.Range("V" & CurrRow).Value = "NA"
        
        ElseIf PR_Nature = "AUC" Then
        
            If DecisionSheet.Range("K" & AssetClassFound.Row) Is Empty Then
                
                DecisionSheet.Range("K" & AssetClassFound.Row).Formula = "=Data!V" & CurrRow
                    
            Else
        
                DataSheet.Range("V" & CurrRow).Value = DecisionSheet.Range("K" & AssetClassFound.Row).Value
                
            End If
            
            DataSheet.Range("U" & CurrRow).Value = "NA"
        
        End If
    
    Else
    
        'if Asset Type not found then add the Details which user will input
        DecisionSheet.Range("H" & lastRow(DecisionSheet, 8) + 1).Value = AssetType
        
        If PR_Nature = "LIVE" Then
        
            If LvaNlva = "LVA" Then
            
                DecisionSheet.Range("J" & lastRow(DecisionSheet, 8)).Formula = "=Data!U" & CurrRow
            
            ElseIf LvaNlva = "NLVA" Then
            
                DecisionSheet.Range("I" & lastRow(DecisionSheet, 8)).Formula = "=Data!U" & CurrRow
            
            End If
        
        ElseIf PR_Nature = "AUC" Then
        
            DecisionSheet.Range("K" & lastRow(DecisionSheet, 8)).Formula = "=Data!V" & CurrRow
        
        End If
    
    End If
    
End Sub
