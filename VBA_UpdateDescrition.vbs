Attribute VB_Name = "VBA_UpdateDescrition"
Sub UpdateDescription()
    
    'Description and Qty
    If unitPrice > 50000 Then
        
        If Unit = "Each" Then
            
            DataSheet.Range("P" & CurrRow).Value = DataSheet.Range("N" & CurrRow) & "-" & Description & "-QTY-1"
            DataSheet.Range("AL" & CurrRow).Value = Qty
            
        Else
        
            DataSheet.Range("P" & CurrRow).Value = DataSheet.Range("N" & CurrRow) & "-" & Description & "-QTY-1 " & Unit
            DataSheet.Range("AL" & CurrRow).Value = Qty & Unit
            
        End If
        
        DataSheet.Range("AK" & CurrRow).Value = Qty
        
    Else
    
        If Unit = "Each" Then
            
            DataSheet.Range("P" & CurrRow).Value = DataSheet.Range("N" & CurrRow) & "-" & Description & "-QTY-" & Qty
            DataSheet.Range("AL" & CurrRow).Value = Qty
            
        Else
        
            DataSheet.Range("P" & CurrRow).Value = DataSheet.Range("N" & CurrRow) & "-" & Description & "-QTY-1 " & Unit
            DataSheet.Range("AL" & CurrRow).Value = Qty & Unit
            
        End If
        
        DataSheet.Range("AK" & CurrRow).Value = 1
    
    End If
    
End Sub
