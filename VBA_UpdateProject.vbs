Attribute VB_Name = "VBA_UpdateProject"
Sub UpdateProject()

    'Update Project
    If Not IsError(DataSheet.Range("AD" & CurrRow)) Then
    
        DataSheet.Range("AM" & CurrRow).Value = "230" & Right(DataSheet.Range("AD" & CurrRow).Value, 5)
    
    Else
    
        DataSheet.Range("AM" & CurrRow).Formula = "=CONCATENATE(" & "230" & ",RIGHT(AD" & CurrRow & ",5))"
    
    End If

End Sub
