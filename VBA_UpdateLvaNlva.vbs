Attribute VB_Name = "VBA_UpdateLvaNlva"
Sub UpdateLvaNlva()

    'Update LVA/NLVA based on Unit Price
    If unitPrice > 50000 Then
        DataSheet.Range("T" & CurrRow).Value = "NLVA"
    Else
        DataSheet.Range("T" & CurrRow).Value = "LVA"
    End If

End Sub
