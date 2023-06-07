Attribute VB_Name = "VBA_UpdateCapitalOrLease"
Sub UpdateCapitalOrLease()
    
    'Update capital/lease as Capital as default - need to change this in UIpath when one lease PR appears
    DataSheet.Range("S" & CurrRow).Value = "CAPITAL"

End Sub
