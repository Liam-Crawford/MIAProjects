Attribute VB_Name = "Format_Filesforemail"
Sub Format_FullyEquipped()
    Dim i As Integer
    
    For i = 2 To 20
        Call Format(i)
    Next
    
End Sub

Sub Format_Autofile()
    For i = 1 To 9
        Call Format(i)
    Next
End Sub

Sub Format_EV()
    For i = 1 To 3
        Call Format(i)
    Next
End Sub

Sub Format(ByVal i As Integer)
    Worksheets(i).Activate
    
    Range("B1", Range("C1").End(xlToRight)).Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Columns.AutoFit
        
    Functions.AutoFitFirstColumnForSATTUS
End Sub

Sub test()
    Worksheets(2).Activate
    Range("B7", Range("C7").End(xlToRight)).Select
    Selection.Columns.AutoFit

End Sub
