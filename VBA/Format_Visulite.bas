Attribute VB_Name = "Format_Visulite"
Sub Format_Visulite()
    Dim month As String
    month = "Dec 17"
    
    Worksheets(1).Activate
    Format_Sheet "Total by Marque - " & month
    Worksheets(2).Activate
    Format_Sheet "Total by Marque - " & month
    Worksheets(3).Activate
    Format_Sheet "Passenger by Marque - " & month
    Worksheets(4).Activate
    Format_Sheet "Passenger by Marque - " & month
    Worksheets(5).Activate
    Format_Sheet "Commercial by Marque - " & month
    Worksheets(6).Activate
    Format_Sheet "Commercial by Marque - " & month
    Worksheets(7).Activate
    Format_Sheet "Passenger by Model - " & month
    Worksheets(8).Activate
    Format_Sheet "Passenger by Model - " & month
    Worksheets(9).Activate
    Format_Sheet "Commercial by Model - " & month
    Worksheets(10).Activate
    Format_Sheet "Commercial by Model - " & month
    Worksheets(11).Activate
    Format_Sheet "Total by Model - " & month
    Worksheets(12).Activate
    Format_Sheet "Total by Model - " & month
    Worksheets(13).Activate
    Format_Sheet "Total by Segment - " & month
    Worksheets(14).Activate
    Format_Sheet "Total by Segment - " & month
    Worksheets(15).Activate
    Format_Sheet "Total by Model - " & month
    Worksheets(16).Activate
    Format_Sheet "Total by Model - " & month
    
    Worksheets(1).Activate
    Range("A1").Select
    
End Sub

Sub Format_Sheet(title As String)
    Range("A1").FormulaR1C1 = title
    Range("A10", Range("A11").End(xlDown)).Columns.AutoFit
    Range("A1").Select
End Sub

Sub test()
    Range("A1:C1").Select
    Application.CommandBars.ExecuteMso ("AlignLeft")
End Sub
