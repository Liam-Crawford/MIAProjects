Attribute VB_Name = "Format_ComSeg"
Sub FormatCommonSeg()
    Introduction
    OnePage "Total Market Segmentation"
    OnePage "Retail Sales By Marque"
    OnePage "Retail Share By Marque"
    OnePage "Retail Sales By Buyer Type"
    MultiPage "Retail Sales By Buyer Type Fuel", False
    MultiPage "Segment Model Passenger", False
    Marque "Marque Passenger"
    Marque "Marque SUV"
    Marque "Marque Passenger + SUV"
    Marque "Marque Light Commercial"
    Marque "Marque Heavy Commercial"
    SegmentModel "Segment Model SUV"
    SegmentModel "Segment Model Light Commercial"
    SegmentModel "Segment Model Heavy Commercial"
    MultiPage "Marque & Model (Segmented)", True
    MultiPage "Marque & Model (Para|Low Vol)", True
    OnePage "Marque & Model (Unsegmented)"
    
End Sub

Sub Introduction()
    Worksheets("Introduction").Activate
    Functions.PageSetupPortrait1Page
End Sub

Sub OnePage(s As String)
    Worksheets(s).Activate
    Functions.PageSetupPortrait1Page
    Functions.MergeAndCentre Range("A1:B1")
End Sub

Sub MultiPage(s As String, b As Boolean)
    Worksheets(s).Activate
    Functions.PageSetupPortrait
    Functions.MergeAndCentre Range("A1:B1")
    
    If b Then
        row = ActiveSheet.UsedRange.Rows.Count - 3
        Functions.AutoFitColumn Range("A7", "A" & row)
    End If
End Sub

Sub Marque(s As String)
    Worksheets(s).Activate
    Functions.PageSetupPortrait1Page
    Functions.MergeAndCentre Range("A1:B1")
    
    Dim row As Integer
    row = 7
    Do While True
        If Not IsEmpty(ActiveSheet.Cells(row, 1)) Then
            row = row + 1
        Else
            Exit Do
        End If
    Loop
    row = row - 1
    
    Range("A7", ("K" & row)).Select
    Application.CommandBars.ExecuteMso ("SortAscendingExcel")
    Application.CommandBars.ExecuteMso ("SortDescendingExcel")
    Application.CommandBars.ExecuteMso ("SortAscendingExcel")
    
    row = ActiveSheet.UsedRange.Rows.Count - 3
    Functions.AutoFitColumn Range("A7", "A" & row)
    
    Range("A1").Select
End Sub

Sub SegmentModel(s As String)
    Worksheets(s).Activate
    Functions.PageSetupPortrait
    Functions.MergeAndCentre Range("A1:B1")
    
    Dim rowFirst, rowLast, row As Integer
    rowFirst = 7
    rowLast = 7
    
    Do While True
        Do While True
            If Not IsEmpty(ActiveSheet.Cells(rowLast, 1)) Then
                rowLast = rowLast + 1
            Else
                Exit Do
            End If
        Loop
        
        rowLast = rowLast - 1
        
        Range(("A" & rowFirst), ("K" & rowLast)).Select
        ' useless sorting functions
        Application.CommandBars.ExecuteMso ("SortAscendingExcel")
        Application.CommandBars.ExecuteMso ("SortDescendingExcel")
        Application.CommandBars.ExecuteMso ("SortAscendingExcel")
        
        rowFirst = rowLast + 5
        rowLast = rowLast + 5
        
        If IsEmpty(ActiveSheet.Cells(rowLast, 1)) Then
            Exit Do
        End If
    Loop
    
    row = ActiveSheet.UsedRange.Rows.Count - 3
    Functions.AutoFitColumn Range("A7", "A" & row)
    
    Range("A1").Select
End Sub
