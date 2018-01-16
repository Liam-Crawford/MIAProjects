Attribute VB_Name = "Format_QuarterlyPDI"
Sub CopyRows()
    Dim i, j, rowCount As Integer
    Dim oldMake, newMake, distributor, filePath1, filePath2, fileName As String
    Dim header As Range

    ' *** REMEMBER TO CHANGE FILEPATH ***
    filePath1 = "Z:\MIA 2017\Product and Safety Committee\PDI lists ex NZTA and VIRMs for light vehicles\"
    filePath2 = "PDI Lists 2016\4 December\"
    
    ' *** REMEMBER TO CHANGE MONTH AND YEAR ***
    fileName = " PDI Report 31 December 2016.xlsx"
    
    Dim dict As New Scripting.Dictionary
        
    i = 2
    rowCount = Worksheets(1).UsedRange.Rows.Count
    Set header = Range("A1:G1")
    
    newMake = ActiveSheet.Cells(i, 5).Value
    oldMake = newMake
    
    Do While i <= rowCount
    distributor = Functions.findDistributor(newMake)
    If dict.Exists(distributor) Then
        j = dict(distributor)
    Else
        dict.Add distributor, 2
        j = 2
        
        Dim ws As Worksheet
        Set ws = Sheets.Add(After:=Sheets(Worksheets.Count))
        ws.Name = distributor
        
        Worksheets(1).Activate
        header.Copy
        ActiveSheet.Paste Destination:=Worksheets(distributor).Range("A1:G1")
        
        Worksheets(distributor).Activate
        Columns(1).Select
        Selection.ColumnWidth = 13
        Columns(2).Select
        Selection.ColumnWidth = 8
        Columns(3).Select
        Selection.ColumnWidth = 8
        Columns(4).Select
        Selection.ColumnWidth = 21
        Columns(5).Select
        Selection.ColumnWidth = 15
        Columns(6).Select
        Selection.ColumnWidth = 20
        Columns(7).Select
        Selection.ColumnWidth = 11
        
        Range("A1").Select
    End If
        Do While True
            Worksheets(1).Activate
            
            newMake = ActiveSheet.Cells(i, 5).Value
            If oldMake <> newMake Then
                oldMake = newMake
                Exit Do
            End If
            
            Range(("A" & i), ("G" & i)).Copy
            ActiveSheet.Paste Destination:=Worksheets(distributor).Range(("A" & j), ("G" & j))
            j = j + 1
            i = i + 1
        Loop
        dict(distributor) = j
    Loop
    
    Set dict = Nothing
    
End Sub
