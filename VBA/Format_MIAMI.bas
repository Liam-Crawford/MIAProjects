Attribute VB_Name = "Format_MIAMI"
Sub FormatMIAMIOutput()
    Dim currentDate As Long
    Dim rowCount As Integer
    currentDate = 20171001
    
    Range("A:M").Select
    Selection.Delete
    Range("C:I").Select
    Selection.Delete
    Range("D:P").Select
    Selection.Delete
    Range("F:I").Select
    Selection.Delete
    Range("G:L").Select
    Selection.Delete
    
    Range("A2", Range("F" & Rows.Count).End(xlUp).Address).Select
    
    Functions.SortDataDescending ("E2")
    Functions.SortDataDescending ("D2")
    
    rowCount = DeleteRows(currentDate)
    DeleteInvalidVFACTS rowCount
    
    Range("A2", Range("F" & Rows.Count).End(xlUp).Address).Select
    
    Functions.SortDataAscending ("C2")
    Functions.SortDataAscending ("B2")
    Functions.SortDataAscending ("A2")
    
End Sub

Function DeleteRows(currentDate As Long) As Integer
    Dim dateToCheck As Long
    Dim i As Integer
        
    i = 3
    
    While CLng(Range("D" & i).Value) > currentDate
         i = i + 1
    Wend
    
    Range("A" & i, Range("F" & Rows.Count).End(xlUp).Address).Select
    Selection.Delete
    
    DeleteRows = i
End Function

Sub DeleteInvalidVFACTS(rowCount As Integer)
    Dim i As Integer
    i = 2
    
    For j = 2 To rowCount
        If IsNumeric(Range("F" & i).Value) Then
            If (Range("F" & i).Value < 1) Then
                Rows(i).Delete
            ElseIf (Range("F" & i).Value > 46) Then
                Rows(i).Delete
            Else
                i = i + 1
            End If
        Else
            Rows(i).Delete
        End If
    Next
End Sub
