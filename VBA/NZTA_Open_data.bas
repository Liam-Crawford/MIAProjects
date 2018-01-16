Attribute VB_Name = "NZTA_Open_data"
' vbCrLf = new line
' 3 - Body Type
' 7 - GVM
' 14 - Motive Power
' 27 - Vehicle Type

Sub RunAll()
    CopyLightByType
    CopyLightCByType
    CopyHeavyCByType
    CopyMotorcycleByBody
    CopyOtherByType
End Sub

Sub CopyLightByType()
    Dim rowCount, i, j As Long, k As Integer, s As String
    s = "PASSENGER CAR/VAN"
    
    Workbooks(2).Activate
    rowCount = Worksheets(1).UsedRange.Rows.Count
    j = 1
    
    For i = 1 To rowCount
        If (Worksheets(1).Cells(i, 27).Value = s) Then
            If (Worksheets(1).Cells(i, 7).Value < 3501) Then
                j = j + 1
                Range(("A" & i), ("AF" & i)).Copy
                ActiveSheet.Paste Destination:=Worksheets(2).Range(("A" & j), ("AF" & j))
            End If
        End If
    Next
    'MsgBox j
End Sub

Sub CopyLightCByType()
    Dim rowCount, i, j As Long, k As Integer, s As String
    Dim vehicles(1 To 3) As String
    vehicles(1) = "GOODS VAN/TRUCK/UTILITY"
    vehicles(2) = "BUS"
    vehicles(3) = "MOTOR CARAVAN"
    
    Workbooks(2).Activate
    rowCount = Worksheets(1).UsedRange.Rows.Count
    j = 1
    
    For i = 1 To rowCount
        For k = 1 To 3
            s = vehicles(k)
            If (Worksheets(1).Cells(i, 27).Value = s) Then
                If (Worksheets(1).Cells(i, 7).Value < 3501) Then
                    j = j + 1
                    Range(("A" & i), ("AF" & i)).Copy
                    ActiveSheet.Paste Destination:=Worksheets(3).Range(("A" & j), ("AF" & j))
                End If
            End If
        Next
    Next
    'MsgBox j
End Sub

Sub CopyHeavyCByType()
    Dim rowCount, i, j As Long, k As Integer, s As String
    Dim vehicles(1 To 3) As String
    vehicles(1) = "GOODS VAN/TRUCK/UTILITY"
    vehicles(2) = "BUS"
    vehicles(3) = "MOTOR CARAVAN"
    
    Workbooks(2).Activate
    rowCount = Worksheets(1).UsedRange.Rows.Count
    j = 1
    
    For i = 1 To rowCount
        For k = 1 To 3
            s = vehicles(k)
            If (Worksheets(1).Cells(i, 27).Value = s) Then
                If (Worksheets(1).Cells(i, 7).Value > 3500) Then
                    j = j + 1
                    Range(("A" & i), ("AF" & i)).Copy
                    ActiveSheet.Paste Destination:=Worksheets(4).Range(("A" & j), ("AF" & j))
                End If
            End If
        Next
    Next
    'MsgBox j
End Sub

Sub CopyMotorcycleByBody()
    Dim rowCount, i, j As Long, k As Integer, s As String
    s = "MOTORCYCLE"
    
    Workbooks(2).Activate
    rowCount = Worksheets(1).UsedRange.Rows.Count
    j = 1
    
    For i = 1 To rowCount
        If (Worksheets(1).Cells(i, 3).Value = s) Then
            j = j + 1
            Range(("A" & i), ("AF" & i)).Copy
            ActiveSheet.Paste Destination:=Worksheets(5).Range(("A" & j), ("AF" & j))
        End If
    Next
    'MsgBox j
End Sub

Sub CopyOtherByType()
    Dim rowCount, i, j As Long, k As Integer, s As String
    Dim vehicles(1 To 7) As String
    vehicles(1) = "AGRICULTURAL MACHINE"
    vehicles(2) = "HIGH SPEED AGRICULTURAL VEHICLE"
    vehicles(3) = "MOBILE MACHINE"
    vehicles(4) = "SPECIAL PURPOSE VEHICLE"
    vehicles(5) = "TRACTOR"
    vehicles(6) = "TRAILER NOT DESIGNED FOR H/WAY USE"
    vehicles(7) = "TRAILER/CARAVAN"
    
    Workbooks(2).Activate
    rowCount = Worksheets(1).UsedRange.Rows.Count
    j = 1
    
    For i = 1 To rowCount
        For k = 1 To 7
            s = vehicles(k)
            If (Worksheets(1).Cells(i, 27).Value = s) Then
                j = j + 1
                Range(("A" & i), ("AF" & i)).Copy
                ActiveSheet.Paste Destination:=Worksheets(6).Range(("A" & j), ("AF" & j))
            End If
        Next
    Next
    'MsgBox j
End Sub

Sub CopyRowsElectric()
    Dim rowCount, i, j As Long, k As Integer
    
    rowCount = Worksheets(1).UsedRange.Rows.Count
    j = 1
    
    For i = 1 To rowCount
        If (Worksheets(1).Cells(i, 14).Value = "ELECTRIC") Then
            j = j + 1
            Range(("A" & i), ("AF" & i)).Copy
            ActiveSheet.Paste Destination:=Worksheets(3).Range(("A" & j), ("AF" & j))
        End If
    Next
End Sub

Sub FindAllTypes()
    Dim rowCount, i As Long, s(1 To 50), output As String, j, t As Integer
    t = 27
    
    Workbooks(2).Activate
    rowCount = Worksheets(1).UsedRange.Rows.Count
    j = 1
    s(j) = Cells(2, t).Value
    
    For i = 3 To rowCount
        If Cells(i, t).Value <> s(j) Then
            j = j + 1
            s(j) = Cells(i, t).Value
        End If
    Next
    
    For i = 1 To j
        output = output + s(i) + vbCrLf
    Next
    MsgBox output
    
End Sub

