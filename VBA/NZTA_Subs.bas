Attribute VB_Name = "NZTA_Subs"
Sub NZTA_001_002()
Attribute NZTA_001_002.VB_Description = "Format 001"
Attribute NZTA_001_002.VB_ProcData.VB_Invoke_Func = " \n14"
' Format 001, 001N, 002, 002N

    Functions.MergeAndCentre Range("A1:Y1")
    Functions.AutoFit Range("A1:Y1")
    
    Columns("C:X").ColumnWidth = 4.6
    Range("A1:Y1").Select
    
    Functions.PageSetupLandscape
End Sub

Sub NZTA_001A()
' Format 001A

    Functions.MergeAndCentre Range("A1:D1")
    Functions.AutoFit Range("A1:D1")

    Columns("A:D").Select
    Selection.ColumnWidth = 16
    
    Range("A1:D1").Select
    
    Functions.PageSetupPortrait1Page
End Sub

Sub NZTA_002A()
' Format 002A

    Functions.MergeAndCentre Range("A1:D1")
    Functions.AutoFit Range("A1:D1")
    
    Columns("A:D").Select
    Selection.ColumnWidth = 17
    
    Range("A1:D1").Select
    
    Functions.PageSetupPortrait1Page
End Sub

Sub NZTA_006()
' Format 006, 006N, 006X

    Functions.MergeAndCentre Range("A1:T1")
    Functions.AutoFit Range("A1:T1")
    Functions.WrapText Range("B2:S2")
    
    Columns("C:R").ColumnWidth = 6
    Range("B2").ColumnWidth = 11
    Range("S2").ColumnWidth = 9
    Rows("2:2").EntireRow.AutoFit
    
    Range("A2", Range("A3").End(xlDown)).Columns.AutoFit
    Range("T2").Columns.AutoFit
    
    Range("A1:T1").Select
    
    Functions.PageSetupLandscape
End Sub

Sub NZTA_008()
' Format 008, 008N, 008X

    Functions.MergeAndCentre Range("A1:V1")
    Functions.AutoFit Range("A1:V1")
    Functions.WrapText Range("B2:U2")

    Columns("C:R").ColumnWidth = 6
    Range("B2").ColumnWidth = 11
    Range("S:U").ColumnWidth = 10
    Rows("2:2").EntireRow.AutoFit
    
    Range("A2", Range("A3").End(xlDown)).Columns.AutoFit
    Range("V2").Columns.AutoFit
    
    Range("A1:V1").Select
    
    Functions.PageSetupLandscape
End Sub

Sub NZTA_051()
' Format 051

    Functions.MergeAndCentre Range("B1:E1")
    Functions.AutoFit Range("A1:E1")

    Range("A1").Select
    
    Functions.PageSetupPortrait1Page
End Sub

Sub NZTA_054()
' Format 054

    Functions.MergeAndCentre Range("A1:F1")
    Functions.AutoFit Range("A1:F1")

    Range("A1:F1").Select
    
    Functions.PageSetupPortrait1Page
End Sub

Sub NZTA_064()
' Format 064N, 064X, 065N, 065X

    Functions.MergeAndCentre Range("A1:Z1")
    Functions.AutoFit Range("A1:Z1")

    Range("A1:Z1").Select
    
    Functions.PageSetupLandscape
End Sub

Sub NZTA_Y065N()
' Format Y-065N

    Functions.MergeAndCentre Range("B1:Z1")
    Functions.AutoFit Range("A1:Z1")

    Range("B1:Z1").Select
    
    Functions.PageSetupLandscape
End Sub

Sub NZTA_MIA_DEREG_MONTHLY()
' Format MIA_DEREG_MONTHLY

    Functions.AutoFit Range("A1:D1")

    Range("A1").Select
    
    Functions.PageSetupPortrait
End Sub

Sub NZTA_N7USG()
' Format N7-USG

    Functions.AutoFit Range("A1:D1")

    Range("A1").Select
    
    Functions.PageSetupPortrait1Page
End Sub

Sub NZTA_UMM_AGE()
' Format U7MM_AGE, U8MM_AGE

    Functions.AutoFit Range("A1:Q1")

    Range("A1").Select
    
    Functions.PageSetupLandscape
End Sub

Sub NZTA_VTyp1013()
' Format VTyp10-13, YTD_RENTALS_NEW

    Functions.AutoFit Range("A1:E1")
    
    Range("A1").Select
    
    Functions.PageSetupPortrait1Page
End Sub

Sub NZTA_X085N()
' Format X-085N

    Functions.MergeAndCentre Range("A1:R1")
    Functions.AutoFit Range("A1:R1")
    Columns("B:L").ColumnWidth = 8
    
    Range("A1").Select
    
    Functions.PageSetupLandscape1Page
End Sub

Sub NZTA_Y_MPC_A()
' Format Y_MPC_A
    Functions.MergeAndCentre Range("B1:O1")
    Range("O2").FormulaR1C1 = "YTD"
    Functions.AutoFit Range("A1:O1")
    Columns("C:N").ColumnWidth = 5
    
    Range("A1").Select
    
    Functions.PageSetupPortrait1Page
End Sub

Sub NZTA_Y00()
' Format Y-001AN, Y-001AN_2AN, Y-001AX, Y-002AN, Y-002AX
    Functions.MergeAndCentre Range("B1:N1")
    Range("N2").FormulaR1C1 = "YTD"
    Functions.AutoFit Range("A1:N1")
    Columns("B:M").ColumnWidth = 6
    
    Range("A1").Select
    
    Functions.PageSetupPortrait1Page
End Sub

Sub NZTA_Y080N()
' Format Y-080N
    Range("X1").FormulaR1C1 = "YTD"
    Functions.AutoFit Range("A1:X1")
    
    Range("A1").Select
    
    Functions.PageSetupLandscape
End Sub

Sub NZTA_Y081N()
' Format Y-081N
    Dim rowCount As Integer
    
    Range("V1").Select
    Selection.AutoFill Destination:=Range("V1:W1"), Type:=xlFillDefault
    Range("W1").FormulaR1C1 = "YTD"
    
    rowCount = Worksheets(1).UsedRange.Rows.Count + 1
    Range("I" & (rowCount)).Select
    ActiveCell.FormulaR1C1 = "TOTALS"
    
    Range(("K" & 2), ("W" & rowCount)).Select
    Application.CommandBars.ExecuteMso ("AutoSum")
    
    Range("W2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Font.Bold = True
    
    Range("I" & rowCount, "W" & rowCount).Select
    Selection.Font.Bold = True
    
    Functions.AutoFit Range("A1:W1")
    Columns("K:V").ColumnWidth = 4.3
    Range("A1").Select
    
    Functions.PageSetupLandscape
End Sub

Sub NZTA_Y084N()
' Format Y-084N, Y-085N, YTD_USED_CARS, YTD_USED_COM
    Dim rowCount As Integer
    
    Range("N1").Select
    Selection.AutoFill Destination:=Range("N1:O1"), Type:=xlFillDefault
    Range("O1").FormulaR1C1 = "YTD"
    
    rowCount = Worksheets(1).UsedRange.Rows.Count + 1
    Range("A" & (rowCount)).Select
    ActiveCell.FormulaR1C1 = "TOTALS"
    
    Range(("C" & 2), ("O" & rowCount)).Select
    Application.CommandBars.ExecuteMso ("AutoSum")
    
    Range("O2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Font.Bold = True
    
    Range("A" & rowCount, "O" & rowCount).Select
    Selection.Font.Bold = True
    
    Functions.AutoFit Range("A1:O1")
    Columns("C:O").ColumnWidth = 5.3
    Range("A1").Select
    
    Functions.PageSetupPortrait
End Sub

Sub NZTA_YMPC50()
' Format Y-MPC50
    Functions.MergeAndCentre Range("B1:P1")
    Range("P2").FormulaR1C1 = "YTD"
    Functions.AutoFit Range("A1:P1")
    Columns("D:P").ColumnWidth = 5
    
    Range("A1").Select
    
    Functions.PageSetupPortrait1Page
End Sub

Sub NZTA_YMPC51()
' Format Y-MPC51
    Functions.MergeAndCentre Range("B1:P1")
    Range("P2").FormulaR1C1 = "YTD"
    Functions.AutoFit Range("A1:P1")
    Columns("D:P").ColumnWidth = 5
    
    Range("A1").Select
    
    Functions.PageSetupLandscape
End Sub

Sub NZTA_YRYCOMMSM1()
' Format YRY-COMMS_M1
    Range("B1").Cut Range("A1")
    Functions.MergeAndCentre Range("A1:AA1")
    Functions.AutoFit Range("A1:AA1")
    
    Range("A1").Select
    
    Functions.PageSetupLandscape
End Sub
