Attribute VB_Name = "Functions"
Function AutoFit(r As Range)
    r.Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Rows.AutoFit
    Selection.Columns.AutoFit
End Function

Function MergeAndCentre(r As Range)
    r.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
End Function

Function WrapText(r As Range)
    r.Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Function

Function PageSetupLandscape()
'
' Sets up the print gridlines, page width, orientation, and margins
' for a landscape page
'

'
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.236220472440945)
        .RightMargin = Application.InchesToPoints(0.236220472440945)
        .TopMargin = Application.InchesToPoints(0.354330708661417)
        .BottomMargin = Application.InchesToPoints(0.354330708661417)
        .HeaderMargin = Application.InchesToPoints(0.31496062992126)
        .FooterMargin = Application.InchesToPoints(0.31496062992126)
        .PrintHeadings = False
        .PrintGridlines = True
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
End Function

Function PageSetupLandscape1Page()
'
' Sets up the print gridlines, page width, orientation, and margins
' for a landscape page with only 1 page
'

'
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.236220472440945)
        .RightMargin = Application.InchesToPoints(0.236220472440945)
        .TopMargin = Application.InchesToPoints(0.354330708661417)
        .BottomMargin = Application.InchesToPoints(0.354330708661417)
        .HeaderMargin = Application.InchesToPoints(0.31496062992126)
        .FooterMargin = Application.InchesToPoints(0.31496062992126)
        .PrintHeadings = False
        .PrintGridlines = True
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
End Function

Function PageSetupPortrait()
'
' Sets up the print gridlines, page width, orientation, and margins
' for a portrait page
'

'
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.236220472440945)
        .RightMargin = Application.InchesToPoints(0.236220472440945)
        .TopMargin = Application.InchesToPoints(0.354330708661417)
        .BottomMargin = Application.InchesToPoints(0.354330708661417)
        .HeaderMargin = Application.InchesToPoints(0.31496062992126)
        .FooterMargin = Application.InchesToPoints(0.31496062992126)
        .PrintHeadings = False
        .PrintGridlines = True
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
End Function

Function PageSetupPortrait1Page()
'
' Sets up the print gridlines, page width, orientation, and margins
' for a portrait page with only 1 page
'

'
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.236220472440945)
        .RightMargin = Application.InchesToPoints(0.236220472440945)
        .TopMargin = Application.InchesToPoints(0.354330708661417)
        .BottomMargin = Application.InchesToPoints(0.354330708661417)
        .HeaderMargin = Application.InchesToPoints(0.31496062992126)
        .FooterMargin = Application.InchesToPoints(0.31496062992126)
        .PrintHeadings = False
        .PrintGridlines = True
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
End Function

Function AutoFitFirstColumnForSATTUS()
    Range("A7", Range("A8").End(xlDown)).Columns.AutoFit
    Range("A1").Select
End Function

Function AutoFitColumn(r As Range)
    r.Columns.AutoFit
End Function

Function DeleteZeroRowsForMotorcycleInputSheet()
    Dim i, rowCount As Integer
    rowCount = Worksheets(1).UsedRange.Rows.Count
    i = 5
    
    For j = 5 To rowCount
        If (Worksheets(1).Cells(i, 8).Value = 0) Then
            Worksheets(1).Rows(i).Delete
        Else
            i = i + 1
        End If
    Next
    
End Function

Function findDistributor(ByVal make As String) As String
    Dim dict As New Scripting.Dictionary
    
    dict.Add "APRILIA", "Triumph"
    dict.Add "GAS GAS", "Triumph"
    dict.Add "KEEWAY", "Triumph"
    dict.Add "MOTO GUZZI", "Triumph"
    dict.Add "PGO", "Triumph"
    dict.Add "PIAGGIO", "Triumph"
    dict.Add "SYM", "Triumph"
    dict.Add "TRIUMPH", "Triumph"
    dict.Add "VESPA", "Triumph"
    
    dict.Add "ALFA ROMEO", "Ateco"
    dict.Add "CHERY", "Ateco"
    dict.Add "CHRYSLER", "Ateco"
    dict.Add "DODGE", "Ateco"
    dict.Add "FIAT", "Ateco"
    dict.Add "JEEP", "Ateco"
    dict.Add "RAM", "Ateco"
    dict.Add "MASERATI", "Ateco"
    
    dict.Add "BMW", "BMW"
    dict.Add "MINI", "BMW"
    
    dict.Add "AUDI", "EMD"
    dict.Add "PORSCHE", "EMD"
    dict.Add "SKODA", "EMD"
    dict.Add "VOLKSWAGEN", "EMD"
    
    dict.Add "LDV", "Great Lake"
    dict.Add "SSANGYONG", "Great Lake"
    
    dict.Add "HUSQVARNA", "KTM"
    dict.Add "KTM", "KTM"
    
    dict.Add "JAGUAR", "Motorcorp"
    dict.Add "LAND ROVER", "Motorcorp"
    
    dict.Add "INDIAN", "Polaris"
    dict.Add "POLARIS", "Polaris"
    dict.Add "VICTORY", "Polaris"
    
    dict.Add "PEUGEOT", "Sime Darby"
    dict.Add "CITROEN", "Sime Darby"
    
    dict.Add "TOYOTA", "Toyota"
    dict.Add "LEXUS", "Toyota"
    
    dict.Add "HONDA", "Honda"
    dict.Add "CAN-AM", "BRP"
    dict.Add "DUCATI", "Ducati"
    dict.Add "FORD", "Ford"
    dict.Add "HARLEY DAVIDSON", "Harley Davidson"
    dict.Add "HOLDEN", "Holden"
    dict.Add "HYUNDAI", "Hyundai"
    dict.Add "ISUZU", "Isuzu Utes"
    dict.Add "KIA", "Kia"
    dict.Add "KAWASAKI", "Kawasaki"
    dict.Add "MAZDA", "Mazda"
    dict.Add "MERCEDES-BENZ", "Mercedes-Benz"
    dict.Add "MITSUBISHI", "Mitsubishi"
    dict.Add "NISSAN", "Nissan"
    dict.Add "SUBARU", "Subaru"
    dict.Add "SUZUKI", "Suzuki"
    dict.Add "YAMAHA", "Yamaha"
    
    If dict.Exists(make) Then
        findDistributor = dict(make)
    Else
        findDistributor = "Other"
    End If
    Set dict = Nothing
End Function

'Make sure to have the data you want selected before calling
Function SortDataDescending(cell As String)
    Selection.Sort key1:=Range(cell, Range(cell).End(xlDown)), _
    order1:=xlDescending, header:=xlNo
End Function

'Make sure to have the data you want selected before calling
Function SortDataAscending(cell As String)
    Selection.Sort key1:=Range(cell, Range(cell).End(xlDown)), _
    order1:=xlAscending, header:=xlNo
End Function
