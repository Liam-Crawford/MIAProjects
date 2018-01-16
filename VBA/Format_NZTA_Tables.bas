Attribute VB_Name = "Format_NZTA_Tables"
Sub Format_NZTA_Tables()
' Opens each NZTA table and runs appropriate formatting on them
    Dim path As String, book As Integer
    book = 3
    path = "Z:\MIA 2018\Registration Committee\Motor Registration Tables - current month\"
    
    ' Format 001, 001N, 002, 002N
    Workbooks.Open (path & "001.xls")
    Call NZTA_Subs.NZTA_001_002
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "001N.xls")
    Call NZTA_Subs.NZTA_001_002
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "002.xls")
    Call NZTA_Subs.NZTA_001_002
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "002N.xls")
    Call NZTA_Subs.NZTA_001_002
    Workbooks(book).Close SaveChanges:=True
    
    ' Format 001A
    Workbooks.Open (path & "001A.xls")
    Call NZTA_Subs.NZTA_001A
    Workbooks(book).Close SaveChanges:=True
    
    ' Format 002A
    Workbooks.Open (path & "002A.xls")
    Call NZTA_Subs.NZTA_002A
    Workbooks(book).Close SaveChanges:=True
    
    ' Format 006, 006N, 006X
    Workbooks.Open (path & "006.xls")
    Call NZTA_Subs.NZTA_006
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "006N.xls")
    Call NZTA_Subs.NZTA_006
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "006X.xls")
    Call NZTA_Subs.NZTA_006
    Workbooks(book).Close SaveChanges:=True
    
    ' Format 008, 008N, 008X
    Workbooks.Open (path & "008.xls")
    Call NZTA_Subs.NZTA_008
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "008N.xls")
    Call NZTA_Subs.NZTA_008
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "008X.xls")
    Call NZTA_Subs.NZTA_008
    Workbooks(book).Close SaveChanges:=True
    
    ' Format 051
    Workbooks.Open (path & "051.xls")
    Call NZTA_Subs.NZTA_051
    Workbooks(book).Close SaveChanges:=True
    
    ' Format 054
    Workbooks.Open (path & "054.xls")
    Call NZTA_Subs.NZTA_054
    Workbooks(book).Close SaveChanges:=True
    
    ' Format 064N, 064X, 065N, 065X
    Workbooks.Open (path & "064N.xls")
    Call NZTA_Subs.NZTA_064
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "064X.xls")
    Call NZTA_Subs.NZTA_064
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "065N.xls")
    Call NZTA_Subs.NZTA_064
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "065X.xls")
    Call NZTA_Subs.NZTA_064
    Workbooks(book).Close SaveChanges:=True
    
    ' Format MIA_DEREG_MONTHLY
    Workbooks.Open (path & "MIA_DEREG_MONTHLY.xls")
    Call NZTA_Subs.NZTA_MIA_DEREG_MONTHLY
    Workbooks(book).Close SaveChanges:=True
    
    ' Format N7-USG
    Workbooks.Open (path & "N7-USG.xls")
    Call NZTA_Subs.NZTA_N7USG
    Workbooks(book).Close SaveChanges:=True
    
    ' Format U7MM_AGE, U8MM_AGE
    Workbooks.Open (path & "U7MM_AGE.xls")
    Call NZTA_Subs.NZTA_UMM_AGE
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "U8MM_AGE_Report.xls")
    Call NZTA_Subs.NZTA_UMM_AGE
    Workbooks(book).Close SaveChanges:=True
    
    ' Format VTyp10-13, YTD_RENTALS_NEW
    Workbooks.Open (path & "VTyp10-13.xls")
    Call NZTA_Subs.NZTA_VTyp1013
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "YTD_RENTALS_NEW.xls")
    Call NZTA_Subs.NZTA_VTyp1013
    Workbooks(book).Close SaveChanges:=True
    
    ' Format X-085N
    Workbooks.Open (path & "X-085N.xls")
    Call NZTA_Subs.NZTA_X085N
    Workbooks(book).Close SaveChanges:=True
    
    ' Format Y_MPC_A
    Workbooks.Open (path & "Y_MPC_A.xls")
    Call NZTA_Subs.NZTA_Y_MPC_A
    Workbooks(book).Close SaveChanges:=True
    
    ' Format Y-001AN, Y-001AN_2AN, Y-001AX, Y-002AN, Y-002AX
    Workbooks.Open (path & "Y-001AN.xls")
    Call NZTA_Subs.NZTA_Y00
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "Y-002AN.xls")
    Call NZTA_Subs.NZTA_Y00
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "Y-001AN_2AN.xls")
    Call NZTA_Subs.NZTA_Y00
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "Y-001AX.xls")
    Call NZTA_Subs.NZTA_Y00
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "Y-002AX.xls")
    Call NZTA_Subs.NZTA_Y00
    Workbooks(book).Close SaveChanges:=True
    
    ' Format Y-065N
    Workbooks.Open (path & "Y-065N.xls")
    Call NZTA_Subs.NZTA_Y065N
    Workbooks(book).Close SaveChanges:=True
    
    ' Format Y-080N
    Workbooks.Open (path & "Y-080N.xls")
    Call NZTA_Subs.NZTA_Y080N
    Workbooks(book).Close SaveChanges:=True
    
    ' Format Y-081N
    Workbooks.Open (path & "Y-081N.xls")
    Call NZTA_Subs.NZTA_Y081N
    Workbooks(book).Close SaveChanges:=True
    
    ' Format Y-084N, Y-085N, YTD_USED_CARS, YTD_USED_COM
    Workbooks.Open (path & "Y-084N.xls")
    Call NZTA_Subs.NZTA_Y084N
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "Y-085N.xls")
    Call NZTA_Subs.NZTA_Y084N
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "YTD_USED_CARS.xls")
    Call NZTA_Subs.NZTA_Y084N
    Workbooks(book).Close SaveChanges:=True
    Workbooks.Open (path & "YTD_USED_COM.xls")
    Call NZTA_Subs.NZTA_Y084N
    Workbooks(book).Close SaveChanges:=True
    
    ' Format Y-MPC50
    Workbooks.Open (path & "Y-MPC50.xls")
    Call NZTA_Subs.NZTA_YMPC50
    Workbooks(book).Close SaveChanges:=True
    
    ' Format Y-MPC51
    Workbooks.Open (path & "Y-MPC51.xls")
    Call NZTA_Subs.NZTA_YMPC51
    Workbooks(book).Close SaveChanges:=True
    
    ' Format YRY-COMMS_M1
    Workbooks.Open (path & "YRY-COMMS_M1.xls")
    Call NZTA_Subs.NZTA_YRYCOMMSM1
    Workbooks(book).Close SaveChanges:=True
    
End Sub
