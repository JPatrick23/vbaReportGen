Attribute VB_Name = "genCCS"
Sub ngmCCSreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("CC Temp").Select
        Sheets("CC Temp").Copy
        Sheets("CC Temp").Select
        Sheets("CC Temp").Name = "NGM Change Control Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "Non-Gene Mediated Change Control Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="ccsDS", RefersToR1C1:= _
        "=ccsDS.xlsx!ccs[#All]"
         ActiveWorkbook.Names("ccsDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\NGMCCS.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\ccsDS.xlsx"
    
'Activete workbook
    Windows("ccsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("ccs").Range.AutoFilter Field:=10, Criteria1:= _
        "NGM"
    
    Range("ccs[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("NGMCCS.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste

'B3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'B3 Paste
Windows("NGMCCS.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste
'C3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'C3 Paste
Windows("NGMCCS.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
'D3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'D3 Paste
Windows("NGMCCS.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
'E3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_SD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'E3 Paste
Windows("NGMCCS.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
'F3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'F3 Paste
Windows("NGMCCS.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste

        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],ccsDS,3,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],ccsDS,6,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],ccsDS,7,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],ccsDS,8,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],ccsDS,5,0)"
   
    'Range("B3").Select
    'Selection.AutoFill Destination:=Range("B3:B50"), Type:=xlFillDefault ' changed value to
    'Range("B3:B7").Select
    'Range("C3").Select
    'Selection.AutoFill Destination:=Range("C3:C50"), Type:=xlFillDefault
    'Range("C3:C7").Select
    'Range("D3").Select
    'Selection.AutoFill Destination:=Range("D3:D50"), Type:=xlFillDefault
    'Range("D3:D7").Select
    'Range("E3").Select
    'Selection.AutoFill Destination:=Range("E3:E50"), Type:=xlFillDefault
    'Range("E3:E7").Select
    'Range("F3").Select
    'Selection.AutoFill Destination:=Range("F3:F50"), Type:=xlFillDefault
   
    
    
  Range("A2").Select
Set rngData = Range(Range("A2"), Range("A2").End(xlDown).End(xlToRight))

ActiveSheet.ListObjects.Add(xlSrcRange, rngData, , xlYes).Name = _
"Table2"

Columns("E:E").Select
   
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"

 Columns("F:F").Select
   
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"

    
' Save Report

ActiveWorkbook.Save
Windows("NGMCCS.xlsx").Activate

'Close data source
Workbooks("ccsDS.XLSX").Close SaveChanges:=False

End Sub

Sub gmCCSreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("CC Temp").Select
        Sheets("CC Temp").Copy
        Sheets("CC Temp").Select
        Sheets("CC Temp").Name = "GM Change Control Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "Gene Mediated Change Control Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="ccsDS", RefersToR1C1:= _
        "=ccsDS.xlsx!ccs[#All]"
         ActiveWorkbook.Names("ccsDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\GMCCS.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\ccsDS.xlsx"
    
'Activete workbook
    Windows("ccsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("ccs").Range.AutoFilter Field:=10, Criteria1:= _
        "GM"
    
    Range("ccs[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("GMCCS.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
'B3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'B3 Paste
Windows("GMCCS.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste
'C3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'C3 Paste
Windows("GMCCS.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
'D3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'D3 Paste
Windows("GMCCS.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
'E3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_SD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'E3 Paste
Windows("GMCCS.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
'F3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'F3 Paste
Windows("GMCCS.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste

        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],ccsDS,3,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],ccsDS,6,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],ccsDS,7,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],ccsDS,8,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],ccsDS,5,0)"
   
    'Range("B3").Select
    'Selection.AutoFill Destination:=Range("B3:B50"), Type:=xlFillDefault ' changed value to
    'Range("B3:B7").Select
    'Range("C3").Select
    'Selection.AutoFill Destination:=Range("C3:C50"), Type:=xlFillDefault
    'Range("C3:C7").Select
    'Range("D3").Select
    'Selection.AutoFill Destination:=Range("D3:D50"), Type:=xlFillDefault
    'Range("D3:D7").Select
    'Range("E3").Select
    'Selection.AutoFill Destination:=Range("E3:E50"), Type:=xlFillDefault
    'Range("E3:E7").Select
    'Range("F3").Select
    'Selection.AutoFill Destination:=Range("F3:F50"), Type:=xlFillDefault
    
    
    
  Range("A2").Select
Set rngData = Range(Range("A2"), Range("A2").End(xlDown).End(xlToRight))

ActiveSheet.ListObjects.Add(xlSrcRange, rngData, , xlYes).Name = _
"Table2"

Columns("E:E").Select
  
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"

 Columns("F:F").Select
  
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"
    
' Save Report

ActiveWorkbook.Save
Windows("GMCCS.xlsx").Activate

'Close data source
Workbooks("ccsDS.XLSX").Close SaveChanges:=False

End Sub



Sub vvCCSreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("CC Temp").Select
        Sheets("CC Temp").Copy
        Sheets("CC Temp").Select
        Sheets("CC Temp").Name = "VV Change Control Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "Viral Vector Change Control Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="ccsDS", RefersToR1C1:= _
        "=ccsDS.xlsx!ccs[#All]"
         ActiveWorkbook.Names("ccsDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\VVCCS.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\ccsDS.xlsx"
    
'Activete workbook
    Windows("ccsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("ccs").Range.AutoFilter Field:=10, Criteria1:= _
        "VV"
    
    Range("ccs[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("VVCCS.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
'B3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'B3 Paste
Windows("VVCCS.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste
'C3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'C3 Paste
Windows("VVCCS.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
'D3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'D3 Paste
Windows("VVCCS.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
'E3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_SD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'E3 Paste
Windows("VVCCS.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
'F3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'F3 Paste
Windows("VVCCS.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste

        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],ccsDS,3,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],ccsDS,6,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],ccsDS,7,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],ccsDS,8,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],ccsDS,5,0)"
   
    'Range("B3").Select
    'Selection.AutoFill Destination:=Range("B3:B50"), Type:=xlFillDefault ' changed value to
    'Range("B3:B7").Select
    'Range("C3").Select
    'Selection.AutoFill Destination:=Range("C3:C50"), Type:=xlFillDefault
    'Range("C3:C7").Select
    'Range("D3").Select
    'Selection.AutoFill Destination:=Range("D3:D50"), Type:=xlFillDefault
    'Range("D3:D7").Select
    'Range("E3").Select
    'Selection.AutoFill Destination:=Range("E3:E50"), Type:=xlFillDefault
    'Range("E3:E7").Select
    'Range("F3").Select
    'Selection.AutoFill Destination:=Range("F3:F50"), Type:=xlFillDefault
   

    
    
    
  Range("A2").Select
Set rngData = Range(Range("A2"), Range("A2").End(xlDown).End(xlToRight))

ActiveSheet.ListObjects.Add(xlSrcRange, rngData, , xlYes).Name = _
"Table2"

Columns("E:E").Select
   
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"

 Columns("F:F").Select
   
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"
    
' Save Report

ActiveWorkbook.Save
Windows("VVCCS.xlsx").Activate

'Close data source
Workbooks("ccsDS.XLSX").Close SaveChanges:=False

End Sub


Sub cc3CCSreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("CC Temp").Select
        Sheets("CC Temp").Copy
        Sheets("CC Temp").Select
        Sheets("CC Temp").Name = "CC3 Change Control Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "CC3 Change Control Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="ccsDS", RefersToR1C1:= _
        "=ccsDS.xlsx!ccs[#All]"
         ActiveWorkbook.Names("ccsDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\CC3CCS.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\ccsDS.xlsx"
    
'Activete workbook
    Windows("ccsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("ccs").Range.AutoFilter Field:=10, Criteria1:= _
        "CC3"
    
    Range("ccs[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("CC3CCS.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
'B3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'B3 Paste
Windows("CC3CCS.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste
'C3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'C3 Paste
Windows("CC3CCS.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
'D3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'D3 Paste
Windows("CC3CCS.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
'E3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_SD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'E3 Paste
Windows("CC3CCS.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
'F3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'F3 Paste
Windows("CC3CCS.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste

        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],ccsDS,3,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],ccsDS,6,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],ccsDS,7,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],ccsDS,8,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],ccsDS,5,0)"
   
    'Range("B3").Select
    'Selection.AutoFill Destination:=Range("B3:B50"), Type:=xlFillDefault ' changed value to
    'Range("B3:B7").Select
    'Range("C3").Select
    'Selection.AutoFill Destination:=Range("C3:C50"), Type:=xlFillDefault
    'Range("C3:C7").Select
    'Range("D3").Select
    'Selection.AutoFill Destination:=Range("D3:D50"), Type:=xlFillDefault
    'Range("D3:D7").Select
    'Range("E3").Select
    'Selection.AutoFill Destination:=Range("E3:E50"), Type:=xlFillDefault
    'Range("E3:E7").Select
    'Range("F3").Select
    'Selection.AutoFill Destination:=Range("F3:F50"), Type:=xlFillDefault
   
    
  Range("A2").Select
Set rngData = Range(Range("A2"), Range("A2").End(xlDown).End(xlToRight))

ActiveSheet.ListObjects.Add(xlSrcRange, rngData, , xlYes).Name = _
"Table2"

Columns("E:E").Select

    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"

 Columns("F:F").Select

    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"
    
' Save Report

ActiveWorkbook.Save
Windows("CC3CCS.xlsx").Activate

'Close data source
Workbooks("ccsDS.XLSX").Close SaveChanges:=False

End Sub


Sub toCCSreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("CC Temp").Select
        Sheets("CC Temp").Copy
        Sheets("CC Temp").Select
        Sheets("CC Temp").Name = "TO Change Control Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "Tech Ops Change Control Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="ccsDS", RefersToR1C1:= _
        "=ccsDS.xlsx!ccs[#All]"
         ActiveWorkbook.Names("ccsDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\TOCCS.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\ccsDS.xlsx"
    
'Activete workbook
    Windows("ccsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("ccs").Range.AutoFilter Field:=10, Criteria1:= _
        "TO"
    
    Range("ccs[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("TOCCS.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
'B3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'B3 Paste
Windows("TOCCS.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste
'C3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'C3 Paste
Windows("TOCCS.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
'D3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'D3 Paste
Windows("TOCCS.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
'E3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_SD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'E3 Paste
Windows("TOCCS.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
'F3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'F3 Paste
Windows("TOCCS.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste

        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],ccsDS,3,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],ccsDS,6,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],ccsDS,7,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],ccsDS,8,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],ccsDS,5,0)"
   
    'Range("B3").Select
    'Selection.AutoFill Destination:=Range("B3:B50"), Type:=xlFillDefault ' changed value to
    'Range("B3:B7").Select
    'Range("C3").Select
    'Selection.AutoFill Destination:=Range("C3:C50"), Type:=xlFillDefault
    'Range("C3:C7").Select
    'Range("D3").Select
    'Selection.AutoFill Destination:=Range("D3:D50"), Type:=xlFillDefault
    'Range("D3:D7").Select
    'Range("E3").Select
    'Selection.AutoFill Destination:=Range("E3:E50"), Type:=xlFillDefault
    'Range("E3:E7").Select
    'Range("F3").Select
    'Selection.AutoFill Destination:=Range("F3:F50"), Type:=xlFillDefault
   
    

Columns("E:E").Select
    
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"

 Columns("F:F").Select
   
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"

    
' Save Report

ActiveWorkbook.Save
Windows("TOCCS.xlsx").Activate

'Close data source
Workbooks("ccsDS.XLSX").Close SaveChanges:=False

End Sub



Sub qaCCSreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("CC Temp").Select
        Sheets("CC Temp").Copy
        Sheets("CC Temp").Select
        Sheets("CC Temp").Name = "QA Change Control Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "QA Change Control Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="ccsDS", RefersToR1C1:= _
        "=ccsDS.xlsx!ccs[#All]"
         ActiveWorkbook.Names("ccsDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\QACCS.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\ccsDS.xlsx"
    
'Activete workbook
    Windows("ccsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("ccs").Range.AutoFilter Field:=10, Criteria1:= _
        "QA"
    
    Range("ccs[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("QACCS.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        


'B3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'B3 Paste
Windows("QACCS.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste
'C3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'C3 Paste
Windows("QACCS.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
'D3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'D3 Paste
Windows("QACCS.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
'E3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_SD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'E3 Paste
Windows("QACCS.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
'F3
Windows("ccsDS.xlsx").Activate
Range("ccs[[#Headers],[cc_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'F3 Paste
Windows("QACCS.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste

        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],ccsDS,3,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],ccsDS,6,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],ccsDS,7,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],ccsDS,8,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],ccsDS,5,0)"
   
    'Range("B3").Select
    'Selection.AutoFill Destination:=Range("B3:B50"), Type:=xlFillDefault ' changed value to
    'Range("B3:B7").Select
    'Range("C3").Select
    'Selection.AutoFill Destination:=Range("C3:C50"), Type:=xlFillDefault
    'Range("C3:C7").Select
    'Range("D3").Select
    'Selection.AutoFill Destination:=Range("D3:D50"), Type:=xlFillDefault
    'Range("D3:D7").Select
    'Range("E3").Select
    'Selection.AutoFill Destination:=Range("E3:E50"), Type:=xlFillDefault
    'Range("E3:E7").Select
    'Range("F3").Select
    'Selection.AutoFill Destination:=Range("F3:F50"), Type:=xlFillDefault
   
    
    
  Range("A2").Select
Set rngData = Range(Range("A2"), Range("A2").End(xlDown).End(xlToRight))

ActiveSheet.ListObjects.Add(xlSrcRange, rngData, , xlYes).Name = _
"Table2"

Columns("E:E").Select
    
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"

 Columns("F:F").Select
    
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"

    
' Save Report

ActiveWorkbook.Save
Windows("QACCS.xlsx").Activate

'Close data source
Workbooks("ccsDS.XLSX").Close SaveChanges:=False

End Sub



