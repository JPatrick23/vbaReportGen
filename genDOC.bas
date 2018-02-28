Attribute VB_Name = "genDOC"


Sub ngmreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("Doc Temp").Select
        Sheets("Doc Temp").Copy
        Sheets("Doc Temp").Select
        Sheets("Doc Temp").Name = "NGM Document Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "Non-Gene Mediated Document Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="docDS", RefersToR1C1:= _
        "=docsDS.xlsx!docs[#All]"
         ActiveWorkbook.Names("docDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\NGMDOC.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\docsDS.xlsx"
    
'Activete workbook
    Windows("docsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("docs").Range.AutoFilter Field:=15, Criteria1:= _
        "NGM"
    
    Range("docs[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("NGMDOC.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
  
 'B3
Windows("docsDS.xlsx").Activate
 Range("docs[[#Headers],[doc_PID]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("NGMDOC.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste
'C3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("NGMDOC.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
        
'D3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("NGMDOC.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
'E3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_Step]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("NGMDOC.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
        
'F3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_DO]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("NGMDOC.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
        
        
        
        
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
   ' Range("B3").Select
   ' Application.CutCopyMode = False
   ' ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],docDS,3,0)"
   ' Range("C3").Select
   ' ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],docDS,4,0)"
   ' Range("D3").Select
   ' ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],docDS,14,0)"
   ' Range("E3").Select
   ' ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],docDS,10,0)"
   ' Range("F3").Select
   ' ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],docDS,18,0)"
   ' Range("F3").Select
   ' ActiveCell.FormulaR1C1 = "=TODAY() - VLOOKUP(RC[-5],docDS,9,0)"
   ' Columns("F:F").Select
   Range("F:F").Activate
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
   ' Range("B3").Select
   ' Selection.AutoFill Destination:=Range("B3:B50"), Type:=xlFillDefault ' changed value to
   ' Range("B3:B7").Select
   ' Range("C3").Select
   ' Selection.AutoFill Destination:=Range("C3:C50"), Type:=xlFillDefault
   ' Range("C3:C7").Select
   ' Range("D3").Select
   ' Selection.AutoFill Destination:=Range("D3:D50"), Type:=xlFillDefault
   ' Range("D3:D7").Select
   ' Range("E3").Select
   ' Selection.AutoFill Destination:=Range("E3:E50"), Type:=xlFillDefault
    'Range("E3:E7").Select
    'Range("F3").Select
   ' Selection.AutoFill Destination:=Range("F3:F50"), Type:=xlFillDefault
    
    
    
  Range("A2").Select
Set rngData = Range(Range("A2"), Range("A2").End(xlDown).End(xlToRight))

ActiveSheet.ListObjects.Add(xlSrcRange, rngData, , xlYes).Name = _
"Table2"


    Range("A3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 60"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 90"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13421823
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    
' Save Report

ActiveWorkbook.Save
Windows("NGMDOC.xlsx").Activate






End Sub


Sub gmreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("Doc Temp").Select
        Sheets("Doc Temp").Copy
        Sheets("Doc Temp").Select
        Sheets("Doc Temp").Name = "GM Document Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "Gene Mediated Document Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="docDS", RefersToR1C1:= _
        "=docsDS.xlsx!docs[#All]"
         ActiveWorkbook.Names("docDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\GMDOC.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\docsDS.xlsx"
    
'Activete workbook
    Windows("docsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("docs").Range.AutoFilter Field:=15, Criteria1:= _
        "GM"
    
    Range("docs[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("GMDOC.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste


 'B3
Windows("docsDS.xlsx").Activate
 Range("docs[[#Headers],[doc_PID]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("GMDOC.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste
'C3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("GMDOC.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
        
'D3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("GMDOC.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
'E3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_Step]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("GMDOC.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
        
'F3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_DO]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("GMDOC.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
        
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],docDS,3,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],docDS,4,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],docDS,14,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],docDS,10,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],docDS,18,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=TODAY() - VLOOKUP(RC[-5],docDS,9,0)"
    'Columns("F:F").Select
    Range("F2").Activate
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
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


    Range("A3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 60"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 90"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13421823
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    

    

' Save Report

ActiveWorkbook.Save





End Sub


Sub vvreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("Doc Temp").Select
        Sheets("Doc Temp").Copy
        Sheets("Doc Temp").Select
        Sheets("Doc Temp").Name = "VV Document Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "Viral Vector Document Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="docDS", RefersToR1C1:= _
        "=docsDS.xlsx!docs[#All]"
         ActiveWorkbook.Names("docDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\VVDOC.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\docsDS.xlsx"
    
'Activete workbook
    Windows("docsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("docs").Range.AutoFilter Field:=15, Criteria1:= _
        "VV"
    
    Range("docs[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("VVDOC.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
 'B3
 Windows("docsDS.xlsx").Activate
 Range("docs[[#Headers],[doc_PID]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("VVDOC.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste
'C3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("VVDOC.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
        
'D3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("VVDOC.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
'E3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_Step]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("VVDOC.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
        
'F3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_DO]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("VVDOC.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
        
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],docDS,3,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],docDS,4,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],docDS,14,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],docDS,10,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],docDS,18,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=TODAY() - VLOOKUP(RC[-5],docDS,9,0)"
    'Columns("F:F").Select
   Range("F:F").Activate
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
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


    Range("A3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 60"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 90"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13421823
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    


' Save Report

ActiveWorkbook.Save




End Sub


Sub cc3report()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("Doc Temp").Select
        Sheets("Doc Temp").Copy
        Sheets("Doc Temp").Select
        Sheets("Doc Temp").Name = "CC3 Document Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "CC3 Document Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="docDS", RefersToR1C1:= _
        "=docsDS.xlsx!docs[#All]"
         ActiveWorkbook.Names("docDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\CC3DOC.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\docsDS.xlsx"
    
'Activete workbook
    Windows("docsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("docs").Range.AutoFilter Field:=15, Criteria1:= _
        "CC3"
    
    Range("docs[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("CC3DOC.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste

 'B3
 Windows("docsDS.xlsx").Activate
 Range("docs[[#Headers],[doc_PID]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("CC3DOC.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste
'C3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("CC3DOC.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
        
'D3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("CC3DOC.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
'E3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_Step]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("CC3DOC.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
        
'F3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_DO]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("CC3DOC.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
        
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],docDS,3,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],docDS,4,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],docDS,14,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],docDS,10,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],docDS,18,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=TODAY() - VLOOKUP(RC[-5],docDS,9,0)"
    'Columns("F:F").Select
    Range("F:F").Activate
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
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


    Range("A3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 60"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 90"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13421823
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
  

' Save Report

ActiveWorkbook.Save


End Sub


Sub toreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("Doc Temp").Select
        Sheets("Doc Temp").Copy
        Sheets("Doc Temp").Select
        Sheets("Doc Temp").Name = "TO Document Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "Tech Ops Document Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="docDS", RefersToR1C1:= _
        "=docsDS.xlsx!docs[#All]"
         ActiveWorkbook.Names("docDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\TODOC.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\docsDS.xlsx"
    
'Activete workbook
    Windows("docsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("docs").Range.AutoFilter Field:=15, Criteria1:= _
        "TO"
    
    Range("docs[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("TODOC.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
 'B3
 Windows("docsDS.xlsx").Activate
 Range("docs[[#Headers],[doc_PID]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("TODOC.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste
'C3
Windows("docsDS.xlsx").Activate
Windows("docsDS.xlsx").Activate
    Range("docs[[#Headers],[doc_Title]]").Select
        Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
    
'Paste relevent data to blank'

Windows("TODOC.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
        
'D3
Windows("docsDS.xlsx").Activate

    Range("docs[[#Headers],[doc_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("TODOC.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
'E3
Windows("docsDS.xlsx").Activate

    Range("docs[[#Headers],[doc_Step]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("TODOC.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
        
'F3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_DO]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("TODOC.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
        
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
   ' Range("B3").Select
   ' Application.CutCopyMode = False
   ' ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],docDS,3,0)"
   ' Range("C3").Select
   ' ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],docDS,4,0)"
   ' Range("D3").Select
   ' ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],docDS,14,0)"
   ' Range("E3").Select
   ' ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],docDS,10,0)"
   ' Range("F3").Select
   ' ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],docDS,18,0)"
   ' Range("F3").Select
   ' ActiveCell.FormulaR1C1 = "=TODAY() - VLOOKUP(RC[-5],docDS,9,0)"
   ' Columns("F:F").Select
    Range("F:F").Activate
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
   ' Range("B3").Select
   ' Selection.AutoFill Destination:=Range("B3:B50"), Type:=xlFillDefault ' changed value to
   ' Range("B3:B7").Select
   ' Range("C3").Select
   ' Selection.AutoFill Destination:=Range("C3:C50"), Type:=xlFillDefault
   ' Range("C3:C7").Select
   ' Range("D3").Select
   ' Selection.AutoFill Destination:=Range("D3:D50"), Type:=xlFillDefault
   ' Range("D3:D7").Select
   ' Range("E3").Select
   ' Selection.AutoFill Destination:=Range("E3:E50"), Type:=xlFillDefault
   ' Range("E3:E7").Select
   ' Range("F3").Select
   ' Selection.AutoFill Destination:=Range("F3:F50"), Type:=xlFillDefault
    
    
    
  Range("A2").Select
Set rngData = Range(Range("A2"), Range("A2").End(xlDown).End(xlToRight))

ActiveSheet.ListObjects.Add(xlSrcRange, rngData, , xlYes).Name = _
"Table2"


    Range("A3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 60"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 90"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13421823
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
' Save Report

ActiveWorkbook.Save



End Sub

Sub qareport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("Doc Temp").Select
        Sheets("Doc Temp").Copy
        Sheets("Doc Temp").Select
        Sheets("Doc Temp").Name = "QA Document Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "QA Document Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="docDS", RefersToR1C1:= _
        "=docsDS.xlsx!docs[#All]"
         ActiveWorkbook.Names("docDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\QADOC.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\docsDS.xlsx"
    
'Activete workbook
    Windows("docsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("docs").Range.AutoFilter Field:=15, Criteria1:= _
        "QA"
    
    Range("docs[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("QADOC.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
 'B3
 Windows("docsDS.xlsx").Activate
 Range("docs[[#Headers],[doc_PID]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("QADOC.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste
'C3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("QADOC.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
        
'D3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("QADOC.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
'E3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_Step]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("QADOC.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
        
'F3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_DO]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("QADOC.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
        
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],docDS,3,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],docDS,4,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],docDS,14,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],docDS,10,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],docDS,18,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=TODAY() - VLOOKUP(RC[-5],docDS,9,0)"
    'Columns("F:F").Select
    Range("F:F").Activate
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
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


    Range("A3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 60"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$F4 > 90"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13421823
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    

' Save Report

ActiveWorkbook.Save





End Sub








