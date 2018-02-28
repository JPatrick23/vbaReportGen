Attribute VB_Name = "genCAPA"


Sub ngmCAPAreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("CAPA Temp").Select
        Sheets("CAPA Temp").Copy
        Sheets("CAPA Temp").Select
        Sheets("CAPA Temp").Name = "NGM CAPA Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "Non-Gene Mediated Document Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="capasDS", RefersToR1C1:= _
        "=capasDS.xlsx!capas[#All]"
         ActiveWorkbook.Names("capasDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\NGMCAPA.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\capasDS.xlsx"
    
'Activete workbook
    Windows("capasDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("capas").Range.AutoFilter Field:=11, Criteria1:= _
        "NGM"
    
    Range("capas[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("NGMCAPA.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    Range("B3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],capasDS,3,0)"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],capasDS,2,0)"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],capasDS,4,0)"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],capasDS,7,0)"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],capas,8,0)"
    
    Range("B3").Select
    Selection.AutoFill Destination:=Range("B3:B50"), Type:=xlFillDefault ' changed value to
    Range("B3:B7").Select
    Range("C3").Select
    Selection.AutoFill Destination:=Range("C3:C50"), Type:=xlFillDefault
    Range("C3:C7").Select
    Range("D3").Select
    Selection.AutoFill Destination:=Range("D3:D50"), Type:=xlFillDefault
    Range("D3:D7").Select
    Range("E3").Select
    Selection.AutoFill Destination:=Range("E3:E50"), Type:=xlFillDefault
    Range("E3:E7").Select
    Range("F3").Select
    Selection.AutoFill Destination:=Range("F3:F50"), Type:=xlFillDefault
    
    
    
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
Windows("NGMCAPA.xlsx").Activate

'Close data source
Workbooks("capasDS.XLSX").Close SaveChanges:=False

End Sub



Sub toCAPAreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("CAPA Temp").Select
        Sheets("CAPA Temp").Copy
        Sheets("CAPA Temp").Select
        Sheets("CAPA Temp").Name = "TO CAPA Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "Tech Ops Document Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="capasDS", RefersToR1C1:= _
        "=capasDS.xlsx!capas[#All]"
         ActiveWorkbook.Names("capasDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\TOCAPA.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\capasDS.xlsx"
    
'Activete workbook
    Windows("capasDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("capas").Range.AutoFilter Field:=11, Criteria1:= _
        "TO"
    
    Range("capas[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("TOCAPA.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    Range("B3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],capasDS,3,0)"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],capasDS,2,0)"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],capasDS,4,0)"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],capasDS,7,0)"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],capas,8,0)"
    
    Range("B3").Select
    Selection.AutoFill Destination:=Range("B3:B50"), Type:=xlFillDefault ' changed value to
    Range("B3:B7").Select
    Range("C3").Select
    Selection.AutoFill Destination:=Range("C3:C50"), Type:=xlFillDefault
    Range("C3:C7").Select
    Range("D3").Select
    Selection.AutoFill Destination:=Range("D3:D50"), Type:=xlFillDefault
    Range("D3:D7").Select
    Range("E3").Select
    Selection.AutoFill Destination:=Range("E3:E50"), Type:=xlFillDefault
    Range("E3:E7").Select
    Range("F3").Select
    Selection.AutoFill Destination:=Range("F3:F50"), Type:=xlFillDefault
    
    
    
  Range("A2").Select
Set rngData = Range(Range("A2"), Range("A2").End(xlDown).End(xlToRight))

ActiveSheet.ListObjects.Add(xlSrcRange, rngData, , xlYes).Name = _
"Table2"


    
    
' Save Report

ActiveWorkbook.Save
Windows("TOCAPA.xlsx").Activate

'Close data source
Workbooks("capasDS.XLSX").Close SaveChanges:=False

End Sub


