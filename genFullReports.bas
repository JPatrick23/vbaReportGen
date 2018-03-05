Attribute VB_Name = "genFullReports"

Sub mfgDOCreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("Doc Temp All").Select
        Sheets("Doc Temp All").Copy
        Sheets("Doc Temp All").Select
        Sheets("Doc Temp All").Name = "Manufacturing Document Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "Manufacturing Document Report"

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
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\MFGDOC.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\docsDS.xlsx"

'Activete workbook
    Windows("docsDS.xlsx").Activate

' Copy relevent data
ActiveSheet.ListObjects("docs").Range.AutoFilter Field:=15, Criteria1:= _
    Array("CC3", "GM", "NGM", "TO", "VV"), Operator:=xlFilterValues

    Range("docs[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

'Paste relevent data to blank'
Windows("MFGDOC.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste

 'B3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

'Paste relevent data to blank'
Windows("MFGDOC.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste

'C3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

'Paste relevent data to blank'
Windows("MFGDOC.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste


'D3
Windows("docsDS.xlsx").Activate
Range("docs[[#Headers],[doc_CDD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

'Paste relevent data to blank'
Windows("MFGDOC.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste

'Remove superfluopus information
Rows(3).Select
    Selection.Delete


  Range("A2").Select
Set rngData = Range(Range("A2"), Range("A2").End(xlDown).End(xlToRight))

ActiveSheet.ListObjects.Add(xlSrcRange, rngData, , xlYes).Name = _
"Table2"

Range("D:D").Select
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"
    
    Range("A3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$D4 > TODAY()"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$D4 = TODAY()"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$D4 < TODAY()"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13421823
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
'Format Dates'
    Range("F3").Select
      Range(Selection, Selection.End(xlDown)).Select
      Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"
' Save Report

ActiveWorkbook.Save
Windows("MFGDOC.xlsx").Activate

End Sub

