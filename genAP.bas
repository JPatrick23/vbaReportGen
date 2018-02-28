Attribute VB_Name = "genAP"
Sub ngmAPreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("AP Temp").Select
        Sheets("AP Temp").Copy
        Sheets("AP Temp").Select
        Sheets("AP Temp").Name = "NGM Action Plan Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "NGM Action Plan Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="apsDS", RefersToR1C1:= _
        "=apsDS.xlsx!aps[#All]"
         ActiveWorkbook.Names("apsDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\NGMAP.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\apsDS.xlsx"
    
'Activete workbook
    Windows("apsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("aps").Range.AutoFilter Field:=16, Criteria1:= _
        "NGM"
    ActiveSheet.ListObjects("aps").Range.AutoFilter Field:=12, Criteria1:= _
        "Draft"
    
    Range("aps[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("NGMAP.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste

'Activete workbook
    Windows("apsDS.xlsx").Activate
        
'NCE Number B
Range("aps[[#Headers],[ap_NCE]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("NGMAP.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste

'Activete workbook
    Windows("apsDS.xlsx").Activate
'Title Number C
Range("aps[[#Headers],[ap_APT]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("NGMAP.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
        
'Activete workbook
    Windows("apsDS.xlsx").Activate
'CurrenStep Number D
Range("aps[[#Headers],[ap_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("NGMAP.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste

'Activete workbook
    Windows("apsDS.xlsx").Activate
'Owner Number D
Range("aps[[#Headers],[ap_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("NGMAP.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
        
        
'Activete workbook
    Windows("apsDS.xlsx").Activate
'Due Date F
Range("aps[[#Headers],[ap_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("NGMAP.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
        

'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],apsDS,7,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],apsDS,2,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],apsDS,14,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],apsDS,15,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],apsDS,4,0)"
    'Range("F3").Select
    
    
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


   

    
' Save Report

ActiveWorkbook.Save
Windows("NGMAP.xlsx").Activate

'Close data source
Workbooks("apsDS.XLSX").Close SaveChanges:=False

End Sub


Sub gmAPreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("AP Temp").Select
        Sheets("AP Temp").Copy
        Sheets("AP Temp").Select
        Sheets("AP Temp").Name = "GM Action Plan Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "GM Action Plan Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="apsDS", RefersToR1C1:= _
        "=apsDS.xlsx!aps[#All]"
         ActiveWorkbook.Names("apsDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\GMAP.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\apsDS.xlsx"
    
'Activete workbook
    Windows("apsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("aps").Range.AutoFilter Field:=16, Criteria1:= _
        "GM"
    ActiveSheet.ListObjects("aps").Range.AutoFilter Field:=12, Criteria1:= _
        "Draft"
    
    Range("aps[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("GMAP.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste

'Activete workbook
    Windows("apsDS.xlsx").Activate
        
'NCE Number B
Range("aps[[#Headers],[ap_NCE]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("GMAP.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste

'Activete workbook
    Windows("apsDS.xlsx").Activate
'Title Number C
Range("aps[[#Headers],[ap_APT]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("GMAP.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
        
'Activete workbook
    Windows("apsDS.xlsx").Activate
'CurrenStep Number D
Range("aps[[#Headers],[ap_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("GMAP.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste

'Activete workbook
    Windows("apsDS.xlsx").Activate
'Owner Number D
Range("aps[[#Headers],[ap_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("GMAP.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
        
        
'Activete workbook
    Windows("apsDS.xlsx").Activate
'Due Date F
Range("aps[[#Headers],[ap_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("GMAP.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],apsDS,7,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],apsDS,2,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],apsDS,14,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],apsDS,15,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],apsDS,4,0)"
    'Range("F3").Select
    
    
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


   

    
' Save Report

ActiveWorkbook.Save
Windows("GMAP.xlsx").Activate

'Close data source
Workbooks("apsDS.XLSX").Close SaveChanges:=False

End Sub


Sub vvAPreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("AP Temp").Select
        Sheets("AP Temp").Copy
        Sheets("AP Temp").Select
        Sheets("AP Temp").Name = "VV Action Plan Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "Viral Vector Action Plan Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="apsDS", RefersToR1C1:= _
        "=apsDS.xlsx!aps[#All]"
         ActiveWorkbook.Names("apsDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\VVAP.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\apsDS.xlsx"
    
'Activete workbook
    Windows("apsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("aps").Range.AutoFilter Field:=16, Criteria1:= _
        "VV"
    ActiveSheet.ListObjects("aps").Range.AutoFilter Field:=12, Criteria1:= _
        "Draft"
    
    Range("aps[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("VVAP.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
'Activete workbook
    Windows("apsDS.xlsx").Activate
        
'NCE Number B
Range("aps[[#Headers],[ap_NCE]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("VVAP.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste

'Activete workbook
    Windows("apsDS.xlsx").Activate
'Title Number C
Range("aps[[#Headers],[ap_APT]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("VVAP.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
        
'Activete workbook
    Windows("apsDS.xlsx").Activate
'CurrenStep Number D
Range("aps[[#Headers],[ap_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("VVAP.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste

'Activete workbook
    Windows("apsDS.xlsx").Activate
'Owner Number D
Range("aps[[#Headers],[ap_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("VVAP.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
        
        
'Activete workbook
    Windows("apsDS.xlsx").Activate
'Due Date F
Range("aps[[#Headers],[ap_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("VVAP.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],apsDS,7,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],apsDS,2,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],apsDS,14,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],apsDS,15,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],apsDS,4,0)"
    'Range("F3").Select
    
    
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


   

    
' Save Report

ActiveWorkbook.Save
Windows("VVAP.xlsx").Activate

'Close data source
Workbooks("apsDS.XLSX").Close SaveChanges:=False

End Sub


Sub cc3APreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("AP Temp").Select
        Sheets("AP Temp").Copy
        Sheets("AP Temp").Select
        Sheets("AP Temp").Name = "NGM Action Plan Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "CC3 Action Plan Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="apsDS", RefersToR1C1:= _
        "=apsDS.xlsx!aps[#All]"
         ActiveWorkbook.Names("apsDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\CC3AP.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\apsDS.xlsx"
    
'Activete workbook
    Windows("apsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("aps").Range.AutoFilter Field:=16, Criteria1:= _
        "CC3"
    ActiveSheet.ListObjects("aps").Range.AutoFilter Field:=12, Criteria1:= _
        "Draft"
    
    Range("aps[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("CC3AP.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
'Activete workbook
    Windows("apsDS.xlsx").Activate
        
'NCE Number B
Range("aps[[#Headers],[ap_NCE]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("CC3AP.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste

'Activete workbook
    Windows("apsDS.xlsx").Activate
'Title Number C
Range("aps[[#Headers],[ap_APT]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("CC3AP.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
        
'Activete workbook
    Windows("apsDS.xlsx").Activate
'CurrenStep Number D
Range("aps[[#Headers],[ap_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("CC3AP.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste

'Activete workbook
    Windows("apsDS.xlsx").Activate
'Owner Number D
Range("aps[[#Headers],[ap_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("CC3AP.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
        
        
'Activete workbook
    Windows("apsDS.xlsx").Activate
'Due Date F
Range("aps[[#Headers],[ap_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("CC3AP.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],apsDS,7,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],apsDS,2,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],apsDS,14,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],apsDS,15,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],apsDS,4,0)"
    'Range("F3").Select
    
    
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


   

    
' Save Report

ActiveWorkbook.Save
Windows("CC3AP.xlsx").Activate

'Close data source
Workbooks("apsDS.XLSX").Close SaveChanges:=False

End Sub

Sub toAPreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("AP Temp").Select
        Sheets("AP Temp").Copy
        Sheets("AP Temp").Select
        Sheets("AP Temp").Name = "TO Action Plan Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "TO Action Plan Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="apsDS", RefersToR1C1:= _
        "=apsDS.xlsx!aps[#All]"
         ActiveWorkbook.Names("apsDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\TOAP.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\apsDS.xlsx"
    
'Activete workbook
    Windows("apsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("aps").Range.AutoFilter Field:=16, Criteria1:= _
        "TO"
    ActiveSheet.ListObjects("aps").Range.AutoFilter Field:=12, Criteria1:= _
        "Draft"
    
    Range("aps[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("TOAP.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
'Activete workbook
    Windows("apsDS.xlsx").Activate
        
'NCE Number B
Range("aps[[#Headers],[ap_NCE]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("TOAP.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste

'Activete workbook
    Windows("apsDS.xlsx").Activate
'Title Number C
Range("aps[[#Headers],[ap_APT]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("TOAP.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
        
'Activete workbook
    Windows("apsDS.xlsx").Activate
'CurrenStep Number D
Range("aps[[#Headers],[ap_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("TOAP.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste

'Activete workbook
    Windows("apsDS.xlsx").Activate
'Owner Number D
Range("aps[[#Headers],[ap_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("TOAP.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
        
        
'Activete workbook
    Windows("apsDS.xlsx").Activate
'Due Date F
Range("aps[[#Headers],[ap_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("TOAP.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
   ' Range("B3").Select
   ' Application.CutCopyMode = False
   ' ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],apsDS,7,0)"
   ' Range("C3").Select
   ' ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],apsDS,2,0)"
   ' Range("D3").Select
   ' ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],apsDS,14,0)"
   ' Range("E3").Select
   ' ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],apsDS,15,0)"
   ' Range("F3").Select
   ' ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],apsDS,4,0)"
   ' Range("F3").Select
    
    
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


   

    
' Save Report

ActiveWorkbook.Save
Windows("TOAP.xlsx").Activate

'Close data source
Workbooks("apsDS.XLSX").Close SaveChanges:=False

End Sub

Sub qaAPreport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("AP Temp").Select
        Sheets("AP Temp").Copy
        Sheets("AP Temp").Select
        Sheets("AP Temp").Name = "QA Action Plan Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "QA Action Plan Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="apsDS", RefersToR1C1:= _
        "=apsDS.xlsx!aps[#All]"
         ActiveWorkbook.Names("apsDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\QAAP.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\apsDS.xlsx"
    
'Activete workbook
    Windows("apsDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("aps").Range.AutoFilter Field:=16, Criteria1:= _
        "QA"
    ActiveSheet.ListObjects("aps").Range.AutoFilter Field:=12, Criteria1:= _
        "Draft"
    
    Range("aps[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("QAAP.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
        
'Activete workbook
    Windows("apsDS.xlsx").Activate
        
'NCE Number B
Range("aps[[#Headers],[ap_NCE]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("QAAP.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste

'Activete workbook
    Windows("apsDS.xlsx").Activate
'Title Number C
Range("aps[[#Headers],[ap_APT]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("QAAP.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
        
'Activete workbook
    Windows("apsDS.xlsx").Activate
'CurrenStep Number D
Range("aps[[#Headers],[ap_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("QAAP.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste

'Activete workbook
    Windows("apsDS.xlsx").Activate
'Owner Number D
Range("aps[[#Headers],[ap_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("QAAP.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
        
        
'Activete workbook
    Windows("apsDS.xlsx").Activate
'Due Date F
Range("aps[[#Headers],[ap_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Windows("QAAP.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],apsDS,7,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],apsDS,2,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],apsDS,14,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],apsDS,15,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],apsDS,4,0)"
    'Range("F3").Select
    
    
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


   

    
' Save Report

ActiveWorkbook.Save
Windows("QAAP.xlsx").Activate

'Close data source
Workbooks("apsDS.XLSX").Close SaveChanges:=False

End Sub

