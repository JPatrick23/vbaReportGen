Attribute VB_Name = "genIssue"
Sub ngmIssuereport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("Issue Temp").Select
        Sheets("Issue Temp").Copy
        Sheets("Issue Temp").Select
        Sheets("Issue Temp").Name = "NGM Issue Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "NGM Issue Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="issueDS", RefersToR1C1:= _
        "=issueDS.xlsx!issues[#All]"
         ActiveWorkbook.Names("issueDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\NGMISSUE.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\issueDS.xlsx"
    
'Activete workbook
    Windows("issueDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("issues").Range.AutoFilter Field:=15, Criteria1:= _
        "NGM"
   
    
    Range("issues[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("NGMISSUE.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete
    
'B3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Source]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'B3 Paste
Windows("NGMISSUE.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste

'C3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'C3 Paste
Windows("NGMISSUE.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste
        
'D3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'D3 Paste
Windows("NGMISSUE.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
        
'E3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'E3 Paste
Windows("NGMISSUE.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
        
'F3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'F3 Paste
Windows("NGMISSUE.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],issueDS,6,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],issueDS,4,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],issueDS,2,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],issueDS,10,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],issueDS,12,0)"
    'Range("F3").Select
    
    
    'Range("B3").Select
    'Selection.AutoFill Destination:=Range("B3:B50"), Type:=xlFillDefault ' changed value to
    'Range("B3:B7").Select
    'Range("C3").Select
    'Selection.AutoFill Destination:=Range("C3:C50"), Type:=xlFillDefault
    'Range("C3:C7").Select
   ' Range("D3").Select
   ' Selection.AutoFill Destination:=Range("D3:D50"), Type:=xlFillDefault
   ' Range("D3:D7").Select
   ' Range("E3").Select
   ' Selection.AutoFill Destination:=Range("E3:E50"), Type:=xlFillDefault
   ' Range("E3:E7").Select
   ' Range("F3").Select
    'Selection.AutoFill Destination:=Range("F3:F50"), Type:=xlFillDefault
    
    
    



 Columns("F:F").Select

    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"


   

    
' Save Report

ActiveWorkbook.Save
Windows("NGMISSUE.xlsx").Activate

'Close data source
Workbooks("issueDS.XLSX").Close SaveChanges:=False

End Sub

Sub gmIssuereport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("Issue Temp").Select
        Sheets("Issue Temp").Copy
        Sheets("Issue Temp").Select
        Sheets("Issue Temp").Name = "GM Issue Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "GM Issue Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="issueDS", RefersToR1C1:= _
        "=issueDS.xlsx!issues[#All]"
         ActiveWorkbook.Names("issueDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\GMISSUE.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\issueDS.xlsx"
    
'Activete workbook
    Windows("issueDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("issues").Range.AutoFilter Field:=15, Criteria1:= _
        "GM"
   
    
    Range("issues[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("GMISSUE.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
'B3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Source]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'B3 Paste
Windows("GMISSUE.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste
        
'C3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'C3 Paste
Windows("GMISSUE.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste

'D3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'D3 Paste
Windows("GMISSUE.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
'E3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'E3 Paste
Windows("GMISSUE.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
'F3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'F3 Paste
Windows("GMISSUE.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],issueDS,6,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],issueDS,4,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],issueDS,2,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],issueDS,10,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],issueDS,12,0)"
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
    
    
    




 Columns("F:F").Select
    Range("Table2[[#Headers],[Due Date (12)]]").Activate
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"

    
' Save Report

ActiveWorkbook.Save
Windows("GMISSUE.xlsx").Activate

'Close data source
Workbooks("issueDS.XLSX").Close SaveChanges:=False

End Sub



Sub VVIssuereport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("Issue Temp").Select
        Sheets("Issue Temp").Copy
        Sheets("Issue Temp").Select
        Sheets("Issue Temp").Name = "VV Issue Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "Viral Vector Issue Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="issueDS", RefersToR1C1:= _
        "=issueDS.xlsx!issues[#All]"
         ActiveWorkbook.Names("issueDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\VVISSUE.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\issueDS.xlsx"
    
'Activete workbook
    Windows("issueDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("issues").Range.AutoFilter Field:=15, Criteria1:= _
        "VV"
   
    
    Range("issues[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("VVISSUE.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
'B3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Source]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'B3 Paste
Windows("VVISSUE.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste
        
'C3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'C3 Paste
Windows("VVISSUE.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste

'D3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'D3 Paste
Windows("VVISSUE.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
'E3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'E3 Paste
Windows("VVISSUE.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
'F3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'F3 Paste
Windows("VVISSUE.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],issueDS,6,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],issueDS,4,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],issueDS,2,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],issueDS,10,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],issueDS,12,0)"
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
    
    
    
 


    Columns("F:F").Select
    Range("Table2[[#Headers],[Due Date (12)]]").Activate
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"

    
' Save Report

ActiveWorkbook.Save
Windows("VVISSUE.xlsx").Activate

'Close data source
Workbooks("issueDS.XLSX").Close SaveChanges:=False

End Sub


Sub cc3Issuereport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("Issue Temp").Select
        Sheets("Issue Temp").Copy
        Sheets("Issue Temp").Select
        Sheets("Issue Temp").Name = "CC3 Issue Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "CC3 Issue Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="issueDS", RefersToR1C1:= _
        "=issueDS.xlsx!issues[#All]"
         ActiveWorkbook.Names("issueDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\CC3ISSUE.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\issueDS.xlsx"
    
'Activete workbook
    Windows("issueDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("issues").Range.AutoFilter Field:=15, Criteria1:= _
        "CC3"
   
    
    Range("issues[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("CC3ISSUE.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
'B3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Source]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'B3 Paste
Windows("CC3ISSUE.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste
        
'C3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'C3 Paste
Windows("CC3ISSUE.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste

'D3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'D3 Paste
Windows("CC3ISSUE.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
'E3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'E3 Paste
Windows("CC3ISSUE.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
'F3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'F3 Paste
Windows("CC3ISSUE.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],issueDS,6,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],issueDS,4,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],issueDS,2,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],issueDS,10,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],issueDS,12,0)"
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
    


 Columns("F:F").Select
    Range("Table2[[#Headers],[Due Date (12)]]").Activate
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"
   

    
' Save Report

ActiveWorkbook.Save
Windows("CC3ISSUE.xlsx").Activate

'Close data source
Workbooks("issueDS.XLSX").Close SaveChanges:=False

End Sub


Sub toIssuereport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("Issue Temp").Select
        Sheets("Issue Temp").Copy
        Sheets("Issue Temp").Select
        Sheets("Issue Temp").Name = "TO Issue Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "TO Issue Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="issueDS", RefersToR1C1:= _
        "=issueDS.xlsx!issues[#All]"
         ActiveWorkbook.Names("issueDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\TOISSUE.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\issueDS.xlsx"
    
'Activete workbook
    Windows("issueDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("issues").Range.AutoFilter Field:=15, Criteria1:= _
        "TO"
   
    
    Range("issues[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("TOISSUE.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
'B3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Source]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'B3 Paste
Windows("TOISSUE.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste
        
'C3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'C3 Paste
Windows("TOISSUE.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste

'D3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'D3 Paste
Windows("TOISSUE.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
'E3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'E3 Paste
Windows("TOISSUE.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
'F3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'F3 Paste
Windows("TOISSUE.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],issueDS,6,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],issueDS,4,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],issueDS,2,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],issueDS,10,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],issueDS,12,0)"
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
    
    
    

 Columns("F:F").Select
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"
   

    
' Save Report

ActiveWorkbook.Save
Windows("TOISSUE.xlsx").Activate

'Close data source
Workbooks("issueDS.XLSX").Close SaveChanges:=False

End Sub


Sub qaIssuereport()
'copy blank template for export

    Windows("templates.xlsx").Activate
        Sheets("Issue Temp").Select
        Sheets("Issue Temp").Copy
        Sheets("Issue Temp").Select
        Sheets("Issue Temp").Name = "QA Issue Report"
'Rename header
    Range("A1:G1").Select
        ActiveCell.FormulaR1C1 = "QA Issue Report"

'Setup Named Ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="issueDS", RefersToR1C1:= _
        "=issueDS.xlsx!issues[#All]"
         ActiveWorkbook.Names("issueDS").Comment = ""
         
        
'Save blank
    ActiveWorkbook.SaveAs "T:\James Patrick\Report Generation\exports\QAISSUE.xlsx"

'Open Data Source
    Workbooks.Open "T:\James Patrick\Report Generation\data\issueDS.xlsx"
    
'Activete workbook
    Windows("issueDS.xlsx").Activate
    
    

' Copy relevent data
    ActiveSheet.ListObjects("issues").Range.AutoFilter Field:=15, Criteria1:= _
        "QA"
   
    
    Range("issues[[#Headers],[Document Number]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Paste relevent data to blank'

Windows("QAISSUE.xlsx").Activate
    Range("A3").Select
        ActiveSheet.Paste
        
'B3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Source]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'B3 Paste
Windows("QAISSUE.xlsx").Activate
    Range("B3").Select
        ActiveSheet.Paste
        
'C3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Title]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'C3 Paste
Windows("QAISSUE.xlsx").Activate
    Range("C3").Select
        ActiveSheet.Paste

'D3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_Per]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'D3 Paste
Windows("QAISSUE.xlsx").Activate
    Range("D3").Select
        ActiveSheet.Paste
'E3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_CS]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'E3 Paste
Windows("QAISSUE.xlsx").Activate
    Range("E3").Select
        ActiveSheet.Paste
'F3
Windows("issueDS.xlsx").Activate
Range("issues[[#Headers],[iss_DD]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
'F3 Paste
Windows("QAISSUE.xlsx").Activate
    Range("F3").Select
        ActiveSheet.Paste
        
'Remove superfluopus information
Rows(3).Select
    Selection.Delete

'Input formulas
        'Formula Fill
    'Range("B3").Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],issueDS,6,0)"
    'Range("C3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],issueDS,4,0)"
    'Range("D3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],issueDS,2,0)"
    'Range("E3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],issueDS,10,0)"
    'Range("F3").Select
    'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],issueDS,12,0)"
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
    
    



    Columns("F:F").Select
    
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"

    
' Save Report

ActiveWorkbook.Save
Windows("QAISSUE.xlsx").Activate

'Close data source
Workbooks("issueDS.XLSX").Close SaveChanges:=False

End Sub






