Attribute VB_Name = "Setup"
Sub workbenchSetup()
Workbooks.Open ("T:\James Patrick\Report Generation\ml.xlsx")
Workbooks.Open ("T:\James Patrick\Report Generation\UserNames.xlsx")
Workbooks.Open ("T:\James Patrick\Report Generation\Templates.xlsx")

Workbooks("ReportGenv1.0").Activate


End Sub
Sub docSetup()
'
' docSetup Macro
'

'open
    Workbooks.Open ("T:\James Patrick\Report Generation\docs.xlsx")
    
    Windows("docs.xlsx").Activate
    
'rename worksheet
    
    ActiveSheet.Name = "docs"
    
'name ml and pertable ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
'create and name table
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$M$290"), , xlYes).Name = _
        "docs"
            
         
'Insert created columns

Range("N1").Select
    ActiveCell.FormulaR1C1 = "doc_Per"
Range("N2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP([@[User ID]],perTable,4,0)"
Range("O1").Select
    ActiveCell.FormulaR1C1 = "doc_Dept"
Range("O2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP([@[User ID]],perTable,6,0)"
Range("P1").Select
    ActiveCell.FormulaR1C1 = "doc_DO"
Range("P2").Select
    ActiveCell.FormulaR1C1 = "=(today() - [@[Notification Date]])"
    

'Change Range names
Range("C1").Select
    ActiveCell.FormulaR1C1 = "doc_PID"
Range("D1").Select
    ActiveCell.FormulaR1C1 = "doc_Title"
Range("N1").Select
    ActiveCell.FormulaR1C1 = "doc_Per"
Range("J1").Select
    ActiveCell.FormulaR1C1 = "doc_Step"


    
'save data source
ActiveWorkbook.SaveAs ("T:\James Patrick\Report Generation\Data\DocsDS.xlsx")
'close data source
ActiveWorkbook.Close

End Sub


Sub apSetup()
'
'Action Plan setup
'

'open
    Workbooks.Open ("T:\James Patrick\Report Generation\aps.xlsx")
    
    Windows("aps.xlsx").Activate
    
'rename worksheet
    
    ActiveSheet.Name = "aps"
    
'create and name table
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$M$290"), , xlYes).Name = _
        "aps"
        
'name ml and pertable ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
         
'cut and paste fist column
Columns("A:A").Select
    Selection.Cut
Columns("O:O").Select
    Selection.Insert Shift:=xlToRight

'Rename Ranges
Range("E1").Select
    ActiveCell.FormulaR1C1 = "User ID"

Range("A1").Select
    ActiveCell.FormulaR1C1 = "Document Number"
    
Range("G1").Select
    ActiveCell.FormulaR1C1 = "ap_NCE"
    
Range("B1").Select
    ActiveCell.FormulaR1C1 = "ap_APT"
    
Range("N1").Select
    ActiveCell.FormulaR1C1 = "ap_CS"
    
Range("D1").Select
    ActiveCell.FormulaR1C1 = "ap_DD"
    

    

    
'Insert created columns

Range("O1").Select
    ActiveCell.FormulaR1C1 = "ap_Per"
Range("O2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP([@[User ID]],perTable,4,0)"
Range("P1").Select
    ActiveCell.FormulaR1C1 = "ap_Dept"
Range("P2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP([@[User ID]],perTable,6,0)"
    
'save data source
ActiveWorkbook.SaveAs ("T:\James Patrick\Report Generation\Data\apsDS.xlsx")
'close data source
ActiveWorkbook.Close


End Sub

Sub issueSetup()
'
'issue setup
'

'open
    Workbooks.Open ("T:\James Patrick\Report Generation\issues.xlsx")
    
    Windows("issues.xlsx").Activate
    
'rename worksheet
    
    ActiveSheet.Name = "issues"
    
'create and name table
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$M$290"), , xlYes).Name = _
        "issues"
        
'name ml and pertable ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
         


'Rename B1
Range("B1").Select
ActiveCell.FormulaR1C1 = "User ID"

Range("A1").Select
    ActiveCell.FormulaR1C1 = "Document Number"
    
    
'Insert created columns

Range("N1").Select
    ActiveCell.FormulaR1C1 = "iss_Per"
Range("N2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP([@[User ID]],perTable,4,0)"
Range("O1").Select
    ActiveCell.FormulaR1C1 = "iss_Dept"
Range("O2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP([@[User ID]],perTable,6,0)"

Range("F1").Select
    ActiveCell.FormulaR1C1 = "iss_Source"
Range("I1").Select
    ActiveCell.FormulaR1C1 = "iss_CS"
Range("L1").Select
    ActiveCell.FormulaR1C1 = "iss_DD"
Range("O1").Select
    ActiveCell.FormulaR1C1 = "iss_Dept"
Range("D1").Select
    ActiveCell.FormulaR1C1 = "iss_Title"

    
'save data source
ActiveWorkbook.SaveAs ("T:\James Patrick\Report Generation\Data\issueDS.xlsx")
'close data source
ActiveWorkbook.Close


End Sub

Sub ccsSetup()
'
' ccsSetup Macro
'

'open
    Workbooks.Open ("T:\James Patrick\Report Generation\ccs.xlsx")
    
    Windows("ccs.xlsx").Activate
    
'rename worksheet
    
    ActiveSheet.Name = "ccs"
    
'create and name table
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$M$290"), , xlYes).Name = _
        "ccs"
        
'name ml and pertable ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
         
'Rename F1
    Range("F1").Select
        ActiveCell.FormulaR1C1 = "User ID"
    Range("A1").Select
        ActiveCell.FormulaR1C1 = "Document Number"
    
         

         
'Insert created columns

Range("I1").Select
    ActiveCell.FormulaR1C1 = "cc_Per"
Range("I2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP([@[User ID]],perTable,4,0)"
Range("J1").Select
    ActiveCell.FormulaR1C1 = "cc_Dept"
Range("J2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP([@[User ID]],perTable,6,0)"

Range("C1").Select
    ActiveCell.FormulaR1C1 = "cc_Title"
Range("G1").Select
    ActiveCell.FormulaR1C1 = "cc_CS"
Range("H1").Select
    ActiveCell.FormulaR1C1 = "cc_SD"
Range("E1").Select
    ActiveCell.FormulaR1C1 = "cc_DD"
    
'save data source
ActiveWorkbook.SaveAs ("T:\James Patrick\Report Generation\Data\ccsDS.xlsx")
'close data source
ActiveWorkbook.Close

End Sub

Sub capaSetup()
'
' docSetup Macro
'

'open
    Workbooks.Open ("T:\James Patrick\Report Generation\capas.xlsx")
    
    Windows("capas.xlsx").Activate
    
'rename worksheet
    
    ActiveSheet.Name = "capas"
    
'name ml and pertable ranges
    ActiveWorkbook.Names.Add Name:="ml", RefersToR1C1:="=ml.xlsx!ml[#All]"
        ActiveWorkbook.Names("ml").Comment = ""
    
    ActiveWorkbook.Names.Add Name:="perTable", RefersToR1C1:= _
        "=UserNames.xlsx!Table3[#All]"
         ActiveWorkbook.Names("perTable").Comment = ""
    
'create and name table
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$M$290"), , xlYes).Name = _
        "capas"
        
'Rename F1
    Range("B1").Select
        ActiveCell.FormulaR1C1 = "User ID"
    Range("A1").Select
        ActiveCell.FormulaR1C1 = "Document Number"
         
'Insert created columns

Range("J1").Select
    ActiveCell.FormulaR1C1 = "Personnel"
Range("J2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP([@[User ID]],perTable,4,0)"
Range("K1").Select
    ActiveCell.FormulaR1C1 = "Dept"
Range("K2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP([@[User ID]],perTable,6,0)"
    
'save data source
ActiveWorkbook.SaveAs ("T:\James Patrick\Report Generation\Data\capasDS.xlsx")
'close data source
ActiveWorkbook.Close

End Sub

