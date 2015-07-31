Attribute VB_Name = "Module9"
Sub Report_by_Name()
    '******************************************************************
    '  Author - Tom Uyemura
    '  Language - Visual Basic for Applications - VBA
    '     1/29/2015 - Changed this to sort by NAME.  No longer to Sort by Account Number  TAU
    ' .
    '******************************************************************
    
    'Create new worksheet
    'object method
    '******************************************************************
    '  This section will make a copy of the " Space #" sheet and
    '    it at the end.
    '******************************************************************
    Dim wk_SheetName As String
    
    Sheets("Space #").Select
    Sheets("Space #").Copy After:=Sheets(1)
    wk_SheetName = ActiveSheet.Name
    ' ** For Debugging **
    ' MsgBox "The New sheet name is " & wk_SheetName
    Sheets(wk_SheetName).Select
    Sheets(wk_SheetName).Name = "by_Name"
    
    '******************************************************************
    '  This section will do some cleaning up,
    '    like hard code the Lot Names
    '******************************************************************
    
    ActiveWorkbook.Worksheets("by_Name").Select
    ' Unprotect the new sheet
    ActiveSheet.Unprotect Password:="RV"
    
    Range("C2:C500").Select
    Selection.Cut
    Range("C501").Select
    ActiveSheet.Paste
    
    Range("C501:C1000").Select
    Selection.Copy
    Range("C2").Select
    ActiveSheet.PasteSpecial Format:=3, Link:=1, DisplayAsIcon:=False, _
        IconFileName:=False
    Range("A1").Select
    
    Range("C501:C1000").Select
    Selection.Delete
    Range("A1").Select
    
    
    '******************************************************************
    '  This section will do the sorting
    '******************************************************************
    
    Cells.Select
    ActiveWorkbook.Worksheets("by_Name").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("by_Name").Sort.SortFields.Add Key:=Range("C2:C430" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("by_Name").Sort.SortFields.Add Key:=Range("A2:A430" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("by_Name").Sort.SortFields.Add Key:=Range("B2:B430" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("by_Name").Sort
        .SetRange Range("A1:AI430")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A1").Select
          
End Sub
Sub MySort()
Attribute MySort.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MySort Macro
'

'
    Cells.Select
    ActiveWorkbook.Worksheets("by_Acct").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("by_Acct").Sort.SortFields.Add Key:=Range("A2:A430" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("by_Acct").Sort.SortFields.Add Key:=Range("C2:C430" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("by_Acct").Sort.SortFields.Add Key:=Range("B2:B430" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("by_Acct").Sort
        .SetRange Range("A1:AI430")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
End Sub
Sub Report_by_Space()
    '******************************************************************
    '  Author - Tom Uyemura
    '  Language - Visual Basic for Applications - VBA
    ' .
    '******************************************************************


    'Create new worksheet
    'object method
    '******************************************************************
    '  This section will make a copy of the " Space #" sheet and
    '    it at the end.
    '******************************************************************
    Dim ws As Worksheet
    Dim wk_SheetName As String
    Dim wk_OrigSheet As String
    Dim i As Integer
    
    Sheets("by_Acct_by_TAU").Select
    Sheets("by_Acct_by_TAU").Copy After:=Sheets(2)
    wk_SheetName = ActiveSheet.Name
    ' ** For Debugging **
    ' MsgBox "The New sheet name is " & wk_SheetName
    Sheets(wk_SheetName).Select
    Sheets(wk_SheetName).Name = "by_Space"
    wk_SheetName = ActiveSheet.Name
'    ws = ActiveSheet
    
    ActiveWorkbook.Worksheets("by_Space").Select
    
    '******************************************************************
    '  This section will do the sorting  ** SORT BY LOCATION & LOT #
    '******************************************************************
    ActiveWorkbook.Worksheets("by_Space").Select
    Range("A1").Select
    
    Cells.Select
    ActiveWorkbook.Worksheets("by_Space").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("by_Space").Sort.SortFields.Add Key:=Range("C2:C430" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("by_Space").Sort.SortFields.Add Key:=Range("B2:B430" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
    With ActiveWorkbook.Worksheets("by_Space").Sort
        .SetRange Range("A1:AI430")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A1").Select
          
    '******************************************************************
    '  This section Changes tha Layout of to
    '    Combine the Lot Name with the Space Number
    '******************************************************************
              
    Range("AA2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(LEFT(RC3,1),""00"",RC2)"
    Range("AA2").Select
    Selection.Copy
    Range("AA3:AA13").Select
    ActiveSheet.Paste
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    
    Range("AA14").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(LEFT(RC3,1),""0"",RC2)"
    Range("AA14").Select
    Selection.Copy
    Range("AA15:AA117").Select
    ActiveSheet.Paste
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("AA118").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(LEFT(RC3,1),"""",RC2)"
    Range("AA118").Select
    Selection.Copy
    Range("AA119:AA308").Select
    ActiveSheet.Paste
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("AA309").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(LEFT(RC3,1),""00"",RC2)"
    Range("AA309").Select
    Selection.Copy
    Range("AA310:AA317").Select
    ActiveSheet.Paste
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    
    Range("AA318").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(LEFT(RC3,1),""0"",RC2)"
    Range("AA318").Select
    Selection.Copy
    Range("AA319:AA416").Select
    ActiveSheet.Paste
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
        
    Range("AA2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B2").Select
    Selection.EntireColumn.Delete
    Range("A2").Select
    
    'Now let's clean up alittle bit
    Range("Z2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete
    Application.CutCopyMode = False
    Range("A2").Select
   
End Sub
