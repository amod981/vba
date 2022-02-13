Sub pivotpractice()
    
    Dim pc As PivotCache
    Dim pt As pivottable
    Dim pf As PivotField
    Dim filter_array As Variant
    
    Sheets("macro_pivot_output").Delete
    Sheets.Add.Name = "macro_pivot_output"
    
    
    
    Set pc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Sheets("Sheet1").Range("A1:G44"))
    
    Set pt = pc.CreatePivotTable(tabledestination:=Sheets("macro_pivot_output").Range("A1:A1"), tablename:="pivot_of_sales")


    'filters
    filter_array = Array("Region", "Rep")
    'columns
    column_array = Array("Item")
    'rows
    row_array = Array("OrderDate")
    'values
    data_array = Array("Units", "UnitCost", "Total")
    
    
    
    For Each Item In filter_array
        pt.PivotFields(Item).Orientation = xlPageField
    Next
    
    For Each Item In column_array
        pt.PivotFields(Item).Orientation = xlColumnField
    Next
    
    For Each Item In row_array
        pt.PivotFields(Item).Orientation = xlRowField
    Next
    
    For Each Item In data_array
        With pt.PivotFields(Item)
            .Orientation = xlDataField
            .Function = Sum
        End With
        
    Next
