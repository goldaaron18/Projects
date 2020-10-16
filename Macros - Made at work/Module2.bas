Attribute VB_Name = "Module2"

'Description: Formats excel sheet after DDS export and makes pivot table. "Anyformat3b" formats digital, and 3a formats everything else
'
'Features:
' -If document was already formated it skips formating and makes a pivot table, ie if you run it a 2nd time on the same doc
' it will just make another pivot table and not format.
' -Inserts a row for "parent" in which a formula is pasted from another sheet(hidden in recap). It gets parent using the network names from an index.
' -Renames all the headers on top
' -Replaces certain words/ networks with diffeent names. For some words, it filters by daypart(usualy 3-20 dayparts) so it only changes them for the
' specific daypart
' -Finds word "Fee" in specific column and changes column 5 cells to the left to fee so we can use it with the filter in the pivot table
' -For 3b, it does similar to "fee" but with "SCA" scatter
' -Creates pivot table on new sheet and assigns appropriate pivot fields
' -If > 2 workbooks are open it does not run.
'
'Instructions: user should have Recap file(for hidden headers and fomulas tab) and DDS pull file open
' User starts in pull file clicks anywhere, then runs macro. if desired, can run again to make a second pivot table.
'
'Improvements: we should be able to combine ths into 1 macro since its easy to know if its digital or not based on the words in the doc
'Instead of using "Estimate #" to filter, It would be better to use "Estimate Name" bc Estimate #s do change.
' Any other formating changes that need to be made to formating should be documented here.
' May want to use some features here to format "program" names as well for index bemchmarks.

Sub Anyformat3a()
Attribute Anyformat3a.VB_ProcData.VB_Invoke_Func = "a\n14"

    Application.ScreenUpdating = False

    Dim Worksheet1 As Worksheet
    Dim Worksheet2 As Worksheet
    

If Range("G1").Value = "Program" And Workbooks.Count > 3 Then
MsgBox "Please close 3rd open workbook"
Exit Sub
End If

If Range("G1").Value = "Program" Then
Set Worksheet1 = ActiveSheet
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Selection.Address), , xlYes).name = _
        "Table1"
    Range("Table1[#All]").Select
    ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleLight1"
    ActiveWindow.ActivatePrevious
    Set Worksheet2 = ActiveSheet
    Sheets("Headers").Visible = True
    Worksheets("Headers").Activate
    Range("A1:S1").Select
    Selection.Copy
    Worksheet1.Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("J:J").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
     Range("J1").Select
    ActiveCell.FormulaR1C1 = "Brand"
    ActiveWindow.ActivatePrevious
    Range("B2").Select
    Selection.Copy
    Worksheet1.Activate
    Range("J2").Select
    ActiveSheet.Paste
     Columns("B:B").Select
     Application.CutCopyMode = False
     Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
     Range("B1").Select
     ActiveCell.FormulaR1C1 = "Parent"
     ActiveWindow.ActivatePrevious
     Range("B3").Select
     Selection.Copy
     Worksheet1.Activate
     Range("B2").Select
     ActiveSheet.Paste
    ActiveWindow.ActivatePrevious
    Worksheet2.Activate
    Sheets("Headers").Visible = False
    Worksheet1.Activate
    LastRow = Range("A" & Rows.Count).End(xlUp).row
    Rows(LastRow).Select
    Selection.Delete Shift:=xlUp
     Cells.Replace What:="NBCS", Replacement:="NBC", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Cells.Replace What:="ESPU", Replacement:="ESPN", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
      Cells.Replace What:="ESPT", Replacement:="ESPN", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
      Cells.Replace What:="FSO", Replacement:="FOX", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

     With ActiveSheet.Range("D1:Z15000")
      Set C = .Find("PRO FEE", LookIn:=xlValues)
      If Not C Is Nothing Then
        firstAddress = C.Address
        Do
            C.Offset(0, -5).Value = "FEE"
            Set C = .FindNext(C)
        If C Is Nothing Then
            GoTo DoneFinding
        End If
        Loop While C.Address <> firstAddress
       End If
DoneFinding:
      End With

 With ActiveSheet.Range("D1:Z15000")
     Set C = .Find("FEE", LookIn:=xlValues)
     If Not C Is Nothing Then
        firstAddress = C.Address
        Do
            C.Offset(0, -5).Value = "FEE"
            Set C = .FindNext(C)
        If C Is Nothing Then
            GoTo DoneFinding2
        End If
        Loop While C.Address <> firstAddress
      End If
DoneFinding2:
     End With
     
     ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=6, Criteria1:= _
        Array("80", "85", "89", "93", "49", "54", "58", "62"), Operator:=xlFilterValues
      Cells.Replace What:="ABC", Replacement:="ESPN", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
      Cells.Replace What:="GOLF", Replacement:="NBC", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
     
     ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=6, Criteria1:= _
        Array("86", "55"), Operator:=xlFilterValues
      Cells.Replace What:="ABC", Replacement:="NCAA BB", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
      Cells.Replace What:="CBS", Replacement:="NCAA BB", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
      Cells.Replace What:="ESPN", Replacement:="NCAA BB", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
      ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=6


End If



    pivot = "PivotTable" & Sheets.Count
    Sheet = "Pivot" & Sheets.Count
    Dest = Sheet & "!R4C1"
    Application.CutCopyMode = False
    Sheets.Add.name = Sheet
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Table1", Version:=6).CreatePivotTable TableDestination:=Dest, _
        TableName:=pivot, DefaultVersion:=6
    Sheets(Sheet).Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables(pivot)
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .errorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables(pivot).PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables(pivot).RepeatAllLabels xlRepeatLabels

     With ActiveSheet.PivotTables(pivot).PivotFields("Net")
        .Orientation = xlRowField
        .Position = 1
     End With
   
    With ActiveSheet.PivotTables(pivot).PivotFields("Brand")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(pivot).PivotFields("Month")
        .Orientation = xlColumnField
        .Position = 2
    End With
    ActiveSheet.PivotTables(pivot).PivotFields("Month").AutoGroup
    ActiveSheet.PivotTables(pivot).AddDataField ActiveSheet.PivotTables( _
        pivot).PivotFields("Net Cost"), "Sum of Net Cost", xlSum
    With ActiveSheet.PivotTables(pivot).PivotFields("Buy Type")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(pivot).PivotFields("Est Name")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables(pivot).PivotFields("Buy Type").CurrentPage = _
        "Upfront"
  


End Sub






Sub Anyformat3b()

Dim Range1 As Range
Application.ScreenUpdating = False

If Workbooks.Count > 3 Then
MsgBox "Please close 3rd open workbook"
Else
If Range("H1").Value = "PRISMA" Then

    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
        
    
    LastRow = Range("A" & Rows.Count).End(xlUp).row
    Rows(LastRow - 1 & ":" & LastRow).Select
    Selection.Delete Shift:=xlUp
    
    Columns("D:D").Select
    Selection.NumberFormat = "m/d/yyyy"
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ActiveWindow.ActivatePrevious
    Sheets("Headers").Visible = True
    Worksheets("Headers").Activate
    Range("B6").Select
    Selection.Copy
    ActiveWindow.ActivatePrevious
    Range("E2").Select
    ActiveSheet.Paste
    LastRow = Range("A" & Rows.Count).End(xlUp).row
    Selection.AutoFill Destination:=Range("E2:E" & LastRow)
    
    Columns("E:E").Select
    Selection.Copy
    Columns("D:D").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    
    ActiveWindow.ActivatePrevious
    Range("A5:I5").Select
    Selection.Copy
    Sheets("Headers").Visible = False
    ActiveWindow.ActivatePrevious
    Range("A1").Select
    ActiveSheet.Paste
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Selection.Address), , xlYes).name = _
        "Table1"
    Range("Table1[#All]").Select
    ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleLight1"
    
    
    Cells.Replace What:="DISNEY ONLINE", Replacement:="ABC.COM", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="FOX SPORTS", Replacement:="FOX.COM", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="HULU LATINO", Replacement:="HULU.COM", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="NBC SPORTS", Replacement:="NBC UNIVERSAL", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="NBC UNIVERSAL-", Replacement:="NBC UNIVERSAL", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="TELEMUNDO NETWORK", Replacement:="NBC UNIVERSAL", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="CNN INTERACTIVE", Replacement:="THE TURNER NETWORK", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="A+E TELEVISION NWK", Replacement:="AETN", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
   Cells.Replace What:="ABC.COM", Replacement:="ABC", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="FOX.COM", Replacement:="FOX", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="FX NETWORK.COM", Replacement:="FOX", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="HULU.COM", Replacement:="HULU", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="NBC UNIVERSAL", Replacement:="NBC", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="THE TURNER NETWORK", Replacement:="Turner", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="CWTV.COM", Replacement:="CW", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="DISCOVERY.COM", Replacement:="Discovery", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="ESPN.COM", Replacement:="ESPN", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="CBS.COM", Replacement:="CBS", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
 
 With ActiveSheet.Range("G:G")
     Set C = .Find("SCA", LookIn:=xlValues)
     If Not C Is Nothing Then
        firstAddress = C.Address
        Do
            C.Offset(0, -2).Value = "SCATTER"
            Set C = .FindNext(C)
        If C Is Nothing Then
            GoTo DoneFinding
        End If
        Loop While C.Address <> firstAddress
      End If
DoneFinding:
End With
    


 
End If
End If


    pivot = "PivotTable" & Sheets.Count
    Sheet = "Pivot" & Sheets.Count
    Dest = Sheet & "!R4C1"
    Application.CutCopyMode = False
    Sheets.Add.name = Sheet
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Table1", Version:=6).CreatePivotTable TableDestination:=Dest, _
        TableName:=pivot, DefaultVersion:=6
    Sheets(Sheet).Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables(pivot)
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .errorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables(pivot).PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables(pivot).RepeatAllLabels xlRepeatLabels
    
    With ActiveSheet.PivotTables(pivot).PivotFields("Vendor Name")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(pivot).PivotFields("Insertion Date ")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables(pivot).PivotFields("Insertion Date ").AutoGroup
    ActiveSheet.PivotTables(pivot).AddDataField ActiveSheet.PivotTables( _
         pivot).PivotFields("Net Ordered"), "Sum of Net Ordered", xlSum
    With ActiveSheet.PivotTables(pivot).PivotFields("Creative Type")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(pivot).PivotFields("Client Code")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables(pivot).PivotFields("Creative Type").CurrentPage = _
        "DISPLAY"

End Sub










