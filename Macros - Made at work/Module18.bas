Attribute VB_Name = "Module18"

'Use: Helps Update Index Benchmarking document faster
'Description: This macro takes data from a pivot table with 2 columns and updates the data in onother sheet.
' It matches each rows network name in the pivot table and finds it in the other sheet and updates the corresponding columns with the data from the pivot table.
'
'Features:
' -Clears data from other sheet so result is only new data from the pivot table.
' -Takes range input from user to update a specific 2 columns(ie "Net Dollars" and the column next to it "IMPS").
' -Macro starts on sheet that needs to be updated, gets the tab name(such as Chevy), and filters for the tab name in the pivot table.
' -The column/ index is set to culumn C which is the networks and rows we are updating if they match the networks in the pivot.
' -If a network in the pivot table is not found, a message box pops up and says " ___  network not found"
'
'Instructions: user should have Index Benchmark file open and pivot file open. The pivot file data should be filtered correctly
'but filtering by brand is not neccesary since the macro does it for the user.
' User starts in Index benchmark file and finds the part of the excel sheet that needs to be updated(ie. 1Q20 or 2Q20)
' User runs macro and then when prompted select the cell with name "Net Dollars"
'
'
' Improvements: Macro relies on Index and many times there are new networks and all of the networks in the index may not be included.
' Also sometimes if there is words in CAPs the word is not found
' Have a macro to formats words in the Pivot table if the index is not sufficient.


Sub IndexUpdate()
Attribute IndexUpdate.VB_ProcData.VB_Invoke_Func = "i\n14"

Dim Column1 As String
Dim First As Range
Dim Range1 As Range
Dim net2 As Range
Dim LRow2 As Long
Dim LRow1 As Long
Dim PT As PivotTable

Dim Range2 As Range
Dim Count As Long
Dim net As String
Dim Line As Range
Dim Netrng As Range
Dim Dollars As Range
Dim DollarColumn As Range
Dim i As Long

Scolumn = "C"

Application.ScreenUpdating = False

If Workbooks.Count > 3 Then
MsgBox "Please close 3rd open workbook"
Exit Sub
End If

    Set Dollars = Application.InputBox("Range", Type:=8)
    
    If Dollars.Value = "Net Dollars" Then
       
       With Application.FindFormat.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
       End With
    Column = (Split(Dollars.Address, "$")(1))
    Set DollarColumn = Range(Column & "1", Column & "1000")

    lastdollars = DollarColumn.Find(What:="SUM", After:=Dollars, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=True).row
    MsgBox lastdollars
    Range(Dollars.Offset(1, 0), Dollars.Offset(lastdollars - 5, 1)).Select
    Selection.ClearContents
    Application.FindFormat.clear
    
    Set First = Dollars.Offset(1, 0)
    Column1 = First.Column
    Set net2 = Range(Scolumn & First.row)
    Set Range1 = Range(net2, net2.End(xlDown))
    
    
    If ActiveSheet.name = "Cadillac" Or ActiveSheet.name = "Cadillac Prime" Then
    ActiveWindow.ActivatePrevious
    For Each PT In ActiveSheet.PivotTables
    PTname = PT.name
    Next
    With ActiveSheet.PivotTables(PTname).PivotFields("Brand")
        .ClearAllFilters
        .EnableMultiplePageItems = True
        For Each pvtitem In .PivotItems
        If pvtitem.name = "Cadillac" Or pvtitem.name = "Cadillac Retail" Then
            pvtitem.Visible = True
            Else
            pvtitem.Visible = False
           End If
          Next
    End With
    ElseIf ActiveSheet.name = "Chevy" Or ActiveSheet.name = "Chevy Prime" Then
    ActiveWindow.ActivatePrevious
    For Each PT In ActiveSheet.PivotTables
    PTname = PT.name
    Next
    With ActiveSheet.PivotTables(PTname).PivotFields("Brand")
        .ClearAllFilters
        .EnableMultiplePageItems = True
        For Each pvtitem In .PivotItems
        If pvtitem.name = "Chevy" Or pvtitem.name = "Chevy Retail" Then
            pvtitem.Visible = True
            Else
            pvtitem.Visible = False
           End If
          Next
    End With
    ElseIf ActiveSheet.name = "Buick" Or ActiveSheet.name = "Buick Prime" Then
    ActiveWindow.ActivatePrevious
    For Each PT In ActiveSheet.PivotTables
    PTname = PT.name
    Next
    With ActiveSheet.PivotTables(PTname).PivotFields("Brand")
        .ClearAllFilters
        .EnableMultiplePageItems = True
        For Each pvtitem In .PivotItems
        If pvtitem.name = "Buick" Then
            pvtitem.Visible = True
            Else
            pvtitem.Visible = False
           End If
          Next
    End With
    ElseIf ActiveSheet.name = "GMC" Or ActiveSheet.name = "GMC Prime" Then
    ActiveWindow.ActivatePrevious
    For Each PT In ActiveSheet.PivotTables
    PTname = PT.name
    Next
    With ActiveSheet.PivotTables(PTname).PivotFields("Brand")
        .ClearAllFilters
        .EnableMultiplePageItems = True
        For Each pvtitem In .PivotItems
        If pvtitem.name = "GMC" Then
            pvtitem.Visible = True
            Else
            pvtitem.Visible = False
           End If
          Next
    End With
    ElseIf ActiveSheet.name = "OnStar" Then
    ActiveWindow.ActivatePrevious
    For Each PT In ActiveSheet.PivotTables
    PTname = PT.name
    Next
    With ActiveSheet.PivotTables(PTname).PivotFields("Brand")
        .ClearAllFilters
        .EnableMultiplePageItems = True
        For Each pvtitem In .PivotItems
        If pvtitem.name = "OnStar" Then
            pvtitem.Visible = True
            Else
            pvtitem.Visible = False
           End If
          Next
    End With
    
    Else
    MsgBox "Please select correct tab"
    Exit Sub
    End If
    
    
    Range("A1").Activate
    Cells.Find(What:="Row Labels", After:=ActiveCell, LookIn:=xlValues, LookAt:= _
        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=True).Activate
    LRow1 = Cells(Rows.Count, 1).End(xlUp).row - ActiveCell.row - 1
    Set Range2 = Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(LRow1, 0))
    
      Count = 1
      While Count <= Range2.Count
      Set Line = Range2.Rows(Count)
      net = Line.Value
      Line.Offset(0, 1).Copy
      ActiveWindow.ActivatePrevious
      Range1.Activate
      Set Netrng = Range1.Find(What:=net, After:=ActiveCell, LookIn:=xlValues, _
      LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
      MatchCase:=False, SearchFormat:=False)
      If Not Netrng Is Nothing Then
       Netrng.Offset(0, Column1 - Range1.Column).PasteSpecial Paste:=xlPasteValues
       ActiveWindow.ActivatePrevious
       Line.Offset(0, 2).Copy
       ActiveWindow.ActivatePrevious
       Netrng.Offset(0, (Column1 - Range1.Column + 1)).PasteSpecial Paste:=xlPasteValues
       ActiveWindow.ActivatePrevious
      Else
      ActiveWindow.ActivatePrevious
      MsgBox net & " Not found"
      
      End If
      Count = Count + 1
      Wend
    Else
    MsgBox "Please click on Net Dollars"
  End If
     
   
End Sub



    
    

    
       
       







