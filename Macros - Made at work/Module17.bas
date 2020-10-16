Attribute VB_Name = "Module17"

'Use: Updates Productivity recap brand allocations with data from pivot table faster.
'Description: This Macro takes data from a pivot table(ex. $$ or GRPs) with brand allocations by month and drops the values in the
' corresponding rows/columns in the Productivity Recap doc.
'
'Features:
' -Works for updating Pivot tabel with $$ or when updating GRPs by knowing what pivot table it being used.
'  It updates a different column in the recap($$ or GRPs) depending on which pivot table is open.
' -Macro puts data in correct columns even for networks that don't have data for all 3 months, ie Jan, Feb or Jan, March ect.
' -Only updates the brands/months that are actually in the pivot(by searching for the words in the pivot table)
' -Clears data previously in recap, and puts "0" in blank cells
' -Code does not run if column 2 is not selected ie the columns with network names in the recap/ dayparts
'
'Instructions: User should have Recap file and pivot table file open with desired data(correct filters).
' The order of Networks in the Recap needs to be in the same exact order of the networks in the pivot table(use "Reorder" macro if necessary or manually reorder)
' User starts in Recap file and clicks on first network name in the daypart that will be updated(ie the same first as the pivot)
'
' Improvements: include code for instances when there is only 1 month of data in a Qaurter.
' Also we may want to use this macro with another macro to update All or multiple dayparts at once.
' The macro is long and the code is repetitive making it harder to edit. should cut down code and limit repetitive code.


Sub Brandupdate3()
Attribute Brandupdate3.VB_ProcData.VB_Invoke_Func = "b\n14"

Application.ScreenUpdating = False

If ActiveCell.Column = 2 Then
 ActiveWindow.ActivatePrevious
 If (WorksheetFunction.CountIf(Range("A1:Z50"), "Sum of AD2554 GRPs")) > 0 Or (WorksheetFunction.CountIf(Range("A1:Z50"), "Sum of CALC GRP")) > 0 Then
    ActiveWindow.ActivatePrevious
    ActiveCell.Offset(0, 28).Select
Else
    ActiveWindow.ActivatePrevious
    ActiveCell.Offset(0, 25).Select
End If
    
    ActiveWindow.ActivatePrevious
    Range("A1").Select
    Dim Range1 As Range
    Dim Months As Range
    Dim LRow As Long
    Dim clear As Range
    Dim Zero1 As Range
    
    
      If (WorksheetFunction.CountIf(Range("A1:Z50"), "Buick")) > 0 Then
    Cells.Find(What:="Buick", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
     LRow = Cells(Rows.Count, 1).End(xlUp).row - 6
     Set Range1 = Range(ActiveCell.Offset(3, 0), ActiveCell.Offset(LRow, 2))
     Set Months = Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(1, 2))
     End If
    If (WorksheetFunction.CountIf(Range("A1:Z50"), "Buick")) = 3 Then
     Range1.Columns(1).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     Application.Run "PERSONAL.XLSB!PasteValue"
     Range1.Columns(2).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 4).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
     Range1.Columns(3).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 4).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
    ElseIf (WorksheetFunction.CountIf(Range("A1:Z50"), "Buick")) = 2 Then
     If Months.Columns(1).Value = "Jan" Or Months.Columns(1).Value = "Apr" Or Months.Columns(1).Value = "Jul" Or Months.Columns(1).Value = "Oct" Then
      Range1.Columns(1).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      Application.Run "PERSONAL.XLSB!PasteValue"
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 4).Select
      ActiveWindow.ActivatePrevious
     Else
      ActiveWindow.ActivatePrevious
      Set clear = Range(ActiveCell.Offset(0, 0), ActiveCell.Offset((LRow - 3), 0))
      clear.ClearContents
      ActiveCell.Offset(0, 4).Select
      ActiveWindow.ActivatePrevious
      Range1.Columns(1).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      Application.Run "PERSONAL.XLSB!PasteValue"
     End If
     If Months.Columns(2).Value = "Feb" Or Months.Columns(2).Value = "May" Or Months.Columns(2).Value = "Aug" Or Months.Columns(2).Value = "Nov" Then
      Range1.Columns(2).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      Application.Run "PERSONAL.XLSB!PasteValue"
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 4).Select
      Set clear = Range(ActiveCell.Offset(0, 0), ActiveCell.Offset((LRow - 3), 0))
      clear.ClearContents
      ActiveWindow.ActivatePrevious
     Else
      Range1.Columns(2).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 4).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
     End If
    Else
     ActiveWindow.ActivatePrevious
     ActiveCell.Offset(0, 8).Select
    ActiveWindow.ActivatePrevious
    End If
    
      If (WorksheetFunction.CountIf(Range("A1:Z50"), "Cadillac")) > 0 Then
    Cells.Find(What:="Cadillac", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
     LRow = Cells(Rows.Count, 1).End(xlUp).row - 6
     Set Range1 = Range(ActiveCell.Offset(3, 0), ActiveCell.Offset(LRow, 2))
     Set Months = Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(1, 2))
     End If
    If (WorksheetFunction.CountIf(Range("A1:Z50"), "Cadillac")) = 3 Then
     Range1.Columns(1).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 8).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
     Range1.Columns(2).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 4).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
     Range1.Columns(3).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 4).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
    ElseIf (WorksheetFunction.CountIf(Range("A1:Z50"), "Cadillac")) = 2 Then
     If Months.Columns(1).Value = "Jan" Or Months.Columns(1).Value = "Apr" Or Months.Columns(1).Value = "Jul" Or Months.Columns(1).Value = "Oct" Then
      Range1.Columns(1).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 8).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 4).Select
      ActiveWindow.ActivatePrevious
     Else
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 8).Select
      Set clear = Range(ActiveCell.Offset(0, 0), ActiveCell.Offset((LRow - 3), 0))
      clear.ClearContents
      ActiveWindow.ActivatePrevious
      Range1.Columns(1).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 4).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
     End If
     If Months.Columns(2).Value = "Feb" Or Months.Columns(2).Value = "May" Or Months.Columns(2).Value = "Aug" Or Months.Columns(2).Value = "Nov" Then
      Range1.Columns(2).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      Application.Run "PERSONAL.XLSB!PasteValue"
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 4).Select
      Set clear = Range(ActiveCell.Offset(0, 0), ActiveCell.Offset((LRow - 3), 0))
      clear.ClearContents
      ActiveWindow.ActivatePrevious
     Else
      Range1.Columns(2).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 4).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
     End If
    Else
    ActiveWindow.ActivatePrevious
    ActiveCell.Offset(0, 16).Select
    ActiveWindow.ActivatePrevious
    End If
    
       If (WorksheetFunction.CountIf(Range("A1:Z50"), "Cadillac Retail")) > 0 Then
    Cells.Find(What:="Cadillac Retail", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
     LRow = Cells(Rows.Count, 1).End(xlUp).row - 6
     Set Range1 = Range(ActiveCell.Offset(3, 0), ActiveCell.Offset(LRow, 2))
     Set Months = Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(1, 2))
     End If
    If (WorksheetFunction.CountIf(Range("A1:Z50"), "Cadillac Retail")) = 3 Then
     Range1.Columns(1).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 8).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
     Range1.Columns(2).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 4).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
     Range1.Columns(3).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 4).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
    ElseIf (WorksheetFunction.CountIf(Range("A1:Z50"), "Cadillac Retail")) = 2 Then
     If Months.Columns(1).Value = "Jan" Or Months.Columns(1).Value = "Apr" Or Months.Columns(1).Value = "Jul" Or Months.Columns(1).Value = "Oct" Then
      Range1.Columns(1).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 8).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 4).Select
      ActiveWindow.ActivatePrevious
     Else
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 8).Select
      Set clear = Range(ActiveCell.Offset(0, 0), ActiveCell.Offset((LRow - 3), 0))
      clear.ClearContents
      ActiveWindow.ActivatePrevious
      Range1.Columns(1).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 4).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
     End If
     If Months.Columns(2).Value = "Feb" Or Months.Columns(2).Value = "May" Or Months.Columns(2).Value = "Aug" Or Months.Columns(2).Value = "Nov" Then
      Range1.Columns(2).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      Application.Run "PERSONAL.XLSB!PasteValue"
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 4).Select
      Set clear = Range(ActiveCell.Offset(0, 0), ActiveCell.Offset((LRow - 3), 0))
      clear.ClearContents
      ActiveWindow.ActivatePrevious
     Else
      Range1.Columns(2).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 4).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
     End If
    Else
    ActiveWindow.ActivatePrevious
    ActiveCell.Offset(0, 16).Select
    ActiveWindow.ActivatePrevious
    End If
    
       If (WorksheetFunction.CountIf(Range("A1:Z50"), "Chevy")) > 0 Then
    Cells.Find(What:="Chevy", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
     LRow = Cells(Rows.Count, 1).End(xlUp).row - 6
     Set Range1 = Range(ActiveCell.Offset(3, 0), ActiveCell.Offset(LRow, 2))
     Set Months = Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(1, 2))
     End If
    If (WorksheetFunction.CountIf(Range("A1:Z50"), "Chevy")) = 3 Then
     Range1.Columns(1).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 8).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
     Range1.Columns(2).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 4).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
     Range1.Columns(3).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 4).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
    ElseIf (WorksheetFunction.CountIf(Range("A1:Z50"), "Chevy")) = 2 Then
     If Months.Columns(1).Value = "Jan" Or Months.Columns(1).Value = "Apr" Or Months.Columns(1).Value = "Jul" Or Months.Columns(1).Value = "Oct" Then
      Range1.Columns(1).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 8).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 4).Select
      ActiveWindow.ActivatePrevious
     Else
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 8).Select
      Set clear = Range(ActiveCell.Offset(0, 0), ActiveCell.Offset((LRow - 3), 0))
      clear.ClearContents
      ActiveWindow.ActivatePrevious
      Range1.Columns(1).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 4).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
     End If
     If Months.Columns(2).Value = "Feb" Or Months.Columns(2).Value = "May" Or Months.Columns(2).Value = "Aug" Or Months.Columns(2).Value = "Nov" Then
      Range1.Columns(2).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      Application.Run "PERSONAL.XLSB!PasteValue"
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 4).Select
      Set clear = Range(ActiveCell.Offset(0, 0), ActiveCell.Offset((LRow - 3), 0))
      clear.ClearContents
      ActiveWindow.ActivatePrevious
     Else
      Range1.Columns(2).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 4).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
     End If
    Else
    ActiveWindow.ActivatePrevious
    ActiveCell.Offset(0, 16).Select
    ActiveWindow.ActivatePrevious
    End If
    
       If (WorksheetFunction.CountIf(Range("A1:Z50"), "Chevy Retail")) > 0 Then
    Cells.Find(What:="Chevy Retail", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
     LRow = Cells(Rows.Count, 1).End(xlUp).row - 6
     Set Range1 = Range(ActiveCell.Offset(3, 0), ActiveCell.Offset(LRow, 2))
     Set Months = Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(1, 2))
     End If
    If (WorksheetFunction.CountIf(Range("A1:Z50"), "Chevy Retail")) = 3 Then
     Range1.Columns(1).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 8).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
     Range1.Columns(2).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 4).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
     Range1.Columns(3).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 4).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
    ElseIf (WorksheetFunction.CountIf(Range("A1:Z50"), "Chevy Retail")) = 2 Then
     If Months.Columns(1).Value = "Jan" Or Months.Columns(1).Value = "Apr" Or Months.Columns(1).Value = "Jul" Or Months.Columns(1).Value = "Oct" Then
      Range1.Columns(1).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 8).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 4).Select
      ActiveWindow.ActivatePrevious
     Else
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 8).Select
      Set clear = Range(ActiveCell.Offset(0, 0), ActiveCell.Offset((LRow - 3), 0))
      clear.ClearContents
      ActiveWindow.ActivatePrevious
      Range1.Columns(1).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 4).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
     End If
     If Months.Columns(2).Value = "Feb" Or Months.Columns(2).Value = "May" Or Months.Columns(2).Value = "Aug" Or Months.Columns(2).Value = "Nov" Then
      Range1.Columns(2).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      Application.Run "PERSONAL.XLSB!PasteValue"
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 4).Select
      Set clear = Range(ActiveCell.Offset(0, 0), ActiveCell.Offset((LRow - 3), 0))
      clear.ClearContents
      ActiveWindow.ActivatePrevious
     Else
      Range1.Columns(2).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 4).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
     End If
    Else
     ActiveWindow.ActivatePrevious
    ActiveCell.Offset(0, 16).Select
    ActiveWindow.ActivatePrevious
    End If
    
       If (WorksheetFunction.CountIf(Range("A1:AQ50"), "GMC")) > 0 Then
    Cells.Find(What:="GMC", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
     LRow = Cells(Rows.Count, 1).End(xlUp).row - 6
     Set Range1 = Range(ActiveCell.Offset(3, 0), ActiveCell.Offset(LRow, 2))
     Set Months = Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(1, 2))
     End If
    If (WorksheetFunction.CountIf(Range("A1:Z50"), "GMC")) = 3 Then
     Range1.Columns(1).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 24).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
     Range1.Columns(2).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 4).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
     Range1.Columns(3).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 4).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
    ElseIf (WorksheetFunction.CountIf(Range("A1:AQ50"), "GMC")) = 2 Then
      If Months.Columns(1).Value = "Jan" Or Months.Columns(1).Value = "Apr" Or Months.Columns(1).Value = "Jul" Or Months.Columns(1).Value = "Oct" Then
      Range1.Columns(1).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 24).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 4).Select
      ActiveWindow.ActivatePrevious
     Else
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 24).Select
      Set clear = Range(ActiveCell.Offset(0, 0), ActiveCell.Offset((LRow - 3), 0))
      clear.ClearContents
      ActiveWindow.ActivatePrevious
      Range1.Columns(1).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 4).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
     End If
     If Months.Columns(2).Value = "Feb" Or Months.Columns(2).Value = "May" Or Months.Columns(2).Value = "Aug" Or Months.Columns(2).Value = "Nov" Then
      Range1.Columns(2).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      Application.Run "PERSONAL.XLSB!PasteValue"
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 4).Select
      Set clear = Range(ActiveCell.Offset(0, 0), ActiveCell.Offset((LRow - 3), 0))
      clear.ClearContents
      ActiveWindow.ActivatePrevious
     Else
      Range1.Columns(2).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 4).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
     End If
    Else
    ActiveWindow.ActivatePrevious
    ActiveCell.Offset(0, 32).Select
    ActiveWindow.ActivatePrevious
    End If
    
    
       If (WorksheetFunction.CountIf(Range("A1:Z50"), "Onstar")) > 0 Then
    Cells.Find(What:="Onstar", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
     LRow = Cells(Rows.Count, 1).End(xlUp).row - 6
     Set Range1 = Range(ActiveCell.Offset(3, 0), ActiveCell.Offset(LRow, 2))
     Set Months = Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(1, 2))
     End If
    If (WorksheetFunction.CountIf(Range("A1:Z50"), "Onstar")) = 3 Then
     Range1.Columns(1).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 8).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
     Range1.Columns(2).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 4).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
     Range1.Columns(3).Select
     Application.Run "PERSONAL.XLSB!CopySwitch"
     ActiveCell.Offset(0, 4).Select
     Application.Run "PERSONAL.XLSB!PasteValue"
     ActiveWindow.ActivatePrevious
    ElseIf (WorksheetFunction.CountIf(Range("A1:Z50"), "Onstar")) = 2 Then
     If Months.Columns(1).Value = "Jan" Or Months.Columns(1).Value = "Apr" Or Months.Columns(1).Value = "Jul" Or Months.Columns(1).Value = "Oct" Then
      Range1.Columns(1).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 8).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 4).Select
      ActiveWindow.ActivatePrevious
     Else
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 8).Select
      Set clear = Range(ActiveCell.Offset(0, 0), ActiveCell.Offset((LRow - 3), 0))
      clear.ClearContents
      ActiveWindow.ActivatePrevious
      Range1.Columns(1).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 4).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
     End If
     If Months.Columns(2).Value = "Feb" Or Months.Columns(2).Value = "May" Or Months.Columns(2).Value = "Aug" Or Months.Columns(2).Value = "Nov" Then
      Range1.Columns(2).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      Application.Run "PERSONAL.XLSB!PasteValue"
      ActiveWindow.ActivatePrevious
      ActiveCell.Offset(0, 4).Select
      Set clear = Range(ActiveCell.Offset(0, 0), ActiveCell.Offset((LRow - 3), 0))
      clear.ClearContents
      ActiveWindow.ActivatePrevious
     Else
      Range1.Columns(2).Select
      Application.Run "PERSONAL.XLSB!CopySwitch"
      ActiveCell.Offset(0, 4).Select
      Application.Run "PERSONAL.XLSB!PasteValue"
      ActiveWindow.ActivatePrevious
     End If
    Else
    ActiveWindow.ActivatePrevious
    ActiveCell.Offset(0, 16).Select
    End If
    

    Application.ScreenUpdating = True
    Range("AA" & ActiveCell.row).Select
    Set Zero1 = Range(ActiveCell.Offset(0, 0), ActiveCell.Offset((LRow - 3), 120))
     Zero1.Replace What:="", Replacement:="0", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Else
MsgBox "Please select correct column"
End If
     
   
End Sub




