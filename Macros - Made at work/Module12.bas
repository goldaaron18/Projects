Attribute VB_Name = "Module12"

'Use/Desription: Similar to "Brandupdate3" but for Digital. Updates Productivity recap brand allocations with data from pivot table.
'Does not work with Digital GRPs since there are no GRPs for digital via DDS pull
'
'Features:
' -Similar to "Brandupdate3"
' -Loops through the filter in pivot table for each brand and places values in their coresponding cells in recap
' -Searches for name of networks so they do not need to be in the same order as the pivot. Messagebox with network name apears if not found
'
'Instructions: Similar to "Brandupdate3" but networks do not need to be in same order
'
' Improvements: Runs slow, figure out how it can run faster.
' may want to use named ranges in recap to go to cell ocations instead of searching for the word
' Consider changing the pivot so the brand is the colomn of pivot and not in the filter, and using macro "brandupdate3"
' This takes about 25 seconds but other one takes 5 seconds
' Use methods here to cut down repetitive code of brandupdate3"

  
Sub Brandupdate3b()
Attribute Brandupdate3b.VB_ProcData.VB_Invoke_Func = "n\n14"

Dim Column1 As String
Dim First As Range
Dim Range1 As Range
Dim Zero1 As Range
Dim Finalcell As Range

Application.ScreenUpdating = False

If ActiveCell.Column = 2 Then
    Set First = ActiveCell
    Set Range1 = Range(First, First.End(xlDown))
    ActiveWindow.ActivatePrevious
    
    workitems = Array("BUN-AA", "CAX-AQ", "CNF-BG", "CVN-BW", "GMN-DS", "XCD-EY")
    For Each workitem In workitems
    nm = Split(workitem, "-")(0)
    clm = Split(workitem, "-")(1)
      For Each pvtitem In ActiveSheet.PivotTables("PivotTable1").PivotFields("Client Code").PivotItems
        If pvtitem.name = nm Then
            ActiveSheet.PivotTables("PivotTable1").PivotFields("Client Code").CurrentPage _
            = nm
            Column1 = clm
            Call Brand(Range1, Column1)
            Exit For
        End If
      Next
    Next

   ActiveWindow.ActivatePrevious
   Set Zero1 = Range(First.Offset(0, 0), First.Offset((Range1.Count - 1), 164))
     Zero1.Replace What:="", Replacement:="0", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
   Application.ScreenUpdating = True
   Range(First.Offset(0, 25), First.Offset(0, 25)).Activate


Else
MsgBox "Please select correct column"
End If
     
   
End Sub


Sub Brand(ByRef Range1 As Range, ByVal Column1 As String)
    
    Dim Range2 As Range
    Dim LRow As Long
    Dim Count As Long
    Dim net As String
    Dim Line As Range
    Dim Netrng As Range

    ActiveSheet.Range("B5").Activate
    LRow = Cells(Rows.Count, 1).End(xlUp).row - 6
    Set Range2 = Range(ActiveCell.Offset(2, -1), ActiveCell.Offset(LRow, -1))
    Set Months = Range(ActiveCell.Offset(0, 0), ActiveCell.Offset(0, 2))
     
If Months.Columns(3).Value <> "Grand Total" Then
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
       Set Netrng = Range(Column1 & Netrng.row)
       Netrng.PasteSpecial Paste:=xlPasteValues
       ActiveWindow.ActivatePrevious
       Line.Offset(0, 2).Copy
       ActiveWindow.ActivatePrevious
       Netrng.Offset(0, 4).PasteSpecial Paste:=xlPasteValues
       ActiveWindow.ActivatePrevious
       Line.Offset(0, 3).Copy
       ActiveWindow.ActivatePrevious
       Netrng.Offset(0, 8).PasteSpecial Paste:=xlPasteValues
      Else
      MsgBox net & " not Found"
      End If
      ActiveWindow.ActivatePrevious
      Count = Count + 1
      Wend
     
ElseIf Months.Columns(3).Value = "Grand Total" Then
    If (Months.Columns(1).Value = "Oct") Or (Months.Columns(1).Value = "Jan") Or (Months.Columns(1).Value = "Apr") Or (Months.Columns(1).Value = "Jul") Then
       If (Months.Columns(2).Value = "Nov") Or (Months.Columns(2).Value = "Feb") Or (Months.Columns(2).Value = "May") Or (Months.Columns(2).Value = "Aug") Then
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
        Set Netrng = Range(Column1 & Netrng.row)
        Netrng.PasteSpecial Paste:=xlPasteValues
        ActiveWindow.ActivatePrevious
        Line.Offset(0, 2).Copy
        ActiveWindow.ActivatePrevious
        Netrng.Offset(0, 4).PasteSpecial Paste:=xlPasteValues
        Netrng.Offset(0, 8).Value = 0
       Else
       MsgBox net & " not Found"
       End If
       ActiveWindow.ActivatePrevious
       Count = Count + 1
       Wend
       ElseIf (Months.Columns(2).Value = "Dec") Or (Months.Columns(2).Value = "Mar") Or (Months.Columns(2).Value = "Jun") Or (Months.Columns(2).Value = "Sep") Then
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
        Set Netrng = Range(Column1 & Netrng.row)
        Netrng.PasteSpecial Paste:=xlPasteValues
        Netrng.Offset(0, 4).Value = 0
        ActiveWindow.ActivatePrevious
        Line.Offset(0, 2).Copy
        ActiveWindow.ActivatePrevious
        Netrng.Offset(0, 8).PasteSpecial Paste:=xlPasteValues
       Else
       MsgBox net & " not Found"
       End If
       ActiveWindow.ActivatePrevious
       Count = Count + 1
       Wend
       End If
     Else
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
        Set Netrng = Range(Column1 & Netrng.row)
        Netrng.Offset(0, 4).PasteSpecial Paste:=xlPasteValues
        Netrng.Value = 0
        ActiveWindow.ActivatePrevious
        Line.Offset(0, 2).Copy
        ActiveWindow.ActivatePrevious
        Netrng.Offset(0, 8).PasteSpecial Paste:=xlPasteValues
       Else
       MsgBox net & " not Found"
       End If
       ActiveWindow.ActivatePrevious
       Count = Count + 1
       Wend
       End If
  
ElseIf Months.Columns(2).Value = "Grand Total" Then
End If
       
       
     
End Sub





