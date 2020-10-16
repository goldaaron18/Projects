Attribute VB_Name = "Module16"

'Description/use: The Productivity Recap file network names for a given daypart may not be in the same order as the pivot table whith new data.
'this macro makes the recap order the same as the pivot table
'
'Instructions: User makes sure desired networks are visible in pivot, then clicks on first network in daypart and run the macro
'
'Improvements: Change "ActiveSheet.Range("B5").Activate" to finding location of a word in the sheet so it doesnt matter what position
'the pivot table is(ie if the top of the pivot is in row 2 or 5), currently only works if pivot is in specific position but is an easy fix.


Sub Reorder()
Attribute Reorder.VB_ProcData.VB_Invoke_Func = " \n14"


Dim First As String
Dim Range1 As Range
Dim Range2 As Range
Dim LRow As Long
Dim Count As Long
Dim net As String
Dim Line As Range
Dim Netrng As Range
Dim Swith As Range

    If ActiveCell.Column = 2 Then
    First = ActiveCell.Address
    ActiveWindow.ActivatePrevious
    
    ActiveSheet.Range("B5").Activate
    LRow = Cells(Rows.Count, 1).End(xlUp).row - 6
    Set Range2 = Range(ActiveCell.Offset(3, -1), ActiveCell.Offset(LRow, -1))

    
    Count = 1
      While Count <= Range2.Count
      Set Line = Range2.Rows(Count)
      net = Line.Value

      
      ActiveWindow.ActivatePrevious
      
      Set Range1 = Range(First, Range(First).End(xlDown))
      Range1.Activate
      Set Netrng = Range1.Find(What:=net, After:=ActiveCell, LookIn:=xlValues, _
      LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
      MatchCase:=False, SearchFormat:=False)
      
      
      If Not Netrng Is Nothing And Netrng.row <> Range1.Rows(Count).row Then
       Rows(Netrng.row & ":" & Netrng.row).Select
       Selection.Cut
       Rows(Range1.Rows(Count).row & ":" & Range1.Rows(Count).row).Select
       Selection.Insert Shift:=xlDown
      End If
      ActiveWindow.ActivatePrevious
      Count = Count + 1
      Wend
 Else
MsgBox "Please select correct column"
End If
        
    

End Sub



'Description/use: The macro below shows that we can execute different actions based on the name of the active workbook.
' we can run different macros depending on the name of the file. This could make it easier to run macros in excel since
' we can just use one keyboard key.

Sub AnyUpdate()

Dim Filename As String
Dim Value As Long

Filename = Application.ActiveWorkbook.Path

  If InStr(Filename, "Index Benchmarks") <> 0 Then
  Application.Run "PERSONAL.XLSB!IndexUpdate"
  Else
  If InStr(Filename, "Planned V Actuals") <> 0 Then
  MsgBox "PVA"
  Else
  MsgBox "Recap"
  End If
  End If

   
End Sub




'Description/use: Macro below assigns a named range to a cell. It names it based on the row and column.
'"namedrangeall" uses method "namedrange" to name many ranges.
' Features:
' -Gets rid of special charactors and adds "_" in spaces
' -Assignes "Gross" or "Net" depending on the column
' -"namedrangeall" can name a secpfic amount of rows in a given column

'Instructions: user selects last row that he wants to assign a named range to in a given column and runs macro (works for any
'column but here we want Net or Gross "Post Options")
'
'Improvements: the macro was made using the named range options menu but we are probably better off just using the cell adress box
' to assign the cell a named range. We didnt do this before because we wanted to have named ranges be tab specific, but we are better of
' putting the name of the tab in front of the named range names.(we wanted tab specific names b4 because some networks are the name for
' different dayparts)
' We also would want to create a new method to make named ranges in the bible. There will be a larger range with a few sizes depending on
' the daypart

Sub NamerangeAll()

Application.ScreenUpdating = True

LastRow = ActiveCell.row

ActiveCell.Offset(-LastRow + 6, 0).Select


Count = 1
While Count < LastRow - 1
Application.Run "PERSONAL.XLSB!NameRange"
Count = Count + 1
Selection.Offset(1, 0).Select
Wend

End Sub

Sub NameRange()
  
    Dim name As String
    Column1 = ActiveCell.Column
    
    If IsEmpty(Range("B" & ActiveCell.row).Value) = False Then

    rng1 = Range(ActiveCell.Address).Address(ReferenceStyle:=xlR1C1)
    name1 = Range("B" & ActiveCell.row).Value
    name1 = Replace(name1, "-", "")
    name1 = Replace(name1, "+", "")
    name1 = Replace(name1, ",", "")
    name1 = Replace(name1, "'", "")
    name1 = Replace(name1, "(", "")
    name1 = Replace(name1, ")", "")
    name1 = Replace(name1, "*", "")
    name1 = Replace(name1, "&", "")
    name1 = Replace(name1, " ", "_")
    name1 = Replace(name1, "/", "_")
    name1 = Replace(name1, "__", "_")
    name1 = Replace(name1, "E!", "Ent")
    name1 = Replace(name1, "@", "at")
    
    If Column1 > 27 Then
    name1 = "Net_" + name1
    Else
    name1 = "Gross_" + name1
    End If
    
    charactor1 = Left(name1, 1)
    If IsNumeric(charactor1) = False Then
    ActiveWorkbook.Worksheets(ActiveSheet.name).Names.Add name:=name1, _
        RefersToR1C1:="='" + ActiveSheet.name + "'!" + rng1
    ActiveWorkbook.Worksheets(ActiveSheet.name).Names(name1).Comment = ""
    End If
    
    End If
   
End Sub




