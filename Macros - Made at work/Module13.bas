Attribute VB_Name = "Module13"
'Descriptiom/ use: This macro makes it easier to update forcasted #s from the bible to the recap.
' "Forcastupdate2" uses the "forcastbrand" method.
'
'Features:
'-Updates numbers based on user selection in bible and places them in correct columns for a given network
'-Runs a different macro depending on daypart. For both macros it adds /3 in formula, and for "net" it adds *.85 to the formula
'-Takes into acount diferences in positions of numbers for different files/ dayparts. offets 1, or 2 ect.
'-Updates brand data as well as Sum and Overspend(based on range/cell adress in bible)
'
'Instructions: In bible find desired network and click on first brand column for the qaurter/row that needs updating.
'In recap click on network name in column 2 and run macro
'
'Improvements: Use name ranges with indexes in future to link instead up using macro.




Sub Forcastupdate2()


Application.ScreenUpdating = False

ActiveCell.Offset(0, 41).Select
ActiveWindow.ActivatePrevious

Application.Run "PERSONAL.XLSB!forcastBrand"
ActiveCell.Offset(0, 8).Select
ActiveWindow.ActivatePrevious
ActiveCell.Offset(0, 1).Select

Application.Run "PERSONAL.XLSB!forcastBrand"
ActiveCell.Offset(0, 8).Select
ActiveWindow.ActivatePrevious
ActiveCell.Offset(0, 1).Select

Application.Run "PERSONAL.XLSB!forcastBrand"
ActiveCell.Offset(0, 8).Select
ActiveWindow.ActivatePrevious
ActiveCell.Offset(0, 1).Select

Application.Run "PERSONAL.XLSB!forcastBrand"
ActiveCell.Offset(0, -72).Select
ActiveWindow.ActivatePrevious
ActiveCell.Offset(0, 1).Select

Application.Run "PERSONAL.XLSB!forcastBrand"
ActiveCell.Offset(0, 88).Select
ActiveWindow.ActivatePrevious
ActiveCell.Offset(0, 1).Select

Application.Run "PERSONAL.XLSB!forcastBrand"
ActiveCell.Offset(0, 8).Select
ActiveWindow.ActivatePrevious
If InStr(ActiveWorkbook.name, "SPORTS") <> 0 Then
  ActiveCell.Offset(0, 2).Select
  Else
  ActiveCell.Offset(0, 1).Select
End If

Application.Run "PERSONAL.XLSB!forcastBrand"
ActiveCell.Offset(0, 43).Select
ActiveWindow.ActivatePrevious
If InStr(ActiveWorkbook.name, "CINEMA") Or InStr(ActiveWorkbook.name, "OLV") <> 0 Then
  ActiveCell.Offset(0, 2).Select
  Else
  ActiveCell.Offset(0, 1).Select
End If

Selection.Copy
ActiveWindow.ActivatePrevious
ActiveSheet.Paste Link:=True
ActiveWindow.ActivatePrevious
ActiveCell.Offset(0, 1).Select
Selection.Copy
ActiveWindow.ActivatePrevious
ActiveCell.Offset(0, 1).Select
ActiveSheet.Paste Link:=True

    
End Sub




Sub ForcastBrand()
Attribute ForcastBrand.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro7 Macro
'

'
  If InStr(ActiveWorkbook.name, "OLV") <> 0 Then
    Selection.Copy
    ActiveWindow.ActivatePrevious
    
    ActiveSheet.Paste Link:=True
    Application.Run "PERSONAL.XLSB!forcastNet"
    Selection.Copy
    ActiveCell.Offset(0, 4).Select
    ActiveSheet.Paste Link:=True
    Selection.Copy
    ActiveCell.Offset(0, 4).Select
    ActiveSheet.Paste Link:=True
  
  Else
  Selection.Copy
    ActiveWindow.ActivatePrevious
    
    ActiveSheet.Paste Link:=True
    Application.Run "PERSONAL.XLSB!forcast"
    Selection.Copy
    ActiveCell.Offset(0, 4).Select
    ActiveSheet.Paste Link:=True
    Selection.Copy
    ActiveCell.Offset(0, 4).Select
    ActiveSheet.Paste Link:=True
  End If

    
End Sub






Sub forcast()
    
        'This macro places the entire function in each cell of a selection inside of an IFERROR function.
        'The user may enter the desired argument for value if error, as a string or a number.
            
        Dim String1 As String
        Dim String2 As String
        
        
        Application.Calculation = xlCalculationManual
        
        For Each C In Selection
            resultString = "=((" + Right(C.Formula, Len(C.Formula) - 1) + ")*0.85)/3"
            C.Formula = resultString
        Next
        
        Application.Calculation = xlCalculationAutomatic
    
    End Sub


Sub forcastNet()
    
            
        Dim String1 As String
        Dim String2 As String
        
        
        Application.Calculation = xlCalculationManual
        
        For Each C In Selection
            resultString = "=(" + Right(C.Formula, Len(C.Formula) - 1) + ")/3"
            C.Formula = resultString
        Next
        
        Application.Calculation = xlCalculationAutomatic
    
    End Sub






