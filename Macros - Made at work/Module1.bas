Attribute VB_Name = "Module1"

'Use/Desription: Simple macro that pastes value and then switches to other workbook.
'Usually the paste funtion pastes value & formating so this is usefull if just the value is wanted

Sub PasteValue()
Attribute PasteValue.VB_ProcData.VB_Invoke_Func = " \n14"

    Selection.PasteSpecial Paste:=xlPasteValues
   ActiveWindow.ActivatePrevious

End Sub

'Use/Desription: Simple macro that copys and switches to other worksheet. usefull when copying data from one sheet to another


Sub CopySwitch()
Attribute CopySwitch.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macroswitch Macro
'
    Selection.Copy
    ActiveWindow.ActivatePrevious
    
    
End Sub

'Use/Desription: Simple macro that pastes format

Sub PasteFormat()

Selection.PasteSpecial Paste:=xlPasteFormats

End Sub


'Use/Desription: Copys and Pastes link to the size of range selected in one Worksheet and pastes to active cell in other worksheet


Sub AnypasteLink()

    If Selection.Cells.Count > 1 Then
       Range(Cells(Selection.row, Selection.Column), Cells(Selection.row, Selection.Column + 1)).Select
       ActiveSheet.Paste Link:=True
    
    Else
       ActiveSheet.Paste Link:=True
    End If
    
    ActiveWindow.ActivatePrevious
    
    
End Sub

'Use/Desription:  This macro places the entire function in each cell of a selection inside of an IFERROR function.

Sub IFERROR()
Attribute IFERROR.VB_ProcData.VB_Invoke_Func = "e\n14"

            
        Dim errorString As String
        Dim resultString As String
        
        
        Application.Calculation = xlCalculationManual
        
        For Each C In Selection
            resultString = "=IFERROR(" + Right(C.Formula, Len(C.Formula) - 1) + "," + "0" + ")"
            C.Formula = resultString
        Next
        
        Application.Calculation = xlCalculationAutomatic
    
    End Sub


Sub PasteasLink()

ActiveSheet.Paste Link:=True
ActiveWindow.ActivatePrevious

End Sub

'Use/Desription:  Converts a Sum funtion ex. =Sum(A1:A10) to =Sum(A1,A2,A3.....A9,A10)
'The reason this is usefull is because sometimes we may want to add or change the order or networks in a given daypart,
'but this messes up the sum function. with new sum format it does not get effected
 
 Sub sumchange()
    Dim s As String, s2 As String, s3 As String
    Dim r As Range, rr As Range

    s = ActiveCell.Formula
    s2 = Mid(s, 6, 99)
    s2 = Left(s2, Len(s2) - 1)
    s3 = ""

    Set r = Range(s2)
    For Each rr In r
        s3 = s3 & "," & rr.Address(0, 0)
    Next rr
    ActiveCell.Formula = "=SUM(" & Mid(s3, 2) & ")"

End Sub



'Use/Desription:  Applies a formula to a pattern of cells to match active cell formula.
' note: excel find and replace feature with formulas may be a better option.

Sub Formulaupdate()

Application.ScreenUpdating = False
    
    Selection.Copy
    ActiveCell.Offset(0, 2).Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, 1).Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, 2).Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, 2).Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, 2).Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, 2).Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, 3).Select
    ActiveSheet.Paste

    
    Dim Counter
    Counter = 0
    
    While Counter < 39
    ActiveCell.Offset(0, 3).Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, 1).Select
    ActiveSheet.Paste
    Counter = Counter + 1
    Wend

    ActiveCell.Offset(0, 3).Select
    ActiveSheet.Paste
 
 End Sub
    
 Sub formulaupdate2()

Application.ScreenUpdating = False
    
    Selection.Copy
    ActiveCell.Offset(0, 2).Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, 5).Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, 2).Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, 5).Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, 6).Select
    ActiveSheet.Paste
    
    ActiveCell.Offset(0, 3).Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, 1).Select
    ActiveSheet.Paste

    
    Dim Counter
    Counter = 0
    
    While Counter < 38
    ActiveCell.Offset(0, 3).Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, 1).Select
    ActiveSheet.Paste
    Counter = Counter + 1
    Wend

    ActiveCell.Offset(0, 3).Select
    ActiveSheet.Paste
 
 End Sub




'Use/Desription: This macro helps copy and paste a named range link. Alternatively you could just type in the named range reference.
'
Sub namedrangeref()

    ActiveWindow.ActivatePrevious
    Dim Namedrange As String
    
    Namedrange = ActiveCell.name.name
    Namedrange2 = Right(Namedrange, Len(Namedrange) - 1)
    Selection.Copy
    ActiveWindow.ActivatePrevious
    ActiveSheet.Paste Link:=True
    
        
    Application.Calculation = xlCalculationManual
        
    For Each C In Selection
      endstring = InStr(C.Formula, "x]") + 1
      String1 = Left(C.Formula, endstring)
      String2 = String1 + Namedrange2
      C.Formula = String2
    Next
        
    Application.Calculation = xlCalculationAutomatic
    
    
    
End Sub



'Use/Desription: If Ctrl+[ is not working, this helps go to location of reference within a formula.
'In the case here, It was used to activate 2 refernces so that they could be updated more easy.

Sub Reference()

    str1 = Range(ActiveCell.Address).Formula
    startpos = InStr(str1, "x]") + 2
    endpos = InStr(str1, "'!")
    last = InStr(str1, ",")
    startsum = InStr(str1, "SUM")
    str2 = Mid(str1, startpos, endpos - startpos)
    
    If last <> 0 Then
    str3 = Mid(str1, endpos + 2, last - endpos - 2)
    Else
    str3 = Mid(str1, endpos + 2)
    End If
    
    ActiveWindow.ActivatePrevious
    Sheets(str2).Select
    If str2 <> "OLV-VOD" Then
    Range(str3).Offset(0, -28).Activate
    Else
    Range(str3).Activate
    End If
    ActiveWindow.ActivatePrevious

    ActiveCell.Offset(0, -12).Select

    str1b = Range(ActiveCell.Address).Formula
    startposb = InStr(str1b, "='") + 2
    endposb = InStr(str1b, "'!")
    str2b = Mid(str1b, startposb, endposb - startposb)
    str3b = Mid(str1b, endposb + 2)
    
    ActiveCell.Offset(0, 12).Select

    Sheets(str2b).Select
    If str2 <> "OLV-VOD" Then
    Range(str3b).Select
    Else
    Range(str3b).Offset(0, 1).Select
    End If


End Sub
    
    

'Use/Desription: This helps change a reference to a named range reference from its linked location.
'Improvment: Apply this to named ranges that are larger that 1 cell

Sub ref2nameOVG()

    str1 = Range(ActiveCell.Address).Formula
    startpos = InStr(str1, "x]") + 2
    endpos = InStr(str1, "'!")
    last = InStr(str1, ",")
    startsum = InStr(str1, "SUM")
    str2 = Mid(str1, startpos, endpos - startpos)
    
    If last <> 0 Then
    str3 = Mid(str1, endpos + 2, last - endpos - 2)
    Else
    str3 = Mid(str1, endpos + 2)
    End If
    
    ActiveWindow.ActivatePrevious
    Sheets(str2).Select
    Range(str3).Activate
    ActiveWindow.ActivatePrevious
    
    Application.Run "PERSONAL.XLSB!copypasteName"

    

End Sub
   


