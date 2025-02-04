Attribute VB_Name = "dataValidation"
Sub copyDataValidation(Optional control As IRibbonControl)
    ''' Copies the data validation from a source range to a target range. The source range is required to be
    ''' a single cell and to have existing data validation. If the target range has any data validation, the
    ''' user will be asked to confirm if they want the existing data validation to be overwritten.



    Dim sourceCell As Range
    Dim targetRange As Range
    Dim validationType As Variant: validationType = Null
    On Error Resume Next
    
    ' Prompt user to select a single source cell
    Do While True
        Set sourceCell = Application.InputBox(Prompt:="Select a single cell with data validation.", Title:="Copy Data Validation", Type:=8)
        
        ' Check if the selection contains a single cell
        If sourceCell.Count = 1 Then
            Exit Do
        Else
            MsgBox ("Selection must be a single cell")
        End If
        
        ' Check if source cell is selected
        If sourceCell Is Nothing Then
            Exit Sub
        End If
    Loop
    
    ' Check if the source cell has data validation
    For Each C In sourceCell
        validationType = C.Validation.Type
        If Err.Number <> 0 Then
            MsgBox ("The selected cell does not have data validation.")
            Exit Sub
        End If
    Next C
    
    ' If the source cell has data validation, allow the user to select a range of cells to paste
    Set targetRange = Application.InputBox(Prompt:="Select a range to paste the data validation.", Title:="Paste Data Validation", Type:=8)
    
    ' Check if target range is selected
    If targetRange Is Nothing Then
        Exit Sub
    Else
    
        ' Copy data validation from source cell to selected range
        sourceCell.Copy
        targetRange.Validation.Delete
        targetRange.PasteSpecial Paste:=xlPasteValidation
        Application.CutCopyMode = False
        
    End If
    
End Sub

Sub clearDataValidation(Optional control As IRibbonControl)
    ''' Clears all data validation from the selected cells.



    If MsgBox("Are you sure you want to clear all data validation from the selected cells?", vbYesNo) = vbYes Then Selection.Validation.Delete
    
End Sub

Sub createDropDownList(Optional control As IRibbonControl)
    ''' Creates data validation for a discrete list of values as a dropdown list for a target range of
    ''' cells. The values must be stored in cells and are referred to by a dynamic named range. If the
    ''' target range has existing data validation or non-conformative values, the user will be asked to
    ''' confirm if they want the existing data validation or values to be overwritten.



    Dim sourceRange As Range
    Dim targetCell As Range
    Dim rngName As String
    Dim validationType As Variant: validationType = Null
    Set targetCell = Application.Selection
    On Error Resume Next
    
    ' Select the range that contains the list of values for the drop-down
    Set sourceRange = Application.InputBox(Prompt:="Select the range that contains the list of values for the drop-down", Title:="Select Values", Type:=8)
    If Err.Number <> 0 Then
        Exit Sub ' User canceled
    End If
    
    ' Check if all cells in the range are blank
    If Application.WorksheetFunction.CountA(sourceRange) = 0 Then
        MsgBox "The selected range is empty.", vbInformation, "No Values Found"
        Exit Sub
    End If
    
    ' Create a dynamic named range to refer to the list values
    rngName = "ListValues_" & Format(Now, "yyyy_mm_dd_hh_mm_ss")
    sourceRange.Parent.Names.Add Name:=rngName, RefersTo:=sourceRange
    
    ' Check if any of the selected cells already have data validation.
    For Each C In targetCell
        validationType = C.Validation.Type
        If Err.Number = 0 Then
            If MsgBox("The selected cell(s) already have data validation. Do you want to overwrite it?", vbYesNo + vbQuestion, "Overwrite Data Validation?") = vbNo Then
                Exit Sub
            End If
            Exit For
        End If
    Next C
    
    'Check if any of the selected cells are not empty and if the values are already in the validation list
    For Each C In targetCell
        If C.Value <> "" And IsError(Application.match(C.Value, sourceRange, 0)) Then
            If MsgBox("The selected cell(s) already contain data. Do you want to overwrite it?", vbYesNo + vbQuestion, "Cell Not Empty") = vbNo Then
                Exit Sub
            Else:
                Exit For
            End If
        End If
    Next C

    ' Add the drop-down list to the selected cell using the dynamic named range
    With targetCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & rngName
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
End Sub

Sub validateNumber(Optional control As IRibbonControl)
    ''' Creates data validation for numerical values for a target range of cells. If the target range has
    ''' existing data validation or non-numerical values, the user will be asked to confirm if they want the
    ''' existing data validation or values to be overwritten.
    
    
    
    Dim targetCell As Range
    Dim validationType As Variant: validationType = Null
    On Error Resume Next
    Set targetCell = Application.Selection
    
    ' Check if any of the selected cells already have data validation.
    For Each C In targetCell
        validationType = C.Validation.Type
        If Err.Number = 0 Then
            If MsgBox("The selected cell(s) already have data validation. Do you want to overwrite it?", vbYesNo + vbQuestion, "Overwrite Data Validation?") = vbNo Then
                Exit Sub
            End If
            Exit For
        End If
    Next C
    
    ' Check if any of the selected cells are not empty and if the values are already in the validation list
    For Each C In targetCell
        If C.Value <> "" And IsError(Application.match(C.Value, rng, 0)) Then
            If MsgBox("The selected cell(s) already contain data. Do you want to overwrite it?", vbYesNo + vbQuestion, "Cell Not Empty") = vbNo Then
                Exit Sub
            Else:
                Exit For
            End If
        End If
    Next C
    
    ' Add the validation to the selected cell
    With targetCell.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlLess, Formula1:="=1.79769313486231*10^308"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
End Sub

Sub validateYesNo(Optional control As IRibbonControl)
    ''' Creates data validation for Boolean values in the form of "Yes" and "No" for a target range of
    ''' cells. If the target range has existing data validation or non-Boolean values, the user will
    ''' be asked to confirm if they want the existing data validation or values to be overwritten.



    Dim targetCell As Range
    Dim validationType As Variant: validationType = Null
    On Error Resume Next
    Set targetCell = Application.Selection
    
    ' Check if any of the selected cells already have data validation.
    For Each C In targetCell
        validationType = C.Validation.Type
        If Err.Number = 0 Then
            If MsgBox("The selected cell(s) already have data validation. Do you want to overwrite it?", vbYesNo + vbQuestion, "Overwrite Data Validation?") = vbNo Then
                Exit Sub
            End If
            Exit For
        End If
    Next C
    
    'Check if any of the selected cells are not empty and if the values are already in the validation list
    For Each C In targetCell
        If C.Value <> "" And IsError(Application.match(C.Value, rng, 0)) Then
            If MsgBox("The selected cell(s) already contain data. Do you want to overwrite it?", vbYesNo + vbQuestion, "Cell Not Empty") = vbNo Then
                Exit Sub
            Else:
                Exit For
            End If
        End If
    Next C
    
    ' Add the drop-down list to the selected cells
    With targetCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Yes,No"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
End Sub
