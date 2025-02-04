VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} importTemplateSheetsUserForm 
   Caption         =   "Import Template Sheets"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3720
   OleObjectBlob   =   "importTemplateSheetsUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "importTemplateSheetsUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ''' Clears the list box of a template sheet user form then populates it with values representing the
    ''' available template sheets. Runs automatically upon invoking the "Import Template Sheets" routine.



    ' Clear list box
    Me.templateSheetsListBox.Clear
    
    ' Populate list box with supported template sheets
    With Me.templateSheetsListBox
        .AddItem "Tensile Test (from CSV)"
        .AddItem "Employees (from JSON)"
        .AddItem "Example Data Sheet"
    End With

End Sub

Private Sub importSheetButton_Click()
    ''' Adds the selected template sheet to the current workbook. Runs automatically upon clicking the
    ''' "Import Sheet" button of an "Import Template Sheets" user form. Available sheets include:
    '''     Tensile Test: Sheet name is set based on timestamp of import.
    '''     Employees: Only 1 employees sheet may exist in a given workbook at any time. If an employees
    '''         sheet already exists, no action is performed.
    
    
    
    Dim sheetToImport As Worksheet
    Dim ws As Worksheet
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet
    Dim currentTime As Date: currentTime = Now()
    Dim sheetName As String
    
    ' Select template sheet from list box value
    Select Case Me.templateSheetsListBox.Value
        Case "Tensile Test (from CSV)"
            Set sheetToImport = tensileTestDataSheet
            
            sheetName = _
                "TensileTest " & DatePart("yyyy", currentTime) & _
                "-" & DatePart("m", currentTime) & _
                "-" & DatePart("d", currentTime) & _
                " " & DatePart("h", currentTime) & _
                ";" & DatePart("n", currentTime) & _
                ";" & DatePart("s", currentTime)
            
        Case "Employees (from JSON)"
            Set sheetToImport = employeesDataSheet
            
            sheetName = "Employees"
            
            ' Only 1 "Employees" sheet may exist in a given workbook at any time.
            For Each ws In ActiveWorkbook.Sheets
                If ws.CodeName = "employeesDataSheet" Then
                    dummy = MsgBox("An ""Employees"" sheet already exists.", Title:="Sheet Already Exists")
                    Exit Sub
                End If
            Next
            
        Case "Example Data Sheet"
            Set sheetToImport = exampleDataSheet
            sheetName = "Example Data Sheet"
            
        ' If no list value was selected
        Case Else
            dummy = MsgBox("Select a template sheet to import.", Title:="No template sheet selected")
            Exit Sub
            
    End Select

    ' Import selected sheet after current sheet
    sheetToImport.Copy After:=currentSheet
    ActiveWorkbook.Sheets(sheetToImport.Name).Activate
    ActiveSheet.Name = sheetName
    
    'If current sheet is blank, delete when template sheet is imported
    If WorksheetFunction.CountA(currentSheet.UsedRange) = 0 And currentSheet.Shapes.Count = 0 Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        currentSheet.Delete
        
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        
    End If
    
    'Remove UserForm
    Unload Me
    
End Sub
