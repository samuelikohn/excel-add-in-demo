VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} addEmployeesUserForm 
   Caption         =   "Add Employees"
   ClientHeight    =   8295.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5640
   OleObjectBlob   =   "addEmployeesUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "addEmployeesUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addEmployeeButton_Click()
    ''' Adds extra "name" and "email" fields to an employees user form. If the field extend off the user
    '' form area, the scroll area of the user form is increased.
    
    
    
    Dim i As Integer: i = (Me.Controls.Count - 1) / 4
    
    ' Increase height of page
    If i >= 4 Then
        Me.ScrollHeight = Me.ScrollHeight + 108
        Me.addEmployeeButton.Top = Me.addEmployeeButton.Top + 108
        Me.addToSheetButton.Top = Me.addToSheetButton.Top + 108
    End If
    
    ' Add "Name" label
    With Me.Controls.Add("Forms.Label.1")
        .Name = "nameLabel" & i
        .Left = 18
        .Top = 108 * i - 36
        .Width = 60
        .Height = 24
        .Caption = "Name"
        .TextAlign = 2
        .Font.Size = 12
    End With
    
    ' Add "Email" label
    With Me.Controls.Add("Forms.Label.1")
        .Name = "emailLabel" & i
        .Left = 18
        .Top = 108 * i
        .Width = 60
        .Height = 24
        .Caption = "Email"
        .TextAlign = 2
        .Font.Size = 12
    End With
    
    ' Add "Name" text box
    With Me.Controls.Add("Forms.TextBox.1")
        .Name = "nameTextBox" & i
        .Left = 96
        .Top = 108 * i - 36
        .Width = 156
        .Height = 24
        .Font.Size = 12
    End With
    
    ' Add "Email" text box
    With Me.Controls.Add("Forms.TextBox.1")
        .Name = "emailTextBox" & i
        .Left = 96
        .Top = 108 * i
        .Width = 156
        .Height = 24
        .Font.Size = 12
    End With

End Sub

Private Sub addToSheetButton_Click()
    ''' Adds employee data in the current user form to an "employees" data sheet. If a data sheet does not
    ''' exist, one is created.
    
    
    
    Dim destinationWorkbook As Workbook
    Set destinationWorkbook = ActiveWorkbook
    Dim destinationWorksheet As Worksheet
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet
    Dim ws As Worksheet
    Dim ctrl As control
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Search for empty employees template sheet
    For Each ws In destinationWorkbook.Worksheets
        If ws.CodeName = "employeesDataSheet" Then
            Set destinationWorksheet = ws
            Exit For
        End If
    Next
        
    ' If no blank sheet found, import one
    If destinationWorksheet Is Nothing Then
        Set ws = employeesDataSheet
        
        ' Import selected sheet after current sheet
        ws.Copy After:=currentSheet
        
        ' Set name to "Employees"
        Set destinationWorksheet = destinationWorkbook.Sheets("Template Employees")
        destinationWorksheet.Name = "Employees"
        
        'If current sheet is blank, delete when template sheet is imported
        If WorksheetFunction.CountA(currentSheet.UsedRange) = 0 And currentSheet.Shapes.Count = 0 Then
            currentSheet.Delete
        End If
        
    End If
    
    ' Find first empty column
    If WorksheetFunction.CountA(destinationWorksheet.UsedRange) = 0 Then
        columnIndex = 1
    Else
        columnIndex = destinationWorksheet.UsedRange.Columns.Count + 1
    End If

    ' Add right border
    With destinationWorksheet.Columns(columnIndex + 1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    ' Merge, center, and bold company name
    With destinationWorksheet.Range(Cells(1, columnIndex), Cells(1, columnIndex + 1))
        .HorizontalAlignment = xlCenter
        .Merge
        .Font.Bold = True
    End With
    
    ' Add bottom border and center "Name" and "Email" titles
    With destinationWorksheet.Range(Cells(2, columnIndex), Cells(2, columnIndex + 1))
        .HorizontalAlignment = xlCenter
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    ' Populate company name
    destinationWorksheet.Cells(1, columnIndex) = Me.companyTextBox.Text
    
    ' Populate "Name" and "Email" titles
    destinationWorksheet.Cells(2, columnIndex) = "Name"
    destinationWorksheet.Cells(2, columnIndex + 1) = "Email"
    
    ' Populate employee data
    For Each ctrl In Me.Controls
        If Left(ctrl.Name, 11) = "nameTextBox" Then
            destinationWorksheet.Cells(Right(ctrl.Name, Len(ctrl.Name) - 11) + 2, columnIndex) = ctrl.Text
        ElseIf Left(ctrl.Name, 12) = "emailTextBox" Then
            destinationWorksheet.Cells(Right(ctrl.Name, Len(ctrl.Name) - 12) + 2, columnIndex + 1) = ctrl.Text
        End If
    Next
    
    ' Autofit columns
    destinationWorksheet.Columns(columnIndex).autoFit
    destinationWorksheet.Columns(columnIndex + 1).autoFit
    
    ' Remove UserForm
    Unload Me
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub
