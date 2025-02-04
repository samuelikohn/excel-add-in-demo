VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} viewUDFsUserForm 
   Caption         =   "View UDF Documentation"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3840
   OleObjectBlob   =   "viewUDFsUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "viewUDFsUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ''' Clears the list box of a "view UDF Documentation" user form then populates it with values
    ''' representing the available UDFs. Runs automatically upon invoking the "View UDF Documentation"
    ''' routine.



    ' Clear list box
    Me.viewUDFListBox.Clear
    
    ' Populate list box with supported UDFs
    With Me.viewUDFListBox
        .AddItem "CORRELMATRIX"
    End With

End Sub

Private Sub viewUDFButton_Click()
    ''' Displays the documentation for the selected UDF. If no UDF is selected from the list box, a message
    ''' saying so is displayed.
    
    
    
    Dim docText As String
    
    ' Select UDF from list box value
    Select Case Me.viewUDFListBox.Value
        Case "CORRELMATRIX"
        
            docText = _
                "Calculates a correlation matrix for series of data using Excel's ""CORREL"" worksheet function." & vbCrLf & vbCrLf & vbCrLf & _
                "ARGS:" & vbCrLf & vbTab & _
                    "inputRange (Range): Contiguous range of cells containing the data series. Each column represents a single series." & vbCrLf & vbCrLf & vbCrLf & _
                "RETURNS:" & vbCrLf & vbTab & _
                    "N by N matrix of numbers between 0 and 1 representing correlation coefficients, where N is the number of series in inputRange." & vbCrLf & vbCrLf & vbCrLf & _
                "ERRORS:" & vbCrLf & vbTab & _
                    "Raises a #VALUE! error if any cells in inputRange contain non-numeric values." & vbCrLf & vbTab & _
                    "Raises a #VALUE! error if any data series only contains 1 value." & vbCrLf & vbTab & _
                    "Raises a #SPILL! error if the returned matrix would overflow into occupied cells." & vbCrLf & vbTab & _
                    "Raises a #VALUE! error if the CORREL function returns an error."
            
            docs = MsgBox(docText, Title:="CORRELMATRIX Documentation")
            
        ' If no list value was selected
        Case Else
            dummy = MsgBox("Select a UDF to view documentation.", Title:="No UDF Selected")
            Exit Sub
            
    End Select
    
    'Remove UserForm
    Unload Me
    
End Sub
