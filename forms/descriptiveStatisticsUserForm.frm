VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} descriptiveStatisticsUserForm 
   Caption         =   "Descriptive Statistics"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "descriptiveStatisticsUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "descriptiveStatisticsUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub copyButton_Click()
    ''' Copies the text of a descriptive statistics user form to the clipboard. Runs automatically when the
    ''' form's "Copy" button is clicked.



    Dim DataObj As New MSForms.DataObject
    
    DataObj.SetText Me.statisticsTextBox.Value
    DataObj.PutInClipboard
    Me.copyButton.Caption = "Copied!"

End Sub

Private Sub okButton_Click()
    ''' Closes and unloads the current descriptive statistics user form. Runs automatically when the form's
    ''' "OK" button is clicked.



    Unload Me
    
End Sub
