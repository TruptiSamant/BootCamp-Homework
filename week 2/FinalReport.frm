VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FinalReport 
   Caption         =   "Final Report"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5955
   OleObjectBlob   =   "FinalReport.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FinalReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub UserForm_Click()

End Sub


'Initialize the combo box with worksheets
Private Sub UserForm_Initialize()
    Dim sht As Worksheet
    Dim txt As String
    
    'FinalReport.Show vbModeless

    For Each sht In ActiveWorkbook.Worksheets
        cboWhichSheet.AddItem sht.Name
    Next sht
    
    cboWhichSheet.AddItem "All"
    
    cboWhichSheet.ListIndex = 0

End Sub



'According the click make worksheet activate

Private Sub cboWhichSheet_Change()
    
    If Me.cboWhichSheet.Value = "All" Then
         Worksheets(1).Select
    Else
        Worksheets(Me.cboWhichSheet.Value).Select
    End If
    Range("A1").Select
    
End Sub


'Clear the Data
Private Sub cmdClearAll_Click()
    ClearAllData
End Sub

'Click the report by Sheet.  If All selected Run the report for all
'This is a Easy Homework
Private Sub cmdEasy_Click()
    If Me.cboWhichSheet.Value = "All" Then
         WallStreetEasy
    Else
        CalculateVolume
    End If
        
End Sub

'Click the report by Sheet.  If All selected Run the report for all
'This is a Moderate Homework

Private Sub cmdModerate_Click()
    If Me.cboWhichSheet.Value = "All" Then
         WallStreetModerate
    Else
        CalculateYearlyChange
    End If
End Sub

'Click the report by Sheet.  If All selected Run the report for all
'This is a Moderate + Hard Homework
Private Sub cmdHard_Click()
    If Me.cboWhichSheet.Value = "All" Then
         WallStreetModerate (True)
    Else
        CalculateYearlyChange (True)
    End If
End Sub
