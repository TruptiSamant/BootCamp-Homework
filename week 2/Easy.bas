Attribute VB_Name = "Easy"

'Unit 2 | Assignment - The VBA of Wall Street
'Use VBA scripting to analyze real stock market data


'
'Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
'You will also need to display the ticker symbol to coincide with the total volume.
'

Public Sub WallStreetEasy()
    
    Dim LastRow As Double
    Dim sh As Worksheet

    
    ' Set WS_Count equal to the number of worksheets in the active
    ' workbook.
     
     For Each sh In ActiveWorkbook.Worksheets
        'Set the Active worksheet
        sh.Activate
        'Calculate Volume pet ticker
        CalculateVolume
        
        'Create table around it and set the values to bold
     Next sh
     
    Set sh = ActiveWorkbook.Worksheets(1)
    Range("J1").Select
 
End Sub


'Calculate volume per sheet
Public Sub CalculateVolume()

     Dim ticketrow As Integer
     Dim volume As Double
     Dim k As Double
    
     ticketrow = 2
     volume = 0

    'Get the row count to Ticker which is A row
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
        'Add Header
    AddHeader
        
    For k = 2 To LastRow
        If Cells(k, 1).Value <> Cells(k + 1, 1).Value Then
            Cells(ticketrow, "J").Value = Cells(k, 1).Value
            'add volume for current cell
            Cells(ticketrow, "K").Value = volume + Cells(k, "G")
                
            'Set the next row
            ticketrow = ticketrow + 1
            'reset the volume
            volume = 0
        Else
            volume = volume + Cells(k, "G")
        End If
            
    Next k
        CreateTable
    
End Sub


Private Sub CreateTable()
Attribute CreateTable.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CreateTable Macro
'
    Dim LastRow As Integer
    
    LastRow = Cells(Rows.Count, "J").End(xlUp).Row
    
    Range(Cells(1, "J"), Cells(LastRow, "K")).Select

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Font.Bold = True
    Range("J1").Select
    
End Sub

'Add Header for the table and make it blue
Private Sub AddHeader()
Attribute AddHeader.VB_ProcData.VB_Invoke_Func = " \n14"
'
' AddHeader on each sheet
'
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Volume"
    Range("J1:K1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

End Sub

'Clear the Data for all work sheets
Private Sub ClearData()
    Dim LastRow As Integer
    Dim sh As Worksheet
    
     WS_Count = ActiveWorkbook.Worksheets.Count
     For i = 1 To WS_Count
         'Get the Active worksheet
        Set sh = ActiveWorkbook.Worksheets(i)
        sh.Activate
        
        LastRow = Cells(Rows.Count, "J").End(xlUp).Row
        
        'Range(Cells(1, "J"), Cells(LastRow, "K")).Select
        Range(Cells(1, "J"), Cells(LastRow, "K")).Clear
     Next i
    
    Set sh = ActiveWorkbook.Worksheets(1)
    sh.Select
    
End Sub

