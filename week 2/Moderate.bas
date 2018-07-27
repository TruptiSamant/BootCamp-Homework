Attribute VB_Name = "Moderate"

'Unit 2 | Assignment - The VBA of Wall Street
'Use VBA scripting to analyze real stock market data


'
'Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
'You will also need to display the ticker symbol to coincide with the total volume.
'


Private Sub Workbook_Open()
    ActiveSheet.Shapes("FinalReport").ControlFormat.Value = xlOn
End Sub
Public Sub WallStreetModerate(Optional isHard As Boolean = False)
    
    Dim LastRow As Double
    Dim sh As Worksheet

    
    ' Set WS_Count equal to the number of worksheets in the active
    ' workbook.
     
     For Each sh In ActiveWorkbook.Worksheets
        'Set the Active worksheet
        sh.Activate
        'Calculate Volume pet ticker
        CalculateYearlyChange (isHard)

     Next sh
     
    Set sh = ActiveWorkbook.Worksheets(1)
    sh.Select
    Range("J1").Select
 
End Sub


'Calculate volume per sheet
Public Sub CalculateYearlyChange(Optional isHard As Boolean = False)

     Dim ticketrow As Integer
     Dim volume As Double
     Dim k As Double
     Dim openPrice As Double
     Dim closePrice As Double
     Dim yearlyChange As Double
     ticketrow = 2
     volume = 0
     k = 2

     'Get the row count to Ticker which is A row
     LastRow = Cells(Rows.Count, "A").End(xlUp).Row
     
      'Store opening price
    openPrice = CDbl(Cells(k, "C").Value)
    
     'Add Header
     AddHeader
        
    For k = 2 To LastRow
        If Cells(k, 1).Value <> Cells(k + 1, 1).Value Then
            Cells(ticketrow, "J").Value = Cells(k, 1).Value
            'add volume for current cell
            Cells(ticketrow, "M").Value = volume + Cells(k, "G")
            
            closePrice = CDbl(Cells(k, "F").Value)
            'Add Yearly change
            yearlyChange = closePrice - openPrice
            Cells(ticketrow, "K").Value = yearlyChange

            'Add Percent change
            If Round(openPrice, 0) <> 0 Then
                 Cells(ticketrow, "L").Value = yearlyChange / openPrice
            Else
                Cells(ticketrow, "L").Value = 1
            End If
            'format the cell
            Range("L" & ticketrow).Select
            Selection.NumberFormat = "0.00%"
            
            'Set the open price from new stock
            openPrice = Cells(k + 1, "C").Value
                  
            'Set the next row
            ticketrow = ticketrow + 1
            'reset the volume
            volume = 0
        Else
            volume = volume + Cells(k, "G")
        End If
            
    Next k
    
    If isHard Then
        AddHeaderForHard
        FindMax
        FindMin
        FindMaxVolume
        Columns("O:Q").EntireColumn.AutoFit
        Range("O2").Select
    End If
    
    'Create table around it and set the values to bold
     CreateTable
    ''Format the cell using conditinal statement
    Format
End Sub

Private Sub CreateTable()
'
' CreateTable Macro
'
    Dim LastRow As Double
    
    LastRow = Cells(Rows.Count, "J").End(xlUp).Row
    Range(Cells(1, "J"), Cells(LastRow, "M")).Select

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
'
' AddHeader on each sheet
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Columns("J:J").EntireColumn.AutoFit
    
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Yearly Change"
    Columns("K:K").EntireColumn.AutoFit
    
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Percent Change"
    Columns("L:L").Select
    Columns("L:L").EntireColumn.AutoFit
    
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Total Stock Volume"
    Columns("M:M").EntireColumn.AutoFit
    
    Range("J1:M1").Select
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
        .Bold = True
    End With
    Range("J1").Select
End Sub

'Cells greater the 0 mark them green
'Cells less then 0 mark them red
Private Sub Format()

    Dim LastRow As Double
    
    LastRow = Cells(Rows.Count, "K").End(xlUp).Row
    Range("K2:K" & LastRow).Select
'
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 192
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("K3").Select
End Sub

'find Max Increase
Private Sub FindMax()

    Dim maxstock As Double
    Dim fnd As Range
    Dim rng As Range
    Dim mx As Double
    
    Set rng = Range("L:L")

    'get maximum value in range
    Set AddressOfMax = rng.Cells(WorksheetFunction.Match(WorksheetFunction.max(rng), rng, 0))

    'return address of first matched cell
    If Not AddressOfMax Is Nothing Then
        maxRow = AddressOfMax.Row
        Range("O2").Value = "Greatest % Increase"

        Range("P2").Value = Range("J" & maxRow).Value
        
        Range("Q2").NumberFormat = "0.00%"
        Range("Q2").Value = Range("L" & maxRow).Value
        
    End If

End Sub

'find Min Increase
Private Sub FindMin()

    Dim maxstock As Double
    Dim fnd As Range
    Dim rng As Range
    Dim mx As Double
    
    Set rng = Range("L:L")

    'get maximum value in range
    Set AddressOfMax = rng.Cells(WorksheetFunction.Match(WorksheetFunction.Min(rng), rng, 0))

    'return address of first matched cell
    If Not AddressOfMax Is Nothing Then
        maxRow = AddressOfMax.Row
        Range("O3").Value = "Greatest % Decrease"

        Range("P3").Value = Range("J" & maxRow).Value
        
        Range("Q3").NumberFormat = "0.00%"
        Range("Q3").Value = Range("L" & maxRow).Value
       
    End If

End Sub

'find Max Volume
Private Sub FindMaxVolume()

    Dim maxstock As Double
    Dim fnd As Range
    Dim rng As Range
    Dim mx As Double
    
    Set rng = Range("M:M")

    'get maximum value in range
    Set AddressOfMax = rng.Cells(WorksheetFunction.Match(WorksheetFunction.max(rng), rng, 0))

    'return address of first matched cell
    If Not AddressOfMax Is Nothing Then
        maxRow = AddressOfMax.Row
        Range("O4").Value = "Greatest Total Volume"

        Range("P4").Value = Range("J" & maxRow).Value
        
        Range("Q4").Value = Range("M" & maxRow).Value
        
    End If

End Sub
'Add Header for the table and make it blue
Private Sub AddHeaderForHard()
'
' AddHeader on each sheet
'
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Columns("P:P").EntireColumn.AutoFit
    
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Value"
    Columns("Q:Q").EntireColumn.AutoFit
    
    Range("P1:Q1").Select
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
Public Sub ClearAllData()
    Dim LastRow As Integer
    Dim sh As Worksheet
    
     WS_Count = ActiveWorkbook.Worksheets.Count
     For i = 1 To WS_Count
         'Get the Active worksheet
        Set sh = ActiveWorkbook.Worksheets(i)
        sh.Activate
        
        LastRow = Cells(Rows.Count, "J").End(xlUp).Row
        
        'Range(Cells(1, "J"), Cells(LastRow, "K")).Select
        Range(Cells(1, "J"), Cells(LastRow, "Q")).Clear
     Next i
    
    Set sh = ActiveWorkbook.Worksheets(1)
    sh.Select
    
End Sub



