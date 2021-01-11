Attribute VB_Name = "Module1"
Option Explicit
Sub StockEvaluation()

Dim Current As Worksheet
Dim starting_ws As Worksheet

'Set the current worksheet as active
Set starting_ws = ActiveSheet

'Turn off screen updating and automatic calculations
Application.Calculation = xlManual
Application.ScreenUpdating = False

'Apply following code to one sheet at a time
For Each Current In ThisWorkbook.Worksheets
    Current.Activate

'Declare Variables
Dim DataRange As Range
Dim NumRows As Long
Dim AnswerRange As Range
Dim Answer_NumRows As Long
Dim TickerValue As String
Dim Year As String
Dim FirstDate As String
Dim concat As String
Dim value As String
Dim LastDate As String
Dim SummaryTableRow As Integer
Dim Open_Value As Double
Dim Close_Value As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim TotalStockVolume As Double
Dim VolumeRow As Integer
Dim MaxPercentDecrease As Double
Dim MaxPercentIncrease As Double
Dim MaxVolIncrease As Double

'Sort Data based on Ticker value and Date value
Range("A1").CurrentRegion.Sort Key1:=Range("A1"), Order1:=xlAscending, Key2:=Range("B1"), Order2:=xlAscending, Header:=xlYes

'Create titles for outputs
Range("J1").value = "Ticker"
Range("K1").value = "Open Value"
Range("L1").value = "Close Value"
Range("M1").value = "Yearly Change"
Range("N1").value = "Percent Change"
Range("O1").value = "Total Stock Value"
Cells(2, 18).value = "Greatest % Increase"
Cells(3, 18).value = "Greatest % Decrease"
Cells(4, 18).value = "Greatest Total Volume"
Cells(1, 19).value = "Ticker"
Cells(1, 20).value = "Value"

'Calculate the number of rows of data in the dataset
Set DataRange = Range("A1").CurrentRegion
NumRows = DataRange.Rows.Count

'Keep track of each unique ticker value
SummaryTableRow = 2

'Set First Date and Last Date values
FirstDate = Cells(2, 2).value
Dim b As String
b = "B"
concat = b & NumRows
value = Range(concat).value
LastDate = value

'Set Year variable
Year = Left(Cells(2, 2).value, 4)

'Finding Ticker Value and associated Open Value and Close Value
Dim I As Long
For I = 2 To NumRows
    
    'Check if the current record is within the same year
    If Left(Cells(I, 2).value, 4) = Year Then
    
        'Check if the ticker value is not the same as that of the next record
        If Cells(I + 1, 1).value <> Cells(I, 1).value Then
    
              'Assign Ticker and Print to summary table
              TickerValue = Cells(I, 1).value
              Range("J" & SummaryTableRow).value = TickerValue
              
              'Add a row to summary table
              SummaryTableRow = SummaryTableRow + 1
              
              'Reassign LastDate and FirstDate values
              LastDate = Cells(I, 2).value
              FirstDate = Cells(I + 1, 2).value
              
              'Assign Open Value and Close Value
              Open_Value = Cells(I + 1, 3).value
              Close_Value = Cells(I, 6).value
              
              'Print Open Value and Close Value
              Range("K" & SummaryTableRow).value = Open_Value
              Range("L" & SummaryTableRow - 1).value = Close_Value
              
           'If current ticker value is the same as that of the next row, review date values
            ElseIf Cells(I, 2).value <= FirstDate Then
            
                'Assign First_Date and Open_Value
                FirstDate = Cells(I, 2).value
                Open_Value = Cells(I, 3).value
                
                'Print Open Value to summary table
                Range("K" & SummaryTableRow).value = Open_Value
                
        End If
    End If
Next I
    
'Set the Total Stock Volume to zero
TotalStockVolume = 0

'Keep track of stock volume record for each unique ticker
VolumeRow = 2

'Calculating Total Stock Volume for each unique ticker value
Dim c As Long
For c = 2 To NumRows
        
    'Confirm that record is within same year
    If Left(Cells(c, 2).value, 4) = Year Then
            
        'Check if the Ticker value of current record does not match the next record
        If Cells(c + 1, 1).value <> Cells(c, 1).value Then
            
              'Add stock volume of current record to Total Stock Volume
              TotalStockVolume = TotalStockVolume + Cells(c, 7).value
              
              'Print Total Stock Volume
              Range("O" & VolumeRow).value = TotalStockVolume
              
              'Add one to the Volume Row
              VolumeRow = VolumeRow + 1
            
              'Reset Total Stock Volume for next Ticker
              TotalStockVolume = 0
              
            Else
                'Add stock volume of current record to Total
                TotalStockVolume = TotalStockVolume + Cells(c, 7).value
                
                'Print Total Stock Volume
                Range("O" & VolumeRow).value = TotalStockVolume
                
        End If
    End If
Next c

'Calculate number of unique Ticker Values
Set AnswerRange = Range("J1").CurrentRegion
Answer_NumRows = AnswerRange.Rows.Count

'Delete contents of cell in summary table that does not have an associated Ticker value
Range("K" & Answer_NumRows).Clear

'Calculating Yearly Change and Percent Change for each Ticker Value
Dim a As Integer
For a = 2 To Answer_NumRows
    
    'Check to make sure record has an associated unique Ticker value
    If Cells(a, 10).value <> "" Then
                
            'Set Open Value and Close Value for each Ticker Value
            Open_Value = Cells(a, 11).value
            Close_Value = Cells(a, 12).value
            
            'Calculate Yearly Change and find the greatest total volume
            Yearly_Change = Close_Value - Open_Value
            MaxVolIncrease = Application.WorksheetFunction.Max(Range("O2:O" & Answer_NumRows))
            
            'Print the Yearly Change value and the Greatest Total Volume
            Range("M" & a).value = Yearly_Change
            Cells(4, 20).value = MaxVolIncrease
            
            'Format Yearly Change value to include 2 decimal places
            Columns("M").NumberFormat = "0.00"
            
        'Check if Open Value is not zero
        If Open_Value <> 0 Then
        
                'Calculate Percent Change, the greatest % increase, and the greatest % decrease in stock volume
                Percent_Change = (Yearly_Change / Open_Value)
                MaxPercentIncrease = Application.WorksheetFunction.Max(Range("N2:N" & Answer_NumRows))
                MaxPercentDecrease = Application.WorksheetFunction.Min(Range("N2:N" & Answer_NumRows))
                
                'Print the Percent Change, Greatest % Increase, and Greatest % decrease values
                Range("N" & a).value = Percent_Change
                Cells(2, 20).value = MaxPercentIncrease
                Cells(3, 20).value = MaxPercentDecrease
                
                'Format % cells to present data with two decimal places
                Columns("N").NumberFormat = "0.00%"
                Range("T2:T3").NumberFormat = "0.00%"
            
            'If Open Value is 0
            ElseIf Open_Value = 0 Then
            
                'Percent Change is set to 0 and print
                Percent_Change = 0
                Range("N" & a).value = Percent_Change
                
        End If
    End If
Next a

'Finding Ticker values associated with Greatest % Increase/ Decrease and Greatest Total Volume
Dim v As Integer
For v = 2 To Answer_NumRows
    
    'Find Ticker associated with Greatest % Increase and print
    If Cells(v, 14).value = MaxPercentIncrease Then
            Cells(2, 19).value = Cells(v, 10).value
    
        'Find Ticker associated with Greatest % Decrease and print
        ElseIf Cells(v, 14).value = MaxPercentDecrease Then
            Cells(3, 19).value = Cells(v, 10).value
        
        'Find Ticker associated with Greatest Total Volume and print
        ElseIf Cells(v, 15).value = MaxVolIncrease Then
            Cells(4, 19).value = Cells(v, 10).value
            
    End If
Next v

'Formatting Yearly Change to identify positive vs. negative changes
Dim f As Integer
For f = 2 To Answer_NumRows

    'Identify cells in Yearly Change column with negative values and color Red
    If Cells(f, 13).value < 0 Then
            Range("M" & f).Interior.Color = RGB(235, 0, 0)
        
        'Identify cells in Yearly Change column with positive values and color Green
        ElseIf Cells(f, 13).value > 0 Then
            Range("M" & f).Interior.Color = RGB(0, 188, 85)
            
    End If
Next f

'Format Data within worksheet
Columns.AutoFit
Columns.VerticalAlignment = xlCenter
Rows(1).HorizontalAlignment = xlCenter
Rows(1).Font.Bold = True
Columns(18).Font.Bold = True

'Delete Open Value and Close Value Columns
Columns("K:L").EntireColumn.Delete

'Loop through next sheet and activate the original sheet once done
Next Current
starting_ws.Activate

'Turn on screen updating and automatic calculations
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True

End Sub

