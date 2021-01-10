Attribute VB_Name = "Module1"
Option Explicit
Sub StockEvaluation()

Dim Current As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

'Turn off screen updating and automatic calculations
Application.Calculation = xlManual
Application.ScreenUpdating = False

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
Dim MinPercentIncrease As Double
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
Cells(2, 17).value = "Greatest % Increase"
Cells(3, 17).value = "Greatest % Decrease"
Cells(4, 17).value = "Greatest Total Volume"
Cells(1, 18).value = "Ticker"
Cells(1, 19).value = "Value"


'Calculate the number of rows of data in the dataset
Set DataRange = Range("A1").CurrentRegion
NumRows = DataRange.Rows.Count


'Keep track of each unique ticker value
SummaryTableRow = 2


' Set First as smallest date value
FirstDate = Cells(2, 2).value


'Set Last Date as largest date value
Dim b As String
b = "B"
concat = b & NumRows
value = Range(concat).value
LastDate = value


'Assign Year variable
Year = Left(Cells(2, 2).value, 4)


'Finding next unique Ticker Value and Entering Ticker value in next available cell
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
              
              'Assign Open Value and Close Value and print to summary table
              Open_Value = Cells(I + 1, 3).value
              Range("K" & SummaryTableRow).value = Open_Value

              Close_Value = Cells(I, 6).value
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

    
' Set the Total Stock Volume to zero
TotalStockVolume = 0


'Keep track of stock volume record for each unique ticker
VolumeRow = 2


' Calculating Total Stock Volume for each unique ticker value
Dim c As Long
For c = 2 To NumRows
        
    ' Confirm that record is within same year
    If Left(Cells(c, 2).value, 4) = Year Then
            
        'Check if the Ticker value of current record does not match the next record
        If Cells(c + 1, 1).value <> Cells(c, 1).value Then
            
              ' Add stock volume of current record to Total Stock Volume
              TotalStockVolume = TotalStockVolume + Cells(c, 7).value
              
              'Print Total Stock Volume
              Range("O" & VolumeRow).value = TotalStockVolume
              
              'Add one to the Volume Row
              VolumeRow = VolumeRow + 1
            
              ' Reset Total Stock Volume for next Ticker
              TotalStockVolume = 0
              
            'If the current ticker value matches the next record
            Else
                    
                ' Add stock volume of current record to Total
                TotalStockVolume = TotalStockVolume + Cells(c, 7).value
                
                ' Print Total Stock Volume
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
    
    'Check to make sure a ticker value exists
    If Cells(a, 10).value <> "" Then
                
            'Assign Open Value and Close Value for each Ticker Value
            Open_Value = Cells(a, 11).value
            Close_Value = Cells(a, 12).value
            
            'Calculate Yearly Change
            Yearly_Change = Close_Value - Open_Value
            
            'Print the Yearly Change value
            Range("M" & a).value = Yearly_Change
            
            'Format Yearly Change value to include 2 decimal places
            Columns("M").NumberFormat = "0.00"
            
            'Find the greatest total volume
            MaxVolIncrease = Application.WorksheetFunction.Max(Range("O2:O" & Answer_NumRows))
            
            'Print Greatest Total Volume
            Cells(4, 19).value = MaxVolIncrease
            
        If Open_Value <> 0 Then
        
            'Calculate Percent Change
            Percent_Change = (Yearly_Change / Open_Value)
            
            'Print the Percent Change Value
            Range("N" & a).value = Percent_Change
            
            'Format Percent Change to Percent with 2 decimal places
            Columns("N").NumberFormat = "0.00%"
            
            'Find the greatest % increase in stock value
            MaxPercentIncrease = Application.WorksheetFunction.Max(Range("N2:N" & Answer_NumRows))
            
            'Print Greatest % Increase value
            Cells(2, 19).value = MaxPercentIncrease
            
            'Find the greatest % decrease in stock volume
            MinPercentIncrease = Application.WorksheetFunction.Min(Range("N2:N" & Answer_NumRows))
            
            'Print Greatest % decrease value
            Cells(3, 19).value = MinPercentIncrease
            
            'Format % cells to present data with two decimal places
            Range("S2:S3").NumberFormat = "0.00%"
         
        End If
    End If
Next a


'Finding Ticker values associated with Greatest % Increase/ Decrease and Greatest Total Volume
Dim v As Integer
For v = 2 To Answer_NumRows
    
    'Find Ticker associated with Greatest % Increase
    If Cells(v, 14).value = MaxPercentIncrease Then
    
            'Print Ticker Value in designated cell
            Cells(2, 18).value = Cells(v, 10).value
    
        'Find Ticker associated with Greatest % Decrease
        ElseIf Cells(v, 14).value = MinPercentIncrease Then
            
            'Print Ticker Value in designated cell
            Cells(3, 18).value = Cells(v, 10).value
        
        'Find Ticker associated with Greatest Total Volume
        ElseIf Cells(v, 15).value = MaxVolIncrease Then
            
            'Print Ticker value in designated cell
            Cells(4, 18).value = Cells(v, 10).value
    End If
Next v


'Formatting Yearly Change to identify positive vs. negative changes
Dim f As Integer
For f = 2 To Answer_NumRows
    
    'Identify cells in Yearly Change column with negative values
    If Cells(f, 13).value < 0 Then
        
            'Change cell color to red
            Range("M" & f).Interior.Color = RGB(235, 0, 0)
        
        'Identify cells in Yearly Change column with positive values
        ElseIf Cells(f, 13).value > 0 Then
            
            'Change cell color to green
            Range("M" & f).Interior.Color = RGB(0, 188, 85)
    
    End If
Next f


'Formatting Data within worksheet
Columns.AutoFit
Columns.VerticalAlignment = xlCenter
Rows(1).HorizontalAlignment = xlCenter
Rows(1).Font.Bold = True
Columns(17).Font.Bold = True


'Turn on screen updating and automatic calculations
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True


Next Current

starting_ws.Activate
End Sub

