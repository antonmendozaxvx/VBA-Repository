VBA Homework
Sub TickerSummary()
' Define variables
Dim OpenPrice As Double
Dim SummaryRow As Integer
Dim TotalVolume As Double

' Entering column titles
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
       
' Finding the last row
FinalRow = Cells(Rows.Count, 1).End(xlUp).Row

'Defining the 1st row where summary start from
SummaryRow = 2
    
' Loop that will encompass until the last row of data
For i = 2 To FinalRow
    
    ' Add total stock volume as it goes down the row
    TotalVolume = Cells(i, 7).Value + TotalVolume
    
    ' Sets staring value for Open Price and Total by searching if comparing if current cell value does not match prior cell value
    If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        OpenPrice = Cells(i, 3).Value
        TotalVolume = Cells(i, 7).Value
    End If
    
    ' if cell value does not match to next cell value, then it will enter summary of ticker, yearly change, percent change, and total stock volume
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        Cells(SummaryRow, 9).Value = Cells(i, 1).Value
        Cells(SummaryRow, 10).Value = Cells(i, 6).Value - OpenPrice
    ' In case OpenPrice = 0
        If OpenPrice <> 0 Then
            Cells(SummaryRow, 11).Value = Cells(SummaryRow, 10).Value / OpenPrice
        Else: Cells(SummaryRow, 11).Value = 0
        End If
        Cells(SummaryRow, 12).Value = TotalVolume
        ' Moves summary row indicator to the next row
        SummaryRow = SummaryRow + 1
    End If

Next i

' *** Challenge Section ***
' Column Labels
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

' Finding the last row
FinalRow = Cells(Rows.Count, 9).End(xlUp).Row

' Loop that will encompass until the last row of data
For a = 2 To FinalRow
    
    ' Finding the Greatest % Increase
    If Cells(a, 11).Value > Cells(2, 17).Value Then
        Cells(2, 16).Value = Cells(a, 9).Value
        Cells(2, 17).Value = Cells(a, 11).Value
    End If
    
    ' Finding the Greatest % Decrease
    If Cells(a, 11).Value < Cells(3, 17).Value Then
        Cells(3, 16).Value = Cells(a, 9).Value
        Cells(3, 17).Value = Cells(a, 11).Value
    End If
    
    ' Finding the Greatest Total Volume
    If Cells(a, 12).Value > Cells(4, 17).Value Then
        Cells(4, 16).Value = Cells(a, 9).Value
        Cells(4, 17).Value = Cells(a, 12).Value
    End If

Next a

' *** Column formatting ***
Range("J:J").NumberFormat = "$#,##0.00_)"
Range("K:K").NumberFormat = "0.00%_)"
Range("Q2:Q3").NumberFormat = "0.00%_)"
With Range("J:J")
    .Select
    .FormatConditions.Delete
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    .FormatConditions(1).Interior.Color = RGB(255, 0, 0)
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
    .FormatConditions(2).Interior.Color = RGB(0, 255, 0)
End With
    
' Clear formatting of Yearly Change column label
Cells(1, 10).FormatConditions.Delete
    
' Auto size summary columns
Columns("I:Q").AutoFit

End Sub
