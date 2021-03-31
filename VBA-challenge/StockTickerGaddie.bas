Attribute VB_Name = "Module1"
Sub TickerCounter()

'Set variables for items needed, and starting totals
Dim Ticker As String

Dim Price As Long
Price = 0

Dim open_price As Double
open_price = 0

Dim close_price As Double
close_price = 0

Dim Percent As Double
Percent = 0

Dim Volume As Long
SVolume = 0

'Keep track of ticker data in a separate summary table
Dim SummaryTableRow As Long
SummaryTableRow = 2

'Set-up Header Rows for SummaryTableRow
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Stock Volume"

 'Runs through the number of rows, no matter how many rows
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Use lastrow variable
For i = 2 To lastrow

' Searches for when the value of the next cell is different than that of the current cell
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'Grab the Ticker name to the cell
    Ticker = Cells(i, 1).Value
    
    'Add to the Yearly Closing Difference by calculating the totals for opening and closing price, then generating the difference
    open_price = open_price + Cells(i, 3).Value
    close_price = close_price + Cells(i, 6).Value
    Price = (close_price - open_price)
    
    'Add to the Percentage Difference using the totals generated above
    'Percent = (Price / open_price)
    If (open_price = 0 And close_price = 0) Then
        Percent = 0
        ElseIf (open_price = 0 And close_price <> 0) Then
        Percent = 1
        Else: Percent = (Price / open_price) * 100
        'Print the Percentage Change to the Summary table
            Range("K" & SummaryTableRow).Value = Percent
        
        End If
        
    'Add to the Stock Volume
    SVolume = SVolume + Cells(i, 7).Value
    
    'Print the Ticker Symbols in the summary table
    Range("I" & SummaryTableRow).Value = Ticker
      
    'Print the Yearly closing difference price to the summary table
    Range("J" & SummaryTableRow).Value = Price
 
    'Print the Percentage Change to the Summary table
    'Range("K" & SummaryTableRow).Value = Percent
    
    'Print the total Stock Volume to the Summary Table
    Range("L" & SummaryTableRow).Value = SVolume
    
  'Reset Totals
   open_price = 0
   close_price = 0
   Price = 0
   Percent = 0
   SVolume = 0

    
    'Conditional formatting for yearly change
    If Cells(SummaryTableRow, 10).Value > 0 Then
        Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
    ElseIf Cells(SummaryTableRow, 10).Value < 0 Then
        Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
    ElseIf Cells(SummaryTableRow, 10).Value = 0 Then
        Cells(SummaryTableRow, 10).Interior.ColorIndex = 0
        
    End If
    
   'Increment the table to the next row
   SummaryTableRow = SummaryTableRow + 1
   
Else

    'Add to the Yearly Closing Difference by calculating the totals for opening and closing price, then generating the difference
    open_price = open_price + Cells(i, 3).Value
    close_price = close_price + Cells(i, 6).Value
    
    Price = (close_price - open_price)
    
    'Add to the Percentage Difference using the totals generated above
    'Percent = (Price / open_price)
    
    'Add to the Volume
    SVolume = SVolume + Cells(i, 7).Value
    
    End If

  Next i
    
    'Format to accounting for Price Difference
    Columns("J:J").Select
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    'Format to % for Percentage
    Columns("K:K").Select
    Selection.NumberFormat = "0.00%"
    
    'Format to autofit details
    Columns("I:L").Select
    Columns("I:L").EntireColumn.AutoFit
    
    'Conditional formatting for yearly change
    If Cells(SummaryTableRow, 10).Value > 0 Then
        Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
    ElseIf Cells(SummaryTableRow, 10).Value < 0 Then
        Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
    ElseIf Cells(SummaryTableRow, 10).Value = 0 Then
        Cells(SummaryTableRow, 10).Interior.ColorIndex = 0
        
    End If
    
End Sub


