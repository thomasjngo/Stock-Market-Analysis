Attribute VB_Name = "Module1"
Sub stock_market()

'Worksheet Loop
Dim ws As Worksheet

For Each ws In Worksheets
    ws.Activate

'Part 1 (Ticker Calculation)
Dim ticker_name As String
Dim total_stock_vol As Double
Dim yearly_change As Double
Dim percent_change As Double

Dim year_open As Double
Dim year_close As Double

total_stock_vol = 0
yearly_change = 0
percent_change = 0

Dim row_generator As Integer
row_generator = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Cells(1, 14).Value = "Year Open"
'Cells(1, 15).Value = "Year Close"

    For i = 2 To lastrow

        If total_stock_vol = 0 Then
            year_open = Cells(i, 3).Value
                
        End If
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
            ticker_name = Cells(i, 1).Value
            
            total_stock_vol = total_stock_vol + Cells(i, 7).Value
                            
            Range("I" & row_generator).Value = ticker_name
                
            Range("L" & row_generator).Value = total_stock_vol
            
            'Used as debug code for testing yearly change formula
            'Range("N" & row_generator).Value = year_open
            'Range("O" & row_generator).Value = year_close
            
            yearly_change = year_close - year_open
                                  
            Range("J" & row_generator).Value = yearly_change
            percent_change = yearly_change / year_open
            Range("K" & row_generator).Value = percent_change
        
            row_generator = row_generator + 1
                
            total_stock_vol = 0
            
            Else
            
                total_stock_vol = total_stock_vol + Cells(i, 7).Value
                year_close = Cells(i + 1, 6).Value

        End If
        
    Next i
 
'Part 2 of Challenge (Greatest Calculation)

Dim x As Integer
Dim greatest_increase As Double
Dim greatest_increase_tick As String
Dim greatest_decrease As Double
Dim greatest_decrease_tick As String
Dim greatest_vol As Double
Dim greatest_vol_tick As String

Cells(1, 18).Value = "Ticker"
Cells(1, 19).Value = "Value"
Cells(2, 17).Value = "Greatest % Change"
Cells(3, 17).Value = "Greatest % Decrease"
Cells(4, 17).Value = "Greatest Total Volume"

lastrow_tick = Cells(Rows.Count, 9).End(xlUp).Row

greatest_increase = 0
greatest_decrease = 0
greatest_vol = 0

    For x = 2 To lastrow_tick
    
            If Cells(x, 10).Value >= 0 Then
                Cells(x, 10).Interior.ColorIndex = 4
                'Not sure if this was needed, but the grading rubric said change % needed conditional formatting
                Cells(x, 11).Interior.ColorIndex = 4
                Else
                    Cells(x, 10).Interior.ColorIndex = 3
                    Cells(x, 11).Interior.ColorIndex = 3
                End If
    
            If greatest_increase < Cells(x, 11).Value Then
                greatest_increase = Cells(x, 11).Value
                greatest_increase_tick = Cells(x, 9).Value
            ElseIf greatest_decrease > Cells(x, 11).Value Then
                greatest_decrease = Cells(x, 11).Value
                greatest_decrease_tick = Cells(x, 9).Value
            ElseIf greatest_vol < Cells(x, 12).Value Then
                greatest_vol = Cells(x, 12).Value
                greatest_vol_tick = Cells(x, 9).Value
            Else
                Cells(2, 18).Value = greatest_increase_tick
                Cells(2, 19).Value = greatest_increase
                Cells(3, 18).Value = greatest_decrease_tick
                Cells(3, 19).Value = greatest_decrease
                Cells(4, 18).Value = greatest_vol_tick
                Cells(4, 19).Value = greatest_vol
        End If
        
    Next x

'Formatting code (not sure if it was needed)
Range("K2:K" & lastrow).NumberFormat = "0.00%"
Range("S2:S3").NumberFormat = "0.00%"
Range("S4").NumberFormat = "#,##0"

 Next ws
 
End Sub
