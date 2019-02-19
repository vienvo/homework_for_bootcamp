' Easy: Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
' Moderate: Create a script that will loop through all the stocks and take the following info.
'            Yearly change from what the stock opened the year at to what the closing price was.
'            The percent change from the what it opened the year at to what it closed.
' Hard: Your solution will also be able to locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".
' Challenge: Make the appropriate adjustments to your script that will allow it to run on every worksheet just by running it once.

Sub generateValue()
    'Set initial variables
    
    Dim ticker_name As String
    Dim volumn_total As Double
    Dim ticker_row As Integer
    Dim open_price As Double
    Dim close_price As Double

    
    'set initial value to variable
    volumn_total = 0
    ticker_row = 2
    
    'count rows
    n = WorksheetFunction.CountA(Range("A:A"))
    
    'Set up headers for return area
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("M1").Value = "Yearly Open Price"
    Range("N1").Value = "Yearly Closing Price"
    
    
    
    'Loop to solve ticker and volumn
    For r = 2 To n
        If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
            ticker_name = Cells(r, 1).Value
            close_price = Cells(r, 5).Value
            row_ticker = 0
            volumn_total = volumn_total + Cells(r, 7).Value
            
            'Pass value to cells
            Range("I" & ticker_row).Value = ticker_name
            Range("L" & ticker_row).Value = volumn_total
            Range("N" & ticker_row).Value = close_price
             
            ticker_row = ticker_row + 1
            volumn_total = 0
        Else
            volumn_total = volumn_total + Cells(r, 7).Value
        End If
        
        
        If Cells(r - 1, 1).Value <> Cells(r, 1).Value Then
            open_price = Cells(r, 3).Value
            
            'Pass value to cells
            Range("M" & ticker_row).Value = open_price
        End If
                  
        
    Next r
    
End Sub

'calculate yearly change
Sub yearlyChange()
    'Set initial variables

    Dim n As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim change As Double
    
       
    'count rows
    n = WorksheetFunction.CountA(Range("I:I"))
    
     
    'Loop to calculate change in yearly price
    For r = 2 To n
        Cells(r, 10).Value = (Cells(r, 14).Value - Cells(r, 13).Value)
        
        'Return change percentage
        If Cells(r, 13).Value = 0 Then
            Cells(r, 11).Value = 0
        Else
            Cells(r, 11).Value = Cells(r, 10).Value / Cells(r, 13).Value
        End If
        
        'Highligh change percentage
        If Cells(r, 10).Value < 0 Then
            Cells(r, 10).Interior.ColorIndex = 3
        Else
            Cells(r, 10).Interior.ColorIndex = 4
        End If
        
    Next r
    
    Range("K2:K" & n).NumberFormat = "0.00%"
    
End Sub

'summary result
Sub summary()
'Setup range to show summary result
    Cells(2, "Q").Value = "Greatest % Increase"
    Cells(3, "Q").Value = "Greatest % Decrease"
    Cells(4, "Q").Value = "Greatest Total Volumn"
    Cells(1, "R").Value = "Ticker"
    Cells(1, "S").Value = "Value"

    Dim last_row As Long
    last_row = Cells(Rows.Count, "I").End(xlUp).Row

'Finding value
    Dim volumn_max As Double
    Dim increase_max As Double
    Dim decrease_max As Double

    volumn_max = WorksheetFunction.Max(Range("L2:L" & last_row))
    increase_max = WorksheetFunction.Max(Range("J2:J" & last_row))
    decrease_max = WorksheetFunction.Min(Range("J2:J" & last_row))

'Return value cells
    Cells(2, "R").Value = increase_max
    Cells(3, "R").Value = decrease_max
    Cells(4, "R").Value = volumn_max

'---------
'looping through result
For r = 2 To last_row
    'Most increase ticker
    If Cells(r, "J").Value = increase_max Then
        Cells(2, "S").Value = Cells(r, "I")
    End If
    
    'Most decrease ticker
    If Cells(r, "J").Value = decrease_max Then
        Cells(3, "S").Value = Cells(r, "I")
    End If
    
    'Traded most
    If Cells(r, "L").Value = volumn_max Then
        Cells(4, "S").Value = Cells(r, "I")
    End If
Next r

'Auto fit result
    Columns("I:S").AutoFit
    Columns("M:N").EntireColumn.Hidden = True

End Sub

'call all sub in all sheets
Sub run_all()

    Dim ws As Worksheet
    Application.ScreenUpdating = False

    For Each ws In Worksheets
        ws.Select
        Call generateValue
        Call yearlyChange
        Call summary
    Next ws
    Application.ScreenUpdating = True
    MsgBox ("Job is done!")
End Sub
