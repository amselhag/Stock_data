    
Sub Wall_street_Hard():

'Looping through all sheets
For Each ws In Worksheets

'Finding the last row

Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
'verify last row

'MsgBox (Str(lastrow))
    
'Naming the ticker, Total stock volume, Yearly change, Percent change in the summary table

Range("I1") = "Ticker"
Range("J1") = "Total Stock Volume"
Range("k1") = "Yearly Change"
Range("L1") = "Percent Change"
    
' start the ticker volume counter

Dim Tickervolume As Double
Tickervolume = 0
    

'start the summary table row counter

Dim rows_summary As Long
rows_summary = 2

'define the opening,closing and annual change variables
    
Dim opening As Double
Dim closing As Double
Dim annualchange As Double
Dim percentchange As Double

' J is a row value to hold the opening to the the first value of each ticker
Dim J As Double
J = 2

'Start a loop to go through cells 2 to the last row

    
Dim i As Long
    
    For i = 2 To lastrow
    
    'If the next ticker value is not equal to the previous one
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Tickervolume = Tickervolume + Cells(i, 7).Value
            Cells(rows_summary, 10).Value = Tickervolume
            Cells(rows_summary, 9).Value = Cells(i, 1)
            
        'calculating the annual change
        
            opening = Cells(J, 3).Value
            closing = Cells(i, 6).Value
            annualchange = closing - opening
            Cells(rows_summary, 11).Value = annualchange
            
        'formatting the annual change rows
        If Cells(rows_summary, 11).Value > 0 Then
            Cells(rows_summary, 11).Interior.ColorIndex = 4
        Else
            Cells(rows_summary, 11).Interior.ColorIndex = 3
        End If
            
        'calculating the percentchange
        If opening <> 0 Then
            percentchange = annualchange * 100 / opening
            Cells(rows_summary, 12).Value = percentchange
        End If
                        
            ' move the row summary cell to the next one
            rows_summary = rows_summary + 1
            
            'set the J (row) value to the first row after change
            J = i + 1
            'reset ticker volume to zero
            Tickervolume = 0
        Else
    
            Tickervolume = Tickervolume + Cells(i, 7).Value
        End If
   
   Next i
   
'beginning of Hard part

'finding the last row in the summary table
lastrow_summary = Cells(Rows.Count, 12).End(xlUp).Row

'define the maximum percentage change as max_pc
Dim max_pc As Double

'Define f as a variable to hold the highest value while looping through percentage change column or total volume column
Dim f As Double
f = 2

'name the cells O1 and P1 as ticker and value, respectively
Range("O1") = "Ticker"
Range("P1") = "Value"

'Looping through the percentage change to find the maximum
    For x = 3 To lastrow_summary
        If Cells(f, 12).Value > Cells(x, 12).Value Then
            max_pc = Cells(f, 12).Value
        Else
            f = x
            max_pc = Cells(x, 12).Value
        End If
    Next x
    
'assign the cells to their respective values
Cells(2, 16).Value = max_pc
Cells(2, 15).Value = Cells(f, 9).Value
Range("N2") = "greatest % increase"

'define the minimum percentage change as min_pc
Dim min_pc As Double
'Looping through the percentage change to find the minimum
    For x = 3 To lastrow_summary
        If Cells(f, 12).Value < Cells(x, 12).Value Then
            min_pc = Cells(f, 12).Value
        Else
            f = x
            min_pc = Cells(x, 12)
        End If
    Next x
    
'assign the cells to their respective values
Cells(3, 16).Value = min_pc
Cells(3, 15).Value = Cells(f, 9).Value
Range("N3") = "greatest % decrease"

'define the maximum volume change as max_vol
Dim max_vol As Double
    For x = 3 To lastrow
        If Cells(f, 10).Value > Cells(x, 10).Value Then
            max_vol = Cells(f, 10).Value
        Else
            f = x
            max_vol = Cells(x, 10).Value
        End If
    Next x
    
'assign the cells to their respective values
Cells(4, 16) = max_vol
Cells(4, 15).Value = Cells(f, 9).Value
Range("N4") = "greatest Total Volume"

Next ws

End Sub


