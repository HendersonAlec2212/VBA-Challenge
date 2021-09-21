

Sub VBA_Moderate()



'assign dimensions
Dim total As Double
Dim yearly_change As Double
Dim year_open As Double
Dim year_close As Double
Dim ticker As String
Dim percentChange As Double
Dim z As Integer
Dim row_start As Double
Dim ws As Worksheet

For Each ws In Worksheets

'count rows because there are so many
row_count = ws.Cells(Rows.Count, "A").End(xlUp).Row

'title rows
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
'check your values. check the math
'ws.Range("N1").Value = "Open Value"
'ws.Range("O1").Value = "Close Value"

'set base values of variables
z = 0
total = 0
row_start = 2
yearly_change = 0

'Start Loop for row_count, start from 2 b/c title row-----------------------------------
      For i = 2 To row_count
                'when true, print results, executes once per ticker set
           If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
             'keeps track of total (last line before ticker change in this case)
            total = total + ws.Cells(i, 7).Value
            

                       
            'divide by zero and overflow issue
            If total = 0 Then
                ' print the values when 0 then move on
                ws.Range("i" & 2 + z).Value = ws.Cells(i, 1).Value

            
            Else
' if the yearOpen = 0 then we need to skip past that row and not divide by things?
'make a mini loop to cause i to change when data is not operable?
'if a=0 then for every time a=0 we change i until not =0
'when a =/= 0 exit the mini loop but keep i the same
'dear god this took forever. dont change the end if statements.

                If ws.Cells(row_start, 3) = 0 Then
                For next_row = row_start To i
                If ws.Cells(next_row, 3).Value <> 0 Then
                row_start = next_row
                Exit For
                End If
                   
                Next next_row
            
                End If
'yearOpen written with row_start allows for a dynamic variable to take the place of i
            year_open = ws.Cells(row_start, 3)
            year_close = ws.Cells(i, 6)
            yearly_change = (year_close - year_open)
            percent_change = Round((yearly_change / year_open * 100), 2)
            'yearly_change = (ws.Cells(i, 6) - ws.Cells(row_start, 3))
            'percent_change = Round(((ws.Cells(i, 6) - ws.Cells(row_start, 3)) / ws.Cells * row_start * 100), 2)
            
           
                'MsgBox (year_close)
                'MsgBox (yearly_change)
        

                'print the values
                
                ws.Range("i" & 2 + z).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + z).Value = yearly_change
                ws.Range("K" & 2 + z).Value = "%" & percent_change
                ws.Range("L" & 2 + z).Value = total
                'ws.Range("N" & 2 + z).Value = year_open
                'ws.Range("O" & 2 + z).Value = year_close
            End If
                'reset between tickers
                total = 0
                yearly_change = 0

                'move to next row
                z = z + 1
                row_start = i + 1
                
         
        
        Else 'tabulate running total of stock volume till next row <> previous
            total = total + ws.Cells(i, 7).Value
             
        End If
    '-----------------------Color Change-----------------------------------
    ' this loop takes place after the numbers are calculated so theres a chance
    ' they may not be correct if the code is changed ex: red on positive
    ' run it again and it should fix up nicely.
            If ws.Range("J" & 2 + z).Value >= 0 Then
            ws.Range("J" & 2 + z).Interior.ColorIndex = 4
                 Else
            ws.Range("J" & 2 + z).Interior.ColorIndex = 3
            End If
        
Next i

Next ws

End Sub







