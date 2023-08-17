# VBA-challenge
Sub alphabetical_testing()

' set initial variable as the ticker
'variable calculations


Dim ws As Worksheet
Dim worksheetname As String
Dim ticker As String
Dim ticker_value As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim great_incr As Double
Dim great_decr As Double
Dim great_vol As Double
Dim lastrow As Long
Dim lastrowJ As Long
Dim summary_table_row As Long
Dim rng As Range
Dim tickcount As Long
Dim j As Long
Dim i As Long


For Each ws In Worksheets

'Define worksheetname
worksheetname = ws.Name


'Define column headers
ws.Cells(1, 10).Value = "ticker_value"
ws.Cells(1, 11).Value = "yearly_change"
ws.Cells(1, 12).Value = "percent_change"
ws.Cells(1, 13).Value = "total_stock_volume"
ws.Cells(1, 17).Value = "ticker"
ws.Cells(1, 18).Value = "value"
ws.Cells(2, 16).Value = "greatest percent increase"
ws.Cells(3, 16).Value = "greatest percent decrease"
ws.Cells(4, 16).Value = "greatest total volume"


'define ticker count to first row
tickcount = 2

'set star row to 2
j = 2


'calculate last row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'loop from the beginning of the worksheet to the last used row
For i = 2 To lastrow



'Check if we are still within the same ticker value, if it is not.
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'input ticker in column J
    ws.Cells(tickcount, 10).Value = ws.Cells(i, 1).Value

    
'Calculate and write yearly_change in column K
   ws.Cells(tickcount, 11).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
        
        'Conditional formatting to show positive values in green and negative values in red
        If ws.Cells(tickcount, 11).Value < 0 Then
        
        ws.Cells(tickcount, 11).Interior.ColorIndex = 3
        
    Else
    ws.Cells(tickcount, 11).Interior.ColorIndex = 4
    
    End If
    
   
    
    
    'Calculate and write percentage change in column L
    If ws.Cells(j, 3).Value <> 0 Then
    percent_change = ((ws.Cells(j, 3).Value - ws.Cells(j, 6).Value) / ws.Cells(j, 3).Value)
    'format percentage
    ws.Cells(tickcount, 12).Value = Format(percent_change, "0.00%")
    
    End If
    
 
    
          
    'Calculate and write total stock volume in colum M
    ws.Cells(tickcount, 13).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))

        'increment tickcount
    tickcount = tickcount + 1
    
        'set new start row of ticker block
        j = i + 1
    
    End If
    
    Next i
    
    'Find the last used row in ticker value column J and show is messagebox
    lastrowJ = ws.Cells(Rows.Count, 10).End(xlUp).Row
    'MsgBox ("last row in column J is "&lastrowJ)
    
    'create summary table
    great_vol = ws.Cells(2, 13).Value
    great_incr = ws.Cells(2, 12).Value
    great_decr = ws.Cells(2, 12).Value
    
    'loop for summary
    For i = 2 To lastrowJ
    
 
     
   'to check for greatest total volume check to see if the next value is large, if confirmed then take over the new value and populate with ws.cells
            If ws.Cells(i, 13).Value > great_vol Then
                great_vol = ws.Cells(i, 13).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 10).Value
            End If

            'for greatest increase, check if the next value is larger and if confirmed take over a new value and populate ws.cells
            If ws.Cells(i, 12).Value > great_incr Then
                great_incr = ws.Cells(i, 12).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 10).Value
            End If

            'for greatest decrease, check if the next value is smaller and if confirmed taker over a new value and populate ws.cells
            If ws.Cells(i, 12).Value < great_decr Then
                great_decr = ws.Cells(i, 12).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 10).Value
            End If
            
    Next i

        'create summary results
        ws.Cells(2, 18).Value = Format(great_incr, "0.00%")
        ws.Cells(3, 18).Value = Format(great_decr, "0.00%")
        ws.Cells(4, 18).Value = Format(great_vol, "0.00E+00")
        
    Next ws

    

    
End Sub

  

