Sub stockvolumes()

'set the variable for the stock ticker
'set as a string because it will be a string of text
Dim stockticker As String

'set the initial variable for holding the stock ticker volume
'set as a double because of the size of the number?
Dim stockvolume As Double

'assigning the volume at the base level of zero?
stockvolume = 0

'create the table in which to hold the ticker name and volume
'create as integer because it is going to be a number?
Dim summary_table_row As Integer
summary_table_row = 2


'create a loop to go through all stock tickers
'use the last row formula in order to count the last row, since each sheet is different
    'still don't know why it didn't work when I assigned lastrow as a string and an integer this did not work
For i = 2 To Cells(Rows.Count, "A").End(xlUp).Row

    'in order to check if it is the same stock ticker, create an if statement
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'set the stockticker
        stockticker = Cells(i, 1).Value
        
        'add up all the stock volume
        'stockvolume is already defined as zero, so you need to do this to make it add up
        stockvolume = stockvolume + Cells(i, 7)
        
        'print out the stock ticker in the summary table
        Range("J" & summary_table_row).Value = stockticker
        
        Range("K" & summary_table_row).Value = stockvolume
        
        'add one to the summary table row so that the number appears one row down?
        summary_table_row = summary_table_row + 1
        
        'reset the stock volume
        stockvolume = 0
        
        'if the cell imediately following a row is the same ticker
        Else
        
            'add the ticker volume
            stockvolume = stockvolume + Cells(i, 7)
            
        'end the if statement
        End If
        
        'close the loop
        Next i
        
      

End Sub
