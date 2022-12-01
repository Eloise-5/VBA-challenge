Sub stockcheck()

'hello reader. I would appreciate some feedback on how to reference the array when outputting into our new columns. 
'I was originally referencing the array rather than the range, however for the last row where I check if we are at the end of one ticker
'and the start of another, I could not reference the next row in the array (as it didnt exist). Hope this makes sense. Thx

    'define worksheets and loop through each sheet in workbook
    Dim sht As Worksheet
    For Each sht In ThisWorkbook.Worksheets
    
    
        'work out the number of rows to loop through
        Dim lastrow As Long
        lastrow = sht.Cells(Rows.Count, 1).End(xlUp).Row
        
        'define and store info in array
        Dim x() As Variant
        x = sht.Range("A2").CurrentRegion
    
        'print column headers for output data
        sht.Range("I1").Value = "Ticker"
        sht.Range("J1").Value = "Yearly Change"
        sht.Range("K1").Value = "Percent Change"
        sht.Range("L1").Value = "Total Stock Volume"
    
    
    
        'define looping variables
        Dim k As Integer
        Dim i As Long
        k = 1
        i = 0
        
        'define variables for open & close prices, to allow calculation
        Dim open_price As Double
        open_price = 0
        
        Dim close_price As Double
        prev_close = 0
        
        'define counter for stock volume
        Dim c As Variant
        c = 0
    
        
        'firstly list out all tickers
        For i = 1 To lastrow - 1
        
            If sht.Range("A1").Cells(i + 1, 1).Value <> sht.Range("A1").Cells(i, 1) Then
            
                sht.Range("I2").Cells(k, 1).Value = sht.Range("A1").Cells(i + 1, 1)
        
                k = k + 1
                
            End If
            
        Next i
            
        'reset variables, just in case
        i = 0
        k = 1
        
        
        
        'loop through original array, grab open & close prices and use these to calculate yearly change and percentage change
        For i = 1 To lastrow
        
            'if its a different ticker to the one above (meaning new open)
            If sht.Range("A1").Cells(i + 1, 1).Value = sht.Range("I2").Cells(k, 1).Value And sht.Range("A1").Cells(i + 1, 1).Value <> sht.Range("A1").Cells(i, 1).Value Then
            
                open_price = sht.Range("A1").Cells(i + 1, 3).Value
                c = c + sht.Range("A1").Cells(i + 1, 7).Value
            
            'if its a different ticker to the one below (meaning last one of the year)
            ElseIf sht.Range("A1").Cells(i + 1, 1).Value = sht.Range("I2").Cells(k, 1).Value And sht.Range("A1").Cells(i + 1, 1).Value <> sht.Range("A1").Cells(i + 2, 1).Value Then
            
                prev_close = sht.Range("A1").Cells(i + 1, 6).Value
                c = c + sht.Range("A1").Cells(i + 1, 7).Value
                
                'yearly change
                sht.Range("I2").Cells(k, 2).Value = prev_close - open_price
                
                    If sht.Range("I2").Cells(k, 2).Value < 0 Then
                    
                        sht.Range("I2").Cells(k, 2).Interior.Color = RGB(255, 0, 0)
                        
                    Else: sht.Range("I2").Cells(k, 2).Interior.Color = RGB(0, 255, 50)
                        
                    End If
                
                
                'percentage change
                sht.Range("I2").Cells(k, 3).Value = (prev_close - open_price) / open_price
                sht.Range("I2").Cells(k, 3).NumberFormat = "0.00%"
                
                'stock volume
                sht.Range("I2").Cells(k, 4).Value = c
                c = 0
                
                'move on to next ticker
                k = k + 1
                open_price = 0
                prev_close = 0
            
            'otherwise if its just another day in the year
            ElseIf sht.Range("A1").Cells(i + 1, 1).Value = sht.Range("I2").Cells(k, 1).Value Then
            
                c = c + sht.Range("A1").Cells(i + 1, 7).Value
            
            End If
        
        Next i

        lastrow = 0
        Erase x
        k = 0
        i = 0


    Next sht
      
    
End Sub


