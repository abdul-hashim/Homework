Attribute VB_Name = "my_HW"
Sub homeWork()
    
    startPrice = 1
    closePrice = 1
    totalvolume = 0
    newtick = True
    lastTicker = False
    
    
    'get last row
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    'loops between sheets
    For Each Sheet In Worksheets
    
        Sheet.Activate
        
        'setup the headers for data collected
        Sheet.Cells(1, 9) = "Ticker"
        Sheet.Cells(1, 10) = "Yearly Change"
        Sheet.Cells(1, 11) = "Percent Change"
        Sheet.Cells(1, 12) = "Total Stock Volume"
        
        Sheet.Columns("I:L").AutoFit
        
        'i is for rows
        For i = 2 To lastrow
            If lastTicker = True Or newtick = True Then
                ticker = Cells(i, 1).Value
                startPrice = Cells(i, 3).Value
                
                closePrice = 1
                totalvolume = 0
                newtick = False
                lastTicker = False
            End If
            'checks to see if the ticker changes on the new row
            If Cells(i + 1, 1) <> ticker Then
                lastTicker = True
            End If
            
            'x is for columns. It will go through each column in the row and collect the data
            For x = 1 To 7
                If x = 6 Then
                    If lastTicker = True Then
                        closePrice = Cells(i, 6).Value
                    End If
                ElseIf x = 7 Then
                'gets volume total
                    totalvolume = Cells(i, 7).Value + totalvolume
                End If
            Next x
            'if lasttick is true then print results and clear data
            If lastTicker = True Then
            
                'checks which row to post gathered data on
                templastrow = Cells(Rows.Count, 9).End(xlUp).Row + 1
                
                'calculate the below
                If startPrice = 0 Or closePrice = 0 Then
                    yearlychange = 0
                    percentChange = 0
                Else
                    calcTheChange = ((closePrice - startPrice) / startPrice)
                    yearlychange = (closePrice - startPrice)
                    percentChange = Format(calcTheChange, "#.##%")
                End If
                
                Cells(templastrow, 9).Value = ticker
                Cells(templastrow, 10).Value = yearlychange
                Cells(templastrow, 11).Value = calcTheChange
                Cells(templastrow, 12).Value = totalvolume
                If yearlychange >= 0 Then
                    Cells(templastrow, 10).Interior.Color = 5296274
                Else
                    Cells(templastrow, 10).Interior.Color = 255
                End If
            End If
        Next i
        
       ' MsgBox Sheet.Name
    Next Sheet
    
End Sub


