Sub grabvol()
    Dim CW As Worksheet
    Dim x As Double
    Dim Total As Double
    Dim TotalV As Double
    Dim pct As Double
    
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    x = 2
    Cells(x, 9).Value = Cells(x, 1).Value
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Volume"
    
    For Each CW In Worksheets
        For I = 2 To LastRow
        
            If Cells(I, 1).Value = Cells(x, 9).Value Then
        
            TotalV = TotalV + Cells(I, 7).Value
        
            Cells(x, 10).Value = TotalV
        
            TotalV = Cells(I, 7).Value

            
            Else
            
            x = x + 1
            Cells(x, 9).Value = Cells(I, 1).Value
        
            End If
            'pct change for later
            'If Cells(i, 1).Value = cells(x,9) Then
                'startvalue = cells(i,1).value=stockstartdate 
                'endvalue = cells(i,1).value=stockenddate 
                'pct = endvalue+startvalue
                'pct = cells(i,8).value
            
            
        
        Next I
    Next


End Sub
