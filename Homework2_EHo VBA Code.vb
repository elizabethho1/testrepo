Sub Ticker()
    'set an inital variable to hold the ticker name
    Dim Ticker As String
    
    'set an initial variable to hold the yearly change
    Dim yearlychange As Double
    
    'set initial variables to hold the opening and closing prices
    Dim openprice As Double
    Dim closeprice As Double
    
    'keep track of the location in the summary table
    Dim summarytablerow As Integer
    summarytablerow = 2
    
    
    'set an initial variable to hold the percent change per ticker
    Dim percentchange As Double
    
    'set an initial variable to hold the total stock volume per ticker
    Dim totalvolume As Double
    
    totalvolume = 0
    
    'set initial values for variables
    yearlychange = 0
    percentchange = 0
    openprice = 0
    closeprice = 0
    totalvolume = 0
    
    'loop through all of the ticker data by date
    For i = 2 To 22771
               
        'check if we are within the same ticker, if it is not
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'set the ticker name
            Ticker = Cells(i, 1).Value
            
            'set the opening price. each ticker has 250 days of data. this will grab the first day.
            openprice = Cells(i - 249, 3).Value
            
            'set the closing price
            closeprice = Cells(i, 6).Value
            
            'add to the total stock volume
            totalvolume = totalvolume + Cells(i, 7).Value
            
            'calculate yearly change
            yearlychange = closeprice - openprice
            
            'calculate percent change
            percentchange = (yearlychange / openprice) * 100
            
            'print the ticker name in the summary table
            Range("I" & summarytablerow).Value = Ticker
            
            'print the yearly change in the summary table
            Range("J" & summarytablerow).Value = yearlychange
            
            'print the percent change in the summary table
            Range("K" & summarytablerow).Value = percentchange
            
            'print the total stock volume in the summary table
            Range("L" & summarytablerow).Value = totalvolume
            
            'This nested if statement will format the colors
            If yearlychange < 0 Then
            
                'if the yearly change is negative, highlight the cell red
                Range("J" & summarytablerow).Interior.ColorIndex = 3
            Else
                
                'if the yearly change is positive, highlight the cell green
                Range("J" & summarytablerow).Interior.ColorIndex = 4
            End If
            
            
            'add one to the summary table row counter so the next cell populates with the info for next ticker
            summarytablerow = summarytablerow + 1
            
            'reset totals
            totalvolume = 0
            openprice = 0
            closeprice = 0
            percentchange = 0
                
        
        Else
            'add to the total volume
            totalvolume = totalvolume + Cells(i, 7).Value
            
    
        End If
        
    Next i
    
End Sub

