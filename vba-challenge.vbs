Sub VBA_of_wallstreet():

Dim sheet as Worksheet

For Each sheet in Worksheets

    'set variables
    Dim tick As String
    Dim yearopen As Double
    Dim yearclose As Double
    Dim yearchange As Double
    Dim percentchange As Double
    Dim volume As Double
    Dim summarytablerow As Integer
    Dim openrow As Integer

    'set values
    volume = 0
    summarytablerow = 2
    openrow = 2
    lastrow = sheet.Cells(Rows.Count, 1).End(xlUp).Row
    lastrowsummary = sheet.Cells(Rows.Count, 10).End(xlUp).Row
      
    'set headers
    sheet.Range("I1").Value = "Ticker"
    sheet.Range("L1").Value = "Total Stock Volume"
    sheet.Range("J1").Value = "Yearly Change"
    sheet.Range("K1").Value = "Percent Change"

    'set value for x, the length of our data
    For x = 2 To lastrow

        If sheet.Cells(x + 1, 1).Value <> sheet.Cells(x, 1).Value Then
            'pulls individual tickers
            tick = sheet.Cells(x, 1).Value
            'calculates total volume
            volume = volume + sheet.Cells(x, 7).Value
    
            sheet.Range("I" & summarytablerow).Value = tick
    
            sheet.Range("L" & summarytablerow).Value = volume
    
            volume = 0
    
            yearopen = sheet.Cells(openrow, 3)
    
            yearclose = sheet.Cells(x, 6)
            'calculate yearly change
            yearchange = yearclose - yearopen
    
            sheet.Range("J" & summarytablerow).Value = yearchange
            'calculate percentage change
            percentchange = (yearchange / yearopen)
        
            sheet.Range("K" & summarytablerow).Value = percentchange
        
            sheet.Range("K" & summarytablerow).NumberFormat = "0.00%"
        
            summarytablerow = summarytablerow + 1
    
        Else

            volume = volume + sheet.Cells(x, 7).Value
    
        End If
 
    'set value for y, length of our summary table
    For y = 2 To lastrowsummary
        'conditional formatting for cells under yearly change
        If sheet.Range("J" & y).Value > 0 Then
            sheet.Range("J" & y).Interior.ColorIndex = 4

        ElseIf sheet.Range("J" & y).Value < 0 Then
            sheet.Range("J" & y).Interior.ColorIndex = 3
        
        End If

    Next y   

Next x  

Next sheet

End Sub
