Sub StockInformationMain()

    'Writes Headers for each column
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    'Reformat column K to Percent with two decimal places
    Range("K:K").NumberFormat = "0.00%"
    Range("Q2:Q3").NumberFormat = "0.00%"
    
    Range("J:J").NumberFormat = "0.000000000"

    'Variable Declarations
    Dim LastRow, NewRow, TotalVolume, NewVolume As Long
    Dim OpeningPrice, ClosingPrice, YearlyChange, PercentChange, GLR, GPI, GPD, GTV, NPC, NTV As Double
    Dim OldTicker, NewTicker, GPIT, GPDT, GTVT As String
    
    'Set LastRow
    LastRow = 1 + Cells(Rows.Count, "A").End(xlUp).Row
    
    'Greatest Last Row
    GLR = Cells(Rows.Count, "I").End(xlUp).Row
    
    'Variable Initialization
    OldTicker = Cells(2, 1).Value
    NewRow = 2
    TotalVolume = 0
    OpeningPrice = 0
    GPI = 0 'Greatest Percent Increase
    GPD = 0 'Greatest Percent Decrease
    GTV = 0 'Greatest Total Volume
    
    'Open main program loop
    For i = 2 To LastRow
        
        'Get Ticker of current row
        NewTicker = Cells(i, 1).Value
        
        'Compares the values of NewTicker and OldTicker
        If NewTicker <> OldTicker Then
        
            'Mathy bits
            ClosingPrice = Cells(i - 1, 6).Value
            YearlyChange = ClosingPrice - OpeningPrice
            If OpeningPrice = 0 Then 'Division by zero detector
                PercentChange = 0
            Else
                PercentChange = YearlyChange / OpeningPrice
            End If
            
            'Conditional Color Formatting
            If YearlyChange > 0 Then
                Cells(NewRow, 10).Interior.ColorIndex = 4
            ElseIf YearlyChange < 0 Then
                Cells(NewRow, 10).Interior.ColorIndex = 3
            End If
            
            'Print stock information
            Cells(NewRow, 9).Value = OldTicker 'Print Ticker
            Cells(NewRow, 10).Value = YearlyChange 'Print Yearly Change
            Cells(NewRow, 11).Value = PercentChange 'Print Percent Change
            Cells(NewRow, 12).Value = TotalVolume 'Print total stock volume
            
            'Reset Values
            OldTicker = NewTicker
            NewRow = NewRow + 1
            TotalVolume = Cells(i, 7).Value
            OpeningPrice = Cells(i, 3).Value
            
        End If
        
        'Get volume of next cell and add it to the current total volume
        NewVolume = Cells(i, 7).Value
        TotalVolume = TotalVolume + NewVolume
        
    Next i
    
    'Calculate Greatest Values
    For i = 2 To GLR
    
        'New Percent Change
        NPC = Range("K" & i).Value
        
        'New Total Volume
        NTV = Range("L" & i).Value
        
        'Compare all the % values to find greatest increase
        If NPC > GPI Then
            GPI = NPC
            GPIT = Range("I" & i).Value
            
        'Compare all the % values to find greatest decrease
        ElseIf NPC < GPD Then
            GPD = NPC
            GPDT = Range("I" & i).Value
        End If
        
        'Compare total to find greatest total
        If NTV > GTV Then
            GTV = NTV
            GTVT = Range("I" & i).Value
        End If
    Next i
    
    'Print Tickers
    Range("P2").Value = GPIT
    Range("P3").Value = GPDT
    Range("P4").Value = GTVT
    
    'Print Values
    Range("Q2").Value = GPI
    Range("Q3").Value = GPD
    Range("Q4").Value = GTV

End Sub

Sub AllWS()
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    StockInformationMain
    Next WS
End Sub