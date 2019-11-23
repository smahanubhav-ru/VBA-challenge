Attribute VB_Name = "Module1"
Sub WallStreet()
    
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
        
        '======================
        'First Table
        '======================
        
        'Get last row in each sheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Add new column names
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percentage Change"
        Range("L1").Value = "Total Stock Volume"

        'Add variables, set default values
        Dim TickerSymbol As String
        
        Dim OpeningPrice As Double
            OpeningPrice = Range("C2").Value
        
        Dim ClosingPrice As Double
            ClosingPrice = Range("F2").Value
        
        Dim YearlyChange As Double
        Dim PctChange As Double
        
        Dim TotalVolume As Double
            TotalVolume = 0
        
        Dim TargetRow As Integer
            TargetRow = 2
        
        'Loop through all ticker symbols
        For i = 2 To LastRow
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                'Calculate all values
                TickerSymbol = Cells(i, 1).Value
                ClosingPrice = Cells(i, 6).Value
                YearlyChange = ClosingPrice - OpeningPrice
                If OpeningPrice = 0 Then
                    PctChange = 0
                Else
                    PctChange = YearlyChange / OpeningPrice
                End If
                TotalVolume = TotalVolume + Cells(i, 7).Value
                
                'Paste calculated values and format cells
                Range("I" & TargetRow).Value = TickerSymbol
                
                Range("J" & TargetRow).Value = YearlyChange
                    If Range("J" & TargetRow).Value <= 0 Then
                       Range("J" & TargetRow).Interior.ColorIndex = 3
                    Else
                       Range("J" & TargetRow).Interior.ColorIndex = 4
                    End If
                
                Range("K" & TargetRow).Value = PctChange
                    Range("K" & TargetRow).NumberFormat = "0.00%"
                
                Range("L" & TargetRow).Value = TotalVolume
                
                'Reset or increment counters
                TargetRow = TargetRow + 1
                OpeningPrice = Cells(i + 1, 3).Value
                TotalVolume = 0
            Else
                TotalVolume = TotalVolume + Cells(i, 7).Value
            End If
        
        Next i
        
        '======================
        'Second Table
        '======================
        
        'Get last row of first table
        NewLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Add new row and column names
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        'Add variables, set default values to first row data
        Dim gtIncTickerSymbol As String
            gtIncTickerSymbol = Range("I2").Value
        
        Dim gtDecTickerSymbol As String
            gtDecTickerSymbol = Range("I2").Value
        
        Dim gtVolTickerSymbol As String
            gtVolTickerSymbol = Range("I2").Value
        
        Dim gtInc As Double
            gtInc = Range("K2").Value
        
        Dim gtDec As Double
            gtDec = Range("K2").Value
        
        Dim gtVol As Double
            gtVol = Range("L2").Value
        
        'Loop through first table
        For j = 2 To NewLastRow
            
            'Calculate all values
            If Cells(j, 11).Value > gtInc Then
                gtInc = Cells(j, 11).Value
                gtIncTickerSymbol = Cells(j, 9).Value
            End If
            
            If Cells(j, 11).Value < gtDec Then
                gtDec = Cells(j, 11).Value
                gtDecTickerSymbol = Cells(j, 9).Value
            End If
            
            If Cells(j, 12).Value > gtVol Then
                gtVol = Cells(j, 12).Value
                gtVolTickerSymbol = Cells(j, 9).Value
            End If
            
            'Paste calculated values
            Range("P2").Value = gtIncTickerSymbol
            Range("Q2").Value = gtInc
                Range("Q2").NumberFormat = "0.00%"
            
            Range("P3").Value = gtDecTickerSymbol
            Range("Q3").Value = gtDec
                Range("Q3").NumberFormat = "0.00%"
            
            Range("P4").Value = gtVolTickerSymbol
            Range("Q4").Value = gtVol
            
        Next j
    
    'Autofit all columns
    ws.Cells.EntireColumn.AutoFit
    Next ws

End Sub
