Attribute VB_Name = "Module1"
Sub StockAnalyzer():

'Define variables
    Dim ws As Worksheet
    Dim OldCell As String
    Dim NewCell As String
    Dim NextRow As Integer
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim BeginDate As Double
    Dim CloseDate As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Long
    Dim i As Long
    
    'Prevent my overflow error
    On Error Resume Next
    
    'Run through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        
        'Set headers for summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Define variables for loop
        'Find lastrow
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'set variables to arbitrary values
        NextRow = 2
        BeginDate = 999999999999#
        CloseDate = 1
        TotalVolume = 0
        
            For i = 2 To lastrow

            OldCell = ws.Cells(i, 1).Value
            NewCell = ws.Cells(i + 1, 1).Value
            
                'Create loop
                If OldCell <> NewCell Then
            
                        'Find ClosePrice and OpenPrice values
                        If ws.Cells(i, 2).Value < BeginDate Then
                            'Set OpenPrice
                            OpenPrice = ws.Cells(i, 3).Value
                            'Change BeginDate Value
                            BeginDate = ws.Cells(i, 2).Value
                        ElseIf ws.Cells(i, 2).Value > BeginDate Then
                            'set ClosePrice
                            ClosePrice = ws.Cells(i, 6).Value
                            'set CloseDate
                            CloseDate = ws.Cells(i, 2).Value
                        End If
                        
                    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                    
                    'Make Calculations
                    YearlyChange = ClosePrice - OpenPrice
                    If OpenPrice = 0 Then
                        PercentChange = 1
                    Else: PercentChange = YearlyChange / OpenPrice
                    End If
            
                    'Print values in summary table
                    ws.Cells(NextRow, 9).Value = OldCell
                    ws.Cells(NextRow, 10).Value = YearlyChange
                    ws.Cells(NextRow, 11).Value = PercentChange
                    ws.Cells(NextRow, 12).Value = TotalVolume
                
                    'Reset values for next loop
                    NextRow = NextRow + 1
                    TotalVolume = 0
                    BeginDate = 9999999999999#
                    CloseDate = 1
                    
                Else
                'Set TotalVolume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
                'Find ClosePrice and OpenPrice values
                    If ws.Cells(i, 2).Value < BeginDate Then
                        'Set OpenPrice
                        OpenPrice = ws.Cells(i, 3).Value
                        'Change BeginDate Value
                        BeginDate = ws.Cells(i, 2).Value
                    ElseIf ws.Cells(i, 2).Value > BeginDate Then
                        'set ClosePrice
                        ClosePrice = ws.Cells(i, 6).Value
                        'set CloseDate
                        CloseDate = ws.Cells(i, 2).Value
                    End If
                End If
            'next loop
            Next i
            
        'Format summary table
        ws.Range("K:K").NumberFormat = "0.00%"
        
                'Apply coniditonal formating to Yearly Change "J" column
                'Define variables for j loop
                Dim j As Double
                LastRowJ = ws.Cells(Rows.Count, 10).End(xlUp).Row
                    
                    For j = 2 To LastRowJ
                    
                    'create loop
                    If ws.Cells(j, 10).Value > 0 Then
                        ws.Cells(j, 10).Interior.ColorIndex = "4"
                    Else
                        ws.Cells(j, 10).Interior.ColorIndex = "3"
                    End If
                
                'reset loop
                Next j
            
        'CHALLENGE
        
        'Find Greatest Values
        'Define Variables
        Dim SumTableLastrow As Integer
        Dim MinPercentDecrease As Double
        Dim MaxPercentIncrease As Double
        Dim MaxTotalVolume As Double
        Dim MinPercentDecreaseTicker As String
        Dim MaxPrecentIncreaseTicker As String
        Dim MaxTotalVolumeTicker As String
        Dim x As Integer
        
        'set arbitrary values for variables
        MinPercentDecrease = 9999999999999#
        MaxPercentIncrease = -9999999999999#
        MaxTotalVolume = -9999999999999#
        
        'Find last row
        SumTableLastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
            For x = 2 To SumTableLastrow
            
            'Create loop
                'Find MinPercentDecrease
                If ws.Cells(x, 11).Value < MinPercentDecrease Then
                    'Set MinPercentDecrease Value and Ticker
                    MinPercentDecreaseTicker = ws.Cells(x, 9).Value
                    MinPercentDecrease = ws.Cells(x, 11).Value
                End If
                'Find MaxPercentIncrease
                If ws.Cells(x, 11).Value > MaxPercentIncrease Then
                    'Set MaxPercentIncrease Value and Ticker
                    MaxPercentIncreaseTicker = ws.Cells(x, 9).Value
                    MaxPercentIncrease = ws.Cells(x, 11).Value
                End If
                'Find MaxTotalVolume
                If ws.Cells(x, 12).Value > MaxTotalVolume Then
                    'Set MaxTotalVolume Value and Ticker
                    MaxTotalVolumeTicker = ws.Cells(x, 9).Value
                    MaxTotalVolume = ws.Cells(x, 12).Value
                End If
            'Reset loop
            Next x
            
            'Format values
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            'Print values in Table
            ws.Range("P2").Value = MaxPercentIncreaseTicker
            ws.Range("Q2").Value = MaxPercentIncrease
            ws.Range("P3").Value = MinPercentDecreaseTicker
            ws.Range("Q3").Value = MinPercentDecrease
            ws.Range("P4").Value = MaxTotalVolumeTicker
            ws.Range("Q4").Value = MaxTotalVolume
            'Format values

        
    Next ws

End Sub
