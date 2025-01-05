Sub YearStock()

    ' Declare variables
    Dim ws As Worksheet
    Dim LastRow As LongLong
    Dim i As LongLong
    Dim Ticker As String
    Dim Quarterly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock As LongLong
    Dim Summary_Table_Row As Long
    
    'Second Table
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As LongLong
    Dim MaxIncreaseTicker As String
    Dim MaxDecreaseTicker As String
    Dim MaxVolumeTicker As String
    
    'First and Last Row for Quarterly Percentage
    Dim ClosePrice As Double
    Dim OpenPrice As Double
    
    
    ' Check that it's the first time for
    Dim FirstTime As Boolean
    
    
    
    ' Loop 1
    
    For Each ws In ThisWorkbook.Worksheets
    
    'Find the last row
        While ws.Cells(LastRow + 1, 1).Value <> ""
            LastRow = LastRow + 1
        Wend
        
        Summary_Table_Row = 2
        Quarterly_Change = 0
        Total_Stock = 0
        FirstTime = True
    
    
        'Headers
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Quarterly Change"
    ws.Cells(1, "K").Value = "Percent Change"
    ws.Cells(1, "L").Value = "Total Stock Volume"
    
    ws.Cells(1, "P").Value = "Ticker"
    ws.Cells(1, "Q").Value = "Value"
    
    
       
    
        ' Loop 2 (nested)
    
        For i = 2 To LastRow
        
        
        'Check for new Ticker
        
        If ws.Cells(i, 1).Value <> Ticker Then
        
            If Not FirstTime Then
                Quarterly_Change = ClosePrice - OpenPrice
                
                If OpenPrice <> 0 Then
                    Percent_Change = Quarterly_Change / OpenPrice
                Else
                    Percent_Change = 0
                End If
                
        ws.Cells(Summary_Table_Row, "I").Value = Ticker
        ws.Cells(Summary_Table_Row, "J").Value = Quarterly_Change
        
        If Quarterly_Change > 0 Then
        ws.Cells(Summary_Table_Row, "J").Interior.ColorIndex = 4
        ElseIf Quarterly_Change < 0 Then
        ws.Cells(Summary_Table_Row, "J").Interior.ColorIndex = 3
        End If
        
        ws.Cells(Summary_Table_Row, "K").Value = Percent_Change
        ws.Cells(Summary_Table_Row, "K").NumberFormat = "0.00%"
        ws.Cells(Summary_Table_Row, "L").Value = Total_Stock
        ws.Cells(Summary_Table_Row, "L").NumberFormat = "#,##0"
            
            Summary_Table_Row = Summary_Table_Row + 1
        
        End If
                
            'Ticker Value Calculations and Print WORKING
            Ticker = ws.Cells(i, 1).Value
            OpenPrice = ws.Cells(i, 3).Value
            Total_Stock = 0
            FirstTime = False
            
    End If
    
    Total_Stock = Total_Stock + ws.Cells(i, 7).Value
    
    ClosePrice = ws.Cells(i, 6).Value
    
Next i

Quarterly_Change = ClosePrice - OpenPrice

If OpenPrice <> 0 Then

    Percent_Change = Quarterly_Change / OpenPrice

Else
    Percent_Change = 0

End If

        ws.Cells(Summary_Table_Row, "I").Value = Ticker
        ws.Cells(Summary_Table_Row, "J").Value = Quarterly_Change
        
        If Quarterly_Change > 0 Then
        ws.Cells(Summary_Table_Row, "J").Interior.ColorIndex = 4
        ElseIf Quarterly_Change < 0 Then
        ws.Cells(Summary_Table_Row, "J").Interior.ColorIndex = 3
        End If
        
        ws.Cells(Summary_Table_Row, "K").Value = Percent_Change
        ws.Cells(Summary_Table_Row, "K").NumberFormat = "0.00%"
        ws.Cells(Summary_Table_Row, "L").Value = Total_Stock
        ws.Cells(Summary_Table_Row, "L").NumberFormat = "#,##0"

    
    ' autofit
    ws.Columns("I:L").AutoFit
    
    
    
    MaxIncrease = -1
    MaxDecrease = 1
    MaxVolume = 0
    
    For i = 2 To Summary_Table_Row - 1
        If ws.Cells(i, "K") > MaxIncrease Then
            MaxIncrease = ws.Cells(i, "K").Value
            MaxIncreaseTicker = ws.Cells(i, "I").Value
        
        End If
        
        If ws.Cells(i, "K").Value < MaxDecrease Then
            MaxDecrease = ws.Cells(i, "K").Value
            MaxDecreaseTicker = ws.Cells(i, "I").Value
        
        End If
        If ws.Cells(i, "L").Value > MaxVolume Then
            MaxVolume = ws.Cells(i, "L").Value
            MaxVolumeTicker = ws.Cells(i, "I").Value
        
        End If
    Next i
        
        ws.Cells(2, "O").Value = "Greatest % Increase"
        ws.Cells(3, "O").Value = "Greatest % Decrease"
        ws.Cells(4, "O").Value = "Greatest Total Volume"
        
        ws.Cells(2, "P").Value = MaxIncreaseTicker
        ws.Cells(3, "P").Value = MaxDecreaseTicker
        ws.Cells(4, "P").Value = MaxVolumeTicker
        
        ws.Cells(2, "Q").Value = MaxIncrease
        ws.Cells(2, "Q").NumberFormat = "0.00%"
        ws.Cells(3, "Q").Value = MaxDecrease
        ws.Cells(3, "Q").NumberFormat = "0.00%"
        ws.Cells(4, "Q").Value = MaxVolume
        ws.Cells(4, "Q").NumberFormat = "#,##0"
        
        ws.Columns("O:Q").AutoFit
    
    
    Next ws

End Sub




