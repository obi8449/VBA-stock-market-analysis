Sub stocks()
    ' Credit:Instructor Eli Rosenberg tutored a group of us with starter code
    Dim lastrow As Long
    Dim i As Long
    Dim count As Long
    Dim opening As Double
    Dim closing As Double
    Dim volume As Double
    Dim Ticker As String
    Dim percent_change As Double
    Dim row As Double
    
    For Each WS In Worksheets
        ' loops through each sheet
        lastrow = WS.Cells(Rows.count, 1).End(xlUp).row
            ' goes to last row containing data
        count = 2
            ' excludes column headers
        volume = 0
            ' forces system to start at 0 (Jan 1)
        opening = WS.Cells(2, 3).Value
        Ticker = WS.Cells(2, 1).Value
            ' defines variable
        WS.Cells(1, 9).Value = "Ticker"
        WS.Cells(1, 10).Value = "Yearly Change"
        WS.Cells(1, 11).Value = "Percent Change"
        WS.Cells(1, 12).Value = "Total Stock Volume"
            ' inputs string into defined cell
        For i = 2 To lastrow
            ' defines variable and starts loop
            ' starting in row 2, go to last row with data, then repeat loop starting with 3, etc...
            volume = volume + WS.Cells(i, 7).Value
                ' 0 = 0 + 261879
                '261879 = 261879 + 157...
                
            If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
                'checks if loop is in the same the ticker symbol, if not then set closing, percent_change, ticker
        
                closing = WS.Cells(i, 6).Value
                percent_change = (closing - opening) / opening
                WS.Cells(count, 9).Value = Ticker
                WS.Cells(count, 10).Value = closing - opening
                If (closing - opening > 0) Then
                    WS.Cells(count, 10).Interior.ColorIndex = 4
                    ' sets colors
                ElseIf (closing - opening < 0) Then
                    WS.Cells(count, 10).Interior.ColorIndex = 3
                End If
                WS.Cells(count, 11).Value = percent_change
                WS.Cells(count, 11).NumberFormat = "0.00%"
                WS.Cells(count, 12).Value = volume
                opening = WS.Cells(i + 1, 3).Value
                Ticker = WS.Cells(i + 1, 1).Value
                count = count + 1
                volume = 0
            End If
        Next i
        
        ' Determines the Last Row of Yearly Change for all worksheets
        YCLastRow = WS.Cells(Rows.count, 9).End(xlUp).row
        
        WS.Cells(2, 15).Value = "Greatest % Increase"
        WS.Cells(3, 15).Value = "Greatest % Decrease"
        WS.Cells(4, 15).Value = "Greatest Total Volume"
        WS.Cells(1, 16).Value = "Ticker"
        WS.Cells(1, 17).Value = "Value"
        
        ' set the greatest values and their corresponding tickers
        For Z = 2 To YCLastRow
            If WS.Cells(Z, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                WS.Cells(2, 16).Value = Cells(Z, 9).Value
                WS.Cells(2, 17).Value = Cells(Z, 11).Value
                WS.Cells(2, 17).NumberFormat = "0.00%"
            ElseIf WS.Cells(Z, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
                WS.Cells(3, 16).Value = Cells(Z, 9).Value
                WS.Cells(3, 17).Value = Cells(Z, 11).Value
                WS.Cells(3, 17).NumberFormat = "0.00%"
            ElseIf WS.Cells(Z, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                WS.Cells(4, 16).Value = Cells(Z, 9).Value
                WS.Cells(4, 17).Value = Cells(Z, 12).Value
            End If
        Next Z
        
    Next WS
        
End Sub
