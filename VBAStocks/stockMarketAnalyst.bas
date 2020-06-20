Attribute VB_Name = "Module1"
Sub stockMarketAnalyst():

    Dim current As String
    Dim nextOne As String
    Dim lastRow As Double
    Dim i As Double
    
    Dim count As Integer
    Dim volume As Double
    Dim opening As Double
    Dim closing As Double
    Dim change As Double
    Dim percentChange As Double
    
    Dim highestVolume As Double
    Dim volumeTicker As String
    Dim higeestIncrease As Double
    Dim increaseTicker As String
    Dim highestDecrease As Double
    Dim decreaseTicker As String
     
    'Loop through all the sheets in the worksheet
    For Each ws In Worksheets
        
        'Get the last row of the sheet
        lastRow = ws.Cells(Rows.count, "A").End(xlUp).Row
        
        'Initialize variables for the sheet
        count = 1
        volume = 0
        opening = ws.Cells(2, 3).Value
     
        'Loop through all of the rows on the this sheet
        For i = 2 To lastRow
            
            'Get the ticker in the current row and the next row
            current = ws.Cells(i, 1).Value
            nextOne = ws.Cells(i + 1, 1).Value
            
             'Add the new row's volume to the running total
            volume = volume + ws.Cells(i, 7).Value
            
            'If the ticker has changed, set all the info for the previous ticker into the sheet
            If (current <> nextOne) Then
            
                'Keep track of how many tickers so we know which row to put the info
                count = count + 1
                
                'Get the closing and the change
                closing = ws.Cells(i, 6).Value
                change = closing - opening
                
                'Put the current ticker info in the sheet
                ws.Cells(count, 10).Value = current
                ws.Cells(count, 11).Value = change
                
                'Make sure we don't divide by zero
                If (opening <> 0) Then
                    percentChange = change / opening
                    ws.Cells(count, 12).Value = percentChange
                Else
                    ws.Cells(count, 12).Value = 0
                End If
                
                'Set the total volume for the ticker
                ws.Cells(count, 13).Value = volume
                
                'If the volume is greater than what is currently the highest, change to the new value
                If (volume > highestVolume) Then
                    highestVolume = volume
                    volumeTicker = current
                End If
                
                'If the change is greater than the highest, change it
                If (percentChange > higeestIncrease) Then
                    higeestIncrease = percentChange
                    increaseTicker = current
                End If
                
                'If the change is less than the lowest, change it
                If (percentChange < highestDecrease) Then
                    highestDecrease = percentChange
                    decreaseTicker = current
                End If
                
                'Reset variables for next ticker
                opening = ws.Cells(i + 1, 6).Value
                volume = 0
                change = 0
                percentChange = 0
 
            End If

        Next i
        
        'Now enter the summary data into the sheet
        ws.Cells(2, 17).Value = higeestIncrease
        ws.Cells(2, 16).Value = increaseTicker
        ws.Cells(3, 17).Value = highestDecrease
        ws.Cells(3, 16).Value = decreaseTicker
        ws.Cells(4, 17).Value = highestVolume
        ws.Cells(4, 16).Value = volumeTicker
        
        'Reset the variables for the next sheet
        highestVolume = 0
        higeestIncrease = 0
        highestDecrease = 0
        
    Next ws

End Sub

