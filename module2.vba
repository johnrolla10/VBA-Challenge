Attribute VB_Name = "Module1"
Sub newFunction():
    'RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    'Assign ticker to column H
   
    Dim totalClose As Double
    Dim totalOpen As Double
    Dim totalVolume As Double
    Dim row As Integer
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim increase As Double
    Dim decrease As Double
    Dim maxVolume As Double
    Dim increaseTicker As String
    Dim decreaseTicker As String
    Dim maxVolumeTicker As String
    
    totalOpen = 0
    totalClose = 0
    totalVolume = 0
    row = 1
    increase = 0
    decrease = 0
    maxVolume = 0
    
    For i = 2 To 93001
        If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            totalOpen = totalOpen + Cells(i, 3).Value
            totalClose = totalClose + Cells(i, 6).Value
            totalVolume = totalVolume + Cells(i, 7).Value
        Else
            Cells(row + 1, 9).Value = Cells(i, 1).Value
            totalOpen = totalOpen + Cells(i, 3).Value
            totalClose = totalClose + Cells(i, 6).Value
            totalVolume = totalVolume + Cells(i, 7).Value
            percentChange = (totalClose - totalOpen) / totalOpen
            quarterlyChange = percentChange
            Cells(row + 1, 10).Value = quarterlyChange
            Cells(row + 1, 11).Value = percentChange * 100
            Cells(row + 1, 12).Value = totalVolume
            row = row + 1
            totalOpen = 0
            totalClose = 0
            totalVolume = 0
        End If
            
        If (Cells(row + 1, 10).Value > 0) Then
            Cells(row + 1, 10).Interior.ColorIndex = 4
        ElseIf (Cells(row + 1, 10).Value < 0) Then
            Cells(row + 1, 10).Interior.ColorIndex = 3
        End If
        
        'increase decrease and total volume
        If (Cells(row + 1, 11).Value > increase) Then
            increase = Cells(row + 1, 11).Value
            increaseTicker = Cells(row + 1, 9).Value
        End If
        If (Cells(row + 1, 11).Value < decrease) Then
            decrease = Cells(row + 1, 11).Value
            decreaseTicker = Cells(row + 1, 9).Value
        End If
         If (Cells(row + 1, 12).Value > maxVolume) Then
            maxVolume = Cells(row + 1, 12).Value
            maxVolumeTicker = Cells(row + 1, 9).Value
        End If
    Next i
    
    'Set value and ticker for increase decrease and total volume
    Cells(2, 17).Value = increase
    Cells(3, 17).Value = decrease
    Cells(4, 17).Value = maxVolume
    
    Cells(2, 16).Value = increaseTicker
    Cells(3, 16).Value = decreaseTicker
    Cells(4, 16).Value = maxVolumeTicker
    
    
End Sub
