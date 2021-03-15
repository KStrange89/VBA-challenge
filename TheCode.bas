Attribute VB_Name = "Module1"
    Dim i As Double
    Dim ticker As String
    Dim volume As Double
    Dim total_volume As Double
    Dim j As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim difference As Double
    Dim total_difference As Double
    Dim prop_change As Double
    Dim percent_change As String
    Dim max As Double
    Dim min As Double
    Dim temp_min As String
    Dim temp_max As String
    Dim big_volume As Double
    Dim ws As Worksheet
    

Sub nameCells():
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
End Sub


Sub yearlyVolume():
    volume = Cells(i, 7).Value
    total_volume = total_volume + volume
    ws.Cells(j, 12).Value = total_volume
End Sub

Sub closePrice():
    close_price = ws.Cells(i, 6).Value
    difference = close_price - open_price
    total_difference = total_difference + difference
    ws.Cells(j, 10).Value = total_difference
    If (total_difference < 0) Then
        ws.Cells(j, 10).Interior.ColorIndex = 3
    ElseIf (total_difference > 0) Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
    Else
    End If
    
End Sub

Sub percentChange():
    If (open_price = 0) Then
        ws.Cells(j, 11).Value = "0.00%"
    Else
        prop_change = total_difference / open_price
        percent_change = FormatPercent(prop_change, 2)
        ws.Cells(j, 11).Value = percent_change
        total_difference = 0
    End If
End Sub

Sub findMax():
    If (prop_change > max) Then
        max = prop_change
        temp_max = FormatPercent(max, 2)
        ws.Cells(2, 16).Value = ws.Cells(i, 1).Value
        ws.Cells(2, 17).Value = temp_max
    Else
    End If
End Sub

Sub findMin():
    If (prop_change < min) Then
        min = prop_change
        temp_min = FormatPercent(min, 2)
        ws.Cells(3, 16).Value = ws.Cells(i, 1).Value
        ws.Cells(3, 17).Value = temp_min
    Else
    End If
End Sub

Sub findVolume():
    big_volume = 0
    For i = 2 To Rows.Count
        If (ws.Cells(i, 1).Value = "") Then
            Exit For
        ElseIf (ws.Cells(i, 12).Value > big_volume) Then
            big_volume = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
        Else
        End If
    Next i
End Sub

Sub columnWidth():
    For i = 2 To 20
        ws.Columns(i).AutoFit
    Next i
End Sub

Sub stocks():
    
    
    
    For Each ws In Worksheets
        Call nameCells
        Call columnWidth
    
        ticker = "Jesus Christ Superstar"
    
        j = 1
        max = 0
        min = 0
        
        For i = 2 To Rows.Count
            If (ws.Cells(i, 1).Value = "") Then
                Exit For
            
            ElseIf (ws.Cells(i, 1).Value = ticker) Then
        
                ws.Cells(j, 9).Value = ticker
            
                Call yearlyVolume
            
                If (ws.Cells(i - 1, 1) <> ticker) Then
                    open_price = ws.Cells(i, 3).Value
                
                Else
                'Do nothing
                End If
            
                If (ws.Cells(i + 1, 1) <> ticker) Then
            
                    Call closePrice
                    Call percentChange
                    Call findMax
                    Call findMin
                
                Else
                'Do nothing
            
                End If
            
            Else
                j = j + 1
                total_volume = 0
                ticker = ws.Cells(i, 1).Value
            
                Call yearlyVolume
            
                If (ws.Cells(i - 1, 1) <> ticker) Then
            
                    open_price = ws.Cells(i, 3).Value
                
                Else
                'Do nothing
                End If
            
            End If
        
        Next i
    
        Call findVolume
    
        Call columnWidth
    
    Next ws
        
End Sub
