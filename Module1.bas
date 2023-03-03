Attribute VB_Name = "Module1"
Sub stocksummary()

    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim count As Integer
    Dim closing As Double
    Dim opening As Double
    Dim change As Double
    Dim percent As Double
    Dim volume As Double
    Dim lastRow2 As Integer
    Dim maxvol As Double
    Dim maxvoltick As String
    Dim maxincrease As Double
    Dim maxintick As String
    Dim maxdecrease As Double
    Dim maxdetick As String
    

    For Each ws In Worksheets
     
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
       
        maxvol = 0
        maxincrease = 0
        maxdecrease = 0
       
        lastRow = ws.Cells(Rows.count, "A").End(xlUp).Row
        
        count = 1
       
       opening = ws.Cells(2, 3)
       
        For i = 2 To lastRow
            
            volume = volume + ws.Cells(i, 7)
            
            ticker = ws.Cells(i, 1).Value
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                ws.Cells(count + 1, 9).Value = ticker
                closing = ws.Cells(i, 6)
                change = closing - opening
                percent = change / opening
                ws.Cells(count + 1, 10) = change
                ws.Cells(count + 1, 11) = percent
                ws.Cells(count + 1, 12) = volume
                
                If (change > 0) Then
                    ws.Cells(count + 1, 10).Interior.Color = vbGreen
                ElseIf (change = 0) Then
                    ws.Cells(count + 1, 10).Interior.Color = vbYellow
                Else
                    ws.Cells(count + 1, 10).Interior.Color = vbRed
                
                End If
                
                 If (percent > 0) Then
                    ws.Cells(count + 1, 11).Interior.Color = vbGreen
                ElseIf (percent = 0) Then
                    ws.Cells(count + 1, 11).Interior.Color = vbYellow
                Else
                    ws.Cells(count + 1, 11).Interior.Color = vbRed
                
                End If
                
                If (volume > maxvol) Then
                    maxvol = volume
                    maxvoltick = ticker
                End If
                
                If (percent > maxincrease) Then
                    maxincrease = percent
                    maxintick = ticker
                End If
                
                If (percent < maxdecrease) Then
                    maxdecrease = percent
                    maxdetick = ticker
                End If
                
                opening = ws.Cells(i + 1, 3)
                count = count + 1
                volume = 0
                
            End If
        
        Next i
        
        lastRow2 = ws.Cells(Rows.count, "I").End(xlUp).Row
        
        ws.Range("P2").Value = maxintick
        ws.Range("Q2").Value = maxincrease
        ws.Range("Q2").NumberFormat = "0.00%"
        
        ws.Range("P3").Value = maxdetick
        ws.Range("Q3").Value = maxdecrease
        ws.Range("Q3").NumberFormat = "0.00%"
        
        ws.Range("P4").Value = maxvoltick
        ws.Range("Q4").Value = maxvol
        
        ws.Columns("K:K").NumberFormat = "0.00%"
        
        ws.Columns("O:O").EntireColumn.AutoFit
        ws.Columns("L:L").EntireColumn.AutoFit
        ws.Columns("K:K").EntireColumn.AutoFit
        ws.Columns("J:J").EntireColumn.AutoFit
           
        
        
    Next ws
    
    

End Sub

