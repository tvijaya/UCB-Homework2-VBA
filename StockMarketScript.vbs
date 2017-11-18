Sub CalculateYearlyValues()

    Dim Ticker As String
    Dim Stock_Total As Double
    Dim Summary_Table_Row As Double
    Dim rowCount As Double
    
    Dim YC_RowCount As Double
    Dim startIndex As Long
    Dim destIndex As Long
    Dim changeCounter As Double
    
    Dim lMax As Double
    Dim lMin As Double
    Dim MM_Counter As Double
    Dim MaxCounter As Integer
    Dim MinCounter As Integer
    
    Dim TVCounter As Double
    Dim MTotalCounter As Integer
    
    
    
    For Each ws In Worksheets
    
      Summary_Table_Row = 2
      Stock_Total = 0
      rowCount = ActiveSheet.UsedRange.Rows.Count
      ws.Range("I1").Value = "Ticker"
      ws.Range("J1").Value = "Total Stock Volume"
      ws.Range("K1").Value = "Yearly Change"
      ws.Range("L1").Value = "Percentage Change"
      ws.Range("O1").Value = "Ticker"
      ws.Range("P1").Value = "Value"
      ws.Range("N2").Value = "Greatest % Increase"
      ws.Range("N3").Value = "Greatest % Decrease"
      ws.Range("N4").Value = "Greatest Total Volume"
      
      For i = 2 To rowCount
      
          If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            Ticker = ws.Cells(i, 1).Value
            Stock_Total = Stock_Total + ws.Cells(i, 7).Value
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            ws.Range("J" & Summary_Table_Row).Value = Stock_Total
            Summary_Table_Row = Summary_Table_Row + 1
            Stock_Total = 0
          
          Else
            Stock_Total = Stock_Total + ws.Cells(i, 7).Value
          End If
        Next i
        
        'Calculation for Yearly change and Percentage change
        destIndex = 2
        startIndex = 2
        YC_RowCount = ActiveSheet.UsedRange.Rows.Count
        
          For j = 2 To YC_RowCount
            If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
            
                ' 11 is column number of yearly change
                ws.Cells(destIndex, 11).Value = ws.Cells(j, 6).Value - ws.Cells(startIndex, 3).Value
                ' 12 is row number of Percentage change
                If (ws.Cells(startIndex, 3).Value <> 0) Then
                        ws.Cells(destIndex, 12).Value = ws.Cells(destIndex, 11).Value / ws.Cells(startIndex, 3).Value
                End If
                ws.Cells(destIndex, 12).NumberFormat = "0.00%"
                destIndex = destIndex + 1
                startIndex = j + 1
                
            End If
                
          Next j
          
          changeCounter = 2
          While Not IsEmpty(ws.Cells(changeCounter, 11)) '11 is column number of yearly change
    
            If (ws.Cells(changeCounter, 11).Value < 0) Then
                 ws.Cells(changeCounter, 11).Interior.ColorIndex = 3
             Else
                 ws.Cells(changeCounter, 11).Interior.ColorIndex = 4
             End If
                
            changeCounter = changeCounter + 1
          Wend
          
          
          'Calculation for Greatest % Inc and Dec
          MM_Counter = 2
          IMax = 0
          IMin = 0
          While Not IsEmpty(ws.Cells(MM_Counter, 12)) '12 is column number of precentage change
        
               If (ws.Cells(MM_Counter, 12).Value > IMax) Then
                    IMax = Cells(MM_Counter, 12).Value
                    MaxCounter = MM_Counter
                ElseIf (ws.Cells(MM_Counter, 12).Value < IMin) Then
                    IMin = ws.Cells(MM_Counter, 12).Value
                    MinCounter = MM_Counter
                End If
                
                MM_Counter = MM_Counter + 1
        Wend
                
        ws.Cells(2, 15).Value = ws.Cells(MaxCounter, 9).Value 'max % ticker
        ws.Cells(3, 15).Value = ws.Cells(MinCounter, 9).Value
        ws.Cells(2, 16).Value = ws.Cells(MaxCounter, 12).Value 'max % value
        ws.Cells(3, 16).Value = ws.Cells(MinCounter, 12).Value
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16).NumberFormat = "0.00%"
            
            
            
            
        'Greatest Total Volume Calculation
        TVCounter = 2
        MTotalCounter = 0
    
        While Not IsEmpty(ws.Cells(TVCounter, 10))
            If (ws.Cells(TVCounter, 10).Value > IMax) Then
                MTotalCounter = TVCounter
            End If
            
            TVCounter = TVCounter + 1
        Wend
        
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = ws.Cells(MTotalCounter, 9).Value ' ticker
        ws.Cells(4, 16).Value = ws.Cells(MTotalCounter, 10).Value ' value

    Next ws

End Sub
