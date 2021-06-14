'Create a sub-routine

 Sub stockmarketAnalysis():
 
    For Each ws In Worksheets   'Repeat the following code for all the worksheets
         
        'Declaring variables
        Dim i, j As Integer 'declaring row counters
        Dim openingPrice As Double
        Dim finalPrice
        Dim TotalStockVol
        Dim PercentChange
        Dim FoundRow As Integer
        
        'initializing values
        j = 2
        TotalStockVol = 0 'initializing TotalStockVolume
        openingPrice = ws.Cells(2, 3) 'Initializing opening price
        
    
        'Summary Table Column Headers
        ws.Range("i1:m1").Font.Bold = 1 'Setting Column Header Font to Bold
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
   
        'Find last row in the stock table
        LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
        'MsgBox (LastRow)
 
          For i = 2 To LastRow
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                    'Inserting ticker symbol to Summary table
                    ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
                    'Cells(j, 10).Value = openingPrice 'Adding the opening Price in Summary Table
              
                    'Grabbing Final Value of the Stock
                    finalPrice = ws.Cells(i, 6).Value
              
                    'Cells(j, 11).Value = finalPrice 'Adding the Final value of Stock in Summary Table
            
                    YearlyChange = finalPrice - openingPrice
             
                    'Inserting Value Change of Stock in SummaryTable
                    ws.Cells(j, 10).Value = YearlyChange
             
                    If YearlyChange = 0 Or openingPrice = 0 Then
                        PercentChange = 0
                    Else
                        PercentChange = Round((YearlyChange / openingPrice), 4)
                    End If
             
                    ws.Cells(j, 11).Value = PercentChange
             
                    'Inserting Total Stock Volume in summary Table
                    ws.Cells(j, 12).Value = TotalStockVol
             
                   'Range("j2").Value = finalPrice
                    'MsgBox (i)
             
                   'Grabbing the new opening value
                    openingPrice = ws.Cells(i + 1, 3).Value
            
                    j = j + 1
                    TotalStockVol = 0
        
                Else
                    TotalStockVol = TotalStockVol + ws.Cells(i + 1, 7)
            
                End If
        Next i
    
        'Finding Last Row in Summary Table
        LastRowST = ws.Range("I" & Rows.Count).End(xlUp).Row

        'MsgBox (LastRowST)
    
        For j = 2 To LastRowST
   
            ws.Range("K" & j).NumberFormat = "0.00%"
             If ws.Cells(j, 10).Value < 0 Then
                 ws.Range("J" & j).Interior.ColorIndex = 3
            ElseIf ws.Cells(j, 10).Value > 0 Then
                ws.Range("J" & j).Interior.ColorIndex = 4
            End If
            
        Next j
        
    
    '*********************BONUS*********************
   ws.Range("P1:Q1").Font.Bold = 1 'Setting Column Header Font to Bold
   ws.Range("O2:O4").Font.Bold = 1
    
    ' Setting up Static values in third table
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    GreatestTotalVol = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
    ws.Range("Q4 ").Value = GreatestTotalVol
  '  MsgBox (GreatestTotalVol)
    
      For j = 2 To LastRowST
        If (ws.Cells(j, 12).Value = GreatestTotalVol) Then
            FoundRow = j
        End If
       Next j
    
       ' MsgBox (ws.Cells(FoundRow, 9).Value)
        ws.Range("P4").Value = ws.Cells(FoundRow, 9).Value
   
    
  '  If NewNum > 0 Then
      ' GreatestPercentIncrease = Application.WorksheetFunction.Max(If Format(ws.Cells(j, 11).Value,"General Number") > 0, Format(ws.Cells(j, 11).Value,"General Number"))
       GreatestPercentIncrease = WorksheetFunction.Max(ws.Range("K2:K" & LastRowST))
       GreatestPercentDecrease = WorksheetFunction.Min(ws.Range("K2:K" & LastRowST))
       
       ws.Range("Q2").Value = GreatestPercentIncrease
       ws.Range("Q3").Value = GreatestPercentDecrease
       
       'Formatting column to Percentage
       ws.Range("Q2:Q3").NumberFormat = "0.00%"
       
     '  MsgBox ("GreatestPercentIncrease " & GreatestPercentIncrease)
      ' MsgBox ("GreatestPercentDecrease " & GreatestPercentDecrease)
  '   End If
  
        For j = 2 To LastRowST
         If (ws.Cells(j, 11).Value = GreatestPercentIncrease) Then
            FoundRow1 = j
        End If
       Next j
    
        ws.Range("P2").Value = ws.Cells(FoundRow1, 9).Value
        
        For j = 2 To LastRowST
         If (ws.Cells(j, 11).Value = GreatestPercentDecrease) Then
            FoundRow2 = j
        End If
       Next j
    
        
       ' MsgBox (ws.Cells(FoundRow, 9).Value)
        ws.Range("P3").Value = ws.Cells(FoundRow2, 9).Value
        
  '******************BONUS COMPLETE********************
  
Next ws
    
    MsgBox ("Analysis Complete")
    
 End Sub


