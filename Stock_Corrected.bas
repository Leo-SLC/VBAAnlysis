
Sub Stock():

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets


    [I1] = "Ticker"
    [J1] = "Yearly Change"
    [K1] = "Percent Change"
    [L1] = "Total Stock Volume"
    [O2] = "Greatest % Increase"
    [O3] = "Greatest % Decrease"
    [O4] = "Greatest Total Volume"
    [P1] = "Ticker"
    [Q1] = "Value"

    
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    sIndex = 2
    openVal = 0
    Total = 0
    
    GInc = 0
    Gdec = 0
    GT = " "
    DT = " "
    GV = " "
    GVal = 0
    
    For I = 2 To lastRow
        
        If openVal = 0 Then
            openVal = ws.Cells(I, "C")
        End If
        
        Total = Total + ws.Cells(I, "G")
        
        If Total > GVal Then
            GVal = Total
            GV = ws.Cells(I, "A")
        End If
        
        If ws.Cells(I, "A") <> ws.Cells(I + 1, "A") Then
            
            ws.Cells(sIndex, "I") = ws.Cells(I, "A")
            
            closeVal = ws.Cells(I, "F")
            Change = closeVal - openVal
            ws.Cells(sIndex, "J") = Change
            
            If Change > GInc Then
                GInc = Change
                GT = ws.Cells(I, "A")
            End If
            
            If Change < Gdec Then
                Gdec = Change
                DT = ws.Cells(I, "A")
            End If
            
            If Change > 0 Then
            
                    ws.Cells(sIndex, "J").Interior.ColorIndex = 4
            Else
            
                    ws.Cells(sIndex, "J").Interior.ColorIndex = 3
                    
            End If
            
            ws.Cells(sIndex, "K") = Change / openVal * 100
            
            ws.Cells(sIndex, "L") = Total
            
            openVal = 0
            sIndex = sIndex + 1
            Total = 0
          
        End If
    

    Next I
    
    ws.Range("P2") = GT
    ws.Range("P3") = DT
    ws.Range("Q2") = GInc
    ws.Range("Q3") = Gdec
    ws.Range("P4") = GV
    ws.Range("Q4") = GVal
    
Next ws

End Sub

