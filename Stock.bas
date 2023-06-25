
Sub Stock():

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    ws.Range(

    [I1] = "Ticker"
    [J1] = "Yearly Change"
    [K1] = "Percent Change"
    [L1] = "Total Stock Volume"
    [O2] = "Greatest % Increase"
    [O3] = "Greatest % Decrease"
    [O4] = "Greatest Total Volume"
    [P1] = "Ticker"
    [Q1] = "Value"

    
    lastRow = ws.cells(iows.Count, "A").End(xlUp).Row
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
            openVal = ws.cells(i, "C")
        End If
        
        Total = Total + ws.cells(i, "G")
        
        If Total > GVal Then
            GVal = Total
            GV = ws.cells(i, "A")
        End If
        
        If ws.cells(i, "A") <> ws.cells(i + 1, "A") Then
            
            ws.cells(iIndex, "I") = ws.cells(i, "A")
            
            closeVal = ws.cells(i, "F")
            Change = closeVal - openVal
            ws.cells(iIndex, "J") = Change
            
            If Change > GInc Then
                GInc = Change
                GT = ws.cells(i, "A")
            End If
            
            If Change < Gdec Then
                Gdec = Change
                DT = ws.cells(i, "A")
            End If
            
            If Change > 0 Then
            
                    ws.cells(iIndex, "J").Interior.ColorIndex = 4
            Else
            
                    ws.cells(iIndex, "J").Interior.ColorIndex = 3
                    
            End If
            
            ws.cells(iIndex, "K") = Change / openVal * 100
            
            ws.cells(iIndex, "L") = Total
            
            openVal = 0
            sIndex = sIndex + 1
            Total = 0
          
        End If
    

    Next I
    
    ws.range("P2") = GT
    ws.range("P3") = DT
    ws.range("Q2") = GInc
    ws.range("Q3") = Gdec
    ws.range("P4") = GV
    ws.range("Q4") = GVal
    
Next ws

End Sub
