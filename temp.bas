
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

    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
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
            openVal = Cells(I, "C")
        End If
        
        Total = Total + Cells(I, "G")
        
        If Total > GVal Then
            GVal = Total
            GV = Cells(I, "A")
        End If
        
        If Cells(I, "A") <> Cells(I + 1, "A") Then
            
            Cells(sIndex, "I") = Cells(I, "A")
            
            closeVal = Cells(I, "F")
            Change = closeVal - openVal
            Cells(sIndex, "J") = Change
            
            If Change > GInc Then
                GInc = Change
                GT = Cells(I, "A")
            End If
            
            If Change < Gdec Then
                Gdec = Change
                DT = Cells(I, "A")
            End If
            
            If Change > 0 Then
            
                    Cells(sIndex, "J").Interior.ColorIndex = 4
            Else
            
                    Cells(sIndex, "J").Interior.ColorIndex = 3
                    
            End If
            
            Cells(sIndex, "K") = Change / openVal * 100
            
            Cells(sIndex, "L") = Total
            
            openVal = 0
            sIndex = sIndex + 1
            Total = 0
          
        End If
    

    Next I
    
    Range("P2") = GT
    Range("P3") = DT
    Range("Q2") = GInc
    Range("Q3") = Gdec
    Range("P4") = GV
    Range("Q4") = GVal
    
Next ws

End Sub
