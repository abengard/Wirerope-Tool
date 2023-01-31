Attribute VB_Name = "Calculation1"
Option Explicit

Sub Calculation()

    For i = 2 To 100
        Worksheets("ChartComparison").Activate

        If Not Cells(i, 2) = "" Then
        
      
            Kv = Cells(i, 2).Value
            Kv_ges = n * Kv
        
            Ks = Cells(i, 3).Value
            Ks_ges = n * Ks
        
            statD = ((m * 9.81) / (Kv_ges)) * 1000
        
            precomF = Kv_ges / 1000 * statD
                
                
            With WireRopeTool
                If .OptionButtonfreeFall = True Then
                
                    dynD = ((((9.81 ^ 2 * m ^ 2 + Ks_ges * m * v ^ 2) / Ks_ges ^ 2) ^ 0.5) + (9.81 * m) / Ks_ges) * 1000
                
                ElseIf .OptionButtonShock1 = True Or .OptionButtonShock2 = True Then
                
                    dynD = (((E * 2) / Ks_ges) ^ 0.5) * 1000
                
                End If
            End With
                
        
            ShF = dynD / 1000 * Ks_ges
        
            ReSha = ShF / m
        
            ReShg = ReSha / 9.81
        
            NatFr = ((Kv_ges / m) ^ 0.5) / 3.14 / 2
        
            ShFr = ((Ks_ges / m) ^ 0.5) / 3.14 / 2
    
            With ThisWorkbook.Worksheets("ChartCalculation").Activate
                Cells(i, 1).Value = Kv_ges
                Cells(i, 2).Value = Ks_ges
                Cells(i, 3).Value = statD
                Cells(i, 4).Value = precomF
                Cells(i, 5).Value = dynD
                Cells(i, 6).Value = ShF
                Cells(i, 7).Value = ReSha
                Cells(i, 8).Value = ReShg
                Cells(i, 9).Value = NatFr
                Cells(i, 10).Value = ShFr
            End With

        End If
    Next i
End Sub

