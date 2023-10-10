Sub RunCalc()

    Call SetParameters
    
    'wsd
    Dim d As Long
    Dim ttd As Long
    
    'wsp
    Dim p As Long
    Dim ttp As Long
    
    Dim route As String
    Dim wgt As Double
    Dim vl As Double
    
    ttp = CountRows(wsp, 1)
    ttd = CountRows(wsd, 1)

    For d = 2 To ttd
    
        If d = 138 Then
            DoEvents
        End If
    
        route = GetWorksheetColumnValue(wsd, d, "Z_Route_Name")
        wgt = GetWorksheetColumnValue(wsd, d, "Z_PesoKg")
        vl = GetWorksheetColumnValue(wsd, d, "Valor Mercadoria")
        
        For p = 2 To ttp

            If wsp.Cells(p, 3) = route Then
                Call GetCarrierCalcValues(d, p, wgt, vl)
                Exit For
            End If

        Next p

    Next d
    
            wsc.Activate
        wsc.Range("L5") = Now
        wsc.Range("L11") = ""
        MsgBox "Done!"

End Sub

Sub GetCarrierCalcValues(d As Long, p As Long, wgt As Double, vl As Double)

    Dim c As Long
    Dim ttc As Long
    Dim colName As String
    Dim carrier As String
    Dim freight As Double
    
    Dim i As Long
    
    ttc = CountColumns(wsp, 1)
    
    For c = 8 To ttc
    
        colName = wsp.Cells(1, c)
        
        If colName Like "*T1" Then
            
            carrier = Replace(colName, " - T1", "")
            freight = CalculateFreight(carrier, c, p, wgt, vl)
            
            Call UpdateDeliveryFreightCalc(d, carrier, freight)
        
        End If
    
    Next

End Sub

Function CalculateFreight(carrier As String, initialCol As Long, r As Long, wgt As Double, vl As Double) As Double
    
    Dim c As Long
    Dim ttc As Long
    
    Dim t As Double     'Tarifa
    Dim tp As String         'Tipo Tarifa
    Dim lm As Double          'Limite Tarifa
    
    Dim tp2 As String
    Dim lm2 As Double
        
    Dim freight As Double
    Dim freight_temp  As Double
  
    ttc = initialCol + 10
    
    For c = initialCol To ttc
        
        freight_temp = 0
        t = wsp.Cells(r, c)
        
        tp = wsp.Cells(r + 100, c)
        lm = wsp.Cells(r + 200, c)
        
        tp2 = wsp.Cells(r + 100, c + 1)
        lm2 = wsp.Cells(r + 200, c + 1)
        
        If t > 0 Then
        
            'Limite
            If CheckFareLimit(tp, lm, wgt, vl) Then
                   
                Select Case tp
                    
                    Case "M", "F"
                        freight = t
                        
                    Case "TON"
                        If tp <> tp2 Then
                            freight_temp = wgt * (t / 1000)
                        ElseIf wgt <= lm2 Then
                            freight_temp = wgt * (t / 1000)
                        End If
                        
                        If freight_temp > freight Then
                            freight = freight_temp
                        End If
                        
                    Case "KG"
                        If tp <> tp2 Then
                            freight_temp = wgt * t
                        ElseIf wgt <= lm2 Then
                            freight_temp = wgt * t
                        End If
                        
                        If freight_temp > freight Then
                            freight = freight_temp
                        End If
                        
                    Case "V"
                        If tp <> tp2 Then
                            freight_temp = vl * t
                        ElseIf wgt <= lm2 Then
                            freight_temp = vl * t
                        End If
                        
                        If freight_temp > freight Then
                            freight = freight_temp
                        End If
                        
                    Case "E"
                        freight_temp = (wgt - lm) * t
                        freight = freight + freight_temp
                        
                    Case "G"
                        freight_temp = vl * t
                        freight = freight + freight_temp
                        
                    'Pedágio
                    Case "P KG"
                        freight_temp = wgt * t
                        freight = freight + freight_temp
                        
                    Case "P 100"
                        freight_temp = (wgt / 100)
                        freight_temp = Application.WorksheetFunction.RoundUp(freight_temp, 0) * t
                        freight = freight + freight_temp
                        
                    Case "P FX"
                        freight = freight + t
                
                End Select
            
            End If
            
        End If

    Next c
    
    CalculateFreight = Round(freight, 2)
    
End Function

Function CheckFareLimit(tp As String, lm As Double, wgt As Double, vl As Double) As Boolean

    Dim response As Boolean
    
    response = False
    
    Select Case tp
        
        Case "M", "TON", "KG", "E", "P KG", "P 100", "P FX"
            
            If wgt > lm Then
                response = True
            End If
            
        Case "V", "G"
            
            If vl > lm Then
                response = True
            End If
                
    End Select
    
    CheckFareLimit = response
    
End Function

Sub UpdateDeliveryFreightCalc(d As Long, carrier As String, freight As Double)

    Dim c As Long
    Dim ttc As Long
    
    ttc = CountColumns(wsd, 1)
    
    For c = 2 To ttc
        
        If wsd.Cells(1, c) = carrier Then
        
            wsd.Cells(d, c) = freight
            Exit For
            
        End If
        
    Next

End Sub
