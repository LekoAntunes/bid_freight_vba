Sub UpdateDeliveries()

    Call SetScreenUpdating(False)
    Call SetParameters
    
    If ConfirmUpdateDeliveries Then
    
        'Atualiza as  deliveries
        Call ClearWorksheet(wsd, 2)
        Call ClearHeader(wsd, "R", 1)
        Call GetDeliveries
        
        'Reinicia Price e Resumo
        Call ClearWorksheet(wsp, 1)
        Call ClearWorksheet(wsr, 1)
        Call AddFixedHeader(wsp)
        Call AddFixedHeader(wsr)
        Call AddDeliveriesResume
                
        wsc.Activate
        wsc.Range("L5") = Now
        wsc.Range("L11") = ""
        MsgBox "Done!"
    
    End If
    
    Call SetScreenUpdating(True)
    
End Sub

Function ConfirmUpdateDeliveries() As Boolean

    Dim check As Long
    Dim response As Boolean
    
    check = MsgBox("Tem certeza que deseja atualizar as Deliveries para essa análise?", vbYesNo, "WEG BID Fracionado - Atualização Deliveries")
    
    If check = 6 Then
        response = True
    Else
        response = False
    End If
    
    ConfirmUpdateDeliveries = response

End Function

Sub GetDeliveries()

    Dim key As String
    Dim keyColumn As String
    Dim cols As Long
    Dim i As Long
    
    Dim r As Long
    Dim ttr As Long
    Dim c As Long
    
    key = wsc.Range("C2")
    keyColumn = wsc.Range("C3")
    cols = CountColumns(wsd, 1)
    i = 1
    
    Set wb_temp = Workbooks.Open(wsc.Range("C5"))
    Set ws_temp = wb_temp.Worksheets("Deliveries")
    
    Call SortWorksheet(ws_temp, keyColumn)
    
    c = GetWorksheetColumnIndex(ws_temp, keyColumn)
    ttr = CountRows(ws_temp, 1)
    
    For r = 2 To ttr
        
        If ws_temp.Cells(r, c) = key Then
            
            i = i + 1
            Call GetDeliveriesLine(cols, i, r)
        
        End If
        
    Next r
    
    Call ClearTempParameters
    
End Sub

Sub GetDeliveriesLine(ttc As Long, i As Long, r As Long)

    Dim c As Long
    Dim h As String
    
    For c = 1 To ttc
        
        h = wsd.Cells(1, c)
        Call SetWorksheetColumnValue(wsd, i, h, GetWorksheetColumnValue(ws_temp, r, h))
        
    Next c
    
End Sub

Sub AddFixedHeader(ws As Worksheet)

    ws.Range("A1") = "Análise"
    ws.Range("B1") = "UF"
    ws.Range("C1") = "Itinerário"
    ws.Range("D1") = "Entregas"
    ws.Range("E1") = "Peso Bruto kg"
    ws.Range("F1") = "Valor Merc. BRL"
    ws.Range("G1") = "Dono Itinerário"

End Sub

Sub AddDeliveriesResume()
    
    Dim r As Long
    Dim ttr As Long
    Dim i As Long
    
    Dim route As String
    Dim newRoute As String
    Dim uf As String
    Dim dlv As Double
    Dim wgt As Double
    Dim vl As Double
    
    Call SortWorksheet(wsd, "Z_Route_Name")
    
    ttr = CountRows(wsd, 1)
    i = 1
    
    For r = 2 To ttr
    
        newRoute = GetWorksheetColumnValue(wsd, r, "Z_Route_Name")
        
       ' Tratativa para primeira linha
        If r = 2 Then
            route = newRoute
        End If
        
        ' Tratativa para mudança de itinerário
        If newRoute <> route Then
        
            i = i + 1
            uf = GetWorksheetColumnValue(wsd, r, "Z_UF")
            Call AddRouteResume(wsp, i, uf, route, dlv, wgt, vl)
            Call AddRouteResume(wsr, i, uf, route, dlv, wgt, vl)
            dlv = 0
            wgt = 0
            vl = 0
            
        End If
        
        'Coleta de dados normal
        route = newRoute
        dlv = dlv + GetWorksheetColumnValue(wsd, r, "Z_Entregas")
        wgt = wgt + GetWorksheetColumnValue(wsd, r, "Z_PesoKg")
        vl = vl + GetWorksheetColumnValue(wsd, r, "Valor Mercadoria")
        
        ' Tratativa para última linha
        If r = ttr Then
        
            i = i + 1
            uf = GetWorksheetColumnValue(wsd, r, "Z_UF")
            Call AddRouteResume(wsp, i, uf, route, dlv, wgt, vl)
            Call AddRouteResume(wsr, i, uf, route, dlv, wgt, vl)
            
        End If
            
    Next

End Sub

Sub AddRouteResume(ws As Worksheet, i As Long, uf As String, route As String, dlv As Double, wgt As Double, vl As Double)

    ws.Cells(i, 1) = wsc.Range("C2")
    ws.Cells(i, 2) = uf
    ws.Cells(i, 3) = route
    ws.Cells(i, 4) = Round(dlv, 0)
    ws.Cells(i, 5) = Round(wgt, 0)
    ws.Cells(i, 6) = Round(vl, 0)

End Sub
