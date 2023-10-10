Sub UpdatePrices()

    Call SetScreenUpdating(False)
    Call SetParameters
    
    If ConfirmUpdatePrice Then
    
        Call GetPrices
        'wsc.Activate
        'wsc.Range("L5") = Now
        'wsc.Range("L11") = ""
        MsgBox "Done!"
    
    End If
    
    Call SetScreenUpdating(True)

End Sub

Function ConfirmUpdatePrice() As Boolean

    Dim check As Long
    Dim response As Boolean
    
    check = MsgBox("Tem certeza que deseja atualizar as prices e executar uma nova simulação de valores?", vbYesNo, "WEG BID Fracionado - Atualização Deliveries")
    
    If check = 6 Then
        response = True
    Else
        response = False
    End If
    
    ConfirmUpdatePrice = response

End Function

Sub GetPrices()

    Dim key As String
    Dim path As String
    
    Dim r As Long
    Dim ttr As Long
    Dim carrier As String
    Dim routesCount As Long
   
    key = wsc.Range("C2")
    path = wsc.Range("C6")
    ttr = CountRows(wsc, 3)
    routesCount = CountRows(wsp, 1)
    
    For r = 8 To ttr
    
        carrier = wsc.Cells(r, 3)
        Set wb_temp = Workbooks.Open(path & carrier & ".xlsx")
        Set ws_temp = wb_temp.Worksheets(key)
        
        Call CheckCarrierFile(carrier, r, routesCount)
        Call ClearTempParameters

    Next

End Sub

Sub CheckCarrierFile(carrier As String, carrierRow As Long, routesCount As Long)
    
   
    If CarrierFileisValid() Then
    
        'Carrega o dono do itinerário
        If carrierRow = 8 Then
            Call AddRouteOwner(routesCount)
        End If
        
        'Adicionar headers
        Call AddHeaders(carrier)
                
        'Captura valores da tabela do transportador
        Call AddCarrierFileValues(carrier, routesCount)
       
    End If

End Sub

Function CarrierFileisValid() As Boolean

    Dim response As Boolean
    
    response = False
    
    If ws_temp.Range("D2") = "Itinerário" Then
        response = True
    End If
    
    CarrierFileisValid = response

End Function

Sub AddRouteOwner(ttr As Long)

    Dim r As Long
    Dim route As String
    Dim n As Long
    
    If ws_temp.Range("O2") = "Transportadora" Then
    
        For r = 2 To ttr
            
            route = wsp.Cells(r, 3)
            n = 4
            
            Do
            
                If ws_temp.Cells(n, 4) = "" Then
                    Exit Do
                End If
                
                If ws_temp.Cells(n, 4) = route Then
                
                    wsp.Cells(r, 7) = ws_temp.Cells(n, 15)
                    wsr.Cells(r, 7) = ws_temp.Cells(n, 15)
                    Exit Do
                    
                End If
                
                n = n + 1
            
            Loop
            
        Next
        
    End If

End Sub

Sub AddHeaders(carrier As String)

    Dim c As Long
    Dim h As String

    Call AddHeader(wsd, carrier)
    
    For c = 1 To 10
        h = carrier & " - T" & c
        Call AddHeader(wsp, h)
    Next c
    
    For c = 1 To 10
        h = carrier & " - C" & c
        Call AddHeader(wsp, h)
    Next c
    
    Call AddHeader(wsp, carrier)
    Call AddHeader(wsr, carrier)
    
    h = carrier & " - %"
    Call AddHeader(wsp, h)
    Call AddHeader(wsr, h)

End Sub

Sub AddHeader(ws As Worksheet, h As String)

    Dim c As Long
    
    c = CountColumns(ws, 1)
    
    ws.Activate
    ws.Cells(1, c + 1) = h
    ws.Cells(1, c + 1).Select
    Selection.Font.Bold = True

End Sub

Sub AddCarrierFileValues(carrier As String, routesCount As Long)

    Dim initialCol As Long
    Dim r As Long 'wsp
    Dim c As Long 'wsp
    Dim route As String
    Dim x As Long 'ws_temp
    Dim y As Long 'ws_temp
    
    initialCol = GetInitialCarrrierCol(carrier)
    
    For r = 2 To routesCount
        
        c = initialCol
        route = wsp.Cells(r, 3)
        x = 4
        
        Do
            
            If ws_temp.Cells(x, 4) = "" Then
                Exit Do
            End If
            
            If ws_temp.Cells(x, 4) = route Then
            
                'Loop price transportador
                For y = 5 To 14
                
                    'Adiciona tarifas wsd
                    Call AddCarrierPrice(r, c, x, y)
                    c = c + 1
                
                Next y
            
            End If

            x = x + 1

        Loop
        
    Next
        
End Sub

Function GetInitialCarrrierCol(carrier As String) As Long

    Dim c As Long
    Dim ttc As Long
    
    ttc = CountColumns(wsp, 1)
    
    For c = 8 To ttc
    
        If wsp.Cells(1, c) Like carrier & "*" Then
            Exit For
        End If

    Next
    
    GetInitialCarrrierCol = c
    
End Function

Sub AddCarrierPrice(r As Long, c As Long, x As Long, y As Long)

    Dim vl As String
    Dim tp As String
    Dim lm As String
    
    vl = ws_temp.Cells(x, y)
    tp = ws_temp.Cells(x, y + 16)
    lm = ws_temp.Cells(x, y + 31)
    
    If Not IsNumeric(vl) Then
        
        vl = 0
        tp = ""
        lm = ""
        
    End If
    
    wsp.Cells(r, c) = Round(CDbl(vl), 4)
    wsp.Cells(r + 100, c) = tp
    wsp.Cells(r + 200, c) = lm

End Sub
