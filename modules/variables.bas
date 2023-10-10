'Stream Parameters
Public stream_charset As String
Public stream_type As Long
Public stream_line_separator As Long

'Execution Log
Public timer_start As Date
Public timer_stop As Date

Public wb As Workbook
Public wsc As Worksheet
Public wsd As Worksheet
Public wsp As Worksheet
Public wsr As Worksheet

Public wb_temp As Workbook
Public ws_temp As Worksheet

Sub SetParameters()

    stream_charset = "ISO-8859-1"
    stream_type = "2"
    stream_line_separator = "10"

    Set wb = ActiveWorkbook
    Set wsc = wb.Sheets("Controle")
    Set wsd = wb.Sheets("Deliveries")
    Set wsp = wb.Sheets("Price")
    Set wsr = wb.Sheets("Resumo")

End Sub

Sub ClearTempParameters()
    
    wb_temp.Close False
    
    Set wb_temp = Nothing
    Set ws_temp = Nothing
    
End Sub
