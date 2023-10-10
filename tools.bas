#If VBA7 Then
   Private Declare PtrSafe Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)
#Else
   Private Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)
#End If

Public IndexColumn As Collection
Public CountColumnNames As Collection
Public CellValueChanged As Boolean
Private p_currentCols As Collection
Private p_currentName As String

Public Sub ClearAllFilter()

    Dim ws As Worksheet
    
    For Each ws In p_wb.Worksheets

        If ws.FilterMode Then
            ws.ShowAllData
        End If
    
    Next

End Sub

Public Sub CheckPermission()

    Dim user As String
    
    user = WindowsUserName
    
    Select Case user
    
        Case " "
            DoEvents
        
        Case Else
        
            Call MsgBox("You are not authorised to run this application", vbCritical)
            End

    End Select
    
End Sub

Public Function WindowsUserName() As String

    Dim szBuffer As String * 100
    Dim lBufferLen As Long

    lBufferLen = 100

    If CBool(GetUserName(szBuffer, lBufferLen)) Then

        WindowsUserName = Left$(szBuffer, lBufferLen - 1)

    Else

        WindowsUserName = CStr(Empty)

    End If

End Function

Public Function InCollection(col As Collection, key As String, Optional IsObj As Boolean = False) As Variant
  
    Dim errNumber As Long
  
    Err.Clear
    On Error Resume Next
    
        If Not IsObj Then
            InCollection = col.Item(key)
        Else
            Set InCollection = col.Item(key)
        End If
        
        errNumber = CLng(Err.Number)
        
    On Error GoTo 0
    
    If errNumber = 5 Then
        
        If Not IsObj Then
            InCollection = 0
        Else
            Set InCollection = Nothing
        End If
                
    End If

End Function

Public Function ConvertionOutput(s As String) As String

    If IsNumeric(s) Then
        
        ConvertionOutput = CStr(CDbl(s))
        
    Else
    
        ConvertionOutput = s
        
    End If

End Function

Public Function ConvertionInput(s As String, format As String) As String

    If IsNumeric(s) Then
        
        ConvertionInput = WorksheetFunction.text(CStr(CDbl(s)), format)
    
    Else
    
        ConvertionInput = s
        
    End If

End Function

Public Function ConvertionInputMatnr(m As String) As String

    ConvertionInputMatnr = ConvertionInput(m, "000000000000000000")

End Function

Public Function ConvertionOutputDate(d As String) As Variant

    Dim dt As String
    
    dt = d
    
    If IsNumeric(d) And Len(d) = 8 Then
    
        dt = Mid(d, 7, 2) & "/" & Mid(d, 5, 2) & "/" & Mid(d, 1, 4)
        
    ElseIf IsNumeric(d) And Len(d) = 14 Then
    
        dt = Mid(d, 7, 2) & "/" & Mid(d, 5, 2) & "/" & Mid(d, 1, 4) & " " & Mid(d, 9, 2) & ":" & Mid(d, 11, 2) & ":" & Mid(d, 13, 2)
        
    ElseIf Len(d) = 10 And InStr(1, d, ".") > 0 Then
    
        dt = Replace(d, ".", "/")
        
    ElseIf Len(d) = 10 And InStr(1, d, "\") > 0 Then
    
        dt = Replace(d, ".", "/")
        
    End If
    
    If IsDate(dt) And dt <> "" Then
    
        ConvertionOutputDate = CDate(dt)
    
    Else
        
        ConvertionOutputDate = ""
    
    End If

End Function

Public Sub SetScreenUpdating(show As Boolean)

    If show Then
    
        Calculate
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        DoEvents
                
    Else
    
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    
    End If

End Sub

Public Function ConvertionOutputDecimal(s As String) As Variant

    Dim vl As String
    
    vl = s
    
    If p_wsParam.Range("Q21") <> Application.DecimalSeparator Then
        vl = Replace(vl, p_wsParam.Range("Q20"), "")
        vl = Replace(vl, p_wsParam.Range("Q21"), Application.DecimalSeparator)
    End If

    If IsNumeric(vl) Then
    
        ConvertionOutputDecimal = CDbl(vl)
        
    Else
    
        ConvertionOutputDecimal = ""
            
    End If

End Function

Public Function ToNumber(s As String) As Variant

    If IsNumeric(s) Then
    
        ToNumber = CDbl(s)
        
    Else
    
        ToNumber = 0
            
    End If

End Function

Public Function IsArrayEmpty(arr As Variant) As Boolean

    Dim i As Integer
    
On Error GoTo Catch
    
    i = UBound(arr)
    IsArrayEmpty = i = 0
    GoTo Finally
    
Catch:

    IsArrayEmpty = True
    
Finally:

End Function

Public Function GetColumnLetter(c As Long) As String

    Dim adrc As String
    Dim rg() As String
    
    adrc = Cells(1, c).Address(True, False)
    rg = Split(adrc, "$")
        
    GetColumnLetter = rg(0)
    
End Function

Public Function GetWorksheetColumnValue(ws As Worksheet, rowNum As Long, columnName As String, Optional initial As Variant = "") As Variant

    Dim c As Long
    
    c = GetWorksheetColumnIndex(ws, columnName)
    If c > 0 Then
    
        GetWorksheetColumnValue = ws.Cells(rowNum, c)
        
        If GetWorksheetColumnValue = "" Then
            GetWorksheetColumnValue = initial
        End If
        
    Else
    
        GetWorksheetColumnValue = initial
    
    End If

End Function

Public Sub SetWorksheetColumnValue(ws As Worksheet, rowNum As Long, columnName As String, v As Variant, Optional initial As Variant = "", Optional sum As Boolean = False, Optional append As Boolean = False)

    Dim c As Long
    Dim r As Range
    Dim vl As Variant
    
    c = GetWorksheetColumnIndex(ws, columnName)
    If c > 0 Then
    
        Set r = ws.Cells(rowNum, c)
        
        If append Then
        
            If v <> "" Then
            
                If r = "" Then
                    r = v
                    CellValueChanged = True
                Else
                    For Each vl In Split(v, ";")
                        If vl <> "" Then
                            If InStr(1, r, vl) <= 0 Then
                                r = r & ";" & vl
                                CellValueChanged = True
                            End If
                        End If
                    Next
                End If
            
            End If
        
        Else
        
            If sum Then
            
                If IsNumeric(v) Then
            
                    r = r + v
                    If v <> 0 Then
                        CellValueChanged = True
                    End If
                
                End If
            
            Else
            
                If v = initial Then
                
                    If r <> "" Then
                        r.Clear
                        CellValueChanged = True
                    End If
                
                Else
                
                    If r <> v Then
                        r = v
                        CellValueChanged = True
                    End If
                
                End If
            
            End If
            
        End If
    
    End If
    
    Set r = Nothing

End Sub

Public Sub SetWorksheetColumnFormula(ws As Worksheet, i As Long, columnName As String, f As String)

    Dim c As Long
    Dim r As Range
    
    c = GetWorksheetColumnIndex(ws, columnName)
    If c > 0 Then
    
        Set r = ws.Cells(i, c)
    
        If r.Formula <> f Then
        
            r.Formula = f
            
        End If
    
    End If
    
    Set r = Nothing

End Sub

Public Function GetWorksheetColumnIndex(ws As Worksheet, columnName As String) As Long

    Dim c As Long
    Dim ttc As Long
    
    ttc = CountColumns(ws, 1)
    
    If IndexColumn Is Nothing Then
        Set IndexColumn = New Collection
        p_currentName = ""
    End If
    
    If p_currentName <> ws.Name Then
        Set p_currentCols = Nothing
    End If
        
    If p_currentCols Is Nothing Then
        Set p_currentCols = InCollection(IndexColumn, LCase(ws.Name), True)
        p_currentName = ws.Name
    End If
        
    If p_currentCols Is Nothing Then
    
        Set p_currentCols = New Collection
    
        For c = 1 To ttc
        
            If ws.Cells(1, c) <> "" Then
                Call p_currentCols.Add(c, ws.Cells(1, c))
            End If
            
        Next c
        
        Call IndexColumn.Add(p_currentCols, LCase(ws.Name))
    
    End If
    
    GetWorksheetColumnIndex = InCollection(p_currentCols, columnName)
    
End Function

Public Function CreateIndex(ws As Worksheet, ParamArray cols() As Variant) As Collection

    Dim index As Collection
    
    Dim key As String
    Dim keyOld As String
    
    Dim n As Variant
    Dim i As Long
    Dim c As Long
    
    Set CreateIndex = New Collection
    keyOld = ""
    
    i = 1
    Do
    
        i = i + 1
        
        key = ""
        For Each n In cols
            key = key & "-" & GetWorksheetColumnValue(ws, i, CStr(n))
        Next
        key = Mid(key, 2)
        
        If Replace(key, "-", "") = "" Then
            Exit Do
        End If
        
        If keyOld <> key Then
        
            Call CreateIndex.Add(i, key)
            keyOld = key
            
        End If
    
    Loop

End Function

Public Sub SortWorksheet(ws As Worksheet, ParamArray cols() As Variant)

    Dim n As Variant
    Dim c As String
    
    ws.Select
    ws.Range("A2").Select
    Selection.CurrentRegion.Select
    
    With ws.Sort
        .SortFields.Clear
        
        For Each n In cols
            c = GetColumnLetter(GetWorksheetColumnIndex(ws, CStr(n))) & "1"
            .SortFields.Add key:=Range(c), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        Next
               
        .SetRange Range(Selection.CurrentRegion.Address(False, False))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ws.Range("A2").Select

End Sub

Public Sub ClearWorksheetColumns(ws As Worksheet, ParamArray cols() As Variant)

    Dim n As Variant
    Dim c As Long

    For Each n In cols
    
        c = GetWorksheetColumnIndex(ws, CStr(n))
        If c > 0 Then
        
            ws.Columns(c).ClearContents
            ws.Cells(1, c) = CStr(n)
        
        End If
        
    Next

End Sub

Public Function EncryptDecrypt(text As String) As String

    Const KEY_TEXT As String = "ldkfo55314e25manDSGGH@#$!kdhsdoe!#$adetf"

    Dim byteText() As Byte
    Dim bytePWD() As Byte
    Dim intPWDPos As Integer
    Dim intPWDLen As Integer
    Dim intLoop As Integer
    
    byteText = text
    bytePWD = KEY_TEXT
    intPWDLen = LenB(KEY_TEXT)
    
    For intLoop = 0 To LenB(text) - 1
        intPWDPos = (intLoop Mod intPWDLen)
        byteText(intLoop) = byteText(intLoop) Xor bytePWD(intPWDPos)
    Next intLoop
    
    EncryptDecrypt = byteText
    
End Function

Public Function OnlyDigits(s As String) As String
    
    Dim r As String
    Dim d As String
    Dim i As Integer

    r = ""
    
    For i = 1 To Len(s)
        d = Mid(s, i, 1)
        If d >= "0" And d <= "9" Then
            r = r + d
        End If
    Next

    OnlyDigits = r
    
End Function

Public Sub AddCountColumnNames(n As String)

    Dim c As Long
    
    c = InCollection(CountColumnNames, n)
    If c > 0 Then
        Call CountColumnNames.Remove(n)
    End If
    
    c = c + 1
    Call CountColumnNames.Add(c, n)

End Sub

Public Function GetLastLine(ws As Worksheet, col As Long)

    Dim n As Long
    
    n = ws.Cells(Rows.count, col).End(xlUp).Offset(0, 0).Row
    GetLastLine = n
    
End Function

Sub ClearWorksheet(ws As Worksheet, r As Long)
    
    Dim ln As String
     ln = r & ":1048576"

    ws.Activate
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    Rows(ln).Select
    Selection.Delete Shift:=xlUp
    
End Sub

Public Function CountRows(ws As Worksheet, col As Long)

    Dim n As Long
    
    n = ws.Cells(Rows.count, col).End(xlUp).Offset(0, 0).Row
    CountRows = n
    
End Function

Public Function CountColumns(ws As Worksheet, ln As Long)

    Dim n As Long
    
    n = ws.Cells(ln, Columns.count).End(xlToLeft).Offset(0, 0).Column
    CountColumns = n
    
End Function

Function RangetoHTML(rng As Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         fileName:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
    
End Function

Public Sub SortWorksheetDesc(ws As Worksheet, ParamArray cols() As Variant)

    Dim n As Variant
    Dim c As String
    
    ws.Parent.Activate
    ws.Select
    ws.Range("A2").Select
    Selection.CurrentRegion.Select
    
    With ws.Sort
        .SortFields.Clear
        
        For Each n In cols
            c = GetColumnLetter(GetWorksheetColumnIndex(ws, CStr(n))) & "1"
            .SortFields.Add key:=Range(c), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        Next
               
        .SetRange Range(Selection.CurrentRegion.Address(False, False))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ws.Range("A2").Select

End Sub

Sub ClearHeader(ws As Worksheet, c As String, r As Long)
    
    Dim rng As String
    
    rng = c & r & ":XFD" & r

    ws.Activate
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    ws.Range(rng).Select
    Selection.Delete Shift:=xlUp

End Sub


