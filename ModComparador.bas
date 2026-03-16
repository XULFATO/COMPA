'===========================================================
' COMPARADOR DE EXCELS - pega este bloque en un modulo VBA
' Alt+F11 -> Insertar -> Modulo -> borra lo que haya -> pega
'===========================================================

Sub ImportarHoy1()
    ImportarHoja 1
End Sub

Sub ImportarHoy2()
    ImportarHoja 2
End Sub

'-----------------------------------------------------------
Sub ImportarHoja(slot As Integer)

    ' Referencia a MENU capturada AL INICIO, antes de cualquier Copy
    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Worksheets("MENU")

    ' --- 1) Listar workbooks abiertos (excluir este) ---
    Dim nombres()  As String
    Dim wbLoop     As Workbook
    Dim n          As Integer
    Dim i          As Integer
    n = 0

    For Each wbLoop In Application.Workbooks
        If wbLoop.Name <> ThisWorkbook.Name Then
            ReDim Preserve nombres(n)
            nombres(n) = wbLoop.Name
            n = n + 1
        End If
    Next wbLoop

    If n = 0 Then
        MsgBox "No hay otros ficheros Excel abiertos." & vbCrLf & _
               "Abre primero el fichero que quieres importar.", _
               vbExclamation, "Sin ficheros"
        Exit Sub
    End If

    ' --- 2) Elegir fichero ---
    Dim lista As String
    lista = "Ficheros Excel abiertos:" & vbCrLf & vbCrLf
    For i = 0 To n - 1
        lista = lista & "  " & (i + 1) & "  ->  " & nombres(i) & vbCrLf
    Next i
    lista = lista & vbCrLf & "Escribe el numero del fichero:"

    Dim respWB As Variant
    respWB = Application.InputBox(lista, "Importar HOY " & slot, Type:=1)

    If VarType(respWB) = vbBoolean Then Exit Sub
    Dim idxWB As Integer
    idxWB = CInt(respWB) - 1
    If idxWB < 0 Or idxWB >= n Then
        MsgBox "Numero fuera de rango (1 a " & n & ").", vbExclamation
        Exit Sub
    End If

    Dim wbOrigen As Workbook
    Set wbOrigen = Application.Workbooks(nombres(idxWB))

    ' --- 3) Elegir hoja ---
    Dim hojas() As String
    Dim m       As Integer
    m = wbOrigen.Worksheets.Count
    ReDim hojas(m - 1)

    Dim listaH As String
    listaH = "Hojas de [" & wbOrigen.Name & "]:" & vbCrLf & vbCrLf
    For i = 1 To m
        hojas(i - 1) = wbOrigen.Worksheets(i).Name
        listaH = listaH & "  " & i & "  ->  " & wbOrigen.Worksheets(i).Name & vbCrLf
    Next i
    listaH = listaH & vbCrLf & "Escribe el numero de la hoja:"

    Dim respWS As Variant
    respWS = Application.InputBox(listaH, "Importar HOY " & slot, Type:=1)

    If VarType(respWS) = vbBoolean Then Exit Sub
    Dim idxWS As Integer
    idxWS = CInt(respWS) - 1
    If idxWS < 0 Or idxWS >= m Then
        MsgBox "Numero fuera de rango (1 a " & m & ").", vbExclamation
        Exit Sub
    End If

    Dim wsOrigen As Worksheet
    Set wsOrigen = wbOrigen.Worksheets(hojas(idxWS))

    ' --- 4) Construir nombre de hoja destino ---
    Dim nomBase As String
    nomBase = wbOrigen.Name
    Dim p As Integer
    p = InStrRev(nomBase, ".")
    If p > 0 Then nomBase = Left(nomBase, p - 1)
    If Len(nomBase) > 20 Then nomBase = Left(nomBase, 20)

    Dim nomHoja As String
    nomHoja = nomBase & " [HOY " & slot & "]"

    Dim cars As Variant
    Dim c    As Variant
    cars = Array("/", "\", "?", "*", "[", "]", ":")
    For Each c In cars
        nomHoja = Replace(nomHoja, CStr(c), "_")
    Next c

    ' --- 5) Borrar hoja anterior del mismo slot si existe ---
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(nomHoja).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' --- 6) Copiar hoja al final de este workbook ---
    wsOrigen.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count).Name = nomHoja

    ' --- 7) Guardar referencia (wsMenu ya capturada al inicio) ---
    If slot = 1 Then
        wsMenu.Range("J1").Value = nomHoja
    Else
        wsMenu.Range("J2").Value = nomHoja
    End If

    wsMenu.Activate
    MsgBox "Hoja importada como:" & vbCrLf & vbCrLf & _
           "  << " & nomHoja & " >>", vbInformation, "HOY " & slot & " OK"
End Sub


'===========================================================
' COMPARAR
'===========================================================
Sub CompararHojas()

    ' Referencia a MENU capturada AL INICIO
    Dim wsMenu As Worksheet
    Set wsMenu = ThisWorkbook.Worksheets("MENU")

    Dim nomH1 As String
    Dim nomH2 As String
    nomH1 = Trim(wsMenu.Range("J1").Value)
    nomH2 = Trim(wsMenu.Range("J2").Value)

    If nomH1 = "" Or nomH2 = "" Then
        MsgBox "Importa primero las dos hojas (HOY 1 y HOY 2).", _
               vbExclamation, "Faltan hojas"
        Exit Sub
    End If

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets(nomH1)
    Set ws2 = ThisWorkbook.Worksheets(nomH2)
    On Error GoTo 0

    If ws1 Is Nothing Then
        MsgBox "No encuentro la hoja HOY 1:" & vbCrLf & nomH1, vbCritical
        Exit Sub
    End If
    If ws2 Is Nothing Then
        MsgBox "No encuentro la hoja HOY 2:" & vbCrLf & nomH2, vbCritical
        Exit Sub
    End If

    ' --- Limites de datos ---
    Dim lastRow1 As Long, lastCol1 As Long
    Dim lastRow2 As Long, lastCol2 As Long
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastCol1 = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    lastCol2 = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column

    Dim maxRow As Long
    Dim maxCol As Long
    maxRow = Application.Max(lastRow1, lastRow2)
    maxCol = Application.Max(lastCol1, lastCol2)

    ' --- Borrar COMPARACION anterior si existe ---
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("COMPARACION").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' --- Crear hoja COMPARACION al final ---
    Dim wsC As Worksheet
    Set wsC = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsC.Name = "COMPARACION"

    ' Copiar cabecera de ws1
    ws1.Rows(1).Copy wsC.Rows(1)

    ' Columna DIFERENTE al final
    Dim colDif As Long
    colDif = maxCol + 1
    With wsC.Cells(1, colDif)
        .Value = "DIFERENTE"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With

    ' --- Recorrer filas de datos ---
    Application.ScreenUpdating = False

    Dim fila    As Long
    Dim col     As Long
    Dim v1      As String
    Dim v2      As String
    Dim difFila As Boolean
    Dim filaC   As Long
    filaC = 2

    For fila = 2 To maxRow

        difFila = False
        For col = 1 To maxCol
            v1 = ""
            v2 = ""
            If col <= lastCol1 And fila <= lastRow1 Then v1 = CStr(ws1.Cells(fila, col).Value)
            If col <= lastCol2 And fila <= lastRow2 Then v2 = CStr(ws2.Cells(fila, col).Value)
            If v1 <> v2 Then
                difFila = True
                Exit For
            End If
        Next col

        ' Copiar la fila (de ws1 si existe, si no de ws2)
        If fila <= lastRow1 Then
            ws1.Rows(fila).Copy wsC.Rows(filaC)
        Else
            ws2.Rows(fila).Copy wsC.Rows(filaC)
        End If

        ' Marcar DIFERENTE y colorear celdas cambiadas
        If difFila Then
            wsC.Cells(filaC, colDif).Value = "SI"
            wsC.Cells(filaC, colDif).Font.Color = RGB(192, 57, 43)
            wsC.Cells(filaC, colDif).Font.Bold = True
            wsC.Rows(filaC).Interior.Color = RGB(255, 235, 235)
            For col = 1 To maxCol
                v1 = ""
                v2 = ""
                If col <= lastCol1 And fila <= lastRow1 Then v1 = CStr(ws1.Cells(fila, col).Value)
                If col <= lastCol2 And fila <= lastRow2 Then v2 = CStr(ws2.Cells(fila, col).Value)
                If v1 <> v2 Then
                    wsC.Cells(filaC, col).Interior.Color = RGB(255, 180, 180)
                    wsC.Cells(filaC, col).Font.Bold = True
                End If
            Next col
        Else
            wsC.Cells(filaC, colDif).Value = "NO"
            wsC.Cells(filaC, colDif).Font.Color = RGB(39, 174, 96)
        End If

        filaC = filaC + 1
    Next fila

    ' --- Formato final de COMPARACION ---
    With wsC.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
    End With

    wsC.Rows(1).AutoFilter
    wsC.Cells.EntireColumn.AutoFit

    wsC.Activate
    wsC.Rows(2).Select
    ActiveWindow.FreezePanes = True
    wsC.Range("A1").Select

    Application.ScreenUpdating = True

    ' --- Resumen ---
    Dim totalDif As Long
    totalDif = Application.WorksheetFunction.CountIf(wsC.Columns(colDif), "SI")

    MsgBox "Comparacion completada." & vbCrLf & vbCrLf & _
           "  Filas analizadas : " & (maxRow - 1) & vbCrLf & _
           "  Filas DIFERENTES : " & totalDif & vbCrLf & _
           "  Filas IGUALES    : " & (maxRow - 1 - totalDif) & vbCrLf & vbCrLf & _
           "Filtra DIFERENTE = SI para ver solo los cambios.", _
           vbInformation, "Resultado"
End Sub
