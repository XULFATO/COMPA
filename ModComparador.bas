Attribute VB_Name = "Comparador"

'===========================================================
' IMPORTAR HOY 1
'===========================================================
Sub ImportarHoy1()
    ImportarHoja 1
End Sub

'===========================================================
' IMPORTAR HOY 2
'===========================================================
Sub ImportarHoy2()
    ImportarHoja 2
End Sub

'===========================================================
' SUBRUTINA GENÉRICA DE IMPORTACIÓN
'===========================================================
Sub ImportarHoja(slot As Integer)

    ' 1) Listar workbooks abiertos (excluir este)
    Dim nombres() As String
    Dim wb As Workbook
    Dim n As Integer
    n = 0
    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            ReDim Preserve nombres(n)
            nombres(n) = wb.Name
            n = n + 1
        End If
    Next

    If n = 0 Then
        MsgBox "No hay otros ficheros Excel abiertos." & vbCrLf & _
               "Abre el fichero que quieres importar y vuelve a pulsar el botón.", _
               vbExclamation, "Sin ficheros"
        Exit Sub
    End If

    ' 2) Elegir fichero con InputBox de lista
    Dim lista As String
    Dim i As Integer
    lista = "Ficheros abiertos:" & vbCrLf & vbCrLf
    For i = 0 To n - 1
        lista = lista & "  " & (i + 1) & "  -  " & nombres(i) & vbCrLf
    Next

    Dim respWB As String
    respWB = Application.InputBox( _
        lista & vbCrLf & "Escribe el número del fichero:", _
        "Seleccionar fichero para HOY " & slot, Type:=1)

    If respWB = "False" Then Exit Sub   ' pulsó Cancelar
    Dim idxWB As Integer
    idxWB = CInt(respWB) - 1
    If idxWB < 0 Or idxWB >= n Then
        MsgBox "Número no válido.", vbExclamation
        Exit Sub
    End If

    Dim wbOrigen As Workbook
    Set wbOrigen = Application.Workbooks(nombres(idxWB))

    ' 3) Elegir hoja del workbook seleccionado
    Dim hojas() As String
    Dim m As Integer
    m = wbOrigen.Worksheets.Count
    ReDim hojas(m - 1)
    Dim listaH As String
    listaH = "Hojas de  [" & wbOrigen.Name & "]:" & vbCrLf & vbCrLf
    For i = 1 To m
        hojas(i - 1) = wbOrigen.Worksheets(i).Name
        listaH = listaH & "  " & i & "  -  " & wbOrigen.Worksheets(i).Name & vbCrLf
    Next

    Dim respWS As String
    respWS = Application.InputBox( _
        listaH & vbCrLf & "Escribe el número de la hoja:", _
        "Seleccionar hoja para HOY " & slot, Type:=1)

    If respWS = "False" Then Exit Sub
    Dim idxWS As Integer
    idxWS = CInt(respWS) - 1
    If idxWS < 0 Or idxWS >= m Then
        MsgBox "Número no válido.", vbExclamation
        Exit Sub
    End If

    Dim wsOrigen As Worksheet
    Set wsOrigen = wbOrigen.Worksheets(hojas(idxWS))

    ' 4) Nombre para la hoja destino
    Dim nomBase As String
    nomBase = wbOrigen.Name
    Dim p As Integer
    p = InStrRev(nomBase, ".")
    If p > 0 Then nomBase = Left(nomBase, p - 1)
    If Len(nomBase) > 20 Then nomBase = Left(nomBase, 20)

    Dim etiqueta As String
    etiqueta = "HOY " & slot

    Dim nomHoja As String
    nomHoja = nomBase & " [" & etiqueta & "]"

    ' Sanitizar caracteres ilegales en nombre de hoja
    Dim cars As Variant
    cars = Array("/", "\", "?", "*", "[", "]", ":")
    Dim c As Variant
    For Each c In cars
        nomHoja = Replace(nomHoja, CStr(c), "_")
    Next

    ' 5) Borrar si ya existe y copiar
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(nomHoja).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    wsOrigen.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count).Name = nomHoja

    ' 6) Guardar referencia en MENU!J1 o J2
    Dim wsMENU As Worksheet
    Set wsMENU = ThisWorkbook.Worksheets("MENU")
    If slot = 1 Then
        wsMENU.Range("J1").Value = nomHoja
    Else
        wsMENU.Range("J2").Value = nomHoja
    End If

    ' Volver al MENU
    wsMENU.Activate
    MsgBox "Hoja importada correctamente como:" & vbCrLf & vbCrLf & _
           "  « " & nomHoja & " »", vbInformation, "HOY " & slot & " importado"
End Sub


'===========================================================
' COMPARAR LAS DOS HOJAS
'===========================================================
Sub CompararHojas()

    Dim wsMENU As Worksheet
    Set wsMENU = ThisWorkbook.Worksheets("MENU")

    Dim nomH1 As String, nomH2 As String
    nomH1 = wsMENU.Range("J1").Value
    nomH2 = wsMENU.Range("J2").Value

    If Trim(nomH1) = "" Or Trim(nomH2) = "" Then
        MsgBox "Primero importa las dos hojas (HOY 1 y HOY 2).", _
               vbExclamation, "Faltan hojas"
        Exit Sub
    End If

    Dim ws1 As Worksheet, ws2 As Worksheet
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets(nomH1)
    Set ws2 = ThisWorkbook.Worksheets(nomH2)
    On Error GoTo 0

    If ws1 Is Nothing Then
        MsgBox "No encuentro la hoja HOY 1: " & nomH1, vbCritical
        Exit Sub
    End If
    If ws2 Is Nothing Then
        MsgBox "No encuentro la hoja HOY 2: " & nomH2, vbCritical
        Exit Sub
    End If

    ' Calcular límites
    Dim lastRow1 As Long, lastCol1 As Long
    Dim lastRow2 As Long, lastCol2 As Long
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastCol1 = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    lastCol2 = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column

    Dim maxRow As Long, maxCol As Long
    maxRow = Application.Max(lastRow1, lastRow2)
    maxCol = Application.Max(lastCol1, lastCol2)

    ' Eliminar hoja COMPARACION anterior si existe
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("COMPARACION").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Crear hoja COMPARACION al final
    Dim wsC As Worksheet
    Set wsC = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsC.Name = "COMPARACION"

    ' Copiar cabecera de ws1
    ws1.Rows(1).Copy wsC.Rows(1)

    ' Cabecera columna DIFERENTE
    Dim colDif As Long
    colDif = maxCol + 1
    With wsC.Cells(1, colDif)
        .Value = "DIFERENTE"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With

    ' Recorrer filas de datos
    Application.ScreenUpdating = False

    Dim fila As Long, col As Long
    Dim v1 As String, v2 As String
    Dim difFila As Boolean
    Dim filaC As Long
    filaC = 2

    For fila = 2 To maxRow
        difFila = False

        ' Comprobar si alguna celda difiere
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

        ' Copiar fila base (desde ws1 si existe, si no desde ws2)
        If fila <= lastRow1 Then
            ws1.Rows(fila).Copy wsC.Rows(filaC)
        Else
            ws2.Rows(fila).Copy wsC.Rows(filaC)
        End If

        ' Marcar resultado
        If difFila Then
            wsC.Cells(filaC, colDif).Value = "SI"
            wsC.Cells(filaC, colDif).Font.Color = RGB(192, 57, 43)
            wsC.Cells(filaC, colDif).Font.Bold = True
            wsC.Rows(filaC).Interior.Color = RGB(255, 235, 235)

            ' Colorear celdas que cambian
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

    ' Formato final
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

    ' Resumen
    Dim totalDif As Long
    totalDif = Application.WorksheetFunction.CountIf( _
        wsC.Columns(colDif), "SI")

    MsgBox "Comparación completada." & vbCrLf & vbCrLf & _
           "  Filas analizadas : " & (maxRow - 1) & vbCrLf & _
           "  Filas DIFERENTES : " & totalDif & vbCrLf & _
           "  Filas IGUALES    : " & (maxRow - 1 - totalDif) & vbCrLf & vbCrLf & _
           "La hoja COMPARACION está lista." & vbCrLf & _
           "Filtra DIFERENTE = SI para ver solo los cambios.", _
           vbInformation, "Resultado"
End Sub
